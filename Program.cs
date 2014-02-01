/* Copyright (C) 2010-2014, Manuel Meitinger
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 2 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using Aufbauwerk.Tools.GroupPolicyInstaller.Properties;
using Microsoft.Win32;

namespace Aufbauwerk.Tools.GroupPolicyInstaller
{
    internal static class Program
    {
        private const uint EWX_REBOOT = 0x00000002;
        private const uint EWX_FORCE = 0x00000004;
        private const uint SHTDN_REASON_MAJOR_APPLICATION = 0x00040000;
        private const uint SHTDN_REASON_MINOR_INSTALLATION = 0x00000002;
        private const uint SHTDN_REASON_FLAG_PLANNED = 0x80000000;

        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        private static extern bool ExitWindowsEx
        (
            [In] uint uFlags,
            [In] uint dwReason
        );

        private static readonly SortedDictionary<ulong, Task> tasks = new SortedDictionary<ulong, Task>();
        private static bool isInitialized = false;
        private static XmlSerializer setupSerializer;
        private static XmlReaderSettings readerSettings;
        private static EventLog log = null;
        private static StreamWriter logFile = null;
        private static bool isExclusive = false;
        private static bool doReboot = false;
        private static bool doStop = false;
        private const string eventLogName = "Application";
        private const string eventSource = "Group Policy Installer";

        private static void AddTask(ulong index, string path)
        {
            // initialize the XML serializer and schema validator if not done yet
            if (!isInitialized)
            {
                setupSerializer = new XmlSerializer(typeof(Setup));
                readerSettings = new XmlReaderSettings();
                using (StringReader stringReader = new StringReader(Resources.ProgramSetupSchema))
                using (XmlReader xmlReader = XmlReader.Create(stringReader))
                    readerSettings.Schemas.Add(null, xmlReader);
                readerSettings.ValidationType = ValidationType.Schema;
                isInitialized = true;
            }

            // deserialize the setup object from file
            Setup setup;
            try
            {
                path = Path.GetFullPath(path);
                using (XmlReader reader = XmlReader.Create(path, readerSettings))
                    setup = (Setup)setupSerializer.Deserialize(reader);
            }
            catch (Exception e)
            {
                WriteEvent(Resources.ProgramLoadSetupFailed, EventLogEntryType.Error, index, path, e.Message);
                return;
            }

            // wrap the setup in a task and try adding it to the list if we haven't encountered an exclusive task yet
            Task task = new Task(path, setup);
            if (!isExclusive)
            {
                // if this is an exclusive task...
                if (task.IsExclusive)
                {
                    // ...stop further tasks from being added and add it iff the task list is empty
                    isExclusive = true;
                    if (tasks.Count > 0)
                        return;
                }
                tasks.Add(index, task);
            }
        }

        [STAThread]
        private static void Main(string[] args)
        {
            // set the alternative config file if specified
            if (args.Length > 0)
                ((AppDomainSetup)typeof(AppDomain).GetProperty("FusionStore", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(AppDomain.CurrentDomain, null)).ConfigurationFile = args[0];

            // set text rendering and exception handling
            Application.SetCompatibleTextRenderingDefault(false);
            Application.EnableVisualStyles();
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            if (!Settings.Default.IgnoreReboot)
                SystemEvents.SessionEnded += new SessionEndedEventHandler(SystemEvents_SessionEnded);

            // try to open the event log (and create an event source if necessary)
            try
            {
                if (!EventLog.SourceExists(eventSource))
                    EventLog.CreateEventSource(eventSource, eventLogName);
                log = new EventLog(EventLog.LogNameFromSourceName(eventSource, "."), ".", eventSource);
            }
            catch { }
            try
            {
                // try to open the log file (and create its parent directories if necessary)
                string logFilePath = Environment.ExpandEnvironmentVariables(Settings.Default.LogPath);
                if (logFilePath.Length > 0)
                {
                    try
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));
                        logFile = new StreamWriter(logFilePath, true, Encoding.Unicode);
                    }
                    catch (Exception e) { WriteEvent(Resources.ProgramOpenLogFailed, EventLogEntryType.Warning, logFilePath, e.Message); }
                }
                try
                {
                    // open the root key of the specified registry path
                    string keyPath = Environment.ExpandEnvironmentVariables(Settings.Default.RegPath);
                    RegistryKey key;
                    if (keyPath.StartsWith(@"HKLM\", StringComparison.OrdinalIgnoreCase))
                        key = Registry.LocalMachine;
                    else if (keyPath.StartsWith(@"HKCU\", StringComparison.OrdinalIgnoreCase))
                        key = Registry.CurrentUser;
                    else if (keyPath.StartsWith(@"HKCR\", StringComparison.OrdinalIgnoreCase))
                        key = Registry.ClassesRoot;
                    else if (keyPath.StartsWith(@"HKU\", StringComparison.OrdinalIgnoreCase))
                        key = Registry.Users;
                    else if (keyPath.StartsWith(@"HKCC\", StringComparison.OrdinalIgnoreCase))
                        key = Registry.CurrentConfig;
                    else
                    {
                        WriteEvent(Resources.ProgramReadRegFailed, EventLogEntryType.Error, keyPath);
                        return;
                    }

                    // open the subkey
                    using (key = key.OpenSubKey(keyPath.Substring(keyPath.IndexOf('\\') + 1)))
                    {
                        if (key == null)
                        {
                            // it is worth a warning if the key doesn't exists (?)
                            WriteEvent(Resources.ProgramReadRegFailed, EventLogEntryType.Warning, keyPath);
                            return;
                        }

                        // read all values
                        long index;
                        foreach (string name in key.GetValueNames())
                        {
                            // read as long, but store as ulong (-1 becomes to last task to execute)
                            if (long.TryParse(name, out index))
                            {
                                switch (key.GetValueKind(name))
                                {
                                    case RegistryValueKind.String:
                                        AddTask((ulong)index, (string)key.GetValue(name));
                                        continue;
                                    case RegistryValueKind.ExpandString:
                                        AddTask((ulong)index, Environment.ExpandEnvironmentVariables((string)key.GetValue(name)));
                                        continue;
                                }
                            }
                            WriteEvent(Resources.ProgramReadRegFailed, EventLogEntryType.Error, Path.Combine(keyPath, name));
                            return;
                        }

                        // if no tasks could be loaded, return
                        if (tasks.Count == 0)
                            return;

                        // run the application and perform a reboot afterwards if necessary
                        RuntimeHelpers.PrepareConstrainedRegions();
                        try { Application.Run(new MainForm()); }
                        finally
                        {
                            if (doReboot && !Settings.Default.SuppressReboot)
                            {
                                // aquire the SE_SHUTDOWN_NAME privilege
                                Type privilegeType = Type.GetType("System.Security.AccessControl.Privilege");
                                object privilege = privilegeType.GetConstructor(new Type[] { typeof(string) }).Invoke(new object[] { privilegeType.GetField("Shutdown").GetValue(null) });
                                RuntimeHelpers.PrepareConstrainedRegions();
                                try
                                {
                                    // reboot using user32 API (g_fReadyForShutdown is false)
                                    privilegeType.GetMethod("Enable").Invoke(privilege, null);
                                    if (!ExitWindowsEx(EWX_REBOOT | EWX_FORCE, SHTDN_REASON_FLAG_PLANNED | SHTDN_REASON_MAJOR_APPLICATION | SHTDN_REASON_MINOR_INSTALLATION))
                                        WriteEvent(Resources.ProgramRebootFailed, EventLogEntryType.Error, new Win32Exception().Message);
                                }
                                finally { privilegeType.GetMethod("Revert").Invoke(privilege, null); }
                            }
                        }
                    }
                }
                finally
                {
                    if (logFile != null)
                    {
                        logFile.Close();
                        logFile = null;
                    }
                }
            }
            catch (Exception e) { WriteEvent(e.ToString(), EventLogEntryType.Error); }
            finally
            {
                if (log != null)
                {
                    log.Close();
                    log = null;
                }
            }
        }

        private static void SystemEvents_SessionEnded(object sender, SessionEndedEventArgs e)
        {
            // stop all we're doing
            doStop = true;
        }

        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            // log the error and exit the application
            WriteEvent(e.Exception.ToString(), EventLogEntryType.Error);
            Application.Exit();
        }

        /// <summary>
        /// Gets an ordered list of all tasks.
        /// </summary>
        internal static IEnumerable<Task> Tasks
        {
            get
            {
                foreach (Task task in tasks.Values)
                {
                    if (doStop)
                        break;
                    yield return task;
                }
            }
        }

        /// <summary>
        /// Writes a formatted message into the log file and event log.
        /// </summary>
        /// <param name="message">A composite format string.</param>
        /// <param name="type">One of the <see cref="System.Diagnostics.EventLogEntryType"/> values.</param>
        /// <param name="args">An object array that contains zero or more objects to format.</param>
        internal static void WriteEvent(string message, EventLogEntryType type, params object[] args)
        {
            WriteEvent(string.Format(message, args), type);
        }

        /// <summary>
        /// Writes a message into the log file and event log.
        /// </summary>
        /// <param name="message">The string to write.</param>
        /// <param name="type">One of the <see cref="System.Diagnostics.EventLogEntryType"/> values.</param>
        internal static void WriteEvent(string message, EventLogEntryType type)
        {
            if (logFile != null)
            {
                try
                {
                    logFile.WriteLine(string.Format("{0}\t{1}\t{2}", type.ToString().ToUpper().Substring(0, 4), DateTime.Now.ToString("yyyy-dd-MM HH:mm:ss"), message.Replace("\r\n", " | ").Replace('\t', ' ')));
                    logFile.Flush();
                }
                catch { }
            }
            if (log != null)
            {
                try { log.WriteEntry(message, type); }
                catch { }
            }
        }

        /// <summary>
        /// Sets a flag that the program should reboot the system upon exit.
        /// </summary>
        internal static void ScheduleReboot()
        {
            doReboot = true;
        }

        /// <summary>
        /// Sets a flag that the program should not yield any more tasks.
        /// </summary>
        /// <remarks>
        /// Used only when a setup task has issued a reboot or any other fatal error occured.
        /// </remarks>
        internal static void Stop()
        {
            doStop = true;
        }
    }
}
