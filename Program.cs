/* Copyright (C) 2010-2015, Manuel Meitinger
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
using System.Globalization;
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
        private static EventLog log = null;
        private static StreamWriter logFile = null;
        private static RegistryKey rootKey = null;
        private static string subKeyName = null;
        private static bool doReboot = false;
        private static bool doStop = false;
        private const string eventLogName = "Application";
        private const string eventSource = "Group Policy Installer";

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
                    if (keyPath.StartsWith(@"HKLM\", StringComparison.OrdinalIgnoreCase))
                    {
                        rootKey = Registry.LocalMachine;
                        subKeyName = keyPath.Substring(5);
                    }
                    else if (keyPath.StartsWith(@"HKCU\", StringComparison.OrdinalIgnoreCase))
                    {
                        rootKey = Registry.CurrentUser;
                        subKeyName = keyPath.Substring(5);
                    }
                    else if (keyPath.StartsWith(@"HKCR\", StringComparison.OrdinalIgnoreCase))
                    {
                        rootKey = Registry.ClassesRoot;
                        subKeyName = keyPath.Substring(5);
                    }
                    else if (keyPath.StartsWith(@"HKU\", StringComparison.OrdinalIgnoreCase))
                    {
                        rootKey = Registry.Users;
                        subKeyName = keyPath.Substring(4);
                    }
                    else if (keyPath.StartsWith(@"HKCC\", StringComparison.OrdinalIgnoreCase))
                    {
                        rootKey = Registry.CurrentConfig;
                        subKeyName = keyPath.Substring(5);
                    }
                    else
                    {
                        WriteEvent(Resources.ProgramRegPathRootInvalid, EventLogEntryType.Error, keyPath);
                        return;
                    }

                    // open the subkey
                    using (RegistryKey registryKey = rootKey.OpenSubKey(subKeyName))
                    {
                        if (registryKey == null)
                        {
                            // it is worth a warning if the key doesn't exists (?)
                            WriteEvent(Resources.ProgramRegPathMissing, EventLogEntryType.Warning, rootKey.Name, subKeyName);
                            return;
                        }

                        // initialize the XML serializer and schema validator
                        XmlSerializer setupSerializer = new XmlSerializer(typeof(Setup));
                        XmlReaderSettings readerSettings = new XmlReaderSettings();
                        using (StringReader stringReader = new StringReader(Resources.ProgramSetupSchema))
                        using (XmlReader xmlReader = XmlReader.Create(stringReader))
                            readerSettings.Schemas.Add(null, xmlReader);
                        readerSettings.ValidationType = ValidationType.Schema;

                        // read all values
                        foreach (string name in registryKey.GetValueNames())
                        {
                            // read the name as integer (-1 becomes the last task to execute)
                            long index;
                            if (!long.TryParse(name, NumberStyles.Integer, CultureInfo.InvariantCulture, out index))
                            {
                                WriteEvent(Resources.ProgramRegValueNameInvalid, EventLogEntryType.Error, registryKey.Name, name);
                                continue;
                            }

                            // read the path
                            string path;
                            RegistryValueKind kind = registryKey.GetValueKind(name);
                            switch (kind)
                            {
                                case RegistryValueKind.String:
                                    path = (string)registryKey.GetValue(name);
                                    break;
                                case RegistryValueKind.ExpandString:
                                    path = (string)Environment.ExpandEnvironmentVariables((string)registryKey.GetValue(name));
                                    break;
                                default:
                                    WriteEvent(Resources.ProgramRegValueTypeInvalid, EventLogEntryType.Error, registryKey.Name, name, kind);
                                    continue;
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
                                continue;
                            }

                            // wrap the setup in a task and add it to the list
                            tasks.Add((ulong)index, new Task(name, path, setup));
                        }
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
            // stop going over new tasks
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
                // if there is an exclusive task return it and only it
                foreach (Task task in tasks.Values)
                {
                    if (task.IsExclusive)
                    {
                        if (!doStop)
                            yield return task;
                        yield break;
                    }
                }

                // otherwise return all tasks or stop if asked to
                foreach (Task task in tasks.Values)
                {
                    if (doStop)
                        break;
                    yield return task;
                }
            }
        }

        /// <summary>
        /// Opens the registry key to the task list.
        /// </summary>
        /// <param name="writable">If <c>true</c>, the returned <see cref="Microsoft.Win32.RegistryKey"/> allows modifications to the registry.</param>
        /// <returns>The key or <c>null</c> if the operation failed.</returns>
        internal static RegistryKey OpenRegistryKey(bool writable)
        {
            return rootKey == null || subKeyName == null ? null : rootKey.OpenSubKey(subKeyName, writable);
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
                    logFile.WriteLine(string.Format("{0}\t{1}\t{2}", type.ToString().ToUpper().Substring(0, 4), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), message.Replace("\r\n", " | ").Replace('\t', ' ')));
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
