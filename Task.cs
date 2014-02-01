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
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Aufbauwerk.Tools.GroupPolicyInstaller.Properties;

namespace Aufbauwerk.Tools.GroupPolicyInstaller
{
    /// <summary>
    /// Callback for task progress changes.
    /// </summary>
    /// <param name="action">Short description of the current action.</param>
    /// <param name="progress">The current progress (between 0 and 100) or -1 if it is unknown.</param>
    public delegate void TaskProgressChanged(string action, int progress);

    /// <summary>
    /// Represents a runnable task.
    /// </summary>
    public class Task
    {
        private readonly Setup setup;
        private readonly string name;
        private readonly string directory;
        private readonly Image image;

        #region Win32/MSI/WU/PS/SetupDi

        private enum INSTALLUILEVEL
        {
            NOCHANGE = 0,    // UI level is unchanged
            DEFAULT = 1,    // default UI is used
            NONE = 2,    // completely silent installation
            BASIC = 3,    // simple progress and error handling
            REDUCED = 4,    // authored UI, wizard dialogs suppressed
            FULL = 5,    // authored UI with wizards, progress, errors
            ENDDIALOG = 0x80, // display success/failure dialog at end of install
            PROGRESSONLY = 0x40, // display only progress dialog
            HIDECANCEL = 0x20, // do not display the cancel button in basic UI
            SOURCERESONLY = 0x100, // force display of source resolution even if quiet
        }

        private enum INSTALLMESSAGE
        {
            FATALEXIT = 0x00000000, // premature termination, possibly fatal OOM
            ERROR = 0x01000000, // formatted error message
            WARNING = 0x02000000, // formatted warning message
            USER = 0x03000000, // user request message
            INFO = 0x04000000, // informative message for log
            FILESINUSE = 0x05000000, // list of files in use that need to be replaced
            RESOLVESOURCE = 0x06000000, // request to determine a valid source location
            OUTOFDISKSPACE = 0x07000000, // insufficient disk space message
            ACTIONSTART = 0x08000000, // start of action: action name & description
            ACTIONDATA = 0x09000000, // formatted data associated with individual action item
            PROGRESS = 0x0A000000, // progress gauge info: units so far, total
            COMMONDATA = 0x0B000000, // product info for dialog: language Id, dialog caption
            INITIALIZE = 0x0C000000, // sent prior to UI initialization, no string data
            TERMINATE = 0x0D000000, // sent after UI termination, no string data
            SHOWDIALOG = 0x0E000000, // sent prior to display or authored dialog or wizard
            RMFILESINUSE = 0x19000000, // the list of apps that the user can request Restart Manager to shut down and restart
        }

        private enum INSTALLOGMODE
        {
            FATALEXIT = (1 << (INSTALLMESSAGE.FATALEXIT >> 24)),
            ERROR = (1 << (INSTALLMESSAGE.ERROR >> 24)),
            WARNING = (1 << (INSTALLMESSAGE.WARNING >> 24)),
            USER = (1 << (INSTALLMESSAGE.USER >> 24)),
            INFO = (1 << (INSTALLMESSAGE.INFO >> 24)),
            RESOLVESOURCE = (1 << (INSTALLMESSAGE.RESOLVESOURCE >> 24)),
            OUTOFDISKSPACE = (1 << (INSTALLMESSAGE.OUTOFDISKSPACE >> 24)),
            ACTIONSTART = (1 << (INSTALLMESSAGE.ACTIONSTART >> 24)),
            ACTIONDATA = (1 << (INSTALLMESSAGE.ACTIONDATA >> 24)),
            COMMONDATA = (1 << (INSTALLMESSAGE.COMMONDATA >> 24)),
            PROPERTYDUMP = (1 << (INSTALLMESSAGE.PROGRESS >> 24)), // log only
            VERBOSE = (1 << (INSTALLMESSAGE.INITIALIZE >> 24)), // log only
            EXTRADEBUG = (1 << (INSTALLMESSAGE.TERMINATE >> 24)), // log only
            LOGONLYONERROR = (1 << (INSTALLMESSAGE.SHOWDIALOG >> 24)), // log only	
            PROGRESS = (1 << (INSTALLMESSAGE.PROGRESS >> 24)), // external handler only
            INITIALIZE = (1 << (INSTALLMESSAGE.INITIALIZE >> 24)), // external handler only
            TERMINATE = (1 << (INSTALLMESSAGE.TERMINATE >> 24)), // external handler only
            SHOWDIALOG = (1 << (INSTALLMESSAGE.SHOWDIALOG >> 24)), // external handler only
            FILESINUSE = (1 << (INSTALLMESSAGE.FILESINUSE >> 24)), // external handler only
            RMFILESINUSE = (1 << (INSTALLMESSAGE.RMFILESINUSE >> 24)), // external handler only
        }

        private enum INSTALLLEVEL
        {
            DEFAULT = 0,      // install authored default
            MINIMUM = 1,      // install only required features
            MAXIMUM = 0xFFFF, // install all features
        }

        private enum INSTALLSTATE
        {
            NOTUSED = -7,  // component disabled
            BADCONFIG = -6,  // configuration data corrupt
            INCOMPLETE = -5,  // installation suspended or in progress
            SOURCEABSENT = -4,  // run from source, source is unavailable
            MOREDATA = -3,  // return buffer overflow
            INVALIDARG = -2,  // invalid function argument
            UNKNOWN = -1,  // unrecognized product or feature
            BROKEN = 0,  // broken
            ADVERTISED = 1,  // advertised feature
            REMOVED = 1,  // component being removed (action state, not settable)
            ABSENT = 2,  // uninstalled (or action state absent but clients remain)
            LOCAL = 3,  // installed on local drive
            SOURCE = 4,  // run from source, CD or net
            DEFAULT = 5,  // use default, local or source
        }

        [Flags]
        private enum INSTALLFLAG
        {
            FORCE = 0x00000001,  // Force the installation of the specified driver
            READONLY = 0x00000002,  // Do a read-only install (no file copy)
            NONINTERACTIVE = 0x00000004,  // No UI shown at all. API will fail if any UI must be shown.
        }

        private enum SPOST
        {
            NONE = 0,
            PATH = 1,
            URL = 2,
        }

        [Flags]
        private enum SP_COPY
        {
            DELETESOURCE = 0x0000001,   // delete source file on successful copy
            REPLACEONLY = 0x0000002,   // copy only if target file already present
            NOOVERWRITE = 0x0000008,   // copy only if target doesn't exist
            OEMINF_CATALOG_ONLY = 0x0040000,   // (SetupCopyOEMInf only) don't copy INF--just catalog
        }

        [DllImport("shell32.dll", ExactSpelling = true, CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CommandLineToArgvW
        (
            [In] string lpCmdLine,
            [Out] out int pNumArgs
        );

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)]
        static extern IntPtr LocalFree(IntPtr hMem);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Auto, SetLastError = false)]
        private delegate DialogResult INSTALLUI_HANDLER
        (
            [In] IntPtr pvContext,
            [In] INSTALLMESSAGE iMessageType,
            [In, Optional] string szMessage
        );

        [DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern INSTALLUI_HANDLER MsiSetExternalUI
        (
            [In] INSTALLUI_HANDLER puiHandler,
            [In] INSTALLOGMODE dwMessageFilter,
            [In] IntPtr pvContext
        );

        [DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern INSTALLUILEVEL MsiSetInternalUI
        (
            [In] INSTALLUILEVEL dwUILevel,
            [In, Out] ref IntPtr phWnd
        );

        [DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern int MsiApplyMultiplePatches
        (
            [In] string szPatchPackages,
            [In, Optional] string szProductCode,
            [In, Optional] string szPropertiesList
        );

        [DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern int MsiConfigureProductEx
        (
            [In] string szProduct,
            [In] INSTALLLEVEL iInstallLevel,
            [In] INSTALLSTATE eInstallState,
            [In] string szCommandLine
        );

        [DllImport("msi.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern int MsiInstallProduct
        (
            [In] string szPackagePath,
            [In] string szCommandLine
        );

        [DllImport("newdev.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UpdateDriverForPlugAndPlayDevices
        (
            [In, Optional] IntPtr hwndParent,
            [In] string HardwareId,
            [In] string FullInfPath,
            [In] INSTALLFLAG InstallFlags,
            [Out, Optional] out bool bRebootRequired
        );

        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetupCopyOEMInf
        (
            [In] string SourceInfFileName,
            [In] string OEMSourceMediaLocation,
            [In] SPOST OEMSourceMediaType,
            [In] SP_COPY CopyStyle,
            [In, Optional] IntPtr DestinationInfFileName,
            [In] int DestinationInfFileNameSize,
            [In, Optional] IntPtr RequiredSize,
            [In, Optional] IntPtr DestinationInfFileNameComponent
        );

        private static void ParseWin32ExitCode(int exitCode, out bool isSuccess, out bool rebootRequired, out string additionalMessage)
        {
            switch (exitCode)
            {
                case 1641:
                    Program.Stop();
                    goto case 3010;
                case 3010:
                case 3011:
                    isSuccess = true;
                    rebootRequired = true;
                    additionalMessage = null;
                    break;
                case 0:
                    isSuccess = true;
                    rebootRequired = false;
                    additionalMessage = null;
                    break;
                default:
                    isSuccess = false;
                    rebootRequired = false;
                    additionalMessage = new Win32Exception(exitCode).Message + ".";
                    break;
            }
        }

        private class MSIContext : IDisposable
        {
            private bool isDisposed = false;
            private readonly Task task;
            private readonly TaskProgressChanged progressChanged;
            private readonly INSTALLUI_HANDLER handler;
            private int totalTicks = 0;
            private int stepOnActionData = 0;
            private int currentTicks = 0;
            private bool forward = true;
            private string action = string.Empty;

            internal MSIContext(Task task, TaskProgressChanged progressChanged)
            {
                this.task = task;
                this.progressChanged = progressChanged;
                this.handler = new INSTALLUI_HANDLER(Callback);
            }

            ~MSIContext()
            {
                Dispose(false);
            }

            void IDisposable.Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            public INSTALLUI_HANDLER Handler
            {
                get
                {
                    CheckDisposed();
                    return handler;
                }
            }

            private DialogResult Callback(IntPtr context, INSTALLMESSAGE messageType, string messageText)
            {
                if (isDisposed)
                    return DialogResult.None;
                switch (messageType)
                {
                    case INSTALLMESSAGE.FATALEXIT:
                    case INSTALLMESSAGE.ERROR:
                        if (!string.IsNullOrEmpty(messageText))
                            Program.WriteEvent(Resources.TaskMessage, EventLogEntryType.Error, task, messageText);
                        return DialogResult.None;
                    case INSTALLMESSAGE.WARNING:
                        if (!string.IsNullOrEmpty(messageText))
                            Program.WriteEvent(Resources.TaskMessage, EventLogEntryType.Warning, task, messageText);
                        return DialogResult.None;
                    case INSTALLMESSAGE.ACTIONDATA:
                        action = messageText == null ? string.Empty : messageText;
                        if (stepOnActionData != 0)
                            currentTicks += forward ? stepOnActionData : -stepOnActionData;
                        ReportProgress();
                        return DialogResult.OK;
                    case INSTALLMESSAGE.PROGRESS:
                        if (messageText == null || !HandleProgress(messageText))
                            return DialogResult.None;
                        return DialogResult.OK;
                    default:
                        return DialogResult.None;
                }
            }

            private void ReportProgress()
            {
                progressChanged(action, totalTicks == 0 ? 0 : Math.Min(100, Math.Max(0, (int)Math.Round(currentTicks * 100.0 / totalTicks))));
            }

            private bool HandleProgress(string messageText)
            {
                // see http://msdn.microsoft.com/en-us/library/aa370354(VS.85).aspx
                string[] parts = messageText.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                int progressType;
                int value;
                if
                (
                    (parts.Length & 0x1) != 0 ||
                    parts.Length < 4 ||
                    parts[0] != "1:" || !int.TryParse(parts[1], out progressType) ||
                    parts[2] != "2:" || !int.TryParse(parts[3], out value)
                )
                    return false;
                bool? flag;
                if (parts.Length >= 6)
                {
                    int flagAsInt;
                    if (parts[4] != "3:" || !int.TryParse(parts[5], out flagAsInt))
                        return false;
                    switch (flagAsInt)
                    {
                        case 0:
                            flag = false;
                            break;
                        case 1:
                            flag = true;
                            break;
                        default:
                            return false;
                    }
                }
                else
                    flag = null;
                switch (progressType)
                {
                    case 0: //Resets progress bar and sets the expected total number of ticks in the bar.
                        if (!flag.HasValue)
                            return false;
                        totalTicks = value;
                        forward = !flag.Value;
                        currentTicks = forward ? 0 : totalTicks;
                        ReportProgress();
                        break;
                    case 1: //Provides information related to progress messages to be sent by the current action.
                        if (!flag.HasValue)
                            return false;
                        stepOnActionData = flag.Value ? value : 0;
                        break;
                    case 2: //Increments the progress bar.
                        if (value != 0)
                        {
                            currentTicks += forward ? value : -value;
                            ReportProgress();
                        }
                        break;
                    case 3: //Enables an action (such as CustomAction) to add ticks to the expected total number of progress of the progress bar.
                        totalTicks += value;
                        ReportProgress();
                        break;
                }
                return true;
            }

            protected void CheckDisposed()
            {
                if (isDisposed)
                    throw new ObjectDisposedException(string.Format(Resources.TaskMSIContext, task));
            }

            protected void Dispose(bool isDisposing)
            {
                if (!isDisposed)
                    isDisposed = true;
            }
        }

        private class PSContext : IDisposable
        {
            private readonly Task task;
            private readonly TaskProgressChanged progressChanged;
            private bool isDisposed = false;
            private bool doReboot = false;
            private string action = string.Empty;
            private int progress = -1;

            internal PSContext(Task task, TaskProgressChanged progressChanged)
            {
                this.task = task;
                this.progressChanged = progressChanged;
            }

            ~PSContext()
            {
                Dispose(false);
            }

            void IDisposable.Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            public string Action
            {
                get
                {
                    CheckDisposed();
                    return action;
                }
                set
                {
                    CheckDisposed();
                    if (value == null)
                        value = string.Empty;
                    if (value != action)
                    {
                        action = value;
                        progressChanged(action, progress);
                    }
                }
            }

            public int Progress
            {
                get
                {
                    CheckDisposed();
                    return progress;
                }
                set
                {
                    CheckDisposed();
                    value = Math.Min(100, Math.Max(-1, value));
                    if (value != progress)
                    {
                        progress = value;
                        progressChanged(action, progress);
                    }
                }
            }

            private void Log(EventLogEntryType type, string message)
            {
                if (message == null)
                    throw new ArgumentNullException("message");
                CheckDisposed();
                Program.WriteEvent(Resources.TaskMessage, type, task, message);
            }

            public void Information(string message)
            {
                Log(EventLogEntryType.Information, message);
            }

            public void Warning(string message)
            {
                Log(EventLogEntryType.Warning, message);
            }

            public void Error(string message)
            {
                Log(EventLogEntryType.Error, message);
            }

            public void RequestReboot()
            {
                CheckDisposed();
                doReboot = true;
            }

            public void Stop()
            {
                CheckDisposed();
                Program.Stop();
            }

            internal bool RequiresReboot
            {
                get
                {
                    CheckDisposed();
                    return doReboot;
                }
            }

            protected void CheckDisposed()
            {
                if (isDisposed)
                    throw new ObjectDisposedException(string.Format(Resources.TaskPSContext, task));
            }

            protected void Dispose(bool isDisposing)
            {
                if (!isDisposed)
                    isDisposed = true;
            }
        }

        private class PSHost : System.Management.Automation.Host.PSHost
        {
            private readonly Guid id = Guid.NewGuid();
            private readonly CultureInfo culture = Thread.CurrentThread.CurrentCulture;
            private readonly CultureInfo uiCulture = Thread.CurrentThread.CurrentUICulture;
            private readonly AssemblyName assembly = Assembly.GetExecutingAssembly().GetName();
            private int exitCode = 0;

            public int ExitCode { get { return exitCode; } }
            public override CultureInfo CurrentCulture { get { return culture; } }
            public override CultureInfo CurrentUICulture { get { return uiCulture; } }
            public override void EnterNestedPrompt() { throw new NotImplementedException(); }
            public override void ExitNestedPrompt() { throw new NotImplementedException(); }
            public override Guid InstanceId { get { return id; } }
            public override string Name { get { return assembly.Name; } }
            public override void NotifyBeginApplication() { return; }
            public override void NotifyEndApplication() { return; }
            public override void SetShouldExit(int exitCode) { this.exitCode = exitCode; }
            public override System.Management.Automation.Host.PSHostUserInterface UI { get { return null; } }
            public override Version Version { get { return assembly.Version; } }
        }

        #endregion

        /// <summary>
        /// Creates a new task.
        /// </summary>
        /// <param name="path">The path from which the setup was loaded.</param>
        /// <param name="setup">The underlying setup object.</param>
        internal Task(string path, Setup setup)
        {
            // set the member variables and load the setup image
            this.setup = setup;
            name = Environment.ExpandEnvironmentVariables(setup.Name);
            directory = Path.GetDirectoryName(path);
            try
            {
                string imagePath = Path.Combine(directory, setup.Image);
                image = Image.FromFile(imagePath);
            }
            catch (Exception e)
            {
                Program.WriteEvent(Resources.TaskLoadImageFailed, EventLogEntryType.Warning, this, setup.Image, e.Message);
                image = null;
            }
        }

        /// <summary>
        /// Gets the name of the task.
        /// </summary>
        public string Name { get { return name; } }

        /// <summary>
        /// Gets the task image.
        /// </summary>
        public Image Image { get { return image; } }

        /// <summary>
        /// Indicates whether this tasks should be run exclusively.
        /// </summary>
        public bool IsExclusive { get { return setup.Exclusive; } }

        /// <summary>
        /// Executes the task.
        /// </summary>
        /// <param name="progressChanged">Callback for progress notifications.</param>
        public void Run(TaskProgressChanged progressChanged)
        {
            bool isSuccess;
            bool rebootRequired;
            string additionalMessage;
            Environment.SetEnvironmentVariable(".", directory);
            string oldCurrentDirectory = Environment.CurrentDirectory;
            Environment.CurrentDirectory = directory;
            try
            {
                switch (setup.Type)
                {
                    case SetupType.WindowsInstallerPackage:
                    case SetupType.WindowsInstallerPatch:
                    case SetupType.WindowsInstallerRemoval:
                        {
                            // handle MSI/MSP files and product removals, start by disabling the Windows Installer UI
                            IntPtr prevWindow = IntPtr.Zero;
                            INSTALLUILEVEL prevUILevel = MsiSetInternalUI(INSTALLUILEVEL.NONE, ref prevWindow);
                            try
                            {
                                // create the callback-context
                                using (MSIContext context = new MSIContext(this, progressChanged))
                                {
                                    // register the context with the installer for action data and progress information
                                    INSTALLUI_HANDLER prevHandler = MsiSetExternalUI(context.Handler, INSTALLOGMODE.FATALEXIT | INSTALLOGMODE.ERROR | INSTALLOGMODE.WARNING | INSTALLOGMODE.ACTIONDATA | INSTALLOGMODE.PROGRESS, IntPtr.Zero);
                                    try
                                    {
                                        // launch the installer and parse the results
                                        string commonProperties = "REBOOT=ReallySuppress";
                                        string properties = string.IsNullOrEmpty(setup.Parameters) ? commonProperties : string.Join(" ", new string[] { commonProperties, Environment.ExpandEnvironmentVariables(setup.Parameters) });
                                        string productOrPatches = Environment.ExpandEnvironmentVariables(setup.FileName);
                                        int exitCode;
                                        switch (setup.Type)
                                        {
                                            case SetupType.WindowsInstallerPackage:
                                                exitCode = MsiInstallProduct(productOrPatches, properties);
                                                break;
                                            case SetupType.WindowsInstallerPatch:
                                                exitCode = MsiApplyMultiplePatches(productOrPatches, null, properties);
                                                break;
                                            case SetupType.WindowsInstallerRemoval:
                                                exitCode = MsiConfigureProductEx(productOrPatches, INSTALLLEVEL.DEFAULT, INSTALLSTATE.ABSENT, properties);
                                                break;
                                            default:
                                                throw new NotImplementedException();
                                        }
                                        ParseWin32ExitCode(exitCode, out isSuccess, out rebootRequired, out additionalMessage);
                                    }
                                    finally { MsiSetExternalUI(prevHandler, 0, IntPtr.Zero); }
                                }
                            }
                            finally { MsiSetInternalUI(prevUILevel, ref prevWindow); }
                        }
                        break;

                    case SetupType.WindowsUpdateFile:
                    case SetupType.WindowsUpdateRemoval:
                        {
                            // launch WUSA with the specified update file to install or kb id to remove and parse the results
                            ProcessStartInfo psi = new ProcessStartInfo();
                            psi.UseShellExecute = false;
                            psi.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "wusa.exe");
                            string commonArguments;
                            switch (setup.Type)
                            {
                                case SetupType.WindowsUpdateFile:
                                    commonArguments = "/quiet /norestart \"{0}\"";
                                    break;
                                case SetupType.WindowsUpdateRemoval:
                                    commonArguments = "/uninstall /kb:{0} /quiet /norestart";
                                    break;
                                default:
                                    throw new NotImplementedException();
                            }
                            commonArguments = string.Format(commonArguments, Environment.ExpandEnvironmentVariables(setup.FileName));
                            psi.Arguments = string.IsNullOrEmpty(setup.Parameters) ? commonArguments : string.Join(" ", new string[] { commonArguments, Environment.ExpandEnvironmentVariables(setup.Parameters) });
                            using (Process process = Process.Start(psi))
                            {
                                process.WaitForExit();
                                switch ((uint)process.ExitCode)
                                {
                                    case 1:
                                        isSuccess = false;
                                        rebootRequired = false;
                                        additionalMessage = null;
                                        break;
                                    case 0x240006:
                                        isSuccess = true;
                                        rebootRequired = false;
                                        additionalMessage = Resources.TaskUpdateAlreadyInstalled;
                                        break;
                                    case 0x240007:
                                        isSuccess = true;
                                        rebootRequired = false;
                                        additionalMessage = Resources.TaskUpdateNotInstalled;
                                        break;
                                    case 0x80240017:
                                        isSuccess = false;
                                        rebootRequired = false;
                                        additionalMessage = Resources.TaskUpdateNotApplicable;
                                        break;
                                    default:
                                        ParseWin32ExitCode(process.ExitCode, out isSuccess, out rebootRequired, out additionalMessage);
                                        break;
                                }
                            }
                        }
                        break;

                    case SetupType.PowerShellScript:
                        {
                            // run a PowerShell script
                            PSHost host = new PSHost();
                            using (PSContext context = new PSContext(this, progressChanged))
                            using (Runspace runspace = RunspaceFactory.CreateRunspace(host))
                            {
                                runspace.Open();
                                runspace.SessionStateProxy.SetVariable("Session", context);
                                using (Pipeline pipeline = runspace.CreatePipeline())
                                {
                                    StringBuilder script = new StringBuilder();
                                    script.AppendLine(".{");
                                    script.AppendLine(File.ReadAllText(Environment.ExpandEnvironmentVariables(setup.FileName)));
                                    script.Append("}");
                                    if (!string.IsNullOrEmpty(setup.Parameters))
                                    {
                                        script.Append(' ');
                                        script.Append(Environment.ExpandEnvironmentVariables(setup.Parameters));
                                    }
                                    script.AppendLine();
                                    pipeline.Commands.AddScript(script.ToString());
                                    try { pipeline.Invoke(); }
                                    catch (Exception e)
                                    {
                                        isSuccess = false;
                                        rebootRequired = context.RequiresReboot;
                                        additionalMessage = e.Message;
                                        break;
                                    }
                                    bool win32Reboot;
                                    ParseWin32ExitCode(host.ExitCode, out isSuccess, out win32Reboot, out additionalMessage);
                                    rebootRequired = context.RequiresReboot || win32Reboot;
                                }
                            }
                        }
                        break;

                    case SetupType.Executable:
                    case SetupType.ExecutableNoWindow:
                        {
                            // execute a windows binary file, plain and easy
                            ProcessStartInfo psi = new ProcessStartInfo();
                            psi.UseShellExecute = false;
                            psi.CreateNoWindow = setup.Type == SetupType.ExecutableNoWindow;
                            psi.FileName = Environment.ExpandEnvironmentVariables(setup.FileName);
                            if (!string.IsNullOrEmpty(setup.Parameters))
                                psi.Arguments = Environment.ExpandEnvironmentVariables(setup.Parameters);
                            using (Process process = Process.Start(psi))
                            {
                                process.WaitForExit();
                                ParseWin32ExitCode(process.ExitCode, out isSuccess, out rebootRequired, out additionalMessage);
                            }
                        }
                        break;

                    case SetupType.DeviceDriverFile:
                    case SetupType.DeviceDriverFileInteractive:
                        {
                            // copy the driver
                            string infFile = Environment.ExpandEnvironmentVariables(setup.FileName);
                            if (SetupCopyOEMInf(infFile, null, SPOST.NONE, 0, IntPtr.Zero, 0, IntPtr.Zero, IntPtr.Zero))
                            {
                                // quit if no devices should be updated
                                if (string.IsNullOrEmpty(setup.Parameters))
                                {
                                    isSuccess = true;
                                    rebootRequired = false;
                                    additionalMessage = null;
                                    break;
                                }

                                // calculate the install flag
                                INSTALLFLAG flags = INSTALLFLAG.FORCE;
                                if (setup.Type != SetupType.DeviceDriverFileInteractive)
                                    flags |= INSTALLFLAG.NONINTERACTIVE;

                                // update all matching devices
                                rebootRequired = false;
                                string[] hardwareIds = Environment.ExpandEnvironmentVariables(setup.Parameters).Split(Path.PathSeparator);
                                int i;
                                for (i = 0; i < hardwareIds.Length; i++)
                                {
                                    string hardwareId = hardwareIds[i].Trim();
                                    if (hardwareId.Length == 0)
                                        continue;
                                    bool rebootFlag;
                                    if (!UpdateDriverForPlugAndPlayDevices(IntPtr.Zero, hardwareId, infFile, flags, out rebootFlag))
                                        break;
                                    if (rebootFlag)
                                        rebootRequired = true;
                                }

                                // handle the result
                                if (i == hardwareIds.Length)
                                {
                                    isSuccess = true;
                                    additionalMessage = null;
                                    break;
                                }
                            }
                            else
                                rebootRequired = false;

                            // an error occured
                            int lastError = Marshal.GetLastWin32Error();
                            isSuccess = false;
                            additionalMessage = lastError == 0 ? null : new Win32Exception(lastError).Message + ".";
                        }
                        break;

                    default:
                        throw new Exception(string.Format(Resources.TaskSetupTypeUnknown, setup.Type));
                }
            }
            catch (Exception e)
            {
                Program.WriteEvent(Resources.TaskLaunchFailed, EventLogEntryType.Error, this, e.Message);
                return;
            }
            finally { Environment.CurrentDirectory = oldCurrentDirectory; }

            // schedule the reboot
            if (setup.Reboot == SetupReboot.Always || (setup.Reboot == SetupReboot.IfRequired && rebootRequired))
                Program.ScheduleReboot();

            // set the general status message and log type
            string statusMessage;
            EventLogEntryType type;
            if (isSuccess)
            {
                // quit if we shouldn't log success
                if (!setup.LogSuccess)
                    return;

                type = EventLogEntryType.Information;
                if (rebootRequired)
                    statusMessage = Resources.TaskRebootRequired;
                else
                    statusMessage = Resources.TaskSucceeded;
            }
            else
            {
                type = EventLogEntryType.Warning;
                statusMessage = Resources.TaskSetupFailed;
            }
            statusMessage = string.Format(statusMessage, this);

            // log the result of the operation
            if (additionalMessage == null)
                Program.WriteEvent(statusMessage, type);
            else
                Program.WriteEvent(Resources.TaskDetailedResult, type, statusMessage, additionalMessage);
        }

        public override string ToString()
        {
            return name;
        }
    }
}
