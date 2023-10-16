using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Valve.VR;

namespace OpenVRStartup
{
    class Program
    {
        static readonly string PATH_LOGFILE = "./OpenVRStartup.log";
        static readonly string PATH_BOOTFOLDER = "./boot/";
        static readonly string PATH_STARTFOLDER = "./start/";
        static readonly string PATH_STOPFOLDER = "./stop/";
        static readonly string FILE_PATTERN = "*.cmd";

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        public const int SW_MINIMIZE = 6; // This should minimize the window and let some other app remain active
        // See https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow

        static ManualResetEventSlim _isReady = new ManualResetEventSlim(false);
        static ManualResetEventSlim _isConnected = new ManualResetEventSlim(false);

        static CancellationTokenSource _cts = new CancellationTokenSource();

        static void Main(string[] _)
        {
            // Window setup
            Console.Title = Properties.Resources.AppName;

            // Starting worker
            var t = new Thread(Worker);
            LogUtils.WriteLineToCache($"Application starting ({Properties.Resources.Version})");
            if (!t.IsAlive) t.Start();
            else LogUtils.WriteLineToCache("Error: Could not start worker thread");

            // Check if first run, if so do NOT minimize yet - write instructions and wait for acknowledgement
            if (!LogUtils.LogFileExists(PATH_LOGFILE))
            {
                Utils.PrintInfo("\n========================");
                Utils.PrintInfo(" First Run Instructions ");
                Utils.PrintInfo("========================");
                Utils.Print("\nThis app automatically sets itself to auto-launch with SteamVR.");
                Utils.Print($"\nWhen it runs it will in turn run all {FILE_PATTERN} files in the {PATH_STARTFOLDER} folder.");
                Utils.Print($"\nIf there are {FILE_PATTERN} files in {PATH_STOPFOLDER} it will stay and run those on shutdown.");
                Utils.Print("\nThis message is only shown once, to see it again delete the log file.");
                Utils.Print("\nPress [Enter] in this window to continue execution.\nIf there are shutdown scripts the window will remain in the task bar.");
                Console.ReadLine();
            }

            // Indicate that we're ready for the worker to do its thing, minimize, and wait for an [Enter] key to manually exit
            _isReady.Set();
            Minimize();

            Thread.Sleep(TimeSpan.FromSeconds(2));
            Utils.Print("\nPress [Enter] to force quit");
            Console.ReadLine();

            // If manually exiting, tell the woker thread to cancel and wait for it to wrap up before quitting
            _cts.Cancel();
            t.Join(); // This shouldn't ever finish because the thread should exit the process when it finishes
        }

        private static void Minimize()
        {
            IntPtr winHandle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(winHandle, SW_MINIMIZE);
        }

        private static void Worker()
        {
            Thread.CurrentThread.IsBackground = true;

            // Run BOOT scripts immediately without waiting for OpenVR connection or anything
            // This doesn't seem to make much difference (not much longer before we run START scripts) but might
            // help with powering up base stations a tiny bit quicker; main delay in a 'normal' SteamVR startup
            // seems to be just waiting for it to initialize itself to the point where it starts loading overlays
            // like this little helper
            Utils.Print("Running BOOT scripts");
            RunScripts(PATH_BOOTFOLDER);

            var token = _cts.Token;
            while (!token.IsCancellationRequested)
            {
                if (!_isConnected.IsSet)
                {
                    // If we haven't yet connected to OpenVR, try now
                    InitVR();
                }

                if (_isConnected.IsSet && _isReady.IsSet)
                {
                    // We're connected to OpenVR and ready to run START scripts
                    Utils.Print("Running START scripts");
                    RunScripts(PATH_STARTFOLDER);

                    // If we have STOP scripts, wait to run them until SteamVR is exiting
                    if (WeHaveScripts(PATH_STOPFOLDER))
                    {
                        WaitForQuit();
                        Utils.Print("Running STOP scripts");
                        RunScripts(PATH_STOPFOLDER);
                    }

                    // STOP scripts (if any) have run - signal completion (for anyone else that might be listening) and exit the loop
                    _cts.Cancel();
                    break;
                }

                // Brief sleep while waiting to connect to OpenVR and for the main thread to signal readiness
                Thread.Sleep(100);
            }

            // Flush log to file before we exit
            LogUtils.WriteLineToCache("Application exiting, writing log");
            LogUtils.WriteCacheToLogFile(PATH_LOGFILE, 100);

            if (_isConnected.IsSet)
            {
                OpenVR.Shutdown();
            }
            Environment.Exit(0);
        }

        // Initializing connection to OpenVR
        private static void InitVR()
        {
            var error = EVRInitError.None;
            OpenVR.Init(ref error, EVRApplicationType.VRApplication_Overlay);
            if (error != EVRInitError.None)
            {
                LogUtils.WriteLineToCache($"Error: OpenVR init failed: {Enum.GetName(typeof(EVRInitError), error)}");
            }
            else
            {
                LogUtils.WriteLineToCache("OpenVR init success");

                // Add app manifest and set auto-launch
                var appKey = "boll7708.openvrstartup";
                if (!OpenVR.Applications.IsApplicationInstalled(appKey))
                {
                    var manifestError = OpenVR.Applications.AddApplicationManifest(Path.GetFullPath("./app.vrmanifest"), false);
                    if (manifestError == EVRApplicationError.None) LogUtils.WriteLineToCache("Successfully installed app manifest");
                    else LogUtils.WriteLineToCache($"Error: Failed to add app manifest: {Enum.GetName(typeof(EVRApplicationError), manifestError)}");

                    var autolaunchError = OpenVR.Applications.SetApplicationAutoLaunch(appKey, true);
                    if (autolaunchError == EVRApplicationError.None) LogUtils.WriteLineToCache("Successfully set app to auto launch");
                    else LogUtils.WriteLineToCache($"Error: Failed to turn on auto launch: {Enum.GetName(typeof(EVRApplicationError), autolaunchError)}");
                }

                // Set _isConnected to indicate we've initialized OpenVR and can proceed
                _isConnected.Set();
            }
        }

        // Scripts
        private static void RunScripts(string folder)
        {
            try
            {
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                var files = Directory.GetFiles(folder, FILE_PATTERN);
                LogUtils.WriteLineToCache($"Found: {files.Length} script(s) in {folder}");
                foreach (var file in files)
                {
                    LogUtils.WriteLineToCache($"Executing: {file}");
                    var path = Path.Combine(Environment.CurrentDirectory, file);
                    Process p = new Process();
                    p.StartInfo.CreateNoWindow = true;
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.StartInfo.FileName = Path.Combine(Environment.SystemDirectory, "cmd.exe");
                    p.StartInfo.Arguments = $"/C \"{path}\"";
                    p.Start();
                }
                if (files.Length == 0) LogUtils.WriteLineToCache($"Did not find any {FILE_PATTERN} files to execute in {folder}");
            }
            catch (Exception e)
            {
                LogUtils.WriteLineToCache($"Error: Could not load scripts from {folder}: {e.Message}");
            }
        }

        private static void WaitForQuit()
        {
            Utils.Print("This window remains to wait for the shutdown of SteamVR to run additional scripts on exit.");
            var token = _cts.Token;

            while (!token.IsCancellationRequested)
            {
                var vrEvents = new List<VREvent_t>();
                var vrEvent = new VREvent_t();
                uint eventSize = (uint)Marshal.SizeOf(vrEvent);
                try
                {
                    while (OpenVR.System.PollNextEvent(ref vrEvent, eventSize))
                    {
                        vrEvents.Add(vrEvent);
                    }
                }
                catch (Exception e)
                {
                    Utils.PrintError($"Could not get new events: {e.Message}");
                }

                foreach (var e in vrEvents)
                {
                    if ((EVREventType)e.eventType == EVREventType.VREvent_Quit)
                    {
                        OpenVR.System.AcknowledgeQuit_Exiting();
                        Utils.Print("OpenVR exiting...");
                        _cts.Cancel();
                    }
                }

                // Longer wait when polling for OpenVR events
                Thread.Sleep(1000);
            }
            Utils.Print("WaitForQuit finished");
        }

        private static bool WeHaveScripts(string folder)
        {
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            return Directory.GetFiles(folder, FILE_PATTERN).Length > 0;
        }
    }
}
