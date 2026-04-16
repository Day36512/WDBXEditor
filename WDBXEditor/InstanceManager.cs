using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WDBXEditor.Common;


namespace WDBXEditor
{
    public static class InstanceManager
    {
        public static ConcurrentQueue<string> AutoRun = new ConcurrentQueue<string>();
        public static Action AutoRunAdded;

        public const string NewInstanceArg = "--new-instance";

        private static readonly HashSet<string> NewInstanceArgs = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "--new-instance",
            "/new-instance",
            "-new-instance",
            "--new",
            "/new",
            "-new",
            "--multi-instance",
            "/multi-instance",
            "-multi-instance"
        };

        private static Mutex mutex;
        private static NamedPipeManager pipeServer;

        /// <summary>
        /// Checks a mutex to see if an instance is running and decides how to proceed.
        /// By default the app stays single-instance for normal launches and file-open requests,
        /// but the --new-instance flag bypasses that behavior entirely.
        /// </summary>
        public static void InstanceCheck(string[] args, bool forceNewInstance = false)
        {
            if (forceNewInstance)
            {
                Program.PrimaryInstance = false;
                return;
            }

            if (!ShouldUseSingleInstanceRouting(args))
                return;

            bool isOnlyInstance;
            mutex = new Mutex(true, "WDBXEditorMutex", out isOnlyInstance);
            if (!isOnlyInstance)
            {
                Program.PrimaryInstance = false;
                SendData(args); // Send args to the primary instance
            }
            else
            {
                Program.PrimaryInstance = true;
                pipeServer = new NamedPipeManager();
                pipeServer.ReceiveString += OpenRequest;
                pipeServer.StartServer();
            }
        }

        public static bool HasNewInstanceFlag(string[] args)
        {
            return args != null && args.Any(arg => !string.IsNullOrWhiteSpace(arg) && NewInstanceArgs.Contains(arg.Trim()));
        }

        public static string[] StripControlArgs(string[] args)
        {
            if (args == null || args.Length == 0)
                return Array.Empty<string>();

            return args
                .Where(arg => !string.IsNullOrWhiteSpace(arg) && !NewInstanceArgs.Contains(arg.Trim()))
                .ToArray();
        }

        private static bool ShouldUseSingleInstanceRouting(string[] args)
        {
            if (args == null || args.Length == 0)
                return true;

            return args.Any(File.Exists);
        }

        public static void LoadDll(string lib)
        {
            string startupDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            string stormlibPath = Path.Combine(startupDirectory, lib);
            bool copyDll = true;

            if (File.Exists(stormlibPath)) // If the file exists check if it is the right architecture
            {
                byte[] data = new byte[4096];
                using (Stream s = new FileStream(stormlibPath, FileMode.Open, FileAccess.Read))
                    s.Read(data, 0, 4096);

                int peHeaderAddr = BitConverter.ToInt32(data, 0x3C);
                bool x86 = BitConverter.ToUInt16(data, peHeaderAddr + 0x4) == 0x014c; // 32-bit check
                copyDll = (x86 != !Environment.Is64BitProcess);
            }

            if (copyDll)
            {
                string copyPath = Path.Combine(startupDirectory, Environment.Is64BitProcess ? "x64" : "x86", lib);
                if (File.Exists(copyPath))
                    File.Copy(copyPath, stormlibPath, true);
            }
        }

        /// <summary>
        /// Enqueues received file names and launches the AutoRun delegate.
        /// </summary>
        public static void OpenRequest(string filenames)
        {
            string[] files = (filenames ?? string.Empty).Split((char)3);
            Parallel.For(0, files.Length, f =>
            {
                if (Regex.IsMatch(files[f], Constants.FileRegexPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase))
                    AutoRun.Enqueue(files[f]);
            });

            AutoRunAdded?.Invoke();
        }

        public static void Start()
        {
            pipeServer?.StartServer();
        }

        public static void Stop()
        {
            if (pipeServer != null)
            {
                pipeServer.ReceiveString -= OpenRequest;
                pipeServer.StopServer();
                pipeServer = null;
            }

            if (mutex != null)
            {
                try { mutex.ReleaseMutex(); } catch { }
                try { mutex.Dispose(); } catch { }
                mutex = null;
            }
        }

        /// <summary>
        /// Opens a new version of the application which bypasses the mutex.
        /// </summary>
        public static bool LoadNewInstance(IEnumerable<string> files)
        {
            using (Process p = new Process())
            {
                p.StartInfo.FileName = Application.ExecutablePath;
                p.StartInfo.Arguments = BuildNewInstanceArguments(files);
                return p.Start();
            }
        }

        private static string BuildNewInstanceArguments(IEnumerable<string> files)
        {
            List<string> args = new List<string> { NewInstanceArg };

            if (files != null)
                args.AddRange(files.Where(file => !string.IsNullOrWhiteSpace(file)).Select(QuoteArgument));

            return string.Join(" ", args);
        }

        private static string QuoteArgument(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "\"\"";

            if (!value.Any(ch => char.IsWhiteSpace(ch) || ch == '"'))
                return value;

            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        public static IEnumerable<string> GetFilesToOpen()
        {
            HashSet<string> files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (AutoRun.Count > 0)
            {
                string file;
                if (AutoRun.TryDequeue(out file) && File.Exists(file))
                    files.Add(file);
            }
            return files;
        }

        public static bool IsRunningAsAdmin()
        {
            WindowsPrincipal principal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        #region Send Data
        private static void SendData(string args)
        {
            NamedPipeManager clientPipe = new NamedPipeManager();
            if (clientPipe.Write(args ?? string.Empty))
                Environment.Exit(0);
        }

        private static void SendData(string[] args)
        {
            SendData(args == null ? string.Empty : string.Join(((char)3).ToString(), args));
        }
        #endregion

        #region Flash Methods
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        internal struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        public static bool FlashWindow(Form form)
        {
            if (Type.GetType("Mono.Runtime") != null)
                return false;

            FLASHWINFO fInfo = new FLASHWINFO();

            uint flashwAll = 3;
            uint flashwTimernofg = 12;

            fInfo.cbSize = Convert.ToUInt32(Marshal.SizeOf(fInfo));
            fInfo.hwnd = form.Handle;
            fInfo.dwFlags = flashwAll | flashwTimernofg;
            fInfo.uCount = uint.MaxValue;
            fInfo.dwTimeout = 0;

            return FlashWindowEx(ref fInfo);
        }
        #endregion
    }
}
