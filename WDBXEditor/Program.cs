using System;
using System.Windows.Forms;
using WDBXEditor.ConsoleHandler;

namespace WDBXEditor
{
    static class Program
    {
        public static bool PrimaryInstance = false;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string[] appArgs = InstanceManager.StripControlArgs(args);
            bool forceNewInstance = InstanceManager.HasNewInstanceFlag(args);

            InstanceManager.InstanceCheck(appArgs, forceNewInstance); // Check whether to reuse the primary instance or launch a new one
            InstanceManager.LoadDll("StormLib.dll"); // Loads the correct StormLib library

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            UpdateManager.Clean();

            if (appArgs != null && appArgs.Length > 0)
            {
                ConsoleManager.LoadCommandDefinitions();

                string commandKey = appArgs[0].ToLowerInvariant();
                if (ConsoleManager.CommandHandlers.ContainsKey(commandKey))
                    ConsoleManager.ConsoleMain(appArgs); // Console mode
                else
                    Application.Run(new Main(appArgs)); // Load file(s)
            }
            else
            {
                Application.Run(new Main()); // Default
            }

            InstanceManager.Stop();
        }
    }
}
