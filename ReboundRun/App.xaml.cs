using IWshRuntimeLibrary;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Xaml.Shapes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Activation;
using Windows.ApplicationModel.Background;
using Windows.Foundation;
using Windows.Foundation.Collections;
using WinUIEx;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ReboundRun
{
    /// <summary>
    /// Provides application-specific behavior to supplement the default Application class.
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// Initializes the singleton application object.  This is the first line of authored code
        /// executed, and as such is the logical equivalent of main() or WinMain().
        /// </summary>
        public App()
        {
            this.InitializeComponent();
        }

        public static WindowEx bkgWindow;

        /// <summary>
        /// Invoked when the application is launched.
        /// </summary>
        /// <param name="args">Details about the launch request and process.</param>
        protected override async void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            CreateShortcut();
            try
            {
                bkgWindow = new WindowEx();
                bkgWindow.SystemBackdrop = new TransparentTintBackdrop();
                bkgWindow.IsMaximizable = false;
                bkgWindow.SetExtendedWindowStyle(ExtendedWindowStyle.ToolWindow);
                bkgWindow.SetWindowStyle(WindowStyle.Visible);
                bkgWindow.Activate();
                bkgWindow.MoveAndResize(0, 0, 0, 0);
                bkgWindow.Minimize();
                bkgWindow.SetWindowOpacity(0);
            }
            catch
            {

            }
            m_window = new MainWindow();
            m_window.Activate();
            await Task.Delay(10);
            (m_window as WindowEx).SetForegroundWindow();
            (m_window as WindowEx).BringToFront();
            if (string.Join(" ", Environment.GetCommandLineArgs().Skip(1)).Contains("STARTUP"))
            {
                m_window.Close();
            }
            //RegisterBackgroundTask();
            StartHook();
        }

        private void CreateShortcut()
        {
            string startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = System.IO.Path.Combine(startupFolderPath, "ReboundRun.lnk");
            string appPath = "C:\\Rebound11\\rrunSTARTUP.exe";

            if (!System.IO.File.Exists(shortcutPath))
            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.Description = "Rebound Run";
                shortcut.TargetPath = appPath;
                shortcut.Save();
            }
        }

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int VK_R = 0x52;
        private const int VK_LWIN = 0x5B;

        private static IntPtr hookId = IntPtr.Zero;
        private static LowLevelKeyboardProc keyboardProc;

        public static void StartHook()
        {
            keyboardProc = HookCallback;
            hookId = SetHook(keyboardProc);
        }

        public static void StopHook()
        {
            UnhookWindowsHookEx(hookId);
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                // Check for Win + R combination
                if (vkCode == VK_R && (GetKeyState(VK_LWIN) & 0x8000) != 0)
                {
                    ((WindowEx)m_window).Show();
                    ((WindowEx)m_window).BringToFront();
                    try
                    {
                        ((WindowEx)m_window).Activate();
                    }
                    catch
                    {
                        m_window = new MainWindow();
                        m_window.Show();
                        ((WindowEx)m_window).BringToFront();
                    }
                    // Prevent default behavior of Win + R
                    return (IntPtr)1;
                }
            }

            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true, CallingConvention = CallingConvention.Winapi)]
        public static extern short GetKeyState(int keyCode);

        private void RegisterBackgroundTask()
        {
            var taskRegistered = false;
            foreach (var task in BackgroundTaskRegistration.AllTasks)
            {
                if (task.Value.Name == "ReboundRunBackgroundActivation")
                {
                    taskRegistered = true;
                    break;
                }
            }

            if (!taskRegistered)
            {
                var builder = new BackgroundTaskBuilder
                {
                    Name = "ReboundRunBackgroundActivation",
                    TaskEntryPoint = "ReboundRun.ReboundRunBackgroundActivation"
                };

                // Set a trigger for your background task
                builder.SetTrigger(new TimeTrigger(15, false)); // For example, trigger every 15 minutes

                // Register the background task
                builder.Register();
            }
        }

        public static Window m_window;
    }

    public sealed class ReboundRunBackgroundActivation : IBackgroundTask
    {
        public async void Run(IBackgroundTaskInstance taskInstance)
        {
            var deferral = taskInstance.GetDeferral();

            // Run a loop in the background
            try
            {
                while (true)
                {
                    // Perform some background work
                    // Example: Log the current time
                    System.Diagnostics.Debug.WriteLine($"Background task running at {DateTime.Now}");

                    // Sleep for a specified interval before running the loop again
                    await Task.Delay(TimeSpan.FromMinutes(1)); // Adjust as needed
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
                System.Diagnostics.Debug.WriteLine($"Error in background task: {ex.Message}");
            }
            finally
            {
                deferral.Complete();
            }
        }
    }
}
