using IWshRuntimeLibrary;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.UI.Xaml.Shapes;
using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Activation;
using Windows.ApplicationModel.Background;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.System;
using Windows.UI.Input.Preview.Injection;
using WinUIEx;
using File = System.IO.File;
using Path = System.IO.Path;
using Task = System.Threading.Tasks.Task;
using TimeTrigger = Windows.ApplicationModel.Background.TimeTrigger;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ReboundRun
{
    public class SingleInstanceLaunchEventArgs : EventArgs
    {
        public SingleInstanceLaunchEventArgs(string arguments, bool isFirstLaunch)
        {
            Arguments = arguments;
            IsFirstLaunch = isFirstLaunch;
        }
        public string Arguments { get; private set; } = "";
        public bool IsFirstLaunch { get; private set; }
    }

    public sealed class SingleInstanceDesktopApp : IDisposable
    {
        private readonly string _mutexName = "";
        private readonly string _pipeName = "";
        private readonly object _namedPiperServerThreadLock = new();

        private bool _isDisposed = false;
        private bool _isFirstInstance;

        private Mutex? _mutexApplication;
        private NamedPipeServerStream? _namedPipeServerStream;

        public event EventHandler<SingleInstanceLaunchEventArgs>? Launched;

        public SingleInstanceDesktopApp(string appId)
        {
            _mutexName = "MUTEX_" + appId;
            _pipeName = "PIPE_" + appId;
        }

        public void Launch(string arguments)
        {
            if (string.IsNullOrEmpty(arguments))
            {
                // The arguments from LaunchActivatedEventArgs can be empty, when
                // the user specified arguments (e.g. when using an execution alias). For this reason we
                // alternatively check for arguments using a different API.
                var argList = System.Environment.GetCommandLineArgs();
                if (argList.Length > 1)
                {
                    arguments = string.Join(' ', argList.Skip(1));
                }
            }

            if (IsFirstApplicationInstance())
            {
                CreateNamedPipeServer();
                Launched?.Invoke(this, new SingleInstanceLaunchEventArgs(arguments, isFirstLaunch: true));
            }
            else
            {
                SendArgumentsToRunningInstance(arguments);

                Process.GetCurrentProcess().Kill();
                // Note: needed to kill the process in WinAppSDK 1.0, since Application.Current.Exit() does not work there.
                // OR: Application.Current.Exit();
            }
        }

        public void Dispose()
        {
            if (_isDisposed)
                return;

            _isDisposed = true;

            _namedPipeServerStream?.Dispose();
            _mutexApplication?.Dispose();
        }

        private bool IsFirstApplicationInstance()
        {
            // Allow for multiple runs but only try and get the mutex once
            if (_mutexApplication == null)
            {
                _mutexApplication = new Mutex(true, _mutexName, out _isFirstInstance);
            }

            return _isFirstInstance;
        }

        /// <summary>
        /// Starts a new pipe server if one isn't already active.
        /// </summary>
        private void CreateNamedPipeServer()
        {
            _namedPipeServerStream = new NamedPipeServerStream(
                _pipeName, PipeDirection.In,
                maxNumberOfServerInstances: 1,
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous,
                inBufferSize: 0,
                outBufferSize: 0);

            _namedPipeServerStream.BeginWaitForConnection(OnNamedPipeServerConnected, _namedPipeServerStream);
        }

        private void SendArgumentsToRunningInstance(string arguments)
        {
            try
            {
                using var namedPipeClientStream = new NamedPipeClientStream(".", _pipeName, PipeDirection.Out);
                namedPipeClientStream.Connect(3000); // Maximum wait 3 seconds
                using var sw = new StreamWriter(namedPipeClientStream);
                sw.Write(arguments);
                sw.Flush();
            }
            catch (Exception)
            {
                // Error connecting or sending
            }
        }

        private void OnNamedPipeServerConnected(IAsyncResult asyncResult)
        {
            try
            {
                if (_namedPipeServerStream == null)
                    return;

                _namedPipeServerStream.EndWaitForConnection(asyncResult);

                // Read the arguments from the pipe
                lock (_namedPiperServerThreadLock)
                {
                    using var sr = new StreamReader(_namedPipeServerStream);
                    var args = sr.ReadToEnd();
                    Launched?.Invoke(this, new SingleInstanceLaunchEventArgs(args, isFirstLaunch: false));
                }
            }
            catch (ObjectDisposedException)
            {
                // EndWaitForConnection will throw when the pipe closes before there is a connection.
                // In that case, we don't create more pipes and just return.
                // This will happen when the app is closed and therefor the pipe is closed as well.
                return;
            }
            catch (Exception)
            {
                // ignored
            }
            finally
            {
                // Close the original pipe (we will create a new one each time)
                _namedPipeServerStream?.Dispose();
            }

            // Create a new pipe for next connection
            CreateNamedPipeServer();
        }
    }
    /// <summary>
    /// Provides application-specific behavior to supplement the default Application class.
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// Initializes the singleton application object.  This is the first line of authored code
        /// executed, and as such is the logical equivalent of main() or WinMain().
        /// </summary>
        
        private readonly SingleInstanceDesktopApp _singleInstanceApp;

        public App()
        {
            this.InitializeComponent();

            _singleInstanceApp = new SingleInstanceDesktopApp("REBOUNDRUN");
            _singleInstanceApp.Launched += OnSingleInstanceLaunched;
        }

        public static WindowEx bkgWindow;

        /// <summary>
        /// Invoked when the application is launched.
        /// </summary>
        /// <param name="args">Details about the launch request and process.</param>
        protected override async void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            _singleInstanceApp.Launch(args.Arguments);
        }

        private async void OnSingleInstanceLaunched(object? sender, SingleInstanceLaunchEventArgs e)
        {
            if (e.IsFirstLaunch)
            {
                await LaunchWork();
            }
            else
            {
                // Get the current process
                Process currentProcess = Process.GetCurrentProcess();

                // Start a new instance of the application
                Process.Start(currentProcess.MainModule.FileName);

                // Terminate the current process
                currentProcess.Kill();
                return;
            }
        }

        public async Task<int> LaunchWork()
        {
            CreateShortcut();

            try
            {
                // Background window creation
                bkgWindow = new WindowEx();
                bkgWindow.SystemBackdrop = new TransparentTintBackdrop();
                bkgWindow.IsMaximizable = false;
                bkgWindow.SetExtendedWindowStyle(ExtendedWindowStyle.ToolWindow);
                bkgWindow.SetWindowStyle(WindowStyle.Visible);
                bkgWindow.Activate();  // Activate the background window
                bkgWindow.MoveAndResize(0, 0, 0, 0);
                bkgWindow.Minimize();  // Minimize to make sure it's not in focus
                bkgWindow.SetWindowOpacity(0);  // Set opacity to 0
            }
            catch
            {
                // Handle errors here
            }

            m_window = new MainWindow();

            // Register any background tasks or hooks
            StartHook();

            // If started with the "STARTUP" argument, exit early
            if (string.Join(" ", Environment.GetCommandLineArgs().Skip(1)).Contains("STARTUP"))
            {
                return 0;
            }

            // Activate the main window
            try
            {
                m_window.Activate();
            }
            catch
            {

            }

            // Ensure m_window is brought to the front with more delay
            await Task.Delay(100);  // Increase delay slightly
            try
            {
                m_window.Activate();  // Reactivate the main window to ensure focus
            }
            catch
            {

            }
            m_window.Show();  // Show the window

            ((WindowEx)m_window).BringToFront();

            return 0;
        }

        public void CreateStartupTask()
        {
            string taskName = "ReboundRun"; // Set your desired task name
            string appPath = Path.GetFullPath($@"shell:AppsFolder\{Package.Current.Id.FamilyName}!App"); // Set the path to your application

            using (var ts = new TaskService())
            {
                // Check if the task already exists
                Microsoft.Win32.TaskScheduler.Task existingTask = ts.FindTask(taskName);

                if (existingTask != null)
                {
                    Console.WriteLine("Task already exists. No new task created.");
                    return; // Task already exists, so we do nothing
                }

                // Create a new task
                TaskDefinition td = ts.NewTask();
                td.RegistrationInfo.Description = "Starts Rebound Run as elevated.";
                td.Principal.UserId = Environment.UserDomainName + "\\" + Environment.UserName; // Set the user
                td.Principal.LogonType = TaskLogonType.InteractiveToken; // Allows running under the user's session

                // Run with highest privileges (elevated)
                td.Principal.RunLevel = TaskRunLevel.Highest; // Run with elevated privileges

                // Add a trigger to start at logon
                td.Triggers.Add(new LogonTrigger());

                // Define the action to run your application
                td.Actions.Add(new ExecAction(appPath, "STARTUP", null));

                // Register the task
                ts.RootFolder.RegisterTaskDefinition(taskName, td);
                Console.WriteLine("Scheduled task created successfully.");
            }
        }

        private void CreateShortcut()
        {
            //CreateStartupTask();

            string startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string oldShortcutPath = System.IO.Path.Combine(startupFolderPath, "ReboundRun.lnk");
            try
            {
                File.Delete(oldShortcutPath);
            }
            catch
            {

            }
            string shortcutPath = System.IO.Path.Combine(startupFolderPath, "ReboundRunStartup.lnk");
            string appPath = "C:\\Rebound11\\rrunSTARTUP.exe";
            /*try
            {
                File.Delete(shortcutPath);
            }
            catch
            {

            }
            return;*/

            if (!System.IO.File.Exists(shortcutPath))
            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.Description = "Rebound Run";
                shortcut.TargetPath = "C:\\Rebound11\\rrunSTARTUP.exe";
                shortcut.Save();
            }
        }

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int VK_R = 0x52;
        private const int VK_LWIN = 0x5B;
        private const int VK_RWIN = 0x5C;

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

        public static bool allowCloseOfRunBox { get; set; } = true;

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private static bool winKeyPressed = false;
        private static bool rKeyPressed = false;

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                // Check for keydown events
                if (wParam == (IntPtr)WM_KEYDOWN)
                {
                    // Check if Windows key is pressed
                    if (vkCode is VK_LWIN or VK_RWIN)
                    {
                        winKeyPressed = true;
                    }

                    // Check if 'R' key is pressed
                    if (vkCode is VK_R)
                    {
                        rKeyPressed = true;

                        // If both Win and R are pressed, show the window
                        if (winKeyPressed)
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
                                ((WindowEx)m_window).Activate();
                                ((WindowEx)m_window).BringToFront();
                            }
                            finally
                            {
                                //(m_window as MainWindow).CloseRunBoxMethod();
                            }

                            // Prevent default behavior of Win + R
                            return (IntPtr)1;
                        }
                    }
                }

                // Check for keyup events
                if (wParam == (IntPtr)WM_KEYUP)
                {
                    /*switch (vkCode)
                    {
                        case VK_LWIN | VK_RWIN:
                            {
                                winKeyPressed = false;

                                // Suppress the Windows Start menu if 'R' is still pressed
                                if (rKeyPressed)
                                {
                                    return (IntPtr)1; // Prevent Windows menu from appearing
                                }
                                return 0;
                            }
                        case VK_R:
                            {
                                rKeyPressed = false;

                                // Suppress the Windows Start menu if 'R' is still pressed
                                if (winKeyPressed)
                                {
                                    return (IntPtr)1; // Prevent Windows menu from appearing
                                }
                                return 0;
                            }
                        default:
                            {
                                return 0;
                            }
                    }*/

                    // Check if Windows key is released
                    if (vkCode is VK_LWIN or VK_RWIN)
                    {
                        winKeyPressed = false;

                        // Suppress the Windows Start menu if 'R' is still pressed
                        if (rKeyPressed == true)
                        {
                            ForceReleaseWin();
                            return (IntPtr)1; // Prevent Windows menu from appearing
                        }
                    }

                    // Check if 'R' key is released
                    if (vkCode is VK_R)
                    {
                        rKeyPressed = false;
                    }
                }
            }

            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        public static async void ForceReleaseWin()
        {
            await Task.Delay(10);

            var inj = InputInjector.TryCreate();
            var info = new InjectedInputKeyboardInfo();
            info.VirtualKey = (ushort)VirtualKey.LeftWindows;
            info.KeyOptions = InjectedInputKeyOptions.KeyUp;
            var infocol = new[] { info };

            inj.InjectKeyboardInput(infocol);
        }

        public const int INPUT_KEYBOARD = 1;
        public const int KEYEVENTF_KEYUP = 0x0002;

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public uint type;
            public MOUSEKEYBDHARDWAREINPUT mkhi;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct MOUSEKEYBDHARDWAREINPUT
        {
            [FieldOffset(0)]
            public MOUSEKEYBDHARDWAREINPUT_KBD ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEKEYBDHARDWAREINPUT_KBD
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        public static void ReleaseKey(ushort keyCode)
        {
            INPUT[] inputs = new INPUT[1];
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].mkhi.ki.wVk = keyCode;
            inputs[0].mkhi.ki.dwFlags = KEYEVENTF_KEYUP;
            inputs[0].mkhi.ki.time = 0;
            inputs[0].mkhi.ki.dwExtraInfo = IntPtr.Zero;

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
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
