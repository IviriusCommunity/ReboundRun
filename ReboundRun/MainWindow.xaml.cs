using Microsoft.Graphics.Display;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Appointments.AppointmentsProvider;
using Windows.ApplicationModel.Store;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage.Pickers;
using Windows.System;
using WinUIEx;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ReboundRun
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : WindowEx
    {
        public double Scale()
        {
            // Get the DisplayInformation object for the current view
            DisplayInformation displayInformation = DisplayInformation.CreateForWindowId(this.AppWindow.Id);
            // Get the RawPixelsPerViewPixel which gives the scale factor
            var scaleFactor = displayInformation.RawPixelsPerViewPixel;
            return scaleFactor;
        }

        public MainWindow()
        {
            this.InitializeComponent();
            this.MoveAndResize(25 * Scale(), (WindowsDisplayAPI.Display.GetDisplays().ToList<WindowsDisplayAPI.Display>()[0].CurrentSetting.Resolution.Height - 370 / Scale()) / Scale(), 525, 295);
            this.IsMinimizable = false;
            this.IsMaximizable = false;
            this.IsResizable = false;
            this.AppWindow.DefaultTitleBarShouldMatchAppModeTheme = true;
            this.SetIcon($"{AppContext.BaseDirectory}/Assets/RunBox.ico");
            this.Title = "Run";
            this.SystemBackdrop = new MicaBackdrop();
            Load();
            LoadRunHistory();
        }

        public async void Load()
        {
            await Task.Delay(100);
            RunBox.Focus(FocusState.Keyboard);
        }

        private void LoadRunHistory(bool clear = false)
        {
            string runMRUPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU";
            RegistryKey runMRUKey = Registry.CurrentUser.OpenSubKey(runMRUPath);

            if (runMRUKey != null)
            {
                // Read the MRUList to determine the order of the entries
                string mruList = runMRUKey.GetValue("MRUList")?.ToString();
                if (mruList != null)
                {
                    List<string> runHistory = new List<string>();

                    // Iterate over each character in the MRUList to get the entries in order
                    foreach (char entry in mruList)
                    {
                        string entryValue = runMRUKey.GetValue(entry.ToString())?.ToString();
                        if (!string.IsNullOrEmpty(entryValue))
                        {
                            // Remove the '/1' suffix if it exists
                            if (entryValue.EndsWith("\\1"))
                            {
                                entryValue = entryValue.Substring(0, entryValue.Length - 2);
                            }
                            if (clear == false) runHistory.Add(entryValue);
                            else
                            {
                                runHistory.Remove(entryValue);
                                runHistory.Clear();
                            }
                        }
                    }

                    // Display the ordered entries in the ListBox
                    RunBox.ItemsSource = runHistory;
                    RunBox.SelectedIndex = 0;
                }

                runMRUKey.Close();
            }
            /*try
            {
                // Access the registry key for the Run history
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"))
                {
                    List<string> runHistory = new();
                    if (key != null)
                    {
                        foreach (string valueName in key.GetValueNames())
                        {
                            object value = key.GetValue(valueName);
                            if (value is not null or "MRUList")
                            {
                                runHistory.Add(value.ToString());
                            }
                        }
                    }
                    RunBox.ItemsSource = runHistory;
                    RunBox.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Run history: {ex.Message}");
            }*/
        }

        public async Task Run(bool runLegacy = false, bool admin = false)
        {
            string newEntry = RunBox.Text;
            if (string.IsNullOrEmpty(newEntry)) return;

            // URI
            else if (newEntry.Contains("://"))
            {
                await Launcher.LaunchUriAsync(new Uri(newEntry));

                Close();
            }

            // Settings URI
            else if (newEntry.ToLower() == "settings")
            {
                if (!runLegacy) await Launcher.LaunchUriAsync(new Uri("ms-settings:///"));
                else
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    if (ArgsBox.Text.ToString() != string.Empty) startInfo.Arguments = $"Start-Process -FilePath \"control\" -ArgumentList \"{ArgsBox.Text}\"";
                    else startInfo.Arguments = $"Start-Process -FilePath \"control\"";

                    if (admin)
                    {
                        startInfo.Arguments += " -Verb RunAs";
                        startInfo.Verb = "runas";
                    }

                    try
                    {
                        var res = Process.Start(startInfo);
                        await res.WaitForExitAsync();
                        if (res.ExitCode == 0) Close();
                    }
                    catch (Exception ex)
                    {
                        await this.ShowMessageDialogAsync($"The system cannot find the file specified.");
                    }
                }
            }

            // dfrgui.exe
            else if (newEntry.Contains("dfrgui") && runLegacy != true)
            {
                if (File.Exists("C:\\Rebound11\\rdfrgui.exe"))
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    if (ArgsBox.Text.ToString() != string.Empty) startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\rdfrgui.exe\" -ArgumentList \"{ArgsBox.Text}\"";
                    else startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\rdfrgui.exe\"";

                    if (admin)
                    {
                        startInfo.Arguments += " -Verb RunAs";
                        startInfo.Verb = "runas";
                    }

                    try
                    {
                        var res = Process.Start(startInfo);
                        await res.WaitForExitAsync();
                        if (res.ExitCode == 0) Close();
                        else await Run(true);
                    }
                    catch (Exception ex)
                    {
                        await Run(true);
                    }
                }
                else await Run(true, admin);
            }

            // control.exe
            else if (newEntry.Contains("control") && runLegacy != true)
            {
                if (File.Exists("C:\\Rebound11\\rcontrol.exe"))
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    if (ArgsBox.Text.ToString() != string.Empty) startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\rcontrol.exe\" -ArgumentList \"{ArgsBox.Text}\"";
                    else startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\rcontrol.exe\"";

                    if (admin)
                    {
                        startInfo.Arguments += " -Verb RunAs";
                        startInfo.Verb = "runas";
                    }

                    try
                    {
                        var res = Process.Start(startInfo);
                        await res.WaitForExitAsync();
                        if (res.ExitCode == 0) Close();
                        else await Run(true);
                    }
                    catch (Exception ex)
                    {
                        await Run(true);
                    }
                }
                else await Run(true, admin);
            }

            // control.exe
            else if (newEntry.Contains("UserAccountControlSettings.exe") && runLegacy != true)
            {
                if (File.Exists("C:\\Rebound11\\ruacsettings.exe"))
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    if (ArgsBox.Text.ToString() != string.Empty) startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\ruacsettings.exe\" -ArgumentList \"{ArgsBox.Text}\"";
                    else startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\ruacsettings.exe\"";

                    if (admin)
                    {
                        startInfo.Arguments += " -Verb RunAs";
                        startInfo.Verb = "runas";
                    }

                    try
                    {
                        var res = Process.Start(startInfo);
                        await res.WaitForExitAsync();
                        if (res.ExitCode == 0) Close();
                        else await Run(true);
                    }
                    catch (Exception ex)
                    {
                        await Run(true);
                    }
                }
                else await Run(true, admin);
            }

            // control.exe
            else if (newEntry.Contains("winver") && runLegacy != true)
            {
                if (File.Exists("C:\\Rebound11\\rwinver.exe"))
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    if (ArgsBox.Text.ToString() != string.Empty) startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\rwinver.exe\" -ArgumentList \"{ArgsBox.Text}\"";
                    else startInfo.Arguments = $"Start-Process -FilePath \"C:\\Rebound11\\rwinver.exe\"";

                    if (admin)
                    {
                        startInfo.Arguments += " -Verb RunAs";
                        startInfo.Verb = "runas";
                    }

                    try
                    {
                        var res = Process.Start(startInfo);
                        await res.WaitForExitAsync();
                        if (res.ExitCode == 0) Close();
                        else await Run(true);
                    }
                    catch (Exception ex)
                    {
                        await Run(true);
                    }
                }
                else await Run(true, admin);
            }

            // Task Manager legacy
            else if (newEntry.Contains("taskmgr") && runLegacy == true)
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Arguments = "taskmgr -d"
                };

                if (admin == true) startInfo.Verb = "runas";

                try
                {
                    var res = Process.Start(startInfo);
                    await res.WaitForExitAsync();
                    if (res.ExitCode == 0) Close();
                }
                catch (Exception ex)
                {
                    await this.ShowMessageDialogAsync($"The system cannot find the file specified.");
                }
            }

            // Task Manager legacy
            else if (newEntry.Contains("run") && runLegacy == true)
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Arguments = "(New-Object -ComObject \"Shell.Application\").FileRun()"
                };

                if (admin == true) startInfo.Verb = "runas";

                try
                {
                    var res = Process.Start(startInfo);
                    await res.WaitForExitAsync();
                    if (res.ExitCode == 0) Close();
                }
                catch (Exception ex)
                {
                    await this.ShowMessageDialogAsync($"The system cannot find the file specified.");
                }
            }

            // Task Manager legacy
            else if (newEntry.Contains("run") && runLegacy != true)
            {
                await this.ShowMessageDialogAsync($"The WinUI 3 run box is already opened.", "Error");
                return;
            }

            // Process or executable command
            else
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                // Handle arguments with spaces by wrapping them in double quotes
                string runBoxText = RunBox.Text.ToString();
                string argsBoxText = ArgsBox.Text.ToString();

                if (!string.IsNullOrWhiteSpace(argsBoxText))
                {
                    startInfo.Arguments = $"Start-Process -FilePath '{runBoxText}' -ArgumentList '{argsBoxText}'";
                }
                else
                {
                    startInfo.Arguments = $"Start-Process -FilePath '{runBoxText}' ";
                }

                // Handle running as administrator
                if (admin)
                {
                    startInfo.Arguments += " -Verb RunAs";
                    startInfo.Verb = "runas";
                }

                try
                {
                    var process = Process.Start(startInfo);
                    await process.WaitForExitAsync();

                    if (process.ExitCode == 0)
                    {
                        Close();
                    }
                    else
                    {
                        throw new Exception($"Process exited with code {process.ExitCode}");
                    }
                }
                catch (Exception ex)
                {
                    await this.ShowMessageDialogAsync("The system cannot find the file specified or the command line arguments are invalid.", "Error");
                }
            }

        }

        private async void SplitButton_Click(SplitButton sender, SplitButtonClickEventArgs args)
        {
            string newEntry = RunBox.Text;

            await Run();

            string runMRUPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU";
            RegistryKey runMRUKey = Registry.CurrentUser.OpenSubKey(runMRUPath, true);

            if (runMRUKey != null)
            {
                // Read the current MRUList
                string mruList = runMRUKey.GetValue("MRUList")?.ToString();
                if (mruList != null)
                {
                    // Check for duplicates and remove the existing entry if found
                    char? existingEntryKey = null;
                    foreach (char entry in mruList)
                    {
                        string entryValue = runMRUKey.GetValue(entry.ToString())?.ToString();
                        if (entryValue != null && entryValue.StartsWith(newEntry))
                        {
                            existingEntryKey = entry;
                            break;
                        }
                    }

                    if (existingEntryKey.HasValue)
                    {
                        // Remove the existing entry
                        mruList = mruList.Replace(existingEntryKey.Value.ToString(), string.Empty);
                    }

                    // Determine the new entry key
                    char newEntryKey = 'a';
                    if (mruList.Length > 0)
                    {
                        newEntryKey = (char)(mruList[0] + 1);
                    }

                    // Add the new entry to the registry
                    runMRUKey.SetValue(newEntryKey.ToString(), newEntry);

                    // Update the MRUList
                    mruList = newEntryKey + mruList;
                    runMRUKey.SetValue("MRUList", mruList);

                    runMRUKey.Close();

                    // Reload the Run history to refresh the ListBox
                    LoadRunHistory();
                }
            }

            // Clear the input box
            //RunBox.Text = string.Empty;
        }

        private async void MenuFlyoutItem_Click(object sender, RoutedEventArgs e)
        {
            await Run(true);
        }

        private async void MenuFlyoutItem_Click_1(object sender, RoutedEventArgs e)
        {
            await Run();
        }

        private async void MenuFlyoutItem_Click_2(object sender, RoutedEventArgs e)
        {
            await Run(false, true);
        }

        private async void MenuFlyoutItem_Click_3(object sender, RoutedEventArgs e)
        {
            await Run(true, true);
        }

        private void MenuFlyoutItem_Click_4(object sender, RoutedEventArgs e)
        {
            ClearRunHistory();
        }

        private async void ClearRunHistory()
        {
            const string runMRUPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU";
            RegistryKey runMRUKey = Registry.CurrentUser.OpenSubKey(runMRUPath);
            try
            {
                /*string psCommand = @"Set-ItemProperty -Path ""HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"" -Name ""MRUList"" -Value """"";
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                //startInfo.Arguments += " -Verb RunAs";
                startInfo.Verb = "runas";

                try
                {
                    Process.Start(startInfo);
                }
                catch (Exception ex)
                {
                    await this.ShowMessageDialogAsync($"The system cannot find the file specified.");
                }
                RunBox.ItemsSource = null;*/
                LoadRunHistory(true);
            }
            catch (SecurityException ex)
            {
                throw new InvalidOperationException("Access denied to the registry key.", ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new InvalidOperationException("Unauthorized access to the registry key.", ex);
            }
        }

        private async void RunBox_KeyUp(object sender, KeyRoutedEventArgs e)
        {
            if (e.Key == VirtualKey.Enter)
            {
                await Run();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            // Create a file picker
            var openPicker = new Windows.Storage.Pickers.FileOpenPicker();

            // See the sample code below for how to make the window accessible from the App class.
            var window = this;

            // Retrieve the window handle (HWND) of the current WinUI 3 window.
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(window);

            // Initialize the file picker with the window handle (HWND).
            WinRT.Interop.InitializeWithWindow.Initialize(openPicker, hWnd);

            // Set options for your file picker
            openPicker.ViewMode = PickerViewMode.Thumbnail;
            openPicker.SuggestedStartLocation = PickerLocationId.Desktop;
            openPicker.CommitButtonText = "Select file to run";
            openPicker.FileTypeFilter.Add(".exe");
            openPicker.FileTypeFilter.Add(".pif");
            openPicker.FileTypeFilter.Add(".com");
            openPicker.FileTypeFilter.Add(".bat");
            openPicker.FileTypeFilter.Add(".cmd");
            openPicker.FileTypeFilter.Add("*");

            // Open the picker for the user to pick a file
            var file = await openPicker.PickSingleFileAsync();
            if (file != null)
            {
                RunBox.Text = file.Path;
                await Run();
            }
            else
            {

            }

        }
    }

    public static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct STARTUPINFO
        {
            public uint cb;
            public string lpReserved;
            public string lpDesktop;
            public string lpTitle;
            public uint dwX;
            public uint dwY;
            public uint dwXSize;
            public uint dwYSize;
            public uint dwXCountChars;
            public uint dwYCountChars;
            public uint dwFillAttribute;
            public uint dwFlags;
            public short wShowWindow;
            public short cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public uint dwProcessId;
            public uint dwThreadId;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool CreateProcess(
            string lpApplicationName,
            string lpCommandLine,
            IntPtr lpProcessAttributes,
            IntPtr lpThreadAttributes,
            bool bInheritHandles,
            uint dwCreationFlags,
            IntPtr lpEnvironment,
            string lpCurrentDirectory,
            ref STARTUPINFO lpStartupInfo,
            out PROCESS_INFORMATION lpProcessInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint GetLastError();
    }
}
