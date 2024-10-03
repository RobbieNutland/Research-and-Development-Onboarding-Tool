using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;
using System.IO;

namespace R_D_Onboarding_Tool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string? currentDirectory = Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName);
        string? currentProgram = Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName);
        public MainWindow()
        {
            string[] onboardingFonts = new string[8] { "GT-Walsheim-Pro-Bold.otf", "GT-Walsheim-Pro-Bold-Oblique.otf", "GT-Walsheim-Pro-Light.otf", "GT-Walsheim-Pro-Light-Oblique.otf", "GT-Walsheim-Pro-Medium.otf", "GT-Walsheim-Pro-Medium-Oblique.otf", "GT-Walsheim-Pro-Regular.otf", "GT-Walsheim-Pro-Regular-Oblique.otf" };

            for (int i = 0; i < 8; i++)
            {
                if (!File.Exists(Environment.GetEnvironmentVariable("WINDIR") + "\\fonts\\" + onboardingFonts[i]) && !File.Exists(Environment.GetEnvironmentVariable("LOCALAPPDATA") + "\\Microsoft\\Windows\\Fonts\\" + onboardingFonts[i]))
                {
                    Directory.CreateDirectory(Environment.GetEnvironmentVariable("LOCALAPPDATA") + "\\Microsoft\\Windows\\Fonts");
                    
                    File.Copy(currentDirectory + "\\Resources\\Shared Assets\\Fonts\\GT Walsheim Pro\\" + onboardingFonts[i], Environment.GetEnvironmentVariable("LOCALAPPDATA") + "\\Microsoft\\Windows\\Fonts\\" + onboardingFonts[i]);
                    
                    RegistryKey? fontKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts", true);
                    fontKey?.SetValue(onboardingFonts[i].Replace("-", " ").Replace(".otf", " (TrueType)"), Environment.GetEnvironmentVariable("LOCALAPPDATA") + "\\Microsoft\\Windows\\Fonts\\" + onboardingFonts[i]);
                    fontKey?.Close();
                }
            }

            RegistryKey? browserKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);                  // sets WebBrowser controls to emulate Internet Explorer 11
            if (browserKey?.GetValue(currentProgram) == null)
            {
                browserKey?.SetValue(currentProgram, "11000", RegistryValueKind.DWord);
                browserKey?.Close();
            }

            InitializeComponent();
        }
        private void MainBrowser_Initialized(object sender, EventArgs e)
        {
            if (currentDirectory != null) {
                Uri startPage = new Uri(new Uri(currentDirectory).AbsoluteUri + "/Resources/The%20R&D%20Onboarding%20Tool/The%20R&D%20Onboarding%20Tool.html");
                //MessageBox.Show(startPage.ToString());
                WebBrowser senderBrowser = (WebBrowser)sender;
                senderBrowser.Navigate(startPage);
            }
        }
    }
}
