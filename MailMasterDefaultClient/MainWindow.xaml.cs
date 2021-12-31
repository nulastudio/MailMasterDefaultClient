using Microsoft.Win32;

using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

using AppResources = MailMasterDefaultClient.Properties.Resources;

namespace MailMasterDefaultClient
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();

            this.MinWidth = this.MaxWidth = this.Width;
            this.MinHeight = this.MaxHeight = this.Height;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("是否将网易邮箱大师设置为默认客户端？", "设置默认客户端", MessageBoxButton.YesNo, MessageBoxImage.None, MessageBoxResult.No);

            if (result != MessageBoxResult.Yes) return;

            //为null表示没有安装
            //为“”表示安装不正确
            string progid = null;

            try
            {
                var mailKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Clients\\Mail");

                if (mailKey != null)
                {
                    var keyNames = mailKey.GetSubKeyNames();

                    var masterKeyName = "";

                    foreach (var name in keyNames)
                    {
                        if (!name.Contains("网易邮箱大师")) return;

                        masterKeyName = name;
                        break;
                    }

                    if (masterKeyName != "")
                    {
                        progid = "";

                        var URLAssociations = mailKey.OpenSubKey(masterKeyName + "\\Capabilities\\URLAssociations");

                        var mailto = URLAssociations.GetValue("mailto");

                        if (mailto != null)
                        {
                            progid = mailto.ToString();
                        }
                    }
                }

                if (progid == null)
                {
                    MessageBox.Show("没有安装网易邮箱大师！");
                }
                else if (progid == "")
                {
                    MessageBox.Show("网易邮箱大师安装不正确，请尝试重装！");
                }
                else
                {
                    //释放SetUserFTA

                    var AppDataDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

                    var AppDir = AppDataDir + "\\MailMasterDefaultClient";

                    var SetUserFTAPath = AppDir + "\\SetUserFTA.exe";

                    if (!Directory.Exists(AppDir))
                    {
                        Directory.CreateDirectory(AppDir);
                    }

                    File.WriteAllBytes(SetUserFTAPath, AppResources.SetUserFTA);

                    //运行SetUserFTA

                    var psi = new ProcessStartInfo(SetUserFTAPath)
                    {
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = false,
                        RedirectStandardError = false,
                    };

                    psi.Arguments = "mailto " + progid;

                    var process = Process.Start(psi);

                    process.WaitForExit();
                    
                    var needTest = MessageBox.Show("设置默认客户端成功！是否需要测试一下？\n\n注意事项：\n如果弹出需要选择默认程序的对话框，选一下Win10自带的邮件客户端（或者其他任意的邮件客户端也行），然后再点一次按钮再次设置网易邮箱大师为默认客户端，否则不会成功！！！", "设置默认客户端", MessageBoxButton.YesNo, MessageBoxImage.None, MessageBoxResult.No);

                    if (needTest == MessageBoxResult.Yes)
                    {
                        Process.Start(new ProcessStartInfo("explorer.exe")
                        {
                            Arguments = "mailto:liesauer@hotmail.com",
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = false,
                            RedirectStandardError = false,
                        });
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("设置默认客户端失败！");
            }
        }
    }
}
