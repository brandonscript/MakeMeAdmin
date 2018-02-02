using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Principal;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace mma2
{
    public partial class MainForm : Form
    {
        static string username = Regex.Replace(WindowsIdentity.GetCurrent().Name, @".*\\", "");
        static string computerName = System.Environment.MachineName;
        static string domainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
        static string programFiles = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles);
        static string appPath = programFiles + "\\mma2\\";
        static string appPathiCacls = programFiles + "\\mma2";
        static string adderFile = "Adder.exe";
        static string olxFile = "objectList.olx";
        static string xmlFile = "mma.xml";
        static string srcPath = System.IO.Directory.GetCurrentDirectory();

        public MainForm()
        {
            InitializeComponent();
        }

        private void mainFunction()
        {
            createFolder(appPath);
            createTextFile(taskxml, appPath + xmlFile);
            createTextFile(namelist(), appPath + olxFile);
            copyFile(srcPath + "\\" + adderFile, appPath + adderFile, true);
            icacls(appPathiCacls);
            createTask();
            finish();

        }

        private void finish()
        {
            DialogResult r = new DialogResult();
            r = MessageBox.Show("The specified object(s) should now have been added to the local Administrators group on " + computerName + ". Do you want to open Computer Management to check?",
                "Open Computer Management?", 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);
            if (r == System.Windows.Forms.DialogResult.Yes)
            {
                ProcessStartInfo exe = new ProcessStartInfo();
                exe.FileName = "compmgmt.msc";
                Process p = Process.Start(exe);
            }

            try
            {
                File.Delete(appPath + xmlFile);
            }
            catch (Exception) { };
        }

        private string namelist()
        {
            return textBox1.Text.Replace(";", "\r\n");
        }

        private void icacls(string o)
        {
            try
            {
                ProcessStartInfo exe = new ProcessStartInfo();
                exe.FileName = "icacls.exe";
                exe.Arguments = "\"" + o + "\" /grant Users:(OI)(CI)M /T";
                exe.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                Process p = Process.Start(exe);
                p.WaitForExit();
            }
            catch (Exception)
            { }
            
        }

        private static void createTask()
        {
            try
            {
                ProcessStartInfo exe = new ProcessStartInfo();
                exe.FileName = "schtasks.exe";
                exe.Arguments = "/create /TN mma2 /XML \"" + appPath + xmlFile + "\"";
                exe.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                Process p = Process.Start(exe);
                p.WaitForExit();

                exe.Arguments = "/run /TN mma2";
                Process q = Process.Start(exe);
                q.WaitForExit();
            }
            catch (Exception)
            { }
        }

        private static void createFolder(string f)
        {
            try
            {
                Directory.CreateDirectory(f);
            }
            catch (Exception)
            { }
        }

        static void copyFile(string src, string dest, bool tf)
        {
            try
            {
                File.Copy(src, dest, tf);
            }
            catch (Exception)
            { }
        }


        static void createTextFile(string content, string path)
        {
            try
            {

                StreamWriter sw;
                sw = File.CreateText(path);
                sw.WriteLine(content);
                sw.Close();
            }
            catch (Exception)
            { };
        }

        static string taskxml =
            "<?xml version=\"1.0\" encoding=\"UTF-16\"?>\r\n" +
            "<Task version=\"1.2\" xmlns=\"http://schemas.microsoft.com/windows/2004/02/mit/task\">\r\n" +
                "\t<RegistrationInfo>\r\n" +
                    "\t\t<Date>2014-05-07T13:35:01.5754822</Date>\r\n" +
                    "\t\t<Author>{YOUR-DOMAIN}\\{AUTHOR-ACCOUNT}</Author>\r\n" +
                "\t</RegistrationInfo>\r\n" +
                "\t<Triggers>\r\n" +
                    "\t\t<BootTrigger>\r\n" +
                        "\t\t\t<Repetition>\r\n" +
                            "\t\t\t\t<Interval>PT1M</Interval>\r\n" +
                            "\t\t\t\t<StopAtDurationEnd>false</StopAtDurationEnd>\r\n" +
                        "\t\t\t</Repetition>\r\n" +
                        "\t\t\t<ExecutionTimeLimit>PT5M</ExecutionTimeLimit>\r\n" +
                        "\t\t\t<Enabled>true</Enabled>\r\n" +
                    "\t\t</BootTrigger>\r\n" +
                "\t</Triggers>\r\n" +
                "\t<Principals>\r\n" +
                    "\t\t<Principal id=\"Author\">\r\n" +
                    "\t\t<UserId>SYSTEM</UserId>\r\n" +
                    "\t\t<RunLevel>HighestAvailable</RunLevel>\r\n" +
                "\t</Principal>\r\n" +
                "\t</Principals>\r\n" +
                "\t<Settings>\r\n" +
                    "\t\t<IdleSettings>\r\n" +
                        "\t\t\t<Duration>PT10M</Duration>\r\n" +
                        "\t\t\t<WaitTimeout>PT1H</WaitTimeout>\r\n" +
                        "\t\t\t<StopOnIdleEnd>true</StopOnIdleEnd>\r\n" +
                        "\t\t\t<RestartOnIdle>false</RestartOnIdle>\r\n" +
                    "\t\t</IdleSettings>\r\n" +
                    "\t\t<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>\r\n" +
                    "\t\t<DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>\r\n" +
                    "\t\t<StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>\r\n" +
                    "\t\t<AllowHardTerminate>true</AllowHardTerminate>\r\n" +
                    "\t\t<StartWhenAvailable>false</StartWhenAvailable>\r\n" +
                    "\t\t<RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>\r\n" +
                    "\t\t<AllowStartOnDemand>true</AllowStartOnDemand>\r\n" +
                    "\t\t<Enabled>true</Enabled>\r\n" +
                    "\t\t<Hidden>false</Hidden>\r\n" +
                    "\t\t<RunOnlyIfIdle>false</RunOnlyIfIdle>\r\n" +
                    "\t\t<WakeToRun>false</WakeToRun>\r\n" +
                    "\t\t<ExecutionTimeLimit>P3D</ExecutionTimeLimit>\r\n" +
                    "\t\t<Priority>7</Priority>\r\n" +
                "\t</Settings>\r\n" +
                "\t<Actions Context=\"Author\">\r\n" +
                    "\t\t<Exec>\r\n" +
                    "\t\t\t<Command>\"" + appPath + adderFile + "\"</Command>\r\n" +
                    "\t\t</Exec>\r\n" +
                "\t</Actions>\r\n" +
            "</Task>";

        private void MainForm_Load(object sender, EventArgs e)
        {
            textBox1.Text = username;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            Cursor = Cursors.WaitCursor;
            Timer t = new Timer();
            t.Interval = 4000;
            mainFunction();
            t.Enabled = true;
            t.Stop();
            button1.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button1.PerformClick();
        }

    }
}
