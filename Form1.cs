using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendCmdHyperV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Send();
        }
        public void Send()
        {
            string vmName = textBox1.Text;
            string command = textBox2.Text;
            string path = System.IO.Directory.GetCurrentDirectory();
            string ps1 = path + "\\Sendcmd.ps1";
            string ps2 = path + "\\cmd.ps1";
            string cmd1 = @"$VmMgmt = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_VirtualSystemManagementService" + "\n";
            string cmd2 = @"$vm = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem -Filter { ElementName = '" + vmName + "'}\n";
            string cmd3 = @"$kvpDataItem = ([WMIClass][String]::Format(""\\{0}\{1}:{2}"", $VmMgmt.ClassPath.Server, $VmMgmt.ClassPath.NamespacePath, ""Msvm_KvpExchangeDataItem"")).CreateInstance()" + "\n";
            string cmd4 = "$kvpDataItem.Name = \"SendCmd\"\n";
            string cmd5 = "$kvpDataItem.Data = '" + command + "'\n";
            string cmd6 = "$kvpDataItem.Source = 0\n";
            string cmd7 = "$VmMgmt.ModifyKvpItems($Vm, $kvpDataItem.PSBase.GetText(1))\n";
            string get = "Get-VMKvpGuestToHost -Key ReturnCMD -VMName Windows2016 | write-host";
            string getimp = "Import-module Hyper-V-KVP-Host\n";
            try
            {
                File.Delete(@"Sendcmd.ps1");
            }
            catch { MessageBox.Show("Error deleting file"); }

            File.AppendAllText(@"Sendcmd.ps1", cmd1);
            File.AppendAllText(@"Sendcmd.ps1", cmd2);
            File.AppendAllText(@"Sendcmd.ps1", cmd3);
            File.AppendAllText(@"Sendcmd.ps1", cmd4);
            File.AppendAllText(@"Sendcmd.ps1", cmd5);
            File.AppendAllText(@"Sendcmd.ps1", cmd6);
            File.AppendAllText(@"Sendcmd.ps1", cmd7);
            File.AppendAllText(@"returncmd.ps1", getimp);
            File.AppendAllText(@"returncmd.ps1", get);
            //var shell = PowerShell.Create();
            //shell.Commands.AddScript(ps1);


            System.Diagnostics.Process.Start(@"C:\Stuff\SendCmdHyperV\SendCmdHyperV\bin\Debug\Debug\returncmd.ps1");
            Process scriptProc = new Process();
            scriptProc.StartInfo.FileName = @"cmd";
            scriptProc.StartInfo.WorkingDirectory = @"C:\Stuff\cmd.vbs"; //<---very important 
            scriptProc.StartInfo.Arguments = "//B //Nologo vbscript.vbs";
            scriptProc.Start();
            scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exit
            scriptProc.Close();
            //var results = shell.Invoke();
            //using (Runspace myRunSpace = RunspaceFactory.CreateRunspace())
            //{
            //    myRunSpace.Open();
            //    using (PowerShell powershell = PowerShell.Create())
            //    {
            //        powershell.Runspace = myRunSpace;
            //        powershell.AddScript("$PSVersionTable.PSVersion; "+ "write-host \"yoyoy\"");
            //        //powershell.AddScript("Get-VMKvpGuestToHost -Key ReturnCMD -VMName Windows2016");

            //        //var results = powershell.Invoke();
            //        //richTextBox1.Text = results.ToString();
            //        ////powershell.Invoke();
            //        ICollection<PSObject> results = powershell.Invoke();
            //        foreach (PSObject result in results)
            //        {

            //            richTextBox1.Text = richTextBox1.Text + (result.BaseObject.ToString())+"\n";
            //        }
            //    }
            //    myRunSpace.Close(); }

            //var shell2 = PowerShell.Create();
            //Pipeline pipeline = shell2.Commands.AddScript(ps2);
            ////var results2 = shell2.Invoke();
            //Collection<PSObject> result3 = shell2.Invoke();
            //MessageBox.Show(result3.ToString());

        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }
    }

}
