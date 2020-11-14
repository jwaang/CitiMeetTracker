using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CitiMeetTracker
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        UserControl1 _uc;
        public const string defaultValue = "YEETYOULLNEVERGUESSME";
        public Form1(UserControl1 uc)
        {
            InitializeComponent();
            _uc = uc;
            SetAutoScale(this);
        }

        public static void SetAutoScale(Form f)
        {
            f.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            f.Font = new System.Drawing.Font("Arial", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel, ((byte)(0)));
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string name = nameTextBox.Text;
            string passcode = passcodeTextBox.Text;

            if (name == "" || passcode == "")
            {
                metroLabel3.Text = "Both fields are required.";
                return;
            }
            if (ReadFromRegistry(name, defaultValue) != defaultValue)
            {
                metroLabel3.Text = "This name already exists. Choose another.";
            }
            else
            {
                StoreInRegistry(name, passcode);
                _uc.setMessage(name + " has been added. Restart Outlook to see your changes.");
                _uc.removeControls();
                this.Close();
            }
        }

        public void StoreInRegistry(string name, string passcode)
        {
            RegistryKey rootKey = Registry.CurrentUser;
            string registryPath = @"Software\Citi\CitiMeetingTracker";
            using (RegistryKey rk = rootKey.CreateSubKey(registryPath))
            {
                rk.SetValue(name, passcode, RegistryValueKind.String);
            }
        }

        public string ReadFromRegistry(string keyName, string defaultValue)
        {
            RegistryKey rootKey = Registry.CurrentUser;
            string registryPath = @"Software\Citi\CitiMeetingTracker";
            using (RegistryKey rk = rootKey.OpenSubKey(registryPath, false))
            {
                if (rk == null)
                {
                    return defaultValue;
                }

                var res = rk.GetValue(keyName, defaultValue);
                if (res == null)
                {
                    return defaultValue;
                }

                return res.ToString();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
