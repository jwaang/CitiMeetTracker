using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Outlook;

namespace CitiMeetTracker
{
    public partial class UserControl1 : MetroFramework.Controls.MetroUserControl
    {
        MetroFramework.Controls.MetroTextBox[] txtBox;
        MetroFramework.Controls.MetroLabel[] lbl;
        MetroFramework.Controls.MetroButton[] btn;
        int space = 20;

        public UserControl1()
        {
            InitializeComponent();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            Form1 frm1 = new Form1(this);
            frm1.StartPosition = FormStartPosition.CenterScreen;
            frm1.ShowDialog(this);
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {
            Dictionary<string, string> keys = GetRegistrySubKeys();
            createFields(keys);
        }

        public void removeControls()
        {
            for (int i = this.Controls.Count - 1; i >= 0; i--)
            {
                Control c = this.Controls[i];
                if (c is MetroFramework.Controls.MetroTextBox || c is MetroFramework.Controls.MetroButton || c is MetroFramework.Controls.MetroLabel)
                {
                    if (c.Name == "metroButton1" || c.Name == "metroLabel1") continue;
                    else this.Controls.RemoveAt(i);
                }
            }

            Dictionary<string, string> keys = GetRegistrySubKeys();
            createFields(keys);
        }

        private void createFields(Dictionary<string, string> keys)
        {
            int n = keys.Count;
            var arrKeys = keys.Keys.ToArray();
            var arrValues = keys.Values.ToArray();

            txtBox = new MetroFramework.Controls.MetroTextBox[n];
            lbl = new MetroFramework.Controls.MetroLabel[n];
            btn = new MetroFramework.Controls.MetroButton[n];

            for(int i=0; i<n; ++i)
            {
                lbl[i] = new MetroFramework.Controls.MetroLabel();
                lbl[i].Name = "label" + arrKeys[i];
                lbl[i].Text = arrKeys[i];

                txtBox[i] = new MetroFramework.Controls.MetroTextBox();
                txtBox[i].Name = "txtBox" + arrValues[i];
                txtBox[i].Text = arrValues[i];

                btn[i] = new MetroFramework.Controls.MetroButton();
                btn[i].Name = arrKeys[i];
                btn[i].Text = "X";
            }

            for (int i = 0; i < n; i++)
            {
                lbl[i].Visible = true;
                lbl[i].Location = new Point(10, 60 + space);
                txtBox[i].Visible = true;
                txtBox[i].Location = new Point(100, 60 + space);
                txtBox[i].Size = new System.Drawing.Size(75,23);
                txtBox[i].TabIndex = n + 1;
                txtBox[i].ReadOnly = true;
                txtBox[i].Click += delegate { Clipboard.SetText(this.ActiveControl.Text); };
                btn[i].Visible = true;
                btn[i].Location = new Point(200, 60 + space);
                btn[i].Size = new System.Drawing.Size(25, 25);
                btn[i].TabIndex = n + 2;
                btn[i].Click += delegate { deleteKey(this.ActiveControl.Name); };
                this.Controls.Add(txtBox[i]);
                this.Controls.Add(lbl[i]);
                this.Controls.Add(btn[i]);
                space += 50;
            }

            space = 20;
        }

        private void deleteKey(string key)
        {
            RegistryKey rootKey = Registry.CurrentUser;
            const string REGISTRY_ROOT = @"Software\Citi\CitiMeetingTracker";
            using (RegistryKey rk = rootKey.OpenSubKey(REGISTRY_ROOT, true))
            {
                if (rk == null)
                {
                    return;
                }
                else
                {
                    this.setMessage(key + " has been deleted. Restart Outlook to see your changes.");
                    rk.DeleteValue(key);
                }
            }

            removeControls();
        }

        public void setMessage(string val)
        {
            this.metroLabel2.Text = val;
        }

        private Dictionary<string, string> GetRegistrySubKeys()
        {
            RegistryKey rootKey = Registry.CurrentUser;
            var valuesBynames = new Dictionary<string, string>();
            const string REGISTRY_ROOT = @"Software\Citi\CitiMeetingTracker";
            using (RegistryKey rk = rootKey.OpenSubKey(REGISTRY_ROOT, false))
            {
                if (rk != null)
                {
                    string[] valueNames = rk.GetValueNames();
                    foreach (string currSubKey in valueNames)
                    {
                        string value = (string) rk.GetValue(currSubKey);
                        valuesBynames.Add(currSubKey, value);
                    }
                    rk.Close();
                }

            }
            return valuesBynames;
        }
    }
}
