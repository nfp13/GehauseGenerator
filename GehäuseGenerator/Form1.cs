using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace GehäuseGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeUI("UIMode");
        }

        private void InitializeUI(string key)
        {
            try
            {
                var uiMode = ConfigurationManager.AppSettings[key];
                if (uiMode == "light")
                {
                    btnchangemode.Text = "Dark Mode";
                    lblueberschrift1.Text = "Light Mode";
                    this.ForeColor = Color.FromArgb(47, 54, 64);
                    this.BackColor = Color.FromArgb(245, 246, 250);
                    ConfigurationManager.AppSettings[key] = "dark";
                }
                else
                {
                    btnchangemode.Text = "Light Mode";
                    lblueberschrift1.Text = "Dark Mode";
                    this.ForeColor = Color.FromArgb(245, 246, 250);
                    this.BackColor = Color.FromArgb(47, 54, 64);
                    ConfigurationManager.AppSettings[key] = "light";
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void btnchangemode_Click(object sender, EventArgs e)
        {
            InitializeUI("UIMode");
        }
    }
            
}
    

