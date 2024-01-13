using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.SqlServer.Server;
using System.Xml;
using System.Linq;
using System.IO;
//using Microsoft.Office.Interop.Excel;
//using Inventor;


namespace GehäuseGenerator
{
    public partial class Form1 : Form
    {
        string FilePath = null;

        Platine platine;
        Status status;
        Gehäuse gehäuseOben, gehäuseUnten;
        Normteile normteile;
        BaugruppeZusammenfuegen baugruppeZusammenfuegen;
        Speichern speichern;

        Inventor.Application inventorApp;


        List<Panel> listPanel = new List<Panel>();
        int index;
        public Form1()
        {
            status = new Status();
            status.Progressed += new EventHandler(UpdateStatus);

            InitializeComponent();
            InitializeUI("UIMode");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                dynamic result = MessageBox.Show("Soll das Program beendet werden?", "Test Prog", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    inventorApp.Quit();
                    normteile.CloseExcel();
                    System.Windows.Forms.Application.Exit();
                    
                    //löschen
                    speichern = new Speichern();
                    speichern.deleteFiles();
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void InitializeUI(string key)
        {
            try
            {
                var uiMode = ConfigurationManager.AppSettings[key];
                if (uiMode == "light")
                {
                    btnchangemode.Text = "Dark Mode";
                    this.ForeColor = Color.FromArgb(47, 54, 64);
                    this.BackColor = Color.FromArgb(245, 246, 250);
                    pnlleiste.BackColor = Color.FromArgb(231, 231, 231);
                    ConfigurationManager.AppSettings[key] = "dark";
                }
                else
                {
                    btnchangemode.Text = "Light Mode";
                    this.ForeColor = Color.FromArgb(245, 246, 250);
                    this.BackColor = Color.FromArgb(29, 29, 29);
                    pnlleiste.BackColor = Color.FromArgb(49, 49, 49);
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

        private void btnweiter_Click(object sender, EventArgs e)
        {
            if(index<listPanel.Count-1)
                listPanel[++index].BringToFront();

            if(true)
            {
                speichern = new Speichern();
                gehäuseOben = new Gehäuse(inventorApp, status, 0.5, 0.06, 0.1, platine.BoardW, platine.BoardL, normteile.GetInsertHoleDia(platine.HoleDia * 10) * 0.1, platine.CornerRadius, platine.BoardH, platine.CompHeightTop, 0.5, normteile.GetScrewHeadDia(platine.HoleDia * 10) * 0.1, normteile.GetScrewHeadHeight(platine.HoleDia * 10) * 0.1, true);
                gehäuseUnten = new Gehäuse(inventorApp, status, 0.5, 0.06, 0.1, platine.BoardW, platine.BoardL, normteile.GetScrewDiameter(platine.HoleDia * 10) * 0.1 + 0.06, platine.CornerRadius, platine.BoardH, platine.CompHeightBottom, 0.5, normteile.GetScrewHeadDia(platine.HoleDia * 10) * 0.1, normteile.GetScrewHeadHeight(platine.HoleDia * 10) * 0.1, false);
                foreach (CutOut cutOut in platine.CutOuts)
                {
                    gehäuseOben.AddCutOut(cutOut);
                    gehäuseUnten.AddCutOut(cutOut);
                }
                gehäuseOben.Save(speichern.getPathOben()); 
                gehäuseUnten.Save(speichern.getPathUnten());

                //zusammenfügen
                baugruppeZusammenfuegen = new BaugruppeZusammenfuegen(inventorApp, FilePath);
                baugruppeZusammenfuegen.PlatineHinzufügen(FilePath, platine.GetTransformationMatrix());
                baugruppeZusammenfuegen.UnteresGehäuseHinzufügen(speichern.getPathUnten(), platine.BoardH);
                baugruppeZusammenfuegen.OberesGehäuseHinzufügen(speichern.getPathOben(), platine.BoardH);
                baugruppeZusammenfuegen.SchraubenHinzufügen(normteile.GetScrewDiameter(platine.HoleDia * 10), platine.BoardW, platine.BoardL, platine.CornerRadius, gehäuseUnten.GetScrewOffset());
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listPanel.Add(panel1);
            listPanel.Add(panel2);
            listPanel.Add(panel3);
            listPanel[index].BringToFront();
        }

        private void btnzurueck_Click(object sender, EventArgs e)
        {
            if (index > 0)
                listPanel[--index].BringToFront();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //char ch = e.KeyChar;

            //textBox1= Format(textBox1, "0.00mm");

            //if (ch == 46 && textBox1.Text.IndexOf('.') != -1)
            //{
            //e.Handled = true;
            //return;
            //}

            //if(!Char.IsDigit(ch) && ch != 8 && ch != 46)
            //{
            //e.Handled = true;
            //}
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;

            //  textBox1 = Format(ch, "0.00mm");
            //textBox1 = Char.ToString("0.##");
            //string s = d.ToString("0.##");
            //formatString += ".00";
            //textBox2.Text = char.ToString("0.##");

            if (ch == 46 && textBox2.Text.IndexOf('.') != -1)
            {
                e.Handled = true;
                return;
            }

            if (!Char.IsDigit(ch) && ch != 8 && ch != 46)
            {
                e.Handled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(btnzip.Text == "")
            {
                btnzip.Text = "X";
            }
            else
            {
                btnzip.Text = "";
            }
                
        }

        private void btnSelFile_Click(object sender, EventArgs e)
        {
            string FileName = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Please Select your CircuitBoardModell...";
            openFileDialog.Filter = "Inventor Assembly (*.iam) | *.iam";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                FilePath = openFileDialog.FileName;
                FileName = openFileDialog.SafeFileName;
            }

            if (FilePath == null)
            {
                MessageBox.Show("No File Selected!");
            }
            else
            {
                platine = new Platine(inventorApp, FilePath, status);
                platine.Analyze();
                cmbBoard.DataSource = platine.Parts;
                comboBox4.DataSource = platine.Parts;
                comboBox6.DataSource = platine.Parts;
                textBox5.Text = FileName;
                platine.SavePictureAs("C:\\temp\\Platine.jpg");

            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            status.Name = "Inventor starting";
            status.OnProgess();

            Type inventorAppType = System.Type.GetTypeFromProgID("Inventor.Application");
            inventorApp = System.Activator.CreateInstance(inventorAppType) as Inventor.Application;
            inventorApp.Visible = false;

            status.Name = "Done";
            status.OnProgess();

            normteile = new Normteile();
        }

        private void btnAddCon(object sender, EventArgs e)
        {
            platine.AddConnectorToCutOuts(platine.Parts.ElementAt(comboBox4.SelectedIndex));
        }

        private void btnAddLed(object sender, EventArgs e)
        {
            platine.AddLEDToCutOuts(platine.Parts.ElementAt(comboBox4.SelectedIndex));
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void btnConfirmBoard_Click(object sender, EventArgs e)
        {
            platine.AnalyzeBoard(platine.Parts.ElementAt(cmbBoard.SelectedIndex));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            speichern = new Speichern();

            //export button
            if (speichern.selectedPath != null)
            {
                speichern.exportFiles();
                baugruppeZusammenfuegen.packAndGo(speichern.getPathBaugruppe(), speichern.folderPathCAD);
            }
            else
            {
                MessageBox.Show("Bitte Speicherort wählen.");
            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            speichern = new Speichern();
            FolderBrowserDialog diag = new FolderBrowserDialog();
            if (diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                speichern.selectedPath = diag.SelectedPath;
            }
        }

        private void UpdateStatus(object sender, EventArgs e)
        {
            //MessageBox.Show(status.Name);
            label16.Text = status.Name;
            progressBar1.Value = status.Progress;
            if (status.Progress > 0) {progressBar1.Value = status.Progress - 1;}
            else { progressBar1.Value = 0;}
            progressBar1.Value = status.Progress;
        }
    }
}
    

