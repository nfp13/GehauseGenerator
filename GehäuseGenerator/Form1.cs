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
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Inventor;


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

        //Für Seiten wechseln
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
                    speichern = new Speichern(status);
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
                    this.ForeColor = System.Drawing.Color.FromArgb(47, 54, 64);
                    this.BackColor = System.Drawing.Color.FromArgb(245, 246, 250);
                    pnlleiste.BackColor = System.Drawing.Color.FromArgb(231, 231, 231);
                    ConfigurationManager.AppSettings[key] = "dark";
                }
                else
                {
                    btnchangemode.Text = "Light Mode";
                    this.ForeColor = System.Drawing.Color.FromArgb(245, 246, 250);
                    this.BackColor = System.Drawing.Color.FromArgb(29, 29, 29);
                    pnlleiste.BackColor = System.Drawing.Color.FromArgb(49, 49, 49);
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
                speichern = new Speichern(status);
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
                baugruppeZusammenfuegen = new BaugruppeZusammenfuegen(inventorApp, FilePath, status);
                baugruppeZusammenfuegen.PlatineHinzufügen(FilePath, platine.GetTransformationMatrix());
                baugruppeZusammenfuegen.UnteresGehäuseHinzufügen(speichern.getPathUnten(), platine.BoardH);
                baugruppeZusammenfuegen.OberesGehäuseHinzufügen(speichern.getPathOben(), platine.BoardH);
                baugruppeZusammenfuegen.SchraubenHinzufügen(normteile.GetScrewDiameter(platine.HoleDia * 10), platine.BoardW, platine.BoardL, platine.CornerRadius, gehäuseUnten.GetScrewOffset());

                baugruppeZusammenfuegen.SavePictureAsOben(speichern.getPathScreenGOben());
                picScreenOben.Image = Image.FromFile(speichern.getPathScreenGOben());

                baugruppeZusammenfuegen.SavePictureAsUnten(speichern.getPathScreenGUnten());
                picScreenUnten.Image = Image.FromFile(speichern.getPathScreenGUnten());
            }

        }

       

        private void btnzurueck_Click(object sender, EventArgs e)
        {
            if (index > 0)
                listPanel[--index].BringToFront();
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
            speichern = new Speichern(status);
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
                platine.SavePictureAs(speichern.getPathScreenBoard());
                picScreenBoard.Image = Image.FromFile(speichern.getPathScreenBoard());
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

        //Seiten wechseln
        private void Form1_Load_1(object sender, EventArgs e)   
        {
            listPanel.Add(panel1);
            //listPanel.Add(panel2);
            listPanel.Add(panel3);
            listPanel[index].BringToFront();
        }

        private void btnConfirmBoard_Click(object sender, EventArgs e)
        {
            platine.AnalyzeBoard(platine.Parts.ElementAt(cmbBoard.SelectedIndex));
        }

        private void button3_Click(object sender, EventArgs e)
        {

            //export button
            //if (speichern.saveAs == 1)
            //{
                speichern.exportFiles();
                baugruppeZusammenfuegen.packAndGo(speichern.getPathBaugruppe(), speichern.folderPathCAD);
            //}
            //else
            //{
                //MessageBox.Show("Bitte Speicherort wählen.");
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            speichern = new Speichern(status);
            FolderBrowserDialog diag = new FolderBrowserDialog();
            if (diag.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                speichern.selectedPath = diag.SelectedPath;
            }
        }

        //Textboxen nur zwei Nachkommastellen
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string enteredText = (sender as System.Windows.Forms.TextBox).Text;
            int cursorPosition = (sender as System.Windows.Forms.TextBox).SelectionStart;

            string[] splitByDecimal = enteredText.Split('.');

            if (splitByDecimal.Length > 1 && splitByDecimal[1].Length > 2)
            {
                (sender as System.Windows.Forms.TextBox).Text = enteredText.Remove(enteredText.Length - 1);
                (sender as System.Windows.Forms.TextBox).SelectionStart = cursorPosition - 1;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string enteredText = (sender as System.Windows.Forms.TextBox).Text;
            int cursorPosition = (sender as System.Windows.Forms.TextBox).SelectionStart;

            string[] splitByDecimal = enteredText.Split('.');

            if (splitByDecimal.Length > 1 && splitByDecimal[1].Length > 2)
            {
                (sender as System.Windows.Forms.TextBox).Text = enteredText.Remove(enteredText.Length - 1);
                (sender as System.Windows.Forms.TextBox).SelectionStart = cursorPosition - 1;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string enteredText = (sender as System.Windows.Forms.TextBox).Text;
            int cursorPosition = (sender as System.Windows.Forms.TextBox).SelectionStart;

            string[] splitByDecimal = enteredText.Split('.');

            if (splitByDecimal.Length > 1 && splitByDecimal[1].Length > 2)
            {
                (sender as System.Windows.Forms.TextBox).Text = enteredText.Remove(enteredText.Length - 1);
                (sender as System.Windows.Forms.TextBox).SelectionStart = cursorPosition - 1;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            string enteredText = (sender as System.Windows.Forms.TextBox).Text;
            int cursorPosition = (sender as System.Windows.Forms.TextBox).SelectionStart;

            string[] splitByDecimal = enteredText.Split('.');

            if (splitByDecimal.Length > 1 && splitByDecimal[1].Length > 2)
            {
                (sender as System.Windows.Forms.TextBox).Text = enteredText.Remove(enteredText.Length - 1);
                (sender as System.Windows.Forms.TextBox).SelectionStart = cursorPosition - 1;
            }
        }

        //Textboxen keine Buchstaben und nur ein Punkt
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void picScreenBoard_Paint(object sender, PaintEventArgs e)
        {

            ControlPaint.DrawBorder(e.Graphics, picScreenBoard.ClientRectangle, System.Drawing.Color.White, ButtonBorderStyle.Solid);
        }

        private void picScreenOben_Paint(object sender, PaintEventArgs e)
        {

            ControlPaint.DrawBorder(e.Graphics, picScreenOben.ClientRectangle, System.Drawing.Color.White, ButtonBorderStyle.Solid);
        }

        private void picScreenUnten_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, picScreenUnten.ClientRectangle, System.Drawing.Color.White, ButtonBorderStyle.Solid);
        }

        //MessageBox.Show(status.Name);
        private void UpdateStatus(object sender, EventArgs e)
        {
            label16.Text = status.Name;
            progressBar1.Value = status.Progress;
            if (status.Progress > 0) {progressBar1.Value = status.Progress - 1;}
            else { progressBar1.Value = 0;}
            progressBar1.Value = status.Progress;
        }
    }
}
    

