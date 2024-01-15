using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace GehäuseGenerator
{
    public class Normteile
    {
        public Normteile()
        {
            //Generating Template file, crating new instance of excel, opening Template file

            _GenerateTemplateFile();
            _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _OpenRefFile();
        }

        private void _OpenRefFile()
        {
            //Openening the Reference file, setting the WorkSheet variables

            string NormteilePath = $"{System.IO.Path.GetTempPath()}Normteile.xlsx";
            _wb = _excelApp.Workbooks.Open(NormteilePath);
            _wsInsert = _wb.Worksheets[1];
            _wsScrew = _wb.Worksheets[2];
        }
        public double GetInsertHoleDia(double holeDia)
        {
            //Get the largest posible screw diameter for the given hole size, return corresponding InsertHoleDiameter

            int MaxRowInd = _wsInsert.Rows.Count;
            int i = 2;
            double readExcelDia = 0.0;

            while (i < MaxRowInd && readExcelDia < holeDia)
            {
                i++;
                readExcelDia = _wsInsert.Cells[i, 1].Value;
            }
            i--;
            if (_wsInsert.Cells[i + 1, 1].Value >= holeDia)
            {
                return (_wsInsert.Cells[i, 2].Value);
            }
            else
            {
                MessageBox.Show("Löcher in Platine zu groß!");
                return 0.0;
            }

        }

        public double GetScrewHeadDia(double holeDia)
        {
            //Get the largest posible screw diameter for the given hole size, return corresponding ScrewHeadDiameter

            int MaxRowInd = _wsScrew.Rows.Count;
            int i = 2;
            double readExcelDia = 0.0;

            while (i < MaxRowInd && readExcelDia < holeDia)
            {
                i++;
                readExcelDia = _wsScrew.Cells[i, 1].Value;
            }
            i--;
            if (_wsScrew.Cells[i + 1, 1].Value >= holeDia)
            {
                return (_wsScrew.Cells[i, 2].Value);
            }
            else
            {
                MessageBox.Show("Löcher in Platine zu groß!");
                return 0.0;
            }
        }

        public double GetScrewHeadHeight(double holeDia)
        {
            //Get the largest posible screw diameter for the given hole size, return corresponding ScrewHeadHeight

            int MaxRowInd = _wsScrew.Rows.Count;
            int i = 2;
            double readExcelDia = 0.0;

            while (i < MaxRowInd && readExcelDia < holeDia)
            {
                i++;
                readExcelDia = _wsScrew.Cells[i, 1].Value;
            }
            i--;
            if (_wsScrew.Cells[i + 1, 1].Value >= holeDia)
            {
                return (_wsScrew.Cells[i, 3].Value);
            }
            else
            {
                MessageBox.Show("Löcher in Platine zu groß!");
                return 0.0;
            }
        }

        public double GetScrewDiameter(double holeDia)
        {
            //Get the largest posible screw diameter for the given hole size, return corresponding ScrewDíameter

            int MaxRowInd = _wsScrew.Rows.Count;
            int i = 2;
            double readExcelDia = 0.0;

            while (i < MaxRowInd && readExcelDia < holeDia)
            {
                i++;
                readExcelDia = _wsScrew.Cells[i, 1].Value;
            }
            i--;
            if (_wsScrew.Cells[i + 1, 1].Value >= holeDia)
            {
                return (_wsScrew.Cells[i, 1].Value);
            }
            else
            {
                MessageBox.Show("Löcher in Platine zu groß!");
                return 0.0;
            }
        }

        private void _GenerateTemplateFile()
        {
            //Creating Template from Resources and saving it to the temp folder 

            byte[] templateFile = Properties.Resources.Normteile;
            string tempPath = $"{System.IO.Path.GetTempPath()}Normteile.xlsx";
            using (MemoryStream ms = new MemoryStream(templateFile))
            {
                using (FileStream fs = new FileStream(tempPath, FileMode.OpenOrCreate))
                {
                    ms.WriteTo(fs);
                    fs.Close();
                }
                ms.Close();
            }
        }

        public void CloseExcel()
        {
            //Close the WorkBook and excel

            _wb.Close();
            _excelApp.Quit();
        }

        private Microsoft.Office.Interop.Excel.Application _excelApp;
        private Workbook _wb;
        private Worksheet _wsInsert, _wsScrew;
    }
}