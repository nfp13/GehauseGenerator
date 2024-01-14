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
            _GenerateTemplateFile();
            _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _OpenRefFile();
        }

        private void _OpenRefFile()
        {
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string NormteilePath = $"{System.IO.Path.GetTempPath()}Normteile.xlsx";
            _wb = _excelApp.Workbooks.Open(NormteilePath);
            _wsInsert = _wb.Worksheets[1];
            _wsScrew = _wb.Worksheets[2];
        }
        public double GetInsertHoleDia(double holeDia)
        {
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
            _wb.Close();
            _excelApp.Quit();
        }

        private Microsoft.Office.Interop.Excel.Application _excelApp;
        private Workbook _wb;
        private Worksheet _wsInsert, _wsScrew;
    }
}
    

