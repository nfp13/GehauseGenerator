using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace GehäuseGenerator
{
    public class Normteile
    {
        public Normteile()
        {
            _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _OpenRefFile();
        }

        private void _OpenRefFile()
        {
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string NormteilePath = string.Format("{0}Resources\\Normteile.xlsx", System.IO.Path.GetFullPath(System.IO.Path.Combine(RunningPath, @"..\..\")));
            _wb = _excelApp.Workbooks.Open(NormteilePath);
            _wsInsert = _wb.Worksheets[1];
            _wsScrew = _wb.Worksheets[2];
        }
        public double GetInsertHoleDia(double holeDia)
        {
            int MaxRowInd = _wsInsert.Rows.Count;
            int i = 3;
            double readExcelDia = 0.0;

            while (i <= MaxRowInd && !(readExcelDia >= holeDia))
            {
                readExcelDia = _wsInsert.Cells[i, 1].Value;
                i++;
            }
            i++;

            if (_wsInsert.Cells[i, 1].Value >= holeDia)
            {
                return (_wsInsert.Cells[i, 2].Value * 0.1);
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
            int i = 3;
            double readExcelDia = 0.0;

            while (i <= MaxRowInd && !(readExcelDia >= holeDia))
            {
                readExcelDia = _wsScrew.Cells[i, 1].Value;
                i++;
            }
            i++;

            if (_wsScrew.Cells[i, 1].Value >= holeDia)
            {
                return (_wsScrew.Cells[i, 2].Value * 0.1);
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
            int i = 3;
            double readExcelDia = 0.0;

            while (i <= MaxRowInd && !(readExcelDia >= holeDia))
            {
                readExcelDia = _wsScrew.Cells[i, 1].Value;
                i++;
            }
            i++;

            if (_wsScrew.Cells[i, 1].Value >= holeDia)
            {
                return (_wsScrew.Cells[i, 3].Value * 0.1);
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
            int i = 3;
            double readExcelDia = 0.0;

            while (i <= MaxRowInd && !(readExcelDia >= holeDia))
            {
                readExcelDia = _wsScrew.Cells[i, 1].Value;
                i++;
            }
            i++;

            if (_wsScrew.Cells[i, 1].Value >= holeDia)
            {
                return (_wsScrew.Cells[i, 1].Value);
            }
            else
            {
                MessageBox.Show("Löcher in Platine zu groß!");
                return 0.0;
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
    

