using System;
using System.IO;


namespace GehäuseGenerator
{
    public class Speichern
    {
        public Speichern()
        {
            string result = Path.GetTempPath();
            _tempPath = result;

            //MessageBox.Show("path: " + _tempPath);
        }

        public string getPathOben()
        {
            string[] paths = { _tempPath, "ObereGehäusehälfte.ipt" };
            _pathOben = Path.Combine(paths);
            return _pathOben;
        }

        public string getPathUnten()
        {
            string[] paths = { _tempPath, "UntereGehäusehälfte.ipt" };
            _pathUnten = Path.Combine(paths);
            return _pathUnten;
        }

        public string getPathBaugruppe()
        {
            string[] paths = { _tempPath, "Gesamt.iam" };
            _pathBaugruppen = Path.Combine(paths);
            return _pathBaugruppen;
        }

        public void deleteFiles()
        {
            File.Delete(getPathOben());
            File.Delete(getPathUnten());
            File.Delete(getPathBaugruppe());
            File.Delete(getPathScreenBoard());
            File.Delete(getPathScreenGOben());
            File.Delete(getPathScreenGUnten());
            _tempPath = "";
            _pathOben = "";
            _pathUnten = "";
        }

        public void exportFiles()
        {
            //Hauptordner
            //desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string[] paths = { selectedPath, "Platinen Gehäuse" };
            string folderPath = Path.Combine(paths);
            var dir1 = folderPath;
            if (!Directory.Exists(dir1))
            {
                Directory.CreateDirectory(dir1);
            }

            //3D Druck Ordner
            string[] pathsDruck = { folderPath, "3D-Druck" };
            folderPathDruck = Path.Combine(pathsDruck);
            var dir2 = folderPathDruck;
            if (!Directory.Exists(dir2))
            {
                Directory.CreateDirectory(dir2);
            }

            //CAD Ordner
            string[] pathsCAD = { folderPath, "CAD" };
            folderPathCAD = Path.Combine(pathsCAD);
            var dir3 = folderPathCAD;
            if (!Directory.Exists(dir3))
            {
                Directory.CreateDirectory(dir3);
            }

        }

        public string getPathScreenBoard()
        {
            string[] paths = { _tempPath, "ScreenGOben.jpg" };
            _pathScreenGOben = Path.Combine(paths);
            return _pathScreenGOben;
        }
        public string getPathScreenGOben()
        {
            string[] paths = { _tempPath, "ScreenGUnten.jpg" };
            _pathScreenGUnten = Path.Combine(paths);
            return _pathScreenGUnten;
        }
        public string getPathScreenGUnten()
        {
            string[] paths = { _tempPath, "ScreenBoard.jpg" };
            _pathScreenBoard = Path.Combine(paths);
            return _pathScreenBoard;
        }

        private string _tempPath;
        private string _pathOben;
        private string _pathUnten;
        private string _pathBaugruppen;
        private string _pathScreenGOben;
        private string _pathScreenGUnten;
        private string _pathScreenBoard;
        public string folderPathCAD, folderPathDruck, selectedPath;
    }

}
    

