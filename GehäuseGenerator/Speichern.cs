﻿using System;
using System.IO;
using System.Windows.Forms;
using System.IO.Compression;


namespace GehäuseGenerator
{
    public class Speichern
    {
        public Speichern(Status status)
        {
            _status = status;
            _status.Name = "Getting Temp Path";
            _status.OnProgess();

            //Gets temporary folder of the user 
            string result = Path.GetTempPath();
            _tempPath = result;

            _status.Name = "Done";
            _status.OnProgess();
        }

        //Gets temporary paths for the Inventor parts and the assebley
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

        //Gets temporary paths for the screenshots
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

        //Gets paths for the selected 3D-model file type
        public string getPathObenStl()
        {
            string[] paths = { folderPathDruck, "Obere Gehäusehälfte.stl" };
            _pathObenStl = Path.Combine(paths);
            return _pathObenStl;
        }
        public string getPathUntenStl()
        {
            string[] paths = { folderPathDruck, "Untere Gehäusehälfte.stl" };
            _pathUntenStl = Path.Combine(paths);
            return _pathUntenStl;
        }
        public string getPathObenStp()
        {
            string[] paths = { folderPathDruck, "Obere Gehäusehälfte.stp" };
            _pathObenStp = Path.Combine(paths);
            return _pathObenStp;
        }
        public string getPathUntenStp()
        {
            string[] paths = { folderPathDruck, "Untere Gehäusehälfte.stp" };
            _pathUntenStp = Path.Combine(paths);
            return _pathUntenStp;
        }
        public string getPathObenOBJ()
        {
            string[] paths = { folderPathDruck, "Obere Gehäusehälfte.obj" };
            _pathObenStp = Path.Combine(paths);
            return _pathObenStp;
        }
        public string getPathUntenOBJ()
        {
            string[] paths = { folderPathDruck, "Untere Gehäusehälfte.obj" };
            _pathUntenStp = Path.Combine(paths);
            return _pathUntenStp;
        }

        //Deletes created files in the temp folder
        public void deleteFiles()
        {
            File.Delete(getPathOben());
            File.Delete(getPathUnten());
            File.Delete(getPathBaugruppe());
            _tempPath = "";
            _pathOben = "";
            _pathUnten = "";
        }

        //Exports final model in a specific folder structure
        public void exportFiles()
        {
            _status.Name = "Exporting Files";
            _status.Progress = 0;
            _status.OnProgess();

            //Main folder
            string[] paths = {selectedPath, "Platinen Gehäuse" };
            folderPath = Path.Combine(paths);
            var dir1 = folderPath;
            if (!Directory.Exists(dir1))
            {
                Directory.CreateDirectory(dir1);
            }

            _status.Progress = 30;
            _status.OnProgess();

            //3D-model folder
            string[] pathsDruck = { folderPath, "3D-Druck" };
            folderPathDruck = Path.Combine(pathsDruck);
            var dir2 = folderPathDruck;
            if (!Directory.Exists(dir2))
            {
                Directory.CreateDirectory(dir2);
            }

            _status.Progress = 60;
            _status.OnProgess();

            //CAD folder
            string[] pathsCAD = { folderPath, "CAD" };
            folderPathCAD = Path.Combine(pathsCAD);
            var dir3 = folderPathCAD;
            if (!Directory.Exists(dir3))
            {
                Directory.CreateDirectory(dir3);
            }
            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();

        }

        //Compresses the Main folder into a .zip
        public void makeZip()
        {
            _status.Name = "Creating Zip";
            _status.Progress = 25;
            _status.OnProgess();

            string[] paths = { selectedPath, "PlatinenGehäuse.zip" };
            string zipPath = Path.Combine(paths);
            System.IO.Compression.ZipFile.CreateFromDirectory(folderPath, zipPath);

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        //Initialization of the required variables
        private Status _status;
        private string _tempPath;
        private string _pathOben;
        private string _pathUnten;
        private string _pathBaugruppen;
        private string _pathScreenGOben;
        private string _pathScreenGUnten;
        private string _pathScreenBoard;
        private string _pathObenStl;
        private string _pathUntenStl;
        private string _pathObenStp;
        private string _pathUntenStp;
        public string folderPathCAD, folderPathDruck, selectedPath, folderPath;
    }

}
    

