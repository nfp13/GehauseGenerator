using System.Collections.Generic;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Inventor;
using System;

namespace GehäuseGenerator
{
    public class Platine
    {
        public Platine(Inventor.Application inventorApp, string filePath, Status status)
        {
            //Setting private variables

            _status = status;
            _filePath = filePath;
            _inventorApp = inventorApp;

            _status.Name = "Opening Assembly";
            _status.OnProgess();

            //Opening the PCB assembly Document

            _assemblyDocument = _inventorApp.Documents.Open(_filePath, true) as AssemblyDocument;       
            _assemblyComponentDefinition = _assemblyDocument.ComponentDefinition;

            _status.Name = "Done";
            _status.OnProgess();

        }

        public void Analyze()
        {
            //Creating list containing all component occurence names

            Parts.Clear();
            int NumberOfOccurrences = _assemblyComponentDefinition.Occurrences.Count;
            int CurrentOccurrence = 1;
            _status.Name = "Finding Parts";
            _status.Progress = 0;
            _status.OnProgess();
            foreach (ComponentOccurrence componentOccurrence in _assemblyComponentDefinition.Occurrences)
            {
                Parts.Add(componentOccurrence.Name);
                _status.Progress = Convert.ToInt16((CurrentOccurrence * 1.0) / (NumberOfOccurrences * 1.0)*100);
                CurrentOccurrence++;
                _status.OnProgess();
            }

            _status.Name = "Done";
            _status.OnProgess();

        }

        public void AnalyzeBoard(string boardOccurrenceName)
        {
            _status.Name = "Analyzing Board";
            _status.OnProgess();

            foreach (ComponentOccurrence componentOccurrence in _assemblyComponentDefinition.Occurrences)
            {
                //Finding Board Assembly and calculating size, coordinates of centerpoint, holediameter and radius of corner

                if (componentOccurrence.Name == boardOccurrenceName)
                {
                    Box box = componentOccurrence.Definition.RangeBox;

                    _MinPointBoard = box.MinPoint;
                    _MaxPointBoard = box.MaxPoint;

                    BoardW = (box.MaxPoint.X - box.MinPoint.X);
                    BoardL = (box.MaxPoint.Y - box.MinPoint.Y);
                    BoardH = (box.MaxPoint.Z - box.MinPoint.Z);

                    _XOM = box.MinPoint.X + (BoardW / 2);
                    _YOM = box.MinPoint.Y + (BoardL / 2);
                    _ZOM = box.MinPoint.Z + (BoardH / 2);

                    AssemblyDocument boardAssemblyDocument;
                    boardAssemblyDocument = componentOccurrence.Definition.Document;

                    AssemblyComponentDefinition boardAssemblyComponentDefinition;
                    boardAssemblyComponentDefinition = boardAssemblyDocument.ComponentDefinition;

                    foreach (ComponentOccurrence componentOccurrence1 in boardAssemblyComponentDefinition.Occurrences)
                    {
                        if (componentOccurrence1.Name.Contains("STEP"))
                        {
                            PartDocument boardPartDocument;
                            boardPartDocument = componentOccurrence1.Definition.Document;

                            _status.Progress = 60;
                            _status.OnProgess();

                            foreach (HoleFeature holeFeature in boardPartDocument.ComponentDefinition.Features.HoleFeatures)
                            {
                                HoleDia = holeFeature.HoleDiameter.Value;

                                Inventor.Point minPointHole = holeFeature.RangeBox.MinPoint;

                                if (_MinPointBoard.VectorTo(minPointHole).X + HoleDia / 2 < CornerRadius || CornerRadius == 0)
                                {
                                    CornerRadius = _MinPointBoard.VectorTo(minPointHole).X + HoleDia / 2;
                                }
                            }
                        }
                    }
                }
                _status.Progress = 100;
                _status.Name = "Done";
                _status.OnProgess();
            }

            //Calculating component height for bottom and top of PCB

            _MaxPointGes = _assemblyComponentDefinition.RangeBox.MaxPoint;
            _MinPointGes = _assemblyComponentDefinition.RangeBox.MinPoint;

            CompHeightBottom = _MinPointGes.VectorTo(_MinPointBoard).Z;
            CompHeightTop = _MaxPointBoard.VectorTo(_MaxPointGes).Z;

            _status.Name = "Done";
            _status.OnProgess();

        }

        public Matrix GetTransformationMatrix()
        {
            //Creating transformation Matrix for placing the PCB-centerpoint in the origin of an assembly document

            Vector MinMaxVector = _MinPointBoard.VectorTo(_MaxPointBoard);
            MinMaxVector.ScaleBy(0.5);
            Inventor.Point MiddlePointPlatine = _MinPointBoard;
            MiddlePointPlatine.TranslateBy(MinMaxVector);
            Vector NewVector = MiddlePointPlatine.VectorTo(_inventorApp.TransientGeometry.CreatePoint(0, 0, 0));
            Inventor.Point NewOrigin = _inventorApp.TransientGeometry.CreatePoint(0, 0, 0);
            NewOrigin.TranslateBy(NewVector);
            Matrix OriginConversion = _inventorApp.TransientGeometry.CreateMatrix();
            OriginConversion.SetCoordinateSystem(NewOrigin, _inventorApp.TransientGeometry.CreateVector(1, 0, 0), _inventorApp.TransientGeometry.CreateVector(0, 1, 0), _inventorApp.TransientGeometry.CreateVector(0, 0, 1));
            return OriginConversion;
        }

        public void AddConnectorToCutOuts(string CutOutOccurrenceName)
        {
            //Adding Connector to CutOuts list and getting all nessesary variables

            foreach (ComponentOccurrence componentOccurrence in _assemblyComponentDefinition.Occurrences)
            {
                //Finding specific occurence

                if (componentOccurrence.Name == CutOutOccurrenceName)
                {
                    Box box = componentOccurrence.Definition.RangeBox;

                    Vector Top = box.MaxPoint.VectorTo(_MaxPointBoard);
                    Vector Bottom = box.MinPoint.VectorTo(_MinPointBoard);

                    CutOut _CutOut = new CutOut();

                    _CutOut.Connector = true;

                    //Calculating size of the connector in x, y and z direction

                    _CutOut.XS = box.MaxPoint.X - box.MinPoint.X;
                    _CutOut.YS = box.MaxPoint.Y - box.MinPoint.Y;
                    _CutOut.ZS = box.MaxPoint.Z - box.MinPoint.Z;

                    //Calculating location of connector-centerpoint relative to board-centerpoint in x, y and z direction

                    _CutOut.XP = box.MinPoint.X + (_CutOut.XS / 2) - _XOM;
                    _CutOut.YP = box.MinPoint.Y + (_CutOut.YS / 2) - _YOM;
                    _CutOut.ZP = box.MinPoint.Z + (_CutOut.ZS / 2) - _ZOM;

                    //Checking if connector is located on the top or the bottom of the PCB

                    if (Top.Z <= 0.0)
                    {
                        _CutOut.Top = true;
                    }
                    else if (Bottom.Z >= 0.0)
                    {
                        _CutOut.Top = false;
                    }
                    else
                    {
                        MessageBox.Show("Fehler");
                    }

                    _CutOut.Name = componentOccurrence.Name;
                    CutOuts.Add(_CutOut);
                }
            }
        }

        public void AddLEDToCutOuts(string CutOutOccurrenceName)
        {
            //Adding LED/Display to CutOuts list and getting all nessesary variables

            foreach (ComponentOccurrence componentOccurrence in _assemblyComponentDefinition.Occurrences)
            {
                //Finding specific occurence

                if (componentOccurrence.Name == CutOutOccurrenceName)
                {
                    Box box = componentOccurrence.Definition.RangeBox;

                    CutOut _CutOut = new CutOut();

                    Vector Top = box.MaxPoint.VectorTo(_MaxPointBoard);
                    Vector Bottom = box.MinPoint.VectorTo(_MinPointBoard);

                    _CutOut.Connector = false;

                    //Calculating size of the connector in x, y and z direction

                    _CutOut.XS = box.MaxPoint.X - box.MinPoint.X;
                    _CutOut.YS = box.MaxPoint.Y - box.MinPoint.Y;
                    _CutOut.ZS = box.MaxPoint.Z - box.MinPoint.Z;

                    //Calculating location of connector-centerpoint relative to board-centerpoint in x, y and z direction

                    _CutOut.XP = box.MinPoint.X + (_CutOut.XS / 2) - _XOM;
                    _CutOut.YP = box.MinPoint.Y + (_CutOut.YS / 2) - _YOM;
                    _CutOut.ZP = box.MinPoint.Z + (_CutOut.ZS / 2) - _ZOM;

                    //Checking if connector is located on the top or the bottom of the PCB

                    if (Top.Z <= 0.0)
                    {
                        _CutOut.Top = true;
                    }
                    else if (Bottom.Z >= 0.0)
                    {
                        _CutOut.Top = false;
                    }
                    else
                    {
                        MessageBox.Show("Fehler");
                    }

                    _CutOut.Name = componentOccurrence.Name;
                    CutOuts.Add(_CutOut);
                }
            }
        }

        public void SavePictureAs(string Path)
        {
            //Taking picture of PCB and saving in Path

            _status.Name = "Taking Screenshot";
            _status.Progress = 0;
            _status.OnProgess();

            //Creating the camera

            Inventor.View view = _assemblyDocument.Views[1];
            Inventor.Camera camera = view.Camera;
            camera.Perspective = false;

            _status.Progress = 75;
            _status.OnProgess();

            //Setting camera perspective, fiting the camera to the PCB and exporting picture

            camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
            camera.Fit();
            camera.Apply();
            view.Update();
            camera.SaveAsBitmap(Path, 1080, 1080);

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        private Inventor.Application _inventorApp;
        private string _filePath;
        private AssemblyDocument _assemblyDocument;
        private AssemblyComponentDefinition _assemblyComponentDefinition;
        private AssemblyDocument _boardAssemblyDocument;

        public List<string> Parts = new List<string>();

        public double BoardW, BoardL, BoardH;                       //Grundmaße der Platine
        public double CompHeightTop, CompHeightBottom;              //Höhe der Komponenten auf der Platine
        public double HoleDia, CornerRadius = 0;                    //Durchmesser der Bohrungslöcher und Rundungsradius der Platine
        private double _XOM, _YOM, _ZOM;                            //Coordinates of centerpoint of Board

        private Inventor.Point _MinPointBoard;
        private Inventor.Point _MaxPointBoard;
        private Inventor.Point _MinPointGes;
        private Inventor.Point _MaxPointGes;

        public List<CutOut> CutOuts = new List<CutOut>();

        private Status _status;
    }
}