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
            _status = status;

            _status.Name = "Opening Assembly";
            _status.OnProgess();

            _filePath = filePath;
            _inventorApp = inventorApp;

            _assemblyDocument = _inventorApp.Documents.Open(_filePath, true) as AssemblyDocument;
            _assemblyComponentDefinition = _assemblyDocument.ComponentDefinition;

            _status.Name = "Done";
            _status.OnProgess();

        }

        public void Analyze()
        {
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
                if (componentOccurrence.Name == boardOccurrenceName)
                {
                    //MessageBox.Show("boardFound");
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

                                Inventor.Point point = holeFeature.RangeBox.MinPoint;

                                if (_MinPointBoard.VectorTo(point).X + HoleDia / 2 < CornerRadius || CornerRadius == 0)
                                {
                                    CornerRadius = _MinPointBoard.VectorTo(point).X + HoleDia / 2;
                                }

                            }
                            //MessageBox.Show(CornerRadius.ToString("0.00") + " ; " + HoleDia.ToString("0.00"));
                        }
                    }

                }
                _status.Progress = 100;
                _status.Name = "Done";
                _status.OnProgess();
            }

            _MaxPointGes = _assemblyComponentDefinition.RangeBox.MaxPoint;
            _MinPointGes = _assemblyComponentDefinition.RangeBox.MinPoint;

            CompHeightBottom = _MinPointGes.VectorTo(_MinPointBoard).Z;
            CompHeightTop = _MaxPointBoard.VectorTo(_MaxPointGes).Z;

            //MessageBox.Show(CompHeightTop.ToString("0.00") + " ; " + CompHeightBottom.ToString());

            _status.Name = "Done";
            _status.OnProgess();

        }

        public Matrix GetTransformationMatrix()
        {
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
            foreach (ComponentOccurrence componentOccurrence in _assemblyComponentDefinition.Occurrences)
            {
                if (componentOccurrence.Name == CutOutOccurrenceName)
                {
                    Box box = componentOccurrence.Definition.RangeBox;

                    Vector Top = box.MaxPoint.VectorTo(_MaxPointBoard);
                    Vector Bottom = box.MinPoint.VectorTo(_MinPointBoard);

                    CutOut _CutOut = new CutOut();

                    _CutOut.Connector = true;

                    _CutOut.XS = box.MaxPoint.X - box.MinPoint.X;
                    _CutOut.YS = box.MaxPoint.Y - box.MinPoint.Y;
                    _CutOut.ZS = box.MaxPoint.Z - box.MinPoint.Z;

                    _CutOut.XP = box.MinPoint.X + (_CutOut.XS / 2) - _XOM;
                    _CutOut.YP = box.MinPoint.Y + (_CutOut.YS / 2) - _YOM;
                    _CutOut.ZP = box.MinPoint.Z + (_CutOut.ZS / 2) - _ZOM;

                    //MessageBox.Show(box.MinPoint.Z.ToString("0.00") + " " + _ZOM.ToString("0.00") + box.MaxPoint.Z.ToString("0.00") + " ");

                    if (Top.Z <= 0.0)
                    {
                        _CutOut.Top = true;
                        //MessageBox.Show("Oben");
                    }
                    else if (Bottom.Z >= 0.0)
                    {
                        _CutOut.Top = false;
                        //MessageBox.Show("Unten");
                    }
                    else
                    {
                        MessageBox.Show("Fehler");
                    }

                    _CutOut.Name = componentOccurrence.Name;

                    //MessageBox.Show(_CutOut.XP.ToString("0.00") + " " + _CutOut.YP.ToString("0.00") + " " + _CutOut.ZP.ToString("0.00"));
                    //MessageBox.Show(_CutOut.XS.ToString("0.00") + " " + _CutOut.YS.ToString("0.00") + " " + _CutOut.ZS.ToString("0.00"));


                    CutOuts.Add(_CutOut);
                }
            }
        }

        public void AddLEDToCutOuts(string CutOutOccurrenceName)
        {
            foreach (ComponentOccurrence componentOccurrence in _assemblyComponentDefinition.Occurrences)
            {
                if (componentOccurrence.Name == CutOutOccurrenceName)
                {
                    Box box = componentOccurrence.Definition.RangeBox;

                    CutOut _CutOut = new CutOut();

                    Vector Top = box.MaxPoint.VectorTo(_MaxPointBoard);
                    Vector Bottom = box.MinPoint.VectorTo(_MinPointBoard);

                    _CutOut.Connector = false;

                    _CutOut.XS = box.MaxPoint.X - box.MinPoint.X;
                    _CutOut.YS = box.MaxPoint.Y - box.MinPoint.Y;
                    _CutOut.ZS = box.MaxPoint.Z - box.MinPoint.Z;

                    _CutOut.XP = box.MinPoint.X + (_CutOut.XS / 2) - _XOM;
                    _CutOut.YP = box.MinPoint.Y + (_CutOut.YS / 2) - _YOM;
                    _CutOut.ZP = box.MinPoint.Z + (_CutOut.ZS / 2) - _ZOM;


                    if (Top.Z <= 0.0)
                    {
                        _CutOut.Top = true;
                        //MessageBox.Show("Oben");
                    }
                    else if (Bottom.Z >= 0.0)
                    {
                        _CutOut.Top = false;
                        //MessageBox.Show("Unten");
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
            _status.Name = "Taking Screenshot";
            _status.Progress = 0;
            _status.OnProgess();

            Inventor.View view = _assemblyDocument.Views[1];
            Inventor.Camera camera = view.Camera;
            camera.Perspective = false;

            _status.Progress = 75;
            _status.OnProgess();

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

        public double BoardW, BoardL, BoardH;               //Grundmaße der Platine
        public double CompHeightTop, CompHeightBottom;      //Höhe der Komponenten auf der Platine
        public double HoleDia, CornerRadius = 0;                //Durchmesser der Bohrungslöcher und Rundungsradius der Platine

        private double _XOM, _YOM, _ZOM;           //Coordinates of centerpoint of Board

        private Inventor.Point _MinPointBoard;
        private Inventor.Point _MaxPointBoard;
        private Inventor.Point _MinPointGes;
        private Inventor.Point _MaxPointGes;

        public List<CutOut> CutOuts = new List<CutOut>();

        private Status _status;

    }

}
    

