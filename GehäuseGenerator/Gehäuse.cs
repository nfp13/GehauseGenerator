using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Inventor;


namespace GehäuseGenerator
{
    public class Gehäuse
    {
        public Gehäuse(Inventor.Application inventorApp, Status status, double DW, double TM, double TE, double BP, double LP, double DL, double REP, double HPt, double HPb, double RR, double SKD, double SKH, bool top)
        {
            _inventorApp = inventorApp;
            _status = status;
            _DW = DW;
            _TM = TM;
            _TE = TE;
            _BP = BP;
            _LP = LP;
            _DL = DL;
            _REP = REP;
            _HPt = HPt;
            _HPb = HPb;
            _RR = RR;
            _SKD = SKD;
            _SKH = SKH;
            _Top = top;
        }
        public Gehäuse(Inventor.Application inventorApp)        //For Testing without parameter changes
        {
            _inventorApp = inventorApp;
        }

        private void _OpenTemplate()
        {
            _status.Name = "Opening Gehäuse Template";
            _status.Progress = 20;
            _status.OnProgess();


            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string TemplatePath = string.Format("{0}Resources\\GehauseGenerator_GehauseVorlage.ipt", System.IO.Path.GetFullPath(System.IO.Path.Combine(RunningPath, @"..\..\")));

            _status.Progress = 40;
            _status.OnProgess();

            _partDocument = _inventorApp.Documents.Open(TemplatePath, true) as PartDocument;
            _partComponentDefinition = _partDocument.ComponentDefinition as PartComponentDefinition;

            _status.Name = "Done";
            _status.Progress = 100;
            _status.OnProgess();
        }

        private void _SetParameters()
        {
            _status.Name = "Setting Parameters";
            _status.Progress = 20;
            _status.OnProgess();

            _parameters = _partDocument.ComponentDefinition.Parameters;

            _parameters["DW"].Value = _DW;
            _parameters["TM"].Value = _TM;
            _parameters["BP"].Value = _BP;
            _parameters["LP"].Value = _LP;
            _parameters["DL"].Value = _DL;
            _parameters["REP"].Value = _REP;
            _parameters["HPt"].Value = _HPt;
            _parameters["HPb"].Value = _HPb;
            _parameters["RR"].Value = _RR;
            _parameters["SKD"].Value = _SKD;
            _parameters["SKH"].Value = _SKH;
            if (_Top)
            {
                ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures["ExtrusionSchraubenKopf"];
                extrudeFeature.Suppressed = true;
                extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures["ExtrusionSchraubenLoch"];
                extrudeFeature.Suppressed = true;
            }
            else
            {
                ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures["ExtrusionSchraubenLoch"];
                extrudeFeature.Suppressed = false;
            }

            _status.Name = "Done";
            _status.Progress = 100;
            _status.OnProgess();
        }

        public void Save(string FilePath)
        {
            _OpenTemplate();
            _SetParameters();
            _partDocument.Update();
            _AddCutOutsToModel();
            _partDocument.Update();
            _partDocument.SaveAs(FilePath, true);
            _path = FilePath;
            _partDocument.Close(true);
        }

        public double GetScrewOffset()
        {
            double offset = _HPb + _HPt * 0.5 + _DW - _SKH;
            return offset;
        }

        public void AddCutOut(CutOut _cutOut)
        {
            _CutOuts.Add(_cutOut);
        }

        private void _AddCutOutsToModel()
        {
            int ZählerCutOuts = 1;
            int AnzahlCutOuts = _CutOuts.Count;

            _status.Name = "Creating CutOuts";

            foreach (CutOut cutOut in _CutOuts)
            {
                _status.Progress = Convert.ToInt16((ZählerCutOuts * 1.0) / (AnzahlCutOuts * 1.0) * 100);
                _status.OnProgess();

                Face face;
                if (cutOut.Connector)
                {
                    if (cutOut.Top && _Top)
                    {
                        if (cutOut.XP + cutOut.XS / 2 >= _BP / 2)
                        {
                            face = _FindNamedFace("FlächeRechts");

                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(cutOut.ZP - _HPt / 2, cutOut.YP), transientGeometry.CreatePoint2d(cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2, cutOut.YP + (cutOut.YS / 2) + _TE));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                            //extrudeFeature.Name = "CutOut" + ZählerCutOuts.ToString();
                        }
                        else if (cutOut.XP - cutOut.XS / 2 <= -_BP / 2)
                        {
                            face = _FindNamedFace("FlächeLinks");
                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(-(cutOut.ZP - _HPt / 2), cutOut.YP), transientGeometry.CreatePoint2d(-(cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2), cutOut.YP + (cutOut.YS / 2) + _TE));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                        }
                        else if (cutOut.YP + cutOut.YS / 2 <= _LP / 2)
                        {
                            face = _FindNamedFace("FlächeHinten");
                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(cutOut.XP, -(cutOut.ZP - _HPt / 2)), transientGeometry.CreatePoint2d(cutOut.XP + (cutOut.XS / 2) + _TE, -(cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2)));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                        }
                        else if (cutOut.YP - cutOut.YS / 2 >= -_LP / 2)
                        {
                            face = _FindNamedFace("FlächeVorne");
                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(cutOut.XP, (cutOut.ZP - _HPt / 2)), transientGeometry.CreatePoint2d(cutOut.XP + (cutOut.XS / 2) + _TE, (cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2)));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                        }
                        else
                        {
                            MessageBox.Show("CutOut (" + cutOut.Name + ") konnte keiner Fläche zugeordnet werden!");
                            ZählerCutOuts--;
                        }
                    }
                    else if (!cutOut.Top && !_Top)
                    {
                        MessageBox.Show("CutOutUnten");
                        if (cutOut.XP + cutOut.XS / 2 >= _BP / 2)
                        {
                            face = _FindNamedFace("FlächeRechts");

                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(cutOut.ZP - _HPt / 2, cutOut.YP), transientGeometry.CreatePoint2d(cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2, cutOut.YP + (cutOut.YS / 2) + _TE));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                            //extrudeFeature.Name = "CutOut" + ZählerCutOuts.ToString();
                        }
                        else if (cutOut.XP - cutOut.XS / 2 <= -_BP / 2)
                        {
                            face = _FindNamedFace("FlächeLinks");
                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(-(cutOut.ZP - _HPt / 2), cutOut.YP), transientGeometry.CreatePoint2d(-(cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2), cutOut.YP + (cutOut.YS / 2) + _TE));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                        }
                        else if (cutOut.YP + cutOut.YS / 2 <= _LP / 2)
                        {
                            face = _FindNamedFace("FlächeVorne");
                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(cutOut.XP, -(cutOut.ZP - _HPt / 2)), transientGeometry.CreatePoint2d(cutOut.XP + (cutOut.XS / 2) + _TE, -(cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2)));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                        }
                        else if (cutOut.YP - cutOut.YS / 2 >= -_LP / 2)
                        {
                            face = _FindNamedFace("FlächeHinten");
                            PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                            sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                            sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                            TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                            sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(cutOut.XP, (cutOut.ZP - _HPt / 2)), transientGeometry.CreatePoint2d(cutOut.XP + (cutOut.XS / 2) + _TE, (cutOut.ZP + (cutOut.ZS / 2) + _TE - _HPt / 2)));

                            Profile profil = sketch.Profiles.AddForSolid(true);
                            ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                            extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                            ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                        }
                        else
                        {
                            MessageBox.Show("CutOut (" + cutOut.Name + ") konnte keiner Fläche zugeordnet werden!");
                            ZählerCutOuts--;
                        }
                    }
                }
                else
                {
                    if (cutOut.Top && _Top)
                    {
                        face = _FindNamedFace("FlächeOben");

                        PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                        sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                        sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                        TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                        sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(-cutOut.XP, cutOut.YP), transientGeometry.CreatePoint2d(-cutOut.XP + (cutOut.XS / 2) + _TE, cutOut.YP + (cutOut.YS / 2) + _TE));

                        Profile profil = sketch.Profiles.AddForSolid(true);
                        ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                        extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                        ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                    }
                    else if (!cutOut.Top && !_Top)
                    {
                        face = _FindNamedFace("FlächeOben");

                        PlanarSketch sketch = _partComponentDefinition.Sketches.Add(face, false);
                        sketch.OriginPoint = _partComponentDefinition.WorkPoints[1];
                        sketch.Name = "CutOut" + ZählerCutOuts.ToString();

                        TransientGeometry transientGeometry = _inventorApp.TransientGeometry;

                        sketch.SketchLines.AddAsTwoPointCenteredRectangle(transientGeometry.CreatePoint2d(-cutOut.XP, -cutOut.YP), transientGeometry.CreatePoint2d(-cutOut.XP + (cutOut.XS / 2) + _TE, -cutOut.YP + (cutOut.YS / 2) + _TE));

                        Profile profil = sketch.Profiles.AddForSolid(true);
                        ExtrudeDefinition extrudeDefinition = _partComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(profil, PartFeatureOperationEnum.kCutOperation);
                        extrudeDefinition.SetDistanceExtent(_DW * 4, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);
                        ExtrudeFeature extrudeFeature = _partComponentDefinition.Features.ExtrudeFeatures.Add(extrudeDefinition);
                    }
                }
                ZählerCutOuts++;
            }

            _status.Name = "Done";
            _status.OnProgess();
        }

        public void ExportToStep(string path)
        {
            _partDocument = _inventorApp.Documents.Open(_path, true) as PartDocument;

            TranslatorAddIn stepTranslator = (TranslatorAddIn)_inventorApp.ApplicationAddIns.ItemById["{90AF7F40-0C01-11D5-8E83-0010B541CD80}"];

            if (stepTranslator == null)
            {
                MessageBox.Show("STEP translator not acvessable");
                return;
            }

            TranslationContext translationContext = _inventorApp.TransientObjects.CreateTranslationContext();
            NameValueMap options = _inventorApp.TransientObjects.CreateNameValueMap();

            if (stepTranslator.HasSaveCopyAsOptions[_partDocument, translationContext, options])
            {
                options.Value["ApplicationProtocolType"] = 2;
                options.Value["Author"] = "GehäuseGenerator";

                translationContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                DataMedium dataMedium = _inventorApp.TransientObjects.CreateDataMedium();
                dataMedium.FileName = path;

                stepTranslator.SaveCopyAs(_partDocument, translationContext, options, dataMedium);
            }

            _partDocument.Close();
        }

        public void ExportToStl(string path)
        {
            _partDocument = _inventorApp.Documents.Open(_path, true) as PartDocument;

            TranslatorAddIn stlTranslator = (TranslatorAddIn)_inventorApp.ApplicationAddIns.ItemById["{533E9A98-FC3B-11D4-8E7E-0010B541CD80}"];

            if (stlTranslator == null)
            {
                MessageBox.Show(".stl translator not acvessable");
                return;
            }

            TranslationContext translationContext = _inventorApp.TransientObjects.CreateTranslationContext();
            NameValueMap options = _inventorApp.TransientObjects.CreateNameValueMap();

            if (stlTranslator.HasSaveCopyAsOptions[_partDocument, translationContext, options])
            {
                options.Value["ExportUnits"] = 4;
                options.Value["Resolution"] = 0;
                options.Value["AllowMoveMeshNode"] = false;
                options.Value["SurfaceDeviation"] = 0.0004;
                options.Value["NormalDeviation"] = 22;
                options.Value["MaxEdgeLength"] = 1.5;
                options.Value["AspectRatio"] = 21.5;
                options.Value["ExportFileStructure"] = 0;
                options.Value["OutputFileType"] = 0;
                options.Value["ExportColor"] = false;

                translationContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                DataMedium dataMedium = _inventorApp.TransientObjects.CreateDataMedium();
                dataMedium.FileName = path;

                stlTranslator.SaveCopyAs(_partDocument, translationContext, options, dataMedium);
            }

            _partDocument.Close();
        }

        public void ExportToObj(string path)
        {
            _partDocument = _inventorApp.Documents.Open(_path, true) as PartDocument;

            TranslatorAddIn objTranslator = null;
            foreach (ApplicationAddIn addIn in _inventorApp.ApplicationAddIns)
            {
                if (addIn.DisplayName.Contains("OBJ ex"))
                {
                    objTranslator = (TranslatorAddIn)addIn;
                }
            }

            TranslationContext translationContext = _inventorApp.TransientObjects.CreateTranslationContext();
            NameValueMap options = _inventorApp.TransientObjects.CreateNameValueMap();


            if (objTranslator.HasSaveCopyAsOptions[_partDocument, translationContext, options])
            {
                options.Value["ExportUnits"] = 0;
                options.Value["Resolution"] = 1;
                options.Value["SurfaceDeviation"] = 16;
                options.Value["NormalDeviation"] = 1500;
                options.Value["MaxEdgeLength"] = 100000;
                options.Value["AspectRatio"] = 2150;
                options.Value["ExportFileStructure"] = 0;

                translationContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                DataMedium dataMedium = _inventorApp.TransientObjects.CreateDataMedium();
                dataMedium.FileName = path;

                objTranslator.SaveCopyAs(_partDocument, translationContext, options, dataMedium);
            }

            _partDocument.Close(true);

        }



        private Face _FindNamedFace(string Name)
        {
            foreach (SurfaceBody body in _partComponentDefinition.SurfaceBodies)
            {
                foreach (Face face in body.Faces)
                {
                    if (face.AttributeSets.NameIsUsed["iLogicEntityNameSet"])
                    {
                        AttributeSet attributeSet = face.AttributeSets["iLogicEntityNameSet"];
                        foreach (Inventor.Attribute attribute in attributeSet)
                        {
                            if (attribute.Value == Name)
                            {
                                //MessageBox.Show("Fläche Gefunden! \n" + Name);

                                return (face);

                            }
                        }
                    }
                }
            }
            MessageBox.Show("Fläche nicht Gefunden!");
            return null;
        }

        private Inventor.Application _inventorApp;
        private PartDocument _partDocument;
        private PartComponentDefinition _partComponentDefinition;

        //Parameter [cm]
        private double
            _DW, //Wanddicke
            _TM, //Toleranz zu mechanischen Teilen
            _TE, //Toleranz zu Elektrischen Teilen (zb. Buchsen)
            _BP, //Breite Platine (X Richtung)
            _LP, //Länge Platine  (Y Richtuung)
            _DL, //Durchmesser Loch
            _REP, //Radius Ecke Platine
            _HPt, //Höhe Board
            _HPb, //Höhe Bauteile Platine
            _RR, //Rundungs Radius
            _SKD, //SchraubenKopfDurchmesser
            _SKH; //SchraubenKopfHöhe
        private bool _Top;
        private Parameters _parameters;

        private List<CutOut> _CutOuts = new List<CutOut>();

        private Status _status;

        private string _path;
    }

}
    

