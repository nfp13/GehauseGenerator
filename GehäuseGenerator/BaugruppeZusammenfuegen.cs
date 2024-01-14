using System;
using System.Windows.Forms;
using Inventor;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using PackAndGoLib;


namespace GehäuseGenerator
{
    public class BaugruppeZusammenfuegen
    {

        public BaugruppeZusammenfuegen(Inventor.Application inventorApp, string filePath, Status status)
        {
            _inventorApp = inventorApp;
            _status = status;

            _status.Name = "Creating Assembly";
            _status.OnProgess();
            //Baugruppe erstellen
            _assemblyDocument = _inventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, _inventorApp.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject)) as AssemblyDocument;

            _status.Name = "Done";
            _status.OnProgess();
        }

        public void PlatineHinzufügen(string filePath, Matrix positionMatrix)
        {
            _status.Name = "Adding PCB";
            _status.Progress = 0;
            _status.OnProgess();

            //Platzieren
            ComponentOccurrence platine = _assemblyDocument.ComponentDefinition.Occurrences.Add(filePath, positionMatrix);

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        public void OberesGehäuseHinzufügen(string filePath, double zOffset)
        {
            _status.Name = "Adding Parts";
            _status.Progress = 0;
            _status.OnProgess();

            Matrix positionMatrix = _inventorApp.TransientGeometry.CreateMatrix();

            //Schieben
            Vector trans = _inventorApp.TransientGeometry.CreateVector(0, 0, (zOffset / 2.0));
            positionMatrix.SetTranslation(trans);

            //Platzieren
            ComponentOccurrence oberesGehäuse = _assemblyDocument.ComponentDefinition.Occurrences.Add(filePath, positionMatrix);

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        public void UnteresGehäuseHinzufügen(string filePath, double zOffset)
        {
            _status.Name = "Adding Parts";
            _status.Progress = 0;
            _status.OnProgess();

            Matrix positionMatrix = _inventorApp.TransientGeometry.CreateMatrix();

            //Drehen
            double angle = Math.PI;
            positionMatrix.GetCoordinateSystem(out Inventor.Point Origin, out Vector XAxis, out Vector YAxis, out Vector ZAxis);
            positionMatrix.SetToRotation(angle, XAxis, Origin);

            //Schieben
            Vector trans = _inventorApp.TransientGeometry.CreateVector(0, 0, -(zOffset / 2.0));
            positionMatrix.SetTranslation(trans);

            //Platzieren
            ComponentOccurrence unteresGehäuse = _assemblyDocument.ComponentDefinition.Occurrences.Add(filePath, positionMatrix);

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        public void SchraubenHinzufügen(double diameter, double breite, double länge, double radius, double höheSchrauben)
        {
            _status.Name = "Adding Screws";
            _status.Progress = 0;
            _status.OnProgess();

            Matrix positionMatrix = _inventorApp.TransientGeometry.CreateMatrix();

            //Familie holen
            AssemblyComponentDefinition asmDef = _assemblyDocument.ComponentDefinition;
            ContentTreeViewNode hexHeadNode = _inventorApp.ContentCenter.TreeViewTopNode.ChildNodes["Verbindungselemente"].ChildNodes["Schrauben"].ChildNodes["Rundkopf"];
            ContentFamily family = null;

            _status.Progress = 25;
            _status.OnProgess();

            foreach (ContentFamily checkFamily in hexHeadNode.Families)
            {
                if (checkFamily.DisplayName == "ISO 7045 Z")
                {
                    family = checkFamily;
                    break;
                }
            }
            if (family != null)
            {
                int zeile = 1;
                switch (diameter)
                {
                    case 1.6:
                        zeile = 1;
                        break;
                    case 2:
                        zeile = 10;
                        break;
                    case 2.5:
                        zeile = 20;
                        break;
                    case 3:
                        zeile = 31;
                        break;
                    case 4:
                        zeile = 53;
                        break;
                    case 5:
                        zeile = 65;
                        break;
                    case 6:
                        zeile = 77;
                        break;
                    case 8:
                        zeile = 91;
                        break;
                    case 10:
                        zeile = 104;
                        break;

                }

                _status.Progress = 50;
                _status.OnProgess();

                MemberManagerErrorsEnum failureReason;
                string failureMessage;
                string memberFilename = family.CreateMember(zeile, out failureReason, out failureMessage, ContentMemberRefreshEnum.kRefreshOutOfDateParts);


                //Drehen
                double angle = -(Math.PI / 2);
                positionMatrix.GetCoordinateSystem(out Inventor.Point Origin, out Vector XAxis, out Vector YAxis, out Vector ZAxis);
                positionMatrix.SetToRotation(angle, YAxis, Origin);

                //Schieben
                Vector trans1 = _inventorApp.TransientGeometry.CreateVector(((breite / 2) - radius), ((länge / 2) - radius), -höheSchrauben);
                positionMatrix.SetTranslation(trans1);

                ComponentOccurrence s1 = asmDef.Occurrences.Add(memberFilename, positionMatrix);

                //Schieben
                Vector trans2 = _inventorApp.TransientGeometry.CreateVector(((breite / 2) - radius), -((länge / 2) - radius), -höheSchrauben);
                positionMatrix.SetTranslation(trans2);

                _status.Progress = 75;
                _status.OnProgess();

                ComponentOccurrence s2 = asmDef.Occurrences.Add(memberFilename, positionMatrix);

                //Schieben
                Vector trans3 = _inventorApp.TransientGeometry.CreateVector(-((breite / 2) - radius), ((länge / 2) - radius), -höheSchrauben);
                positionMatrix.SetTranslation(trans3);

                ComponentOccurrence s3 = asmDef.Occurrences.Add(memberFilename, positionMatrix);

                //Schieben
                Vector trans4 = _inventorApp.TransientGeometry.CreateVector(-((breite / 2) - radius), -((länge / 2) - radius), -höheSchrauben);
                positionMatrix.SetTranslation(trans4);

                ComponentOccurrence s4 = asmDef.Occurrences.Add(memberFilename, positionMatrix);

                _status.Progress = 100;
                _status.Name = "Done";
                _status.OnProgess();
            }
        }

        public void Save(string BaugruppePath)
        {
            _status.Name = "Pack and Go";
            _status.Progress = 0;
            _status.OnProgess();

            _assemblyDocument.Update();
            _status.Progress = 50;
            _status.OnProgess();
            _assemblyDocument.SaveAs(BaugruppePath, true);
            _assemblyDocument.Close(true);

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        public void packAndGo(string pathBaugruppe, string pathToSave)
        {
            _status.Name = "Pack and Go";
            _status.Progress = 0;
            _status.OnProgess();

            string desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
            PackAndGoComponent packAndGoComp = new PackAndGoComponent();
            Save(pathBaugruppe);
            PackAndGo packAndGo = packAndGoComp.CreatePackAndGo(pathBaugruppe, pathToSave);

            string[] refFiles = new string[] { };
            object refMissFiles = new object();

            _status.Progress = 50;
            _status.OnProgess();

            // Set the options
            packAndGo.SkipLibraries = false;
            packAndGo.SkipStyles = true;
            packAndGo.SkipTemplates = true;
            packAndGo.CollectWorkgroups = false;
            packAndGo.KeepFolderHierarchy = true;
            packAndGo.IncludeLinkedFiles = true;

            // Get the referenced files
            packAndGo.SearchForReferencedFiles(out refFiles, out refMissFiles);

            // Add the referenced files for package
            packAndGo.AddFilesToPackage(refFiles);

            // Start the pack and go to create the package
            packAndGo.CreatePackage();

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        public void SavePictureAsOben(string Path)
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

        public void SavePictureAsUnten(string Path)
        {
            _status.Name = "Taking Screenshot";
            _status.Progress = 0;
            _status.OnProgess();

            Inventor.View view = _assemblyDocument.Views[1];
            Inventor.Camera camera = view.Camera;
            camera.Perspective = false;

            _status.Progress = 75;
            _status.OnProgess();

            camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoBottomRightViewOrientation;
            camera.Fit();
            camera.Apply();
            view.Update();
            camera.SaveAsBitmap(Path, 1080, 1080);

            _status.Progress = 100;
            _status.Name = "Done";
            _status.OnProgess();
        }

        private Inventor.Application _inventorApp;
        private AssemblyDocument _assemblyDocument;
        private Status _status;

    }
}
    

