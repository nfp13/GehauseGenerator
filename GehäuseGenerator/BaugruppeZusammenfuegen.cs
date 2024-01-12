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

        public BaugruppeZusammenfuegen(Inventor.Application inventorApp, string filePath)
        {
            _inventorApp = inventorApp;

            //Baugruppe erstellen
            _assemblyDocument = _inventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, _inventorApp.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject)) as AssemblyDocument;

        }

        public void PlatineHinzufügen(string filePath, Matrix positionMatrix)
        {
            //Platzieren
            ComponentOccurrence platine = _assemblyDocument.ComponentDefinition.Occurrences.Add(filePath, positionMatrix);
        }

        public void OberesGehäuseHinzufügen(string filePath, double zOffset)
        {
            Matrix positionMatrix = _inventorApp.TransientGeometry.CreateMatrix();

            //Schieben
            Vector trans = _inventorApp.TransientGeometry.CreateVector(0, 0, (zOffset / 2.0));
            positionMatrix.SetTranslation(trans);

            //Platzieren
            ComponentOccurrence oberesGehäuse = _assemblyDocument.ComponentDefinition.Occurrences.Add(filePath, positionMatrix);
        }

        public void UnteresGehäuseHinzufügen(string filePath, double zOffset)
        {
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
        }

        public void SchraubenHinzufügen(double diameter, double breite, double länge, double radius, double höheSchrauben)
        {
            Matrix positionMatrix = _inventorApp.TransientGeometry.CreateMatrix();

            //Familie holen
            AssemblyComponentDefinition asmDef = _assemblyDocument.ComponentDefinition;
            ContentTreeViewNode hexHeadNode = _inventorApp.ContentCenter.TreeViewTopNode.ChildNodes["Verbindungselemente"].ChildNodes["Schrauben"].ChildNodes["Rundkopf"];
            ContentFamily family = null;
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

                ComponentOccurrence s2 = asmDef.Occurrences.Add(memberFilename, positionMatrix);

                //Schieben
                Vector trans3 = _inventorApp.TransientGeometry.CreateVector(-((breite / 2) - radius), ((länge / 2) - radius), -höheSchrauben);
                positionMatrix.SetTranslation(trans3);

                ComponentOccurrence s3 = asmDef.Occurrences.Add(memberFilename, positionMatrix);

                //Schieben
                Vector trans4 = _inventorApp.TransientGeometry.CreateVector(-((breite / 2) - radius), -((länge / 2) - radius), -höheSchrauben);
                positionMatrix.SetTranslation(trans4);

                ComponentOccurrence s4 = asmDef.Occurrences.Add(memberFilename, positionMatrix);
            }
        }

        public void Save(string BaugruppePath)
        {
            _assemblyDocument.Update();
            _assemblyDocument.SaveAs(BaugruppePath, true);
            _assemblyDocument.Close(true);
        }

        public void packAndGo(string pathBaugruppe, string pathToSave)
        {
            string desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
            PackAndGoComponent packAndGoComp = new PackAndGoComponent();
            Save(pathBaugruppe);
            PackAndGo packAndGo = packAndGoComp.CreatePackAndGo(pathBaugruppe, pathToSave);

            string[] refFiles = new string[] { };
            object refMissFiles = new object();

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
        }

        private Inventor.Application _inventorApp;
        private AssemblyDocument _assemblyDocument;

    }
}
    

