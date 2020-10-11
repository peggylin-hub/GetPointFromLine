using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Excel = Microsoft.Office.Interop.Excel;

namespace CreateCogoPointsFromExcel
{
    public class Commands// : IExtensionApplication
    {
        [CommandMethod("CCPFE_XYZ")] //CCPFE = CreateCogoPointsFromExcel

        static public void CCPFE_XYZ()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            CivilDocument Cdoc = CivilApplication.ActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            double wallThk = 0.3;
            double baseThk = 0.3;

            string[,] data = LoadExcel(getFolderPath());
            for (long i = 1; i < data.GetLength(0); i++)
            {

                double x = Convert.ToDouble(data[i, 4]);
                double y = Convert.ToDouble(data[i, 5]);
                double z = Convert.ToDouble(data[i, 6]);

                CogoPointCollection cogoPoints = Cdoc.CogoPoints;
                ObjectId pointId = cogoPoints.Add(new Point3d(x, y, z), true);

                //reverse the angle direction, this input is from 12d
                double rotationAng = -degreeToRad(Convert.ToDouble(data[i, 11]));
                cogoPoints.SetMarkerRotation(pointId, rotationAng);

                //check, this needed to be unique.
                string name = data[i, 2];
                cogoPoints.SetPointName(pointId, name);

                string rawName = data[i, 1];
                cogoPoints.SetRawDescription(pointId, rawName);

                //double scaler = 1.1;
                //cogoPoints.SetScaleXY(pointId, scaler);
            }

            ////start a transaction 
            //using (Transaction trans = db.TransactionManager.StartTransaction())
            //{

            //    // All points in a document are held in a CogoPointCollection object 
            //    // We can access CogoPointCollection through the CivilDocument.CogoPoints property

            //    CogoPointCollection cogoPoints = CivilApplication.ActiveDocument.CogoPoints;

            //    // Adds a new CogoPoint at the given location with the specified description information
            //    ObjectId pointId = cogoPoints.Add(location, "Survey Point");
            //    CogoPoint cogoPoint = pointId.GetObject(OpenMode.ForWrite) as CogoPoint;

            //    // Set Some Properties
            //    cogoPoint.PointName = "Survey_Base_Point";
            //    cogoPoint.RawDescription = "This is Survey Base Point";

            //    trans.Commit();
            //}
        }

        private static string[,] LoadExcel(string filename)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //set the range
            range = xlWorkSheet.UsedRange;
            //pass data into a object array
            object[,] values = range.Value;

            string[,] fileData = ConvertToString(values);//convert object array to string array

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            return fileData;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (System.Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        //convert object array to string array
        private static string[,] ConvertToString(object[,] list)
        {
            string[,] fileData = new string[list.GetUpperBound(0), list.GetUpperBound(1)];

            for (int i = 1; i <= list.GetUpperBound(0); i++)
            {
                for (int k = 1; k <= list.GetUpperBound(1); k++)
                {
                    //ToString for possibly null object
                    string test = (list[i, k] ?? "").ToString();
                    fileData[i - 1, k - 1] = test;
                }
            }

            return fileData;
        }

        private static string getFolderPath()
        {
            //abtain save file path
            OpenFileDialog ODialog = new OpenFileDialog();
            //FolderBrowserDialog ODialog = new FolderBrowserDialog();
            string fileFullame = "";
            if (ODialog.ShowDialog() == DialogResult.OK)
            {
                //folderPath = System.IO.Path.GetDirectoryName(ODialog.FileName);
                //folderPath = System.IO.Path.GetDirectoryName(ODialog.SelectedPath);
                fileFullame = ODialog.FileName;
            }

            return fileFullame;
        }

        private static double degreeToRad(double degree)
        {
            double rad = degree * 2 * Math.PI / 360;
            return rad;
        }

    }

}
