using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System.IO;

using OfficeOpenXml;

namespace TunnalCal.Helper
{
    class Excel
    {
        //export data to excel
        public static void createCSV(List<string> data, string filepath)
        {
            StringBuilder output = new StringBuilder(1000);
            for (int r = 0; r < data.Count; r++)
            {
                output.Append(data[r]);
                output.Append(Environment.NewLine);
            }

            string text = output.ToString();
            DateTime today = DateTime.Now;
            string fileNameStr = $"_PointData_{today.ToString("yy_MM_dd-hh-mm")}.csv";
            string fileName = filepath + fileNameStr;
            System.IO.File.WriteAllText(fileName, text);
        }

        //export data to excel
        public static void createCSV(Dictionary<string, List<string>> data, string filepath)
        {
            StringBuilder output = new StringBuilder(1000);
            foreach(KeyValuePair<string, List<string>> keyValuePair in data)
            {
                string fn = keyValuePair.Key;
                List<string> layers = keyValuePair.Value;
                foreach(string l in layers)
                {
                    output.Append(fn + "," + l + Environment.NewLine);
                }
            }

            string text = output.ToString();
            DateTime today = DateTime.Now;
            string fileNameStr = $"_PointData_{today.ToString("yy_MM_dd-hh-mm")}.csv";
            string fileName = filepath + fileNameStr;
            System.IO.File.WriteAllText(fileName, text);
        }

        /// <summary>
        /// Create Point3d list by importing from excel
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="AssetExcelFile"></param>
        /// <param name="ExcelShtName"></param>
        /// <returns></returns>
        public static List<Point3d> getAllpoint(Document doc, string AssetExcelFile, string ExcelShtName)
        {
            List<Point3d> pts = new List<Point3d>();

            Dictionary<string, Point3d> data = new Dictionary<string, Point3d>();

            FileInfo fi = new FileInfo(AssetExcelFile);

            try
            {
                using (ExcelPackage pk = new ExcelPackage(fi))
                {
                    ExcelWorksheet ws = null; //pk.Workbook.Worksheets[1];

                    foreach (ExcelWorksheet sht in pk.Workbook.Worksheets)
                    {
                        if (sht.Name == ExcelShtName)
                            ws = sht;
                    }

                    int rowCount = ws.Dimension.End.Row;

                    for (int row = 2; row <= rowCount; row++)//read from the second row, instead of first row
                    {
                        double x = 0;
                        double y = 0;
                        double z = 0;

                        if (ws.Cells[row, 1].Value != null)
                            x = Convert.ToDouble(ws.Cells[row, 2].Value.ToString().Trim());
                        if (ws.Cells[row, 2].Value != null)
                            y = Convert.ToDouble(ws.Cells[row, 3].Value.ToString().Trim());
                        if (ws.Cells[row, 3].Value != null)
                            z = Convert.ToDouble(ws.Cells[row, 4].Value.ToString().Trim());

                        string text = $"{x},{y},{z}";
                        if (!data.ContainsKey(text))
                        {
                            Point3d pt = new Point3d(x, y, z);
                            data.Add(text, pt);
                            pts.Add(pt);
                        }
                    }
                    if (pts.Count() > 0)
                        return pts;
                    else
                        return null;
                }
            }
            catch 
            {
                return null;
            }

        }
    }
}
