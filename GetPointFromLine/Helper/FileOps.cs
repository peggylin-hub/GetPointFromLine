using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TunnalCal.Helper
{
    class FileOps
    {
        public static string SelectFile()
        {
            //abtain save file path
            OpenFileDialog ODialog = new OpenFileDialog();
            //FolderBrowserDialog ODialog = new FolderBrowserDialog();
            string fileFullame = "";
            if (ODialog.ShowDialog() == DialogResult.OK)
            {
                fileFullame = ODialog.FileName;
            }

            return fileFullame;
        }

        public static string[] selectFiles(string title, string filter, string dialog)
        {
            //use autodesk windows rather than windows form
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog(title, null, filter, dialog, Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            ofd.ShowDialog();
            string[] result = ofd.GetFilenames();

            return result;
        }
    }
}
