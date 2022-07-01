#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDeleteBackups : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            // set variables
            int counter = 0;
            string logPath = "";

            // create list for log file
            List<string> deletedFileLog = new List<string>();
            deletedFileLog.Add("The following backup files have been deleted:");

            FolderBrowserDialog selectFolder = new FolderBrowserDialog();
            selectFolder.ShowNewFolderButton = false;

            // open folder dialog and only run code if a folder is selected
            if(selectFolder.ShowDialog() == DialogResult.OK)
            {
                // get the selected folder path
                string directory = selectFolder.SelectedPath;

                // get all files from selected folder
                string[] files = Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories);

                // loop through files
                foreach(string file in files)
                {
                    // check if the file is a Revit file
                    if(Path.GetExtension(file) == ".rvt" || Path.GetExtension(file) == ".rfa")
                    {
                        // get the last 9 characters of filename to check if backup
                        string checkString = file.Substring(file.Length - 9, 9);

                        if(checkString.Contains(".0") == true)
                        {
                            // add filename to list
                            deletedFileLog.Add(file);

                            // delete file
                            File.Delete(file);

                            // increment counter
                            counter++;
                        }
                    }
                }

                // output log file
                if (counter > 0)
                {
                    logPath = WriteListToTxt(deletedFileLog, directory);
                }
            }

            // alert user
            TaskDialog td = new TaskDialog("Complete");
            td.MainInstruction = "Deleted " + counter.ToString() + " backup files.";
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Click to view log file");
            td.CommonButtons = TaskDialogCommonButtons.Ok;

            TaskDialogResult result = td.Show();

            if(result == TaskDialogResult.CommandLink1)
            {
                Process.Start(logPath);
            }

            return Result.Succeeded;
        }

        internal string WriteListToTxt(List<string> stringList, string filePath)
        {
            string fileName = "_Delete Backup Files.txt";
            string fullPath = filePath + @"\" + fileName;
            
            File.WriteAllLines(fullPath, stringList);

            return fullPath;
        }
    }
}
