using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CSharpExcelAPplication
{
    class readFiles
    {
        // Function to show the open file dialog box.
        public string openFileDialogBox(string extension)
        {
            string fileName;
            //create an instance of the open file dialog box.
            Microsoft.Win32.OpenFileDialog openfileDialog1 = new Microsoft.Win32.OpenFileDialog();

            // Set the title of the dialog box
            openfileDialog1.Title = "Choose your file ";

            // Set filter options and filter index.
            if (extension == "workbook")
            {
                openfileDialog1.Filter = "Excel Workbook|*.xlsx";
            }
            else if (extension == "csv")
            {
                openfileDialog1.Filter = "CSV File|*.csv";
            }
            bool? userClickedOK = openfileDialog1.ShowDialog();

            //process input if the user clicked ok.
            if (userClickedOK == true)
            {
                fileName = openfileDialog1.FileName;
                return fileName;
            }

            return "No File Selected";
        }

        // Function to read in the file chosen.
        public List<string> readInFile(string fileName, int column)
        {
            List<string> listA = new List<string>();
            List<string> listB = new List<string>();
            List<string> listC = new List<string>();
            //var row;

            StreamReader sr = new StreamReader(fileName);
            //string data = sr.ReadLine();
            while (!sr.EndOfStream)
            {
                //string line = sr.ReadLine();
                var row = sr.ReadLine().Split(',');

                listA.Add(row[0]);
                listB.Add(row[1]);
                listC.Add(row[2]);

            }
            if (column == 0)
            {
                return listA;
            }
            else if (column == 1)
            {
                return listB;
            }
            else if (column == 2)
            {
                return listC;
            }
            else
                return listA;
        }
    }
}
