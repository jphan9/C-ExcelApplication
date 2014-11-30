using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpExcelAPplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelApplication eApp = new ExcelApplication();
        readFiles readFile = new readFiles();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(1);
        }

        /*public string openFileDialogBox()
        {
            string fileName; 
            //create an instance of the open file dialog box.
            Microsoft.Win32.OpenFileDialog openfileDialog1 = new Microsoft.Win32.OpenFileDialog();

            // Set the title of the dialog box
            openfileDialog1.Title = "Choose your file ";

            // Set filter options and filter index.
            openfileDialog1.Filter = "CSV File|*.csv|Excel Workbook|*.xlsx";

            bool? userClickedOK = openfileDialog1.ShowDialog();

            //process input if the user clicked ok.
            if (userClickedOK == true)
            {
                fileName = openfileDialog1.FileName;
                return fileName;
            }

            return "No File Selected";
        }

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

        /*public void ExcelApplication(string fileName)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            List<string> listOfDataA;
            List<string> listOfDataB;
            List<string> listOfDataC;

            int cIndex1 = 1, cIndex2;
            int index = 1;

            // Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            // Get a new workbook.
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(System.Reflection.Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            listOfDataA = readInFile(fileName, 0);
            listOfDataB = readInFile(fileName, 1);
            listOfDataC = readInFile(fileName, 2);

            int rowIndexA = 1;
            int rowIndexB = 1;
            int rowIndexC = 1;

            for (int i = 0; i < listOfDataA.Count; i++)
            {
                oSheet.Cells[rowIndexA, "A"] = listOfDataA[i];
                rowIndexA++;
            }

            for (int i = 0; i < listOfDataB.Count; i++)
            {
                oSheet.Cells[rowIndexB, "B"] = listOfDataB[i];
                rowIndexB++;
            }

            for (int i = 0; i < listOfDataC.Count; i++)
            {
                oSheet.Cells[rowIndexC, "C"] = listOfDataC[i];
                rowIndexC++;
            }

            //gets the total number of cells used in Column C
            int columnTotalC = oSheet.UsedRange.Columns["C:C", Type.Missing].rows.count;
            
            MessageBox.Show(Convert.ToString(columnTotalC));

            for (int i = 0; i < columnTotalC/4; i++)
            {
                cIndex2 = cIndex1 + 3;
                oSheet.Range["G" + Convert.ToString(index)].Formula = "=Average(C" + Convert.ToString(cIndex1) + ":C" + Convert.ToString(cIndex2);
                cIndex1 = cIndex2 + 1;
                index++;
            }
        }*/

        private void heatmapLayoutbutton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("workbook");
            LayoutTextbox.Text = fileName;
        }

        private void dataLowButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataLowTextbox.Text = fileName;
        }

        private void dataMedButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataMedTextbox.Text = fileName;
        }

        private void dataHighButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataHighTextbox.Text = fileName;
        }

        private void dataUltraButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataUltraTextbox.Text = fileName;
        }

        private void goButton_Click(object sender, RoutedEventArgs e)
        {
            string openFileName = LayoutTextbox.Text;
            string fileName = dataLowTextbox.Text;
            string fileName1 = dataMedTextbox.Text;
            string fileName2 = dataHighTextbox.Text;
            string fileName3 = dataUltraTextbox.Text;
            //List<string> listA = readInFile(fileName);
            //eApp.writeToExcelLow(fileName1, fileName);
            MessageBox.Show(fileName);
            eApp.openExcel(openFileName, fileName, fileName1, fileName2, fileName3);
            //eApp.writeToExcelMed(fileName2);
        }

    }
}
