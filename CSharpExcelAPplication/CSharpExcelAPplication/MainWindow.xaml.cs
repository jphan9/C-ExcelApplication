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
        public MainWindow()
        {
            InitializeComponent();
        }

        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(1);
        }

        public string openFileDialogBox()
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

        public List<string[]> readInFile(string fileName)
        {
            List<string[]> listA = new List<string[]>();
            List<string> listB = new List<string>();
            string[] row;

            StreamReader sr = new StreamReader(fileName);
            //string data = sr.ReadLine();
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();
                row = line.Split(',');

                listA.Add(row);
                MessageBox.Show(line);

                return listA;
            }
            return listA;
        }

        public void ExcelApplication(List<string[]> listA)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            // Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = true;

            // Get a new workbook.
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(System.Reflection.Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            // Create an Array to multiple values at once.
            //string[,] values = new string[10,10];
            
            int rowIndex = 1;
            oSheet.Cells[rowIndex, "A"] = listA[0];
        }

        private void heatmapLayoutbutton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = openFileDialogBox();
            heatmapLayoutTextbox.Text = fileName;
        }

        private void dataLowButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = openFileDialogBox();
            dataLowTextbox.Text = fileName;
        }

        private void dataMedButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = openFileDialogBox();
            dataMedTextbox.Text = fileName;
        }

        private void dataHighButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = openFileDialogBox();
            dataHighTextbox.Text = fileName;
        }

        private void dataUltraButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = openFileDialogBox();
            dataUltraTextbox.Text = fileName;
        }

        private void goButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName = dataLowTextbox.Text;
            List<string[]> listA = readInFile(fileName);
            ExcelApplication(listA);
        }

    }
}
