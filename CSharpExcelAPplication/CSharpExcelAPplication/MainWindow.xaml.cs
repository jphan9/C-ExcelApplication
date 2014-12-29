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

        // Function to exit the program when clicking on the Exit button.
        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(1);
        }

        // Function to grab the name of the heatmap file.
        private void heatmapLayoutbutton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("workbook");
            LayoutTextbox.Text = fileName;
        }

        // Function to grab the name of the data low file.
        private void dataLowButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataLowTextbox.Text = fileName;
        }

        // Function to grab the name of the data med file.
        private void dataMedButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataMedTextbox.Text = fileName;
        }

        // Function to grab the name of the data high file.
        private void dataHighButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataHighTextbox.Text = fileName;
        }

        // Function to grab the name of the data ultra file.
        private void dataUltraButton_Click(object sender, RoutedEventArgs e)
        {
            string fileName;
            fileName = readFile.openFileDialogBox("csv");
            dataUltraTextbox.Text = fileName;
        }

        // Function to open all the files gathered and start computing the openExcel Function.
        private void goButton_Click(object sender, RoutedEventArgs e)
        {
            string openFileName = LayoutTextbox.Text;
            string fileName = dataLowTextbox.Text;
            string fileName1 = dataMedTextbox.Text;
            string fileName2 = dataHighTextbox.Text;
            string fileName3 = dataUltraTextbox.Text;
            eApp.openExcel(openFileName, fileName, fileName1, fileName2, fileName3);
        }

    }
}
