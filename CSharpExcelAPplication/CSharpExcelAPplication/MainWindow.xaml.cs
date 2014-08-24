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

        public String openFileDialogBox()
        {
            String fileName; 
            //create an instance of the open file dialog box.
            Microsoft.Win32.OpenFileDialog openfileDialog1 = new Microsoft.Win32.OpenFileDialog();

            bool? userClickedOK = openfileDialog1.ShowDialog();

            //process input if the user clicked ok.
            if (userClickedOK == true)
            {
                fileName = openfileDialog1.FileName;
                return fileName;
            }

            return "No File Selected";
        }

        private void heatmapLayoutbutton_Click(object sender, RoutedEventArgs e)
        {
            String fileName;
            fileName = openFileDialogBox();
            heatmapLayoutTextbox.Text = fileName;
        }
    }
}
