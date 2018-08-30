using Microsoft.Win32;
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

namespace RCSVB
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

        private void Realms_CSV_File_Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            openFileDialog.InitialDirectory = @"C:\";
            openFileDialog.Title = "Please select a CSV file exported from Realm";
            if (openFileDialog.ShowDialog() == true)
            {
                Realms_CSV_File_TextBox.Text = openFileDialog.FileName;
            }
        }

        private void Output_XLSX_File_Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            if (saveFileDialog.ShowDialog() == true)
            {
                Output_XLSX_File_TextBox.Text = saveFileDialog.FileName;
            }
        }

        private void Process_Realms_CSV_File_to_Output_File_Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelBuilder.CreateWorkbook(Realms_CSV_File_TextBox.Text, Output_XLSX_File_TextBox.Text);
        }
    }
}
