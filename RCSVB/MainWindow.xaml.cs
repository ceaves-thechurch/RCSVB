using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
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

        public MainWindow ()
        {
            InitializeComponent();
        }

        private void Realms_CSV_File_Button_Click (object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            openFileDialog.InitialDirectory = @"C:\";
            openFileDialog.Title = "Select source CSV file exported from Realm";
            if (openFileDialog.ShowDialog() == true)
            {
                Realms_CSV_File_TextBox.Text = openFileDialog.FileName;
            }
        }

        private void Output_XLSX_File_Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            saveFileDialog.InitialDirectory = @"C:\";
            saveFileDialog.Title = "Select or create destination file.";
            if (saveFileDialog.ShowDialog() == true)
            {
                Output_XLSX_File_TextBox.Text = saveFileDialog.FileName;
            }
        }

        private async void Process_Realms_CSV_File_to_Output_File_Button_Click(object sender, RoutedEventArgs e) 
        {
            var source = Realms_CSV_File_TextBox.Text;
            var destination = Output_XLSX_File_TextBox.Text;

            Process_Realms_CSV_File_to_Output_File_Button.IsEnabled = false;
            Process_Realms_CSV_File_to_Output_File_Button.Content = "Processing...";

            Realms_CSV_to_XLSX_File_Progress_Label.Visibility = Visibility.Visible;

            Realms_CSV_to_XLSX_File_Progress_Bar.Visibility = Visibility.Visible;
            Realms_CSV_to_XLSX_File_Progress_Bar.IsIndeterminate = true;
            

            var task = Task<int>.Factory.StartNew (() => 
                ExcelBuilder.CreateFromRealmsCSV (source, destination)
            );
            await task;

            Process_Realms_CSV_File_to_Output_File_Button.IsEnabled = true;
            Process_Realms_CSV_File_to_Output_File_Button.Content = "Process";

            Realms_CSV_to_XLSX_File_Progress_Label.Visibility = Visibility.Hidden;
            Realms_CSV_to_XLSX_File_Progress_Bar.Visibility = Visibility.Hidden;
        }
    }
}
