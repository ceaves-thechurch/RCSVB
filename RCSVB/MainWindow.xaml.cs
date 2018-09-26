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
            OpenFile("CSV files (*.csv)|*.csv", @"C:\", "Select source CSV file exported from Realm", Realms_CSV_File_TextBox);
        }

        private void Realms_CSV_File_TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            OpenFile("CSV files (*.csv)|*.csv", @"C:\", "Select source CSV file exported from Realm", Realms_CSV_File_TextBox);
        }

        private void Output_XLSX_File_Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFile("Excel Files (*.xlsx)|*.xlsx", @"C:\", "Select or create destination file.", Output_XLSX_File_TextBox);
        }

        private void Output_XLSX_File_TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            SaveFile("Excel Files (*.xlsx)|*.xlsx", @"C:\", "Select or create destination file.", Output_XLSX_File_TextBox);
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

        private void OpenFile(string filter, string initialDirectory, string title, TextBox target)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                InitialDirectory = initialDirectory,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                target.Text = openFileDialog.FileName;
            }
        }

        private void SaveFile(string filter, string initialDirectory, string title, TextBox target)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Title = title,
                Filter = filter,
                InitialDirectory = initialDirectory,
                RestoreDirectory = true
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                target.Text = saveFileDialog.FileName;
            }
        }
    }
}
