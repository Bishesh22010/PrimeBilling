using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace PrimeInsulationBilling.Views
{
    /// <summary>
    /// Interaction logic for ViewBillsView.xaml
    /// </summary>
    public partial class ViewBillsView : UserControl
    {
        private string generatedBillsDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeneratedBills");

        public ViewBillsView()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            LoadBills();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            LoadBills();
        }

        private void LoadBills()
        {
            try
            {
                if (!Directory.Exists(generatedBillsDirectory))
                {
                    Directory.CreateDirectory(generatedBillsDirectory);
                }

                var billFiles = Directory.GetFiles(generatedBillsDirectory, "*.xlsx")
                                         .Select(filePath => new BillFile
                                         {
                                             FileName = Path.GetFileName(filePath),
                                             FullPath = filePath,
                                             DateCreated = File.GetCreationTime(filePath)
                                         })
                                         .OrderByDescending(f => f.DateCreated) // Show newest first
                                         .ToList();

                BillsListView.ItemsSource = billFiles;
                lblBillCount.Text = $"Total Bills: {billFiles.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading bills: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            OpenSelectedBill();
        }

        private void BillsListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenSelectedBill();
        }

        private void OpenSelectedBill()
        {
            if (BillsListView.SelectedItem is BillFile selectedBill)
            {
                try
                {
                    var p = new Process
                    {
                        StartInfo = new ProcessStartInfo(selectedBill.FullPath)
                        {
                            UseShellExecute = true
                        }
                    };
                    p.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not open file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a bill from the list to open.", "No Bill Selected", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }

    // A small helper class to store bill information
    public class BillFile
    {
        public string FileName { get; set; }
        public string FullPath { get; set; }
        public DateTime DateCreated { get; set; }
    }
}
