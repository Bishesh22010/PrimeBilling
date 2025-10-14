using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace PrimeInsulationBilling.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // This method runs when the window is first opened
            LoadTemplatesIntoComboBox();
            dpInvoiceDate.SelectedDate = DateTime.Now; // Set today's date by default
        }

        // Scans the 'Templates' folder and populates the dropdown list
        private void LoadTemplatesIntoComboBox()
        {
            try
            {
                // Define the directory where templates are stored
                string templateDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");

                if (Directory.Exists(templateDirectory))
                {
                    // Get all files ending with .xlsx
                    var templates = Directory.GetFiles(templateDirectory, "*.xlsx");

                    // Add each file name to the ComboBox
                    foreach (var template in templates)
                    {
                        cmbTemplates.Items.Add(Path.GetFileName(template));
                    }

                    // If templates were found, select the first one by default
                    if (cmbTemplates.Items.Count > 0)
                    {
                        cmbTemplates.SelectedIndex = 0;
                    }
                    else
                    {
                        lblStatus.Text = "No templates found in the 'Templates' folder.";
                    }
                }
                else
                {
                    lblStatus.Text = "Error: 'Templates' folder not found.";
                    MessageBox.Show("The 'Templates' folder could not be found. Please create it and add your .xlsx files.", "Folder Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not load templates: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Handles the click event for the "Generate and Open Bill" button
        private void GenerateBillButton_Click(object sender, RoutedEventArgs e)
        {
            // --- 1. Validate Input ---
            if (cmbTemplates.SelectedItem == null)
            {
                MessageBox.Show("Please select a bill template from the dropdown list.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(txtInvoiceNumber.Text))
            {
                MessageBox.Show("Please enter an Invoice Number.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                lblStatus.Text = "Generating bill, please wait...";

                // --- 2. Gather Data from the UI ---
                var billData = new Dictionary<string, string>
                {
                    { "invoice_number", txtInvoiceNumber.Text },
                    { "e_way_bill", txtEWayBill.Text },
                    { "vehicle_no", txtVehicleNo.Text },
                    { "invoice_date", dpInvoiceDate.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd") },
                    { "total_amount", txtTotalAmount.Text }
                    // Add other key-value pairs for any other fields you add
                };

                string templateName = cmbTemplates.SelectedItem.ToString();
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", templateName);

                // --- 3. Call the Excel Service to Create the Bill ---
                ExcelService excelService = new ExcelService();
                string generatedFilePath = excelService.CreateBill(templatePath, billData);

                // --- 4. Open the Generated File ---
                // UseShellExecute is crucial for opening the file with its default program (e.g., Excel)
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo(generatedFilePath)
                    {
                        UseShellExecute = true
                    }
                };
                process.Start();

                lblStatus.Text = $"Successfully generated: {Path.GetFileName(generatedFilePath)}";
                MessageBox.Show("Bill has been generated and opened successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "An error occurred while generating the bill.";
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
