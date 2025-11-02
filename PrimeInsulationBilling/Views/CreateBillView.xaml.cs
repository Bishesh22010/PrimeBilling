using PrimeInsulationBilling.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace PrimeInsulationBilling.Views
{
    /// <summary>
    /// Interaction logic for CreateBillView.xaml
    /// </summary>
    public partial class CreateBillView : UserControl
    {
        // This variable stores the "Generate and Open Bill" button from the Parent Window
        private Button _generateBillButton;
        // This variable stores the Status Bar TextBlock from the Parent Window
        private TextBlock _lblStatus;

        public CreateBillView()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            // When this control loads, it finds the controls in its parent (MainWindow)
            // This is a common pattern for connecting UserControls to a Shell
            if (Window.GetWindow(this) is MainWindow parentWindow)
            {
                // Find the main button and status label in the parent window
                _generateBillButton = parentWindow.FindName("GenerateBillButton") as Button;
                _lblStatus = parentWindow.FindName("lblStatus") as TextBlock;

                // Attach the click event handler
                if (_generateBillButton != null)
                {
                    // Detach first to prevent multiple handlers
                    _generateBillButton.Click -= GenerateBillButton_Click;
                    // Attach the handler
                    _generateBillButton.Click += GenerateBillButton_Click;
                }
            }

            LoadTemplates();
            dpInvoiceDate.SelectedDate = DateTime.Now; // Set today's date by default
        }

        private void LoadTemplates()
        {
            try
            {
                string templateDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");

                if (!Directory.Exists(templateDirectory))
                {
                    MessageBox.Show("The 'Templates' folder could not be found. Please create it next to the application .exe and add your .xlsx files.", "Folder Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var templates = Directory.GetFiles(templateDirectory, "*.xlsx");

                // Clear any existing items except the placeholder
                while (cmbTemplates.Items.Count > 1)
                {
                    cmbTemplates.Items.RemoveAt(1);
                }

                foreach (var template in templates)
                {
                    cmbTemplates.Items.Add(Path.GetFileName(template));
                }

                cmbTemplates.SelectedIndex = 0; // Default to "Choose a template..."

                if (cmbTemplates.Items.Count <= 1)
                {
                    if (_lblStatus != null) _lblStatus.Text = "No templates found in the 'Templates' folder.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not load templates: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cmbTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_lblStatus == null) return;

            if (cmbTemplates.SelectedIndex <= 0) // "Choose a template..." is selected
            {
                _lblStatus.Text = "Please select a valid template.";
            }
            else if (cmbTemplates.SelectedItem != null)
            {
                _lblStatus.Text = $"Template '{cmbTemplates.SelectedItem}' selected. Ready.";
            }
        }

        private void Calculation_TextChanged(object sender, TextChangedEventArgs e)
        {
            decimal.TryParse(txtAmount.Text, out decimal baseAmount);
            decimal.TryParse(txtCgst.Text, out decimal cgstPercent);
            decimal.TryParse(txtSgst.Text, out decimal sgstPercent);
            decimal.TryParse(txtIgst.Text, out decimal igstPercent);

            decimal cgstAmount = baseAmount * (cgstPercent / 100);
            decimal sgstAmount = baseAmount * (sgstPercent / 100);
            decimal igstAmount = baseAmount * (igstPercent / 100);

            decimal grandTotal = baseAmount + cgstAmount + sgstAmount + igstAmount;

            lblGrandTotal.Text = grandTotal.ToString("C", new CultureInfo("en-IN"));
        }

        private void txtRoff_LostFocus(object sender, RoutedEventArgs e)
        {
            if (decimal.TryParse(txtRoff.Text, out decimal amount))
            {
                txtAmountInWords.Text = NumberToWordsConverter.ToIndianCurrencyWords(amount);
            }
            else
            {
                txtAmountInWords.Text = "";
            }
        }

        private void GenerateBillButton_Click(object sender, RoutedEventArgs e)
        {
            // --- Step 1: Validation ---
            if (cmbTemplates.SelectedIndex <= 0)
            {
                MessageBox.Show("Please select a valid bill template.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtInvoiceNumber.Text) || dpInvoiceDate.SelectedDate == null)
            {
                MessageBox.Show("Invoice Number and Date are required fields.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                if (_lblStatus != null) _lblStatus.Text = "Generating bill...";

                // --- Step 2: Gather All Data from the Form ---
                var billData = new Dictionary<string, string>
                {
                    { "invoice_number", txtInvoiceNumber.Text },
                    { "invoice_date", dpInvoiceDate.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd") },
                    { "e_way_bill", txtEWayBill.Text },
                    { "lr_number", txtLrNumber.Text },
                    { "vehicle_no", txtVehicleNo.Text },
                    { "description_of_goods", txtDescription.Text },
                    { "packing1", txtPacking1.Text },
                    { "packing2", txtPacking2.Text },
                    { "packing3", txtPacking3.Text },
                    { "packing4", txtPacking4.Text },
                    { "hsn_code", txtHsnCode.Text },
                    { "quantity", txtQuantity.Text },
                    { "rate", txtRate.Text },
                    { "total_amount", txtAmount.Text },
                    { "amount_in_words", txtAmountInWords.Text },
                    { "cgst", txtCgst.Text },
                    { "sgst", txtSgst.Text },
                    { "igst", txtIgst.Text },
                    { "roff", txtRoff.Text },
                    { "declaration", txtDeclaration.Text }
                };

                string templateName = cmbTemplates.SelectedItem.ToString();
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", templateName);

                // --- Step 3: Call the Excel Service to Create the Bill ---
                ExcelService excelService = new ExcelService();
                string generatedFilePath = excelService.CreateBill(templatePath, billData);

                // --- Step 4: Open the Generated File ---
                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(generatedFilePath)
                    {
                        UseShellExecute = true
                    }
                };
                p.Start();

                if (_lblStatus != null) _lblStatus.Text = $"Successfully generated: {Path.GetFileName(generatedFilePath)}";
                MessageBox.Show("Bill generated and opened successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                if (_lblStatus != null) _lblStatus.Text = "An error occurred while generating the bill.";
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

