using PrimeInsulationBilling.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq; // Added for LINQ queries
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input; // Added for MouseDoubleClick

namespace PrimeInsulationBilling.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Path to the directory where bills are saved
        private readonly string generatedBillsDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeneratedBills");

        public MainWindow()
        {
            InitializeComponent();
            WindowState = WindowState.Maximized;
        }

        /// <summary>
        /// This method runs once the window has finished loading.
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadTemplates();
            dpInvoiceDate.SelectedDate = DateTime.Now; // Set today's date by default
        }

        /// <summary>
        /// Handles the window resize event.
        /// </summary>
        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // This method is required by the XAML but does not need any code for now.
        }

        // ===================================================================
        // NAVIGATION LOGIC
        // ===================================================================

        private void CreateBillNavButton_Click(object sender, RoutedEventArgs e)
        {
            // Show the Create Bill page
            CreateBillGrid.Visibility = Visibility.Visible;
            BottomBar.Visibility = Visibility.Visible; // Show the "Generate" button bar

            // Hide the View Bills page
            ViewBillsGrid.Visibility = Visibility.Collapsed;

            lblStatus.Text = "Ready to create a new bill.";
        }

        private void ViewBillsNavButton_Click(object sender, RoutedEventArgs e)
        {
            // Show the View Bills page
            ViewBillsGrid.Visibility = Visibility.Visible;

            // Hide the Create Bill page
            CreateBillGrid.Visibility = Visibility.Collapsed;
            BottomBar.Visibility = Visibility.Collapsed; // Hide the "Generate" button bar

            // Load the bills every time we switch to this page
            LoadBills();
            lblStatus.Text = "Viewing generated bills.";
        }

        // ===================================================================
        // CREATE BILL LOGIC (Your existing methods)
        // ===================================================================

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
                    lblStatus.Text = "No templates found in the 'Templates' folder.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not load templates: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cmbTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbTemplates.SelectedIndex <= 0) // "Choose a template..." is selected
            {
                lblStatus.Text = "Please select a valid template.";
            }
            else if (cmbTemplates.SelectedItem != null)
            {
                lblStatus.Text = $"Template '{cmbTemplates.SelectedItem}' selected. Ready.";
            }
        }

        private void Calculation_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Check if UI controls are loaded
            if (lblGrandTotal == null) return;

            decimal.TryParse(txtAmount.Text, out decimal baseAmount);
            decimal.TryParse(txtCgst.Text, out decimal cgstPercent);
            decimal.TryParse(txtSgst.Text, out decimal sgstPercent);
            decimal.TryParse(txtIgst.Text, out decimal igstPercent);

            decimal cgstAmount = baseAmount * (cgstPercent / 100);
            decimal sgstAmount = baseAmount * (sgstPercent / 100);
            decimal igstAmount = baseAmount * (igstPercent / 100);

            decimal grandTotal = baseAmount + cgstAmount + sgstAmount + igstAmount;

            // ONLY update the grand total label.
            lblGrandTotal.Text = grandTotal.ToString("C", new CultureInfo("en-IN"));
        }

        private void txtRoff_LostFocus(object sender, RoutedEventArgs e)
        {
            if (decimal.TryParse(txtRoff.Text, out decimal amount))
            {
                // Convert the manually entered number to words.
                txtAmountInWords.Text = NumberToWordsConverter.ToIndianCurrencyWords(amount);
            }
            else
            {
                txtAmountInWords.Text = ""; // Clear if input is not a valid number.
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
                lblStatus.Text = "Generating bill...";

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

                // FIX: Added '!' (null-forgiving operator) to satisfy compiler.
                // We know this is safe because of the 'SelectedIndex <= 0' check above.
                string templateName = cmbTemplates.SelectedItem.ToString()!;
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", templateName);

                // --- Step 3: Call the Excel Service to Create the Bill ---
                ExcelService excelService = new ExcelService();
                string generatedFilePath = excelService.CreateBill(templatePath, billData);

                // --- Step 4: Open the Generated File ---
                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(generatedFilePath)
                    {
                        UseShellExecute = true // This uses the default program (Excel) to open the file
                    }
                };
                p.Start();

                lblStatus.Text = $"Successfully generated: {Path.GetFileName(generatedFilePath)}";
                MessageBox.Show("Bill generated and opened successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "An error occurred while generating the bill.";
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ===================================================================
        // NEW: VIEW BILLS LOGIC
        // ===================================================================

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            LoadBills();
        }

        private void LoadBills()
        {
            try
            {
                // Ensure the directory exists
                if (!Directory.Exists(generatedBillsDirectory))
                {
                    Directory.CreateDirectory(generatedBillsDirectory);
                }

                // Get all Excel files from the directory
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
            // Check if an item is selected
            if (BillsListView.SelectedItem is BillFile selectedBill)
            {
                try
                {
                    // Open the file using the system's default application (Excel)
                    var p = new Process
                    {
                        // FIX: Added null check on FullPath
                        StartInfo = new ProcessStartInfo(selectedBill.FullPath ?? "")
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

        //
        // *** THE DUPLICATE BillFile CLASS DEFINITION WAS HERE AND IS NOW REMOVED ***
        //
    }
}

