using OfficeOpenXml; // You must have the EPPlus NuGet package installed
using System;
using System.Collections.Generic;
using System.IO;

namespace PrimeInsulationBilling
{
    public class ExcelService
    {
        // =================================================================
        //  THE GUARANTEED FIX: A Static Constructor
        //  This special constructor runs only ONCE when the application first
        //  needs to use ExcelService, ensuring the license is set globally.
        // =================================================================
        static ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Creates a new bill by populating an Excel template with data.
        /// </summary>
        /// <param name="templatePath">The full path to the .xlsx template file.</param>
        /// <param name="data">A dictionary containing the data to insert.</param>
        /// <returns>The file path of the newly created bill.</returns>
        public string CreateBill(string templatePath, Dictionary<string, string> data)
        {
            // The license is now set globally by the static constructor,
            // so we no longer need the line here.

            FileInfo templateFile = new FileInfo(templatePath);

            // Define the name and path for the new, generated bill
            string newFileName = $"Bill-{data["invoice_number"]}-{DateTime.Now:yyyy-MM-dd}.xlsx";
            string generatedBillsDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeneratedBills");

            // Ensure the output directory exists. If not, create it.
            Directory.CreateDirectory(generatedBillsDirectory);

            string newFilePath = Path.Combine(generatedBillsDirectory, newFileName);

            // Create a new Excel package from the template
            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                // Get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // --- Place Data into Specific Cells ---
                // This is where you map your dictionary data to the cells in your template.
                // It's important that these cell addresses match your .xlsx file exactly.
                worksheet.Cells["F14"].Value = data["invoice_number"] + "/25-26";
                worksheet.Cells["H14"].Value = DateTime.Parse(data["invoice_date"]).ToString("dd.MM.yyyy");
                worksheet.Cells["F18"].Value = data["e_way_bill"];
                worksheet.Cells["H26"].Value = data["vehicle_no"];

                // Safely parse the total amount to a decimal before setting the value
                if (decimal.TryParse(data["total_amount"], out decimal amount))
                {
                    worksheet.Cells["J31"].Value = amount;
                }

            }

            // Return the full path of the newly created file
            return newFilePath;
        }
    }
}

