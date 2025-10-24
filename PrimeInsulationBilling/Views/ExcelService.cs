using OfficeOpenXml;
using PrimeInsulationBilling.Services; // Add this line to use the converter
using System;
using System.Collections.Generic;
using System.IO;

namespace PrimeInsulationBilling
{
    public class ExcelService
    {
        static ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public string CreateBill(string templatePath, Dictionary<string, string> data)
        {
            FileInfo templateFile = new FileInfo(templatePath);
            string newFileName = $"Bill-{data["invoice_number"]}-{DateTime.Now:yyyy-MM-dd}.xlsx";
            string generatedBillsDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GeneratedBills");
            Directory.CreateDirectory(generatedBillsDirectory);
            string newFilePath = Path.Combine(generatedBillsDirectory, newFileName);

            using (ExcelPackage package = new ExcelPackage(templateFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // --- HEADER SECTION ---
                worksheet.Cells["F14"].Value = data["invoice_number"] + "/25-26";
                worksheet.Cells["H14"].Value = DateTime.Parse(data["invoice_date"]).ToString("dd.MM.yyyy");
                worksheet.Cells["F18"].Value = data["e_way_bill"];
                worksheet.Cells["F26"].Value = data["lr_number"];
                worksheet.Cells["H26"].Value = data["vehicle_no"];

                // --- ITEM DETAILS SECTION ---
                worksheet.Cells["B31"].Value = data["description_of_goods"];
                worksheet.Cells["E31"].Value = data.ContainsKey("packing1") ? data["packing1"] : "";
                worksheet.Cells["E32"].Value = data.ContainsKey("packing2") ? data["packing2"] : "";
                worksheet.Cells["E34"].Value = data.ContainsKey("packing4") ? data["packing4"] : "";
                worksheet.Cells["F31"].Value = data["hsn_code"];
                worksheet.Cells["G31"].Value = data["quantity"];
                if (decimal.TryParse(data["rate"], out decimal rate)) worksheet.Cells["H31"].Value = rate;
                if (decimal.TryParse(data["total_amount"], out decimal amount)) worksheet.Cells["J31"].Value = amount;

                // --- FOOTER SECTION ---
                worksheet.Cells["A42"].Value = "Indian Rupees: " + data["amount_in_words"];

                // --- GST & TOTALS - THE FIX IS HERE ---
                // We parse the percentage, divide by 100, and then set the cell's format.
                if (decimal.TryParse(data["cgst"], out decimal cgst))
                {
                    worksheet.Cells["G45"].Value = cgst / 100;
                    worksheet.Cells["G45"].Style.Numberformat.Format = "0.00%";
                }
                if (decimal.TryParse(data["sgst"], out decimal sgst))
                {
                    worksheet.Cells["G48"].Value = sgst / 100;
                    worksheet.Cells["G48"].Style.Numberformat.Format = "0.00%";
                }
                if (decimal.TryParse(data["igst"], out decimal igst))
                {
                    worksheet.Cells["G51"].Value = igst / 100;
                    worksheet.Cells["G51"].Style.Numberformat.Format = "0.00%";
                }
                if (decimal.TryParse(data["roff"], out decimal roff))
                {
                    worksheet.Cells["J54"].Value = roff;
                }

                worksheet.Cells["A56"].Value = data["declaration"];
                package.SaveAs(new FileInfo(newFilePath));
            }
            return newFilePath;
        }
    }
}