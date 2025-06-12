using DocumentFormat.OpenXml.Spreadsheet;
using PayrollAutomation.Models;
using System.Diagnostics;
using System.IO;
using System.Windows;


namespace PayrollAutomation.Services
{
    public static class PdfService
    {
        public static string GeneratePdf(EmployeePayroll payroll, string folderPath)
        {
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string fileName = $"{payroll.EmployeeName.Replace(" ", "_")}_PaySlip.pdf";
            string fullPath = Path.Combine(folderPath, fileName);

            // Load HTML Template and replace placeholders
            string htmlTemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template", "index.html");

            if (!File.Exists(htmlTemplatePath))
            {
                System.Windows.MessageBox.Show($"Template not found at:\n{htmlTemplatePath}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return string.Empty;
            }


            string html = File.ReadAllText(htmlTemplatePath);
            string logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template", "algo-logo-2x.png");
            if (File.Exists(logoPath))
            {
                string base64Logo = Convert.ToBase64String(File.ReadAllBytes(logoPath));
                html = html.Replace("{{LogoBase64}}", base64Logo);
            }
            else
            {
                html = html.Replace("{{LogoBase64}}", ""); // fallback if image missing
            }
            string filledHtml = HtmlPlaceholderService.FillPlaceholders(html, payroll);

            string tempHtmlPath = Path.Combine(Path.GetTempPath(), fileName.Replace(".pdf", ".html"));
            File.WriteAllText(tempHtmlPath, filledHtml);

            string exePath = @"D:\OFFICE\Utility\PayrollAutomation\PayrollAutomation\PayrollAutomation\Assests\WkHtmlToPdf\wkhtmltopdf.exe";//Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "WkHtmlToPdf", "wkhtmltopdf.exe");
            if (!File.Exists(exePath))
            {
                System.Windows.MessageBox.Show("wkhtmltopdf.exe not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return string.Empty;
            }

            var psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"\"{tempHtmlPath}\" \"{fullPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            Process.Start(psi)?.WaitForExit();

            File.Delete(tempHtmlPath); // Optional cleanup

            return fullPath;
        }
    }

    public static class HtmlPlaceholderService
    {
        public static string FillPlaceholders(string template, EmployeePayroll p)
        {
            if (p == null)
                throw new ArgumentNullException(nameof(p), "Payroll data is null.");
            
            DateTime now = DateTime.Now;
            string SafeStr(string? value) => string.IsNullOrWhiteSpace(value) ? "-" : value;
            string SafeDec(decimal? value) => value.HasValue && value.Value != 0 ? value.Value.ToString("N2") : "-";
            string SafeYear(int? year) => year.HasValue ? year.Value.ToString() : "-";
            p.PayYear = now.Year;

            // --- Perform Calculations ---
            p.TotalEarnings = (p.GrossSalaryFMO ?? 0) + (p.Arrears ?? 0) + (p.Bonus ?? 0) + (p.LeaveEncashment ?? 0);
            p.TotalTaxDeductions = (p.IncomeTaxDeduction ?? 0) + (p.ProfessionalTax ?? 0);
            p.NonTaxableTotal = (p.Fuel ?? 0) + (p.OPD ?? 0) + (p.OtherNonTaxable ?? 0);
            p.TotalOtherDeductions = (p.MessDeduction ?? 0) + (p.EOBI ?? 0) + (p.OtherDeductions ?? 0);
            p.TotalNetEarnings = p.TotalEarnings + p.NonTaxableTotal;
            p.TotalDeduction = p.TotalTaxDeductions + p.TotalOtherDeductions;
            p.NetSalariesPayable = Math.Round((decimal)(p.TotalNetEarnings - p.TotalDeduction), 2, MidpointRounding.AwayFromZero);

            return template
                .Replace("{{EmployeeID}}", SafeStr(p.EmployeeID?.ToString()))
                .Replace("{{EmployeeName}}", SafeStr(p.EmployeeName))
                .Replace("{{Deptt}}", SafeStr(p.Deptt))
                .Replace("{{CNIC}}", SafeStr(p.CNIC))
                .Replace("{{NTN}}", SafeStr(p.NTN))
                .Replace("{{Designation}}", SafeStr(p.Designation))
                .Replace("{{Grade}}", SafeStr(p.Grade))
                .Replace("{{BankName}}", SafeStr(p.PaymentMode))               
                .Replace("{{PayMonth}}", SafeStr(p.PayMonth))
                .Replace("{{PayYear}}", SafeYear(p.PayYear))
                .Replace("{{PayLocation}}", SafeStr(p.PayLocation))
                
               

                .Replace("{{GrossSalaryFMO}}", SafeDec(p.GrossSalaryFMO))
                .Replace("{{Arrears}}", SafeDec(p.Arrears))
                .Replace("{{Bonus}}", SafeDec(p.Bonus))
                .Replace("{{LeaveEncashment}}", SafeDec(p.LeaveEncashment))
                .Replace("{{TotalEarnings}}", SafeDec(p.TotalEarnings))

                .Replace("{{TaxDeductionFromSalary}}", SafeDec(p.IncomeTaxDeduction))
                .Replace("{{ProfessionalTax}}", SafeDec(p.ProfessionalTax))
                .Replace("{{TotalTaxDeductions}}", SafeDec(p.TotalTaxDeductions))



                .Replace("{{Fuel}}", SafeDec(p.Fuel))
                .Replace("{{OPD}}", SafeDec(p.OPD))
                .Replace("{{OtherNonTaxable}}", SafeDec(p.OtherNonTaxable))
                .Replace("{{NonTaxableTotal}}", SafeDec(p.NonTaxableTotal))

                .Replace("{{MessDeduction}}", SafeDec(p.MessDeduction))
                .Replace("{{EOBI}}", SafeDec(p.EOBI))
                .Replace("{{OtherDeductions}}", SafeDec(p.OtherDeductions))
                .Replace("{{TotalOtherDeductions}}", SafeDec(p.TotalOtherDeductions))

                .Replace("{{TotalNetEarnings}}", SafeDec(p.TotalNetEarnings))
                .Replace("{{TotalDeduction}}", SafeDec(p.TotalDeduction))
                .Replace("{{NetSalariesPayable}}", SafeDec(p.NetSalariesPayable));

        }


    }
}
