using PayrollAutomation.Models;
using ClosedXML.Excel;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PayrollAutomation.Services
{
    public static class ExcelService
    {
        public static string ReadEmployees(string filePath, out List<EmployeePayroll> employees)
        {
            employees = new List<EmployeePayroll>();
            var errors = new StringBuilder();
            var uniqueNames = new HashSet<string>();

            try
            {
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(1);

                var headerRow = 6;
                var dataStartRow = 7;

                // Build a dictionary of column headers and their corresponding indices
                var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var cell in worksheet.Row(headerRow).CellsUsed())
                {
                    string header = cell.GetString().Trim();
                    if (!string.IsNullOrEmpty(header))
                    {
                        headerMap[header] = cell.Address.ColumnNumber;
                    }
                }

                // Validate required columns
                var requiredColumns = new[]
                {
                "Sr.No", "Employee Name", "Payment Mode", "Employee ID", "Designation", "Department", "Grade",
                "Pay Month", "Pay Location", "CNIC No.", "NTN",
                "Gross Salary FMO April", "Arrears", "Bonus", "Total",
                "Tax Deduction from Salary", "Tax Deduction",
                "Mess Deduction", "Other Deduction", "EOBI", "Total Deduction",
                "Fuel", "OPD", "Other", "Net Salaries Payable"
                };

                foreach (var col in requiredColumns)
                {
                    if (!headerMap.ContainsKey(col))
                    {
                        return $"The column '{col}' is missing or incorrect in the Excel file. Please correct it.";
                    }
                }

                var lastRow = worksheet.LastRowUsed().RowNumber();
                for (int row = dataStartRow; row <= lastRow; row++)
                {
                    string empName = worksheet.Cell(row, headerMap["Employee Name"]).GetString().Trim();
                    if (string.IsNullOrEmpty(empName)) continue;

                    if (!uniqueNames.Add(empName))
                    {
                        errors.AppendLine($"Duplicate employee: {empName} at row {row}");
                        continue;
                    }

                    employees.Add(new EmployeePayroll
                    {
                        //Personal Information
                        SrNo = ParseInt(worksheet.Cell(row, headerMap["Sr.No"])),
                        EmployeeName = empName,
                        PaymentMode = worksheet.Cell(row, headerMap["Payment Mode"]).GetString().Trim(),
                        EmployeeID = ParseInt(worksheet.Cell(row, headerMap["Employee ID"])),
                        Designation = worksheet.Cell(row, headerMap["Designation"]).GetValue<string>()?.Trim() ?? string.Empty,
                        Deptt = worksheet.Cell(row, headerMap["Department"]).GetString(),
                        Grade = worksheet.Cell(row, headerMap["Grade"]).GetValue<string>() ?? string.Empty,
                        PayMonth = worksheet.Cell(row, headerMap["Pay Month"]).GetValue<string>() ?? string.Empty,
                        PayLocation = worksheet.Cell(row, headerMap["Pay Location"]).GetValue<string>() ?? string.Empty,
                        CNIC = worksheet.Cell(row, headerMap["CNIC No."]).GetValue<string>() ?? string.Empty,
                        NTN = worksheet.Cell(row, headerMap["NTN"]).GetValue<string>() ?? string.Empty,

                        //Earnings
                        GrossSalaryFMO = ParseDecimal(worksheet.Cell(row, headerMap["Gross Salary FMO April"]).GetString()),
                        Arrears = ParseDecimal(worksheet.Cell(row, headerMap["Arrears"]).GetString()),
                        Bonus = ParseDecimal(worksheet.Cell(row, headerMap["Bonus"]).GetString()),
                        TotalEarnings = ParseDecimal(worksheet.Cell(row, headerMap["Total"]).GetString()),

                        //Tax Decductaion
                        IncomeTaxDeduction = ParseDecimal(worksheet.Cell(row, headerMap["Tax Deduction from Salary"]).GetString()),
                        ProfessionalTax = ParseDecimal(worksheet.Cell(row, headerMap["Tax Deduction"]).GetString()),

                        //Other deduction
                        MessDeduction = ParseDecimal(worksheet.Cell(row, headerMap["Mess Deduction"]).GetString()),
                        OtherDeductions = ParseDecimal(worksheet.Cell(row, headerMap["Other Deduction"]).GetString()),
                        EOBI = ParseDecimal(worksheet.Cell(row, headerMap["EOBI"]).GetString()),
                        TotalOtherDeductions = ParseDecimal(worksheet.Cell(row, headerMap["Total Deduction"]).GetString()),

                        //Non Taxable
                        Fuel = ParseDecimal(worksheet.Cell(row, headerMap["Fuel"]).GetString()),
                        OPD = ParseDecimal(worksheet.Cell(row, headerMap["OPD"]).GetString()),
                        OtherNonTaxable = ParseDecimal(worksheet.Cell(row, headerMap["Other"]).GetString()),

                        //Net Earning
                        NetSalariesPayable = ParseDecimal(worksheet.Cell(row, headerMap["Net Salaries Payable"]).GetString()),
                       // Email = worksheet.Cell(row, headerMap["Email"]).GetValue<string>() ?? string.Empty,
                    });
                }

                return errors.Length > 0 ? $"Validation completed with issues:\n{errors}" : "Validation successful.";
            }
            catch (Exception ex)
            {
                return $"Failed to read Excel file: {ex.Message}";
            }
        }

        private static int ParseInt(IXLCell cell)
        {
            if (cell.DataType == XLDataType.Number)
                return (int)cell.GetDouble();

            return int.TryParse(cell.GetString(), out int result) ? result : 0;
        }

        private static decimal ParseDecimal(string? value)
        {
            return decimal.TryParse(value, out decimal result) ? result : 0;
        }
    }
}
