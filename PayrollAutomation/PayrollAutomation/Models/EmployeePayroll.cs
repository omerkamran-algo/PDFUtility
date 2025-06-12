using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PayrollAutomation.Models
{
    public class EmployeePayroll
    {
        // Personal & Header Info
        public int? SrNo { get; set; }
        public int? EmployeeID { get; set; }
        public string EmployeeName { get; set; }
        public string Deptt { get; set; }
        public string CNIC { get; set; }
        public string NTN { get; set; }
        public string? Designation { get; set; }
        public string Grade { get; set; }
        public string PaymentMode { get; set; }
        public string? PayMonth { get; set; }   
        public int PayYear { get; set; } 
        public string PayLocation { get; set; }
        public string Email { get; set; }


        // Earnings

        public decimal? GrossSalaryFMO { get; set; } //GrossPay
        public decimal? Arrears { get; set; }
        public decimal? Bonus { get; set; }
        public decimal? LeaveEncashment { get; set; }
        public decimal? TotalEarnings { get; set; }

        // Deductions
        public decimal? IncomeTaxDeduction { get; set; } 
        public decimal? ProfessionalTax { get; set; }
        public decimal? TotalTaxDeductions { get; set; }

        // Non Taxable
        public decimal? Fuel { get; set; }
        public decimal? OPD { get; set; }
        public decimal? OtherNonTaxable { get; set; }
        public decimal? NonTaxableTotal { get; set; }


        // Other Deduction
        public decimal? MessDeduction { get; set; }
        public decimal? EOBI { get; set; }
        public decimal? OtherDeductions { get; set; }
        public decimal? TotalOtherDeductions { get; set; }


        // Totals
        public decimal? TotalNetEarnings { get; set; } //Total Earnings
        public decimal? TotalDeduction { get; set; }
        public decimal? NetSalariesPayable { get; set; } //Net Earnings


    }
  
}
