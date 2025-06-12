// MainWindow.xaml.cs
using Microsoft.Win32;
using PayrollAutomation.Models;
using PayrollAutomation.Services;
using System.Collections.ObjectModel;
using System.Windows;
using WinForms = System.Windows.Forms;
using System.IO;

namespace PayrollAutomation
{
    public partial class MainWindow : Window
    {
        private List<EmployeePayroll> employees;
        private ObservableCollection<StatusItem> statusList = new();

        public MainWindow()
        {
            InitializeComponent();
            lvStatus.ItemsSource = statusList;
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinForms.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx"
            };
            if (dialog.ShowDialog()== WinForms.DialogResult.OK)
            {
                txtExcelPath.Text = dialog.FileName;

                var result = ExcelService.ReadEmployees(dialog.FileName, out employees);
                WinForms.MessageBox.Show(result);
            }
        }

        private async void GeneratePdfs_Click(object sender, RoutedEventArgs e)
        {
            if (employees == null || employees.Count == 0)
            {
                WinForms.MessageBox.Show("No valid employee data loaded.");
                return;
            }

            var dialog = new WinForms.FolderBrowserDialog();
            if (dialog.ShowDialog() != WinForms.DialogResult.OK)
                return;

            statusList.Clear();
            string folderPath = dialog.SelectedPath;

            foreach (var emp in employees)
            {
                // Skip the 'TOTAL' row
                if (emp.EmployeeName?.Trim().ToUpper() == "TOTAL")
                    continue;
                try
                {
                    string pdfPath = PdfService.GeneratePdf(emp, folderPath);
                    string status = string.IsNullOrEmpty(pdfPath) ? "Failed" : "PDF Generated";

                    if (chkSendEmail.IsChecked == true && !string.IsNullOrEmpty(pdfPath))
                    {
                        var email = new Email
                        {
                            To = emp.Email,
                            Subject = "Your Payslip",
                            Message = $"Dear {emp.EmployeeName},\n\nPlease find attached your payslip.",
                            HasAttachment = true,
                            AttachmentPath = pdfPath
                        };

                        status = await EmailService.SendEmailAsync(email);
                    }

                    statusList.Add(new StatusItem { EmployeeName = emp.EmployeeName, Status = status });
                }
                catch (Exception ex)
                {
                    statusList.Add(new StatusItem { EmployeeName = emp.EmployeeName, Status = $"Error: {ex.Message}" });
                }
            }

            WinForms.MessageBox.Show("PDF Generation complete.");
        }

        public class StatusItem
        {
            public string EmployeeName { get; set; }
            public string Status { get; set; }
        }
    }
}
