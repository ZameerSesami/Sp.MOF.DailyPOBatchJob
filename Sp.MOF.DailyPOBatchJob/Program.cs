using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using ClosedXML.Excel;

namespace Sp.MOF.DailyPOBatchJob
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Log("===============================================");
                Log("Daily PO Transaction Report - START");
                Log("===============================================");

                // ---------------------------------------------------------
                // 1. Calculate previous day
                // ---------------------------------------------------------

                DateTime reportDate = DateTime.Today.AddDays(-1);
                //DateTime reportDate = DateTime.Today;

                Log("Report Date : " + reportDate.ToString("yyyy-MM-dd"));

                // ---------------------------------------------------------
                // 2. Read configuration
                // ---------------------------------------------------------

                string connectionString =
                    ConfigurationManager.ConnectionStrings["PeppolMOFCon"].ConnectionString;

                string sqlQuery =
                    ConfigurationManager.AppSettings["SqlQuery"];

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    throw new Exception(
                        "PeppolMOF connection string is not configured.");
                }

                if (string.IsNullOrWhiteSpace(sqlQuery))
                {
                    throw new Exception(
                        "SqlQuery is not configured in App.config.");
                }


                // ---------------------------------------------------------
                // 3. Fetch data from SQL
                // ---------------------------------------------------------

                Log("Fetching data from SQL Server...");

                DataTable dataTable =
    GetPOTransactions(
        connectionString,
        sqlQuery);

                Log("Total transactions retrieved : " +
                    dataTable.Rows.Count);


                // ---------------------------------------------------------
                // 4. Create today's report folder
                // ---------------------------------------------------------

                string baseOutputPath =
                    ConfigurationManager.AppSettings["ReportOutputPath"];

                if (string.IsNullOrWhiteSpace(baseOutputPath))
                {
                    throw new Exception(
                        "ReportOutputPath is not configured.");
                }

                // Create today's folder: yyyyMMdd
                string todayFolder =
                    DateTime.Today.ToString("yyyyMMdd");

                string outputPath =
                    Path.Combine(baseOutputPath, todayFolder);

                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                }

                Log("Report output folder : " + outputPath);


                // ---------------------------------------------------------
                // 5. Generate Excel
                // ---------------------------------------------------------

                string fileNameTemplate =
     ConfigurationManager.AppSettings["ReportFileName"];

                if (string.IsNullOrWhiteSpace(fileNameTemplate))
                {
                    fileNameTemplate =
                        "Daily_PO_Transaction_Report_{0:yyyyMMddHHmmss}.xlsx";
                }

                string fileName =
                    string.Format(
                        fileNameTemplate,
                        DateTime.Now);

                string excelFilePath =
                    Path.Combine(outputPath, fileName);

                Log("Generating Excel report...");

                GenerateExcel(
                    dataTable,
                    excelFilePath,
                    reportDate);

                Log("Excel report generated:");
                Log(excelFilePath);


                // ---------------------------------------------------------
                // 6. Send email
                // ---------------------------------------------------------

                string emailEnabled =
                    ConfigurationManager.AppSettings["EmailEnabled"];

                if (string.Equals(
                    emailEnabled,
                    "true",
                    StringComparison.OrdinalIgnoreCase))
                {
                    Log("Sending email...");

                    SendEmail(
                        excelFilePath,
                        reportDate,
                        dataTable.Rows.Count);

                    Log("Email sent successfully.");
                }
                else
                {
                    Log("Email sending is disabled.");
                }


                // ---------------------------------------------------------
                // 7. Finish
                // ---------------------------------------------------------

                Log("===============================================");
                Log("Daily PO Transaction Report - COMPLETED");
                Log("===============================================");
            }
            catch (Exception ex)
            {
                Log("===============================================");
                Log("ERROR");
                Log("===============================================");

                Log(ex.ToString());

                Console.WriteLine();
                Console.WriteLine("ERROR:");
                Console.WriteLine(ex.Message);

                Environment.ExitCode = 1;
            }
        }


        // ================================================================
        // GET DATA FROM SQL SERVER
        // ================================================================

        private static DataTable GetPOTransactions(
       string connectionString,
       string sqlQuery)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {
                connection.Open();

                Log("SQL connection opened.");

                using (SqlCommand command =
                    new SqlCommand(sqlQuery, connection))
                {
                    command.CommandTimeout = 300;

                    using (SqlDataAdapter adapter =
                        new SqlDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                Log("SQL query completed.");
            }

            return dataTable;
        }


        // ================================================================
        // GENERATE EXCEL
        // ================================================================

        private static void GenerateExcel(
      DataTable dataTable,
      string filePath,
      DateTime reportDate)
        {
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet =
                    workbook.Worksheets.Add("PO Transactions");

                // ---------------------------------------------------------
                // Report title
                // ---------------------------------------------------------

                worksheet.Cell(1, 1).Value =
                    "Daily PO Transaction Report";

                int columnCount = dataTable.Columns.Count;

                if (columnCount > 0)
                {
                    worksheet.Range(
                        1,
                        1,
                        1,
                        columnCount).Merge();
                }

                worksheet.Cell(1, 1).Style.Font.Bold = true;
                worksheet.Cell(1, 1).Style.Font.FontSize = 16;
                worksheet.Cell(1, 1).Style.Alignment.Horizontal =
                    XLAlignmentHorizontalValues.Center;


                // ---------------------------------------------------------
                // Report date
                // ---------------------------------------------------------

                worksheet.Cell(2, 1).Value =
                    "Report Date";

                worksheet.Cell(2, 2).Value =
                    reportDate;

                worksheet.Cell(2, 2).Style.DateFormat.Format =
                    "yyyy-MM-dd";


                // ---------------------------------------------------------
                // Transaction count
                // ---------------------------------------------------------

                worksheet.Cell(3, 1).Value =
                    "Transaction Count";

                worksheet.Cell(3, 2).Value =
                    dataTable.Rows.Count;


                // ---------------------------------------------------------
                // Dynamic headers
                // ---------------------------------------------------------

                int headerRow = 5;

                for (int columnIndex = 0;
                     columnIndex < dataTable.Columns.Count;
                     columnIndex++)
                {
                    worksheet.Cell(
                        headerRow,
                        columnIndex + 1).Value =
                        dataTable.Columns[columnIndex].ColumnName;
                }


                // ---------------------------------------------------------
                // Header formatting
                // ---------------------------------------------------------

                if (columnCount > 0)
                {
                    IXLRange headerRange =
                        worksheet.Range(
                            headerRow,
                            1,
                            headerRow,
                            columnCount);

                    headerRange.Style.Font.Bold = true;

                    headerRange.Style.Alignment.Horizontal =
                        XLAlignmentHorizontalValues.Center;

                    headerRange.Style.Alignment.Vertical =
                        XLAlignmentVerticalValues.Center;

                    headerRange.Style.Alignment.WrapText = true;
                }


                // ---------------------------------------------------------
                // Data
                // ---------------------------------------------------------

                int currentRow = headerRow + 1;

                foreach (DataRow dataRow in dataTable.Rows)
                {
                    for (int columnIndex = 0;
                         columnIndex < dataTable.Columns.Count;
                         columnIndex++)
                    {
                        object value =
                            dataRow[columnIndex];

                        IXLCell cell =
                            worksheet.Cell(
                                currentRow,
                                columnIndex + 1);

                        if (value == null ||
                            value == DBNull.Value)
                        {
                            cell.Value = "";
                        }
                        else
                        {
                            DateTime dateValue;

                            if (value is DateTime)
                            {
                                dateValue = (DateTime)value;

                                cell.Value = dateValue;

                                cell.Style.DateFormat.Format =
                                    "yyyy-MM-dd HH:mm:ss";
                            }
                            else
                            {
                                cell.Value = value.ToString();
                            }
                        }
                    }

                    currentRow++;
                }


                // ---------------------------------------------------------
                // Add Excel table
                // ---------------------------------------------------------

                if (dataTable.Rows.Count > 0 &&
                    columnCount > 0)
                {
                    IXLRange tableRange =
                        worksheet.Range(
                            headerRow,
                            1,
                            currentRow - 1,
                            columnCount);

                    IXLTable table =
                        tableRange.CreateTable("POTransactions");

                    table.Theme =
                        XLTableTheme.TableStyleMedium2;
                }


                // ---------------------------------------------------------
                // Freeze header
                // ---------------------------------------------------------

                worksheet.SheetView.FreezeRows(headerRow);


                // ---------------------------------------------------------
                // Autofilter
                // ---------------------------------------------------------

                //if (dataTable.Rows.Count > 0 &&
                //    columnCount > 0)
                //{
                //    worksheet.Range(
                //        headerRow,
                //        1,
                //        currentRow - 1,
                //        columnCount)
                //        .SetAutoFilter();
                //}
                if (dataTable.Rows.Count > 0 &&
                   columnCount > 0)
                {
                    worksheet.Range(
                        headerRow,
                        1,
                        currentRow - 1,
                        columnCount);
                }



                // ---------------------------------------------------------
                // Adjust columns
                // ---------------------------------------------------------

                worksheet.Columns().AdjustToContents();


                // ---------------------------------------------------------
                // Prevent excessively wide columns
                // ---------------------------------------------------------

                for (int i = 1;
                     i <= columnCount;
                     i++)
                {
                    if (worksheet.Column(i).Width > 45)
                    {
                        worksheet.Column(i).Width = 45;
                    }
                }


                // ---------------------------------------------------------
                // Save
                // ---------------------------------------------------------

                workbook.SaveAs(filePath);
            }
        }

        // ================================================================
        // SET NORMAL CELL
        // ================================================================

        private static void SetCellValue(
            IXLCell cell,
            object value)
        {
            if (value == null ||
                value == DBNull.Value)
            {
                cell.Value = "";
            }
            else
            {
                cell.Value = value.ToString();
            }
        }


        // ================================================================
        // SET DATE CELL
        // ================================================================

        private static void SetDateCell(
            IXLCell cell,
            object value,
            string format)
        {
            if (value == null ||
                value == DBNull.Value)
            {
                cell.Value = "";
                return;
            }

            DateTime dateValue;

            if (DateTime.TryParse(
                value.ToString(),
                out dateValue))
            {
                cell.Value = dateValue;

                cell.Style.DateFormat.Format =
                    format;
            }
            else
            {
                cell.Value = value.ToString();
            }
        }


        // ================================================================
        // SEND EMAIL
        // ================================================================

        private static void SendEmail(string attachmentPath, DateTime reportDate, int transactionCount)
        {
            string from = ConfigurationManager.AppSettings["EmailFrom"];
            string to = ConfigurationManager.AppSettings["EmailTo"];
            string cc = ConfigurationManager.AppSettings["EmailCc"];
            string smtpHost = ConfigurationManager.AppSettings["SmtpHost"];
            string smtpPortString = ConfigurationManager.AppSettings["SmtpPort"];
            string enableSslString = ConfigurationManager.AppSettings["EnableSsl"];

            string useDefaultCredentialsString = ConfigurationManager.AppSettings["UseDefaultCredentials"];
            string smtpUsername = ConfigurationManager.AppSettings["SmtpUsername"];
            string smtpPassword = ConfigurationManager.AppSettings["SmtpPassword"];


            if (string.IsNullOrWhiteSpace(from))
            {
                Log("EmailFrom is not configured.");
                return;
            }

            if (string.IsNullOrWhiteSpace(to) && string.IsNullOrWhiteSpace(cc))
            {
                Log("No email recipients are configured.");
                return;
            }

            if (string.IsNullOrWhiteSpace(smtpHost))
            {
                Log("SmtpHost is not configured.");
                return;
            }


            int smtpPort = 25;
            if (!string.IsNullOrWhiteSpace(smtpPortString))
            {
                int.TryParse(smtpPortString, out smtpPort);
            }


            bool enableSsl = false;
            bool.TryParse(enableSslString, out enableSsl);

            bool useDefaultCredentials = true;

            if (!string.IsNullOrWhiteSpace(useDefaultCredentialsString))
            {
                bool.TryParse(useDefaultCredentialsString, out useDefaultCredentials);
            }


            string subjectTemplate = ConfigurationManager.AppSettings["EmailSubject"];
            string bodyTemplate = ConfigurationManager.AppSettings["EmailBody"];


            if (string.IsNullOrWhiteSpace(subjectTemplate))
            {
                subjectTemplate = "Daily PO Transaction Report - {0:yyyy-MM-dd}";
            }

            if (string.IsNullOrWhiteSpace(bodyTemplate))
            {
                bodyTemplate = "Please find attached the Daily PO Transaction Report.";
            }


            string subject = string.Format(subjectTemplate, reportDate);
            string body = string.Format(bodyTemplate, reportDate);


            // Add transaction count to email body
            body += "<br/><br/>" + "Total PO transactions: " + transactionCount;


            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(from);

                AddRecipients(mail.To, to);
                AddRecipients(mail.CC, cc);
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = true;
                mail.BodyEncoding = Encoding.UTF8;
                mail.SubjectEncoding = Encoding.UTF8;


                // ---------------------------------------------------------
                // Attachment
                // ---------------------------------------------------------

                if (!File.Exists(attachmentPath))
                {
                    Log("Excel attachment was not found. " + attachmentPath);
                    return;
                }

                mail.Attachments.Add(new Attachment(attachmentPath));


                // ---------------------------------------------------------
                // SMTP
                // ---------------------------------------------------------

                using (SmtpClient smtp = new SmtpClient(smtpHost, smtpPort))
                {
                    smtp.EnableSsl = enableSsl;
                    smtp.UseDefaultCredentials = useDefaultCredentials;

                    if (!useDefaultCredentials)
                    {
                        smtp.Credentials = new NetworkCredential(smtpUsername, smtpPassword);
                    }

                    smtp.Send(mail);
                }
            }
        }


        // ================================================================
        // ADD TO / CC / BCC RECIPIENTS
        // ================================================================

        private static void AddRecipients(MailAddressCollection collection, string addresses)
        {
            if (string.IsNullOrWhiteSpace(addresses))
            {
                return;
            }

            string[] addressList = addresses.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);


            foreach (string address in addressList)
            {
                string trimmedAddress = address.Trim();
                if (!string.IsNullOrWhiteSpace(trimmedAddress))
                {
                    collection.Add(trimmedAddress);
                }
            }
        }


        // ================================================================
        // LOGGING
        // ================================================================

        private static void Log(string message)
        {
            try
            {
                string logPath = ConfigurationManager.AppSettings["LogPath"];

                if (string.IsNullOrWhiteSpace(logPath))
                {
                    logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                }


                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }


                string logFile = Path.Combine(logPath, "POTransactionDailyReport_" + DateTime.Now.ToString("yyyyMMdd") + ".log");


                string logMessage = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine;


                File.AppendAllText(logFile, logMessage);

                Console.WriteLine(logMessage.TrimEnd());
            }
            catch
            {
                // Do not allow logging failure to
                // terminate the report process.
                Console.WriteLine(message);
            }
        }
    }
}