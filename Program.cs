using CsvHelper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Security;

namespace CSVManipulationsAndEmail
{
    class Programс
    {
        public class CSVLine
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Country { get; set; }
            public string City { get; set; }
            public int Score { get; set; }

            public CSVLine(string inFirstName, string inLastName, string inCountry, string inCity, int inScore)
            {
                FirstName = inFirstName;
                LastName = inLastName;
                Country = inCountry;
                City = inCity;
                Score = inScore;
            }
            
        }
        public class NewFileStructure
        {
            public double AverageScore { get; set; }
            public double MedianScore { get; set; }
            public double MaxScore { get; set; }
            public string MaxScorePerson { get; set; }
            public double MinScore { get; set; }
            public string MinScorePerson { get; set; }
            public int RecordCount { get; set; }

            public NewFileStructure(double inAverageScore, double inMedianScore, double inMaxScore, string inMaxScorePerson, double inMinScore, string inMinScorePerson, int inRecordCount)
            {
                AverageScore = inAverageScore;
                MedianScore = inMedianScore;
                MaxScore = inMaxScore;
                MaxScorePerson = inMaxScorePerson;
                MinScore = inMinScore;
                MinScorePerson = inMinScorePerson;
                RecordCount = inRecordCount;
            }
        }
        public static DataTable jsonStringToTable(string jsonContent)
        {
            var dt = JsonConvert.DeserializeObject<DataTable>(jsonContent.ToString());
            return dt;
        }

        public static string jsonToCSV(string jsonContent, string delimiter)
        {
            StringWriter csvString = new StringWriter();
            using (var csv = new CsvWriter(csvString, CultureInfo.CurrentCulture))
            {
                csv.Configuration.Delimiter = delimiter;

                using (var dt = jsonStringToTable(jsonContent))
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        csv.WriteField(column.ColumnName);
                    }
                    csv.NextRecord();

                    foreach (DataRow row in dt.Rows)
                    {
                        for (var i = 0; i < dt.Columns.Count; i++)
                        {
                            csv.WriteField(row[i]);
                        }
                        csv.NextRecord();
                    }
                }
            }
            return csvString.ToString();
        }

        /*public static string SendMail(string SmtpHost, string Domain, string SenderEmail, SecureString SenderPassword, string ReceiverEmail, bool UseSSL)
        {
            try
            {
                string AppLocation = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(SmtpHost);

                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                mail.From = new MailAddress(SenderEmail);
                mail.To.Add(ReceiverEmail);
                mail.Subject = "CSVFile - ReportByCountry";
                mail.Body = "test";

                Attachment data = new Attachment(Path.Combine(AppLocation, "ReportByCountry.csv"), MediaTypeNames.Application.Octet);
                mail.Attachments.Add(data);

                SmtpServer.Port = 587;
                SmtpServer.Timeout = 20000;
                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.Credentials = new System.Net.NetworkCredential(SenderEmail, SenderPassword);
                SmtpServer.EnableSsl = UseSSL;

                SmtpServer.Send(mail);
                return "Success";
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Authentication Required"))
                {
                    Console.WriteLine("Error with authentication for domain '{0}'", Domain);
                }
                else
                {
                    Console.WriteLine("Error while sending email. Please check the information that you set at the beggining and try again.");
                    Console.WriteLine(ex.Message);
                }
                return "Error";
            }
         } */

        private static SecureString GetConsoleSecurePassword()
        {
            SecureString pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                
                if (!((i.Key == ConsoleKey.Backspace) && (pwd.Length == 0)))
                {
                    if (i.Key == ConsoleKey.Enter)
                    {
                        Console.WriteLine();
                        break;
                    }
                    else if (i.Key == ConsoleKey.Backspace)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                    else
                    {
                        pwd.AppendChar(i.KeyChar);
                        Console.Write("*");
                    }
                }
            }
            return pwd;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("This program will take four input parameters that will be specified at the beggining!");
            Console.WriteLine("Please use Gmail for Sender!");

            Console.Write("CSV File path : ");
            string FilePath = Console.ReadLine();
            while (!File.Exists(FilePath))
            {
                Console.WriteLine("File Not Found!");
                Console.Write("CSV File path : ");
                FilePath = Console.ReadLine();
            }
            Console.Write("Sender email address : ");
            string SenderEmail = Console.ReadLine();
            Console.Write("Sender email password : ");
            SecureString SenderPassword = GetConsoleSecurePassword(); ;
            Console.Write("Receiver email address : ");
            string ReceiverEmail = Console.ReadLine();

            List<CSVLine> CSVLines = new List<CSVLine>();
            using (var reader = new StreamReader(FilePath))
            {
                bool isFirstRow = true;
                while (!reader.EndOfStream)
                {
                    if (isFirstRow) 
                    {
                        var line = reader.ReadLine();
                        isFirstRow = false;
                    }
                    else
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(';');
                        CSVLines.Add(new CSVLine(values[0], values[1], values[2], values[3], int.Parse(values[4])));

                    }
                }
            }

            IEnumerable<IGrouping<string, CSVLine>> Countries = CSVLines
                  .GroupBy(x => x.Country);

            List<NewFileStructure> NewCSV = new List<NewFileStructure>();
            foreach (var item in Countries)
            {
                var AverageScore = (double)Math.Round(CSVLines
                                                            .Where(x => x.Country == item.Key)
                                                            .Select(x => x.Score)
                                                            .DefaultIfEmpty(0)
                                                            .Average(), 2);
                var MaxScore = CSVLines
                                    .Where(x => x.Country == item.Key)
                                    .Select(x => x.Score)
                                    .DefaultIfEmpty(0)
                                    .Max();
                var MaxScorePerson = CSVLines
                                        .Where(x => x.Country == item.Key && x.Score == MaxScore)
                                        .Select(x => x.FirstName + " " + x.LastName)
                                        .First();
                var MinScore = CSVLines
                                    .Where(x => x.Country == item.Key)
                                    .Select(x => x.Score)
                                    .DefaultIfEmpty(0)
                                    .Min();
                var MinScorePerson = CSVLines
                                        .Where(x => x.Country == item.Key && x.Score == MinScore)
                                        .Select(x => x.FirstName + " " + x.LastName)
                                        .First();
                var RecordCount = CSVLines.Count();
                var MedianScore = (CSVLines.ElementAt(RecordCount / 2).Score + CSVLines.ElementAt((RecordCount - 1) / 2).Score) / 2;

                NewCSV.Add(new NewFileStructure(AverageScore, MedianScore, MaxScore, MaxScorePerson, MinScore, MinScorePerson, RecordCount));
            }

            string AppLocation = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));


            var csv = jsonToCSV(JsonConvert.SerializeObject(NewCSV.OrderByDescending(x => x.AverageScore)),";");

            File.WriteAllText(Path.Combine(AppLocation, "ReportByCountry.csv"), csv.TrimEnd(), System.Text.Encoding.UTF8);

           /* var domain = SenderEmail.Substring(SenderEmail.IndexOf("@") + 1);
            if (domain == "gmail.com")
            {
                SendMail("smtp.gmail.com", domain, SenderEmail, SenderPassword, ReceiverEmail, true);
            } */
            var message = new MimeMessage();

            message.From.Add(new MailboxAddress("Tsvetina Kalcheva",SenderEmail));
            message.To.Add(new MailboxAddress("Stela Valcheva", ReceiverEmail));
            message.Subject = "Test";
            message.Body = new TextPart("plain"){ Text = @"ReportByCountry.csv" };

            using (var client = new SmtpClient())
            {
             client.Connect("smtp.gmail.com", 587);

   ////Note: only needed if the SMTP server requires authentication
            client.Authenticate(SenderEmail, SenderPassword);

            client.Send(message);
            client.Disconnect(true);
}

        }
    }
}
