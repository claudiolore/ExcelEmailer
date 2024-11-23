using System;
using System.IO;
using OfficeOpenXml; // Richiede il pacchetto NuGet EPPlus
using System.Net.Mail;
using System.Net.Mime;
using System.Collections.Generic;

namespace ExcelEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            // Impostazione della licenza EPPlus (necessaria dalla versione 5)
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            Console.WriteLine("Excel Email Sender");
            Console.WriteLine("-----------------");

            Console.Write("\nInserisci il percorso del file Excel: ");
            string excelPath = Console.ReadLine().Trim('"');

            if (!File.Exists(excelPath))
            {
                Console.WriteLine("\nATTENZIONE!!!!! Il file Excel specificato non esiste ò il percorso è sbagliato!");
                Console.ReadKey();
                return;
            }

            Console.WriteLine($"\nLettura del file: {excelPath}\n");
            List<string> emailAddresses = ReadEmailsFromExcel(excelPath);

            Console.WriteLine($"\nTrovate {emailAddresses.Count} email da inviare.");

            if (emailAddresses.Count == 0)
            {
                Console.WriteLine("\nATTENZIONE!!! Nessuna email trovata nel file Excel.");
                Console.ReadKey();
                return;
            }

            Console.Write("\nInserisci l'oggetto dell'email: ");
            string subject = Console.ReadLine();

            Console.Write("\nInserisci il corpo dell'email: ");
            string body = Console.ReadLine();

            Console.Write("\nVuoi allegare un PDF? (S/N): ");
            bool attachPdf = Console.ReadLine().ToUpper() == "S";

            Console.WriteLine(); // Riga vuota per separazione
            string pdfPath = "";
            if (attachPdf)
            {
                Console.Write("Inserisci il percorso del file PDF: ");
                pdfPath = Console.ReadLine();
            }

            // Configurazione email mittente
            Console.Write("\nInserisci l'email del mittente: ");
            string senderEmail = Console.ReadLine();
            Console.Write("Inserisci la password dell'email: ");
            string senderPassword = Console.ReadLine();

            Console.WriteLine("\nInizio invio email...\n");

            // Invio email
            foreach (string email in emailAddresses)
            {
                try
                {
                    SendEmail(senderEmail, senderPassword, email, subject, body, attachPdf ? pdfPath : null);
                    Console.WriteLine($"[SUCCESSO] Email inviata con successo a: {email}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERRORE] Errore nell'invio dell'email a {email}: {ex.Message}");
                }
            }

            Console.WriteLine("\nProcesso completato. Premi un tasto per uscire.");
            Console.ReadKey();
        }

        //----------------------------------------------------------------------------------------------------------------------------------------------------
        static List<string> ReadEmailsFromExcel(string filePath)
        {
            List<string> emails = new List<string>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("\nATTYENZIONE!!! Il file Excel non contiene fogli di lavoro!");
                    return emails;
                }

                var worksheet = package.Workbook.Worksheets[0];

                // Verifichiamo se ci sono righe
                if (worksheet.Dimension == null)
                {
                    Console.WriteLine("\nATTYENZIONE!!! Il foglio di lavoro è vuoto!");
                    return emails;
                }

                int rowCount = worksheet.Dimension.Rows;

                // La colonna Email è la 6° (indice 6)
                for (int row = 2; row <= rowCount; row++) // dalla seconda riga per intestaione colonne
                {
                    string email = worksheet.Cells[row, 6].Value?.ToString();
                    if (!string.IsNullOrEmpty(email))
                    {
                        emails.Add(email);
                        Console.WriteLine($"\nATTYENZIONE!!! Email trovata: {email}"); // Aggiunto per debug
                    }
                }
            }

            return emails;
        }

        //------------------------------------------------------------------------------------------------------------------
        static void SendEmail(string senderEmail, string senderPassword, string recipientEmail,
                            string subject, string body, string pdfPath = null)
        {
            using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587))
            {
                smtpClient.EnableSsl = true;
                smtpClient.Credentials = new System.Net.NetworkCredential(senderEmail, senderPassword);

                using (MailMessage mailMessage = new MailMessage())
                {
                    mailMessage.From = new MailAddress(senderEmail);
                    mailMessage.To.Add(recipientEmail);
                    mailMessage.Subject = subject;
                    mailMessage.Body = body;

                    if (!string.IsNullOrEmpty(pdfPath))
                    {
                        Attachment attachment = new Attachment(pdfPath, MediaTypeNames.Application.Pdf);
                        mailMessage.Attachments.Add(attachment);
                    }

                    smtpClient.Send(mailMessage);
                }
            }
        }
    }
}
