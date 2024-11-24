using System;
using System.IO;
using OfficeOpenXml;
using System.Net.Mail;
using System.Net.Mime;
using System.Collections.Generic;

namespace ExcelEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true) 
            { 
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
                string subject = string.Empty;
                string body = string.Empty;
                string senderEmail;
                string senderPassword;

                Console.WriteLine("Excel Email Sender");
                Console.WriteLine("BENVENUTO");
                Console.WriteLine("-----------------");
                
                //INSERIMENTO EXCEL CON EMAIL
                Console.Write("\n--Inserisci il percorso del file Excel in cui ci sono le email a cui inviare: ");
                string excelPath = Console.ReadLine().Trim('"');

                if (!File.Exists(excelPath))
                {
                    Console.WriteLine("\nATTENZIONE!!!!! Il file Excel specificato non esiste ò il percorso è sbagliato!");
                    Console.WriteLine("Premi qualunque tasto per riprovare");
                    Console.ReadKey();
                    continue;
                }

                //LETTURA FILE E NUMERO EMAIL
                Console.WriteLine($"\nLettura del file: {excelPath}\n");
                List<string> emailAddresses = ReadEmailsFromExcel(excelPath);
                
                Console.WriteLine($"\nTrovate {emailAddresses.Count} email da inviare.");
                if (emailAddresses.Count == 0)
                {
                    Console.WriteLine("\nATTENZIONE!!! Nessuna email trovata nel file Excel.");
                    Console.ReadKey();
                    continue;
                }

                //INSERIMENTO OGGETTO EMAIL
                while(true)
                {
                    Console.Write("\n--Inserisci l'oggetto dell'email: ");
                    subject = Console.ReadLine();
                    if(string.IsNullOrEmpty(subject))
                    {
                        Console.WriteLine("\nAttenzione!!! inserire almeno un carattere");
                        Console.WriteLine("premi un tasto qualunque per riprovare");
                        Console.ReadKey();
                        continue;
                    }
                    break;                   
                }

                //INSERIMENTO CORPO EMAIL
                while (true)
                {
                    Console.Write("\n--Inserisci il corpo dell'email: ");
                    body = Console.ReadLine();
                    if (string.IsNullOrEmpty(body))
                    {
                        Console.WriteLine("\nAttenzione!!! inserire almeno un carattere");
                        Console.WriteLine("premi un tasto qualunque per riprovare");
                        Console.ReadKey();
                        continue;
                    }
                    break;
                }

                //ALLEGARE UN PDF
                Console.Write("\nVuoi allegare un PDF? (S/N): ");
                bool attachPdf = Console.ReadLine().ToUpper() == "S";

                string pdfPath = "";
                if (attachPdf)
                {
                    while (true) 
                    { 
                        Console.Write("\nInserisci il percorso del file PDF: ");
                        pdfPath = Console.ReadLine().Trim('"');
                        if (!File.Exists(pdfPath))
                        {
                            Console.WriteLine("\nATTENZIONE!!!!! Il file pdf specificato non esiste ò il percorso è sbagliato!");
                            Console.WriteLine("Premi qualunque tasto per riprovare");
                            Console.ReadKey();
                            continue;
                        }
                        break ;
                    }
                }
                //VALIDAZIONE CREDENZIALI
                while (true)
                {
                    // Configurazione email mittente
                    Console.Write("\n--Inserisci l'email del mittente: ");
                    senderEmail = Console.ReadLine();
                    Console.Write("\nInserisci la password dell'email: ");
                    senderPassword = Console.ReadLine();


                    Console.WriteLine("\nInizio invio email...\n");
                    if (string.IsNullOrEmpty(senderEmail) || string.IsNullOrEmpty(senderPassword)) 
                    {
                        Console.WriteLine("\nAttenzione!!! inserire almeno un carattere");
                        Console.WriteLine("premi un tasto qualunque per riprovare");
                        Console.ReadKey();
                        continue;
                    }
                    break;
                }

                int emailInviate = 0;
                int emailFallite = 0;
                //INVIO EMAIL
                foreach (string email in emailAddresses)
                {
                    try
                    {
                        SendEmail(senderEmail, senderPassword, email, subject, body, smtpClient, attachPdf ? pdfPath : null);
                        Console.WriteLine($"\n[SUCCESSO] Email inviata con successo a: {email}");
                        emailInviate++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\n[ERRORE] Errore nell'invio dell'email a {email}: {ex.Message}");
                        emailFallite++;
                    }
                }

                Console.WriteLine($"\n\tEmail inviate n: {emailInviate}");
                Console.WriteLine($"\tEmail fallite n: {emailFallite}");
                Console.WriteLine("\nProcesso completato. Premi un 'E' per uscire.");
                Console.WriteLine("Oppure un altro tasto per ripetere");
                string risposta = Console.ReadLine().ToLower();
                if (risposta == "e")
                {
                    break ;
                }
            }

        }

        //----------------------------------------------------------------------------------------------------------------------------------------------------
        static List<string> ReadEmailsFromExcel(string filePath)
        {
            List<string> emails = new List<string>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("\nATTENZIONE!!! Il file Excel non contiene fogli di lavoro!");
                    return emails;
                }

                var worksheet = package.Workbook.Worksheets[0];

                // Verifica se il foglio contiene righe
                if (worksheet.Dimension == null)
                {
                    Console.WriteLine("\nATTENZIONE!!! Il foglio di lavoro è vuoto!");
                    return emails;
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Cerca la colonna "Email" nella prima riga
                int emailColumnIndex = -1;
                for (int col = 1; col <= colCount; col++)
                {
                    string header = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(header) && header.Equals("Email", StringComparison.OrdinalIgnoreCase))
                    {
                        emailColumnIndex = col;
                        break;
                    }
                }

                if (emailColumnIndex == -1)
                {
                    Console.WriteLine("\nATTENZIONE!!! Non è stata trovata alcuna colonna 'Email' nella prima riga!");
                    return emails;
                }

                Console.WriteLine($"\nColonna 'Email' trovata all'indice: {emailColumnIndex}");

                // Leggi le email dalla colonna identificata
                for (int row = 2; row <= rowCount; row++) // Dalla seconda riga in poi (saltando l'intestazione)
                {
                    string email = worksheet.Cells[row, emailColumnIndex].Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(email))
                    {
                        emails.Add(email);
                        Console.WriteLine($"-Email trovata: {email}");
                    }
                }
            }

            return emails;
        }


        //------------------------------------------------------------------------------------------------------------------
        static void SendEmail(string senderEmail, string senderPassword, string recipientEmail,
                            string subject, string body, SmtpClient smtpClient, string pdfPath = null)
        {
            using (smtpClient)
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
        //-----------------------------------------------------------------------------------------------------------------------
    }
}
