using System;
using System.IO;
using OfficeOpenXml;
using System.Net.Mail;
using System.Net.Mime;
using System.Collections.Generic;
using email_sender;
using MimeKit;
using MimeKit.Text;
using MailKit.Security;
using MailKit.Net.Smtp;
using SmtpClient = MailKit.Net.Smtp.SmtpClient;

namespace ExcelEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                string subject = string.Empty;
                string body = string.Empty;
                string excelPath = string.Empty;
                string outputPdfPath = string.Empty;
                string pdfPath = string.Empty;
                int emailInviate = 0;
                int emailFallite = 0;
                string risposta = string.Empty;


                // Banner di benvenuto
                WriteColoredText(ConsoleColor.Cyan, "\n\t\t\t\t╔═══════════════════════╗\n");
                WriteColoredText(ConsoleColor.Cyan, "\t\t\t\t║   Excel Email Sender  ║\n");
                WriteColoredText(ConsoleColor.Cyan, "\t\t\t\t║      BENVENUTO        ║\n");
                WriteColoredText(ConsoleColor.Cyan, "\t\t\t\t╚═══════════════════════╝\n\n");

                //VALIDAZIONE CREDENZIALI
                var credentialValidator = new CredentialValidator();
                var (senderEmail, senderPassword) = credentialValidator.ValidateCredentials();

                WriteColoredText(ConsoleColor.Yellow, "--Inserisci il percorso del file Excel in cui ci sono le email a cui inviare: ");
                excelPath = Console.ReadLine().Trim('"');

                if (string.IsNullOrEmpty(excelPath))
                {
                    WriteColoredText(ConsoleColor.Red, "\n⚠ ATTENZIONE! Devi inserire almeno un carattere!\n");
                    WriteColoredText(ConsoleColor.Red, "Premi qualunque tasto per riprovare\n");
                    Console.ReadKey();
                    continue;
                }

                if (!File.Exists(excelPath))
                {
                    WriteColoredText(ConsoleColor.Red, "\n⚠ ATTENZIONE! Il file Excel specificato non esiste o il percorso è sbagliato!\n");
                    WriteColoredText(ConsoleColor.Red, "Premi qualunque tasto per riprovare\n");
                    Console.ReadKey();
                    continue;
                }

                //CARICAMENTO INFORMAZIONI AZIENDE
                List<Azienda> listaAziende = ReadAziendeFromExcel(excelPath);

                WriteColoredText(ConsoleColor.Green, $"\nLettura del file: {excelPath}\n\n");

                List<string> emailAddresses = GetEmails(listaAziende);

                WriteColoredText(ConsoleColor.Cyan, $"\nTrovate {emailAddresses.Count} email da inviare.\n");

                if (emailAddresses.Count == 0)
                {
                    WriteColoredText(ConsoleColor.Red, "\n⚠ ATTENZIONE! Nessuna email trovata nel file Excel.\n");
                    Console.ReadKey();
                    continue;
                }

                //INSERIMENTO OGGETTO EMAIL
                while (true)
                {
                    WriteColoredText(ConsoleColor.Yellow, "\n--Inserisci l'oggetto dell'email: ");
                    subject = Console.ReadLine();

                    if (string.IsNullOrEmpty(subject))
                    {
                        WriteColoredText(ConsoleColor.Red, "\n⚠ ATTENZIONE! Inserire almeno un carattere\n");
                        WriteColoredText(ConsoleColor.Red, "Premi un tasto qualunque per riprovare\n");
                        Console.ReadKey();
                        continue;
                    }
                    break;
                }

                //INSERIMENTO CORPO EMAIL
                while (true)
                {
                    WriteColoredText(ConsoleColor.Yellow, "\n--Inserisci il corpo dell'email: ");
                    body = Console.ReadLine();

                    if (string.IsNullOrEmpty(body))
                    {
                        WriteColoredText(ConsoleColor.Red, "\n⚠ ATTENZIONE! Inserire almeno un carattere\n");
                        WriteColoredText(ConsoleColor.Red, "Premi un tasto qualunque per riprovare\n");
                        Console.ReadKey();
                        continue;
                    }
                    break;
                }

                //ALLEGARE UN PDF
                WriteColoredText(ConsoleColor.Yellow, "\nVuoi allegare un PDF? (S/N): ");
                bool attachPdf = Console.ReadLine().ToUpper() == "S";

                if (attachPdf)
                {
                    while (true)
                    {
                        WriteColoredText(ConsoleColor.Yellow, "\nInserisci il percorso del file PDF (deve avere una colonna nominativo, email e indirizzo): ");
                        pdfPath = Console.ReadLine().Trim('"');

                        if (string.IsNullOrEmpty(pdfPath) || !pdfPath.ToLower().EndsWith(".pdf"))
                        {
                            WriteColoredText(ConsoleColor.Red, "\n⚠ ATTENZIONE! Devi selezionare un file PDF valido!\n");
                            continue;
                        }

                        if (!File.Exists(pdfPath))
                        {
                            WriteColoredText(ConsoleColor.Red, "\n⚠ ATTENZIONE! Il file PDF specificato non esiste o il percorso è sbagliato!\n");
                            continue;
                        }

                        WriteColoredText(ConsoleColor.Green, $"\n[✓] File PDF trovato: {pdfPath}\n");

                        break;
                    }
                }

                PrintResoconto(senderEmail, excelPath, pdfPath, subject, body);

                WriteColoredText(ConsoleColor.Yellow, "\nVuoi continuare? (s/n)\n");

                risposta = string.Empty;
                risposta = Console.ReadLine().ToLower();

                if (!risposta.Equals("s"))
                {
                    break;
                }

                //INVIO EMAIL
                foreach (string email in emailAddresses)
                {
                    try
                    {
                        Azienda azienda = listaAziende.FirstOrDefault(a => a.Email.Equals(email, StringComparison.OrdinalIgnoreCase));

                        if (azienda == null)
                        {
                            WriteColoredText(ConsoleColor.Red, $"\n[⚠] Nessuna azienda trovata per l'email: {email}\n");
                            continue;
                        }

                        string name = azienda.Nome;
                        string address = azienda.Indirizzo;

                        if (attachPdf)
                        {
                            // Genera il PDF in memoria come byte array
                            byte[] pdfBytes = PdfMemoryHandler.GeneratePdfInMemory(pdfPath, name, address, email);

                            // Invia email con il PDF in byte
                            PdfMemoryHandler.SendEmailWithPdfBytes(
                                senderEmail,
                                senderPassword,
                                email,
                                subject,
                                body,
                                pdfBytes,
                                $"{Path.GetFileName(pdfPath)}.pdf"
                            );
                        }
                        else
                        {
                            SendEmail(senderEmail, senderPassword, email, subject, body);
                        }

                        WriteColoredText(ConsoleColor.Green, $"\n[✓] Email inviata con successo a: {email}\n");
                        emailInviate++;
                    }
                    catch (Exception ex)
                    {
                        WriteColoredText(ConsoleColor.Red, $"\n[✗] Errore nell'invio dell'email a {email}: {ex.Message}\n");
                        emailFallite++;
                    }
                }

                PrintFinalReport(emailInviate, emailFallite);
                risposta = Console.ReadLine().ToLower();

                if (risposta == "e")
                {
                    break;
                }
            }
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
        public static List<Azienda> ReadAziendeFromExcel(string filePath)
        {
            List<Azienda> aziende = new List<Azienda>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("\nATTENZIONE!!! Il file Excel non contiene fogli di lavoro!");
                    return aziende;
                }

                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                {
                    Console.WriteLine("\nATTENZIONE!!! Il foglio di lavoro è vuoto!");
                    return aziende;
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                int nomeColumnIndex = -1;
                int indirizzoColumnIndex = -1;
                int emailColumnIndex = -1;

                // Identifica le colonne necessarie
                for (int col = 1; col <= colCount; col++)
                {
                    string header = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(header))
                    {
                        if (header.Equals("Nominativo", StringComparison.OrdinalIgnoreCase))
                            nomeColumnIndex = col;
                        else if (header.Equals("Indirizzo", StringComparison.OrdinalIgnoreCase))
                            indirizzoColumnIndex = col;
                        else if (header.Equals("Email", StringComparison.OrdinalIgnoreCase))
                            emailColumnIndex = col;
                    }
                }

                if (nomeColumnIndex == -1 || indirizzoColumnIndex == -1 || emailColumnIndex == -1)
                {
                    Console.WriteLine("\nATTENZIONE!!! Non tutte le colonne richieste ('Nome', 'Indirizzo', 'Email') sono presenti!");
                    return aziende;
                }

                // Leggi i dati riga per riga
                for (int row = 2; row <= rowCount; row++)
                {
                    string nome = worksheet.Cells[row, nomeColumnIndex].Value?.ToString()?.Trim();
                    string indirizzo = worksheet.Cells[row, indirizzoColumnIndex].Value?.ToString()?.Trim();
                    string email = worksheet.Cells[row, emailColumnIndex].Value?.ToString()?.Trim();

                    if (!string.IsNullOrEmpty(email))
                    {
                        aziende.Add(new Azienda
                        {
                            Id = Guid.NewGuid(),
                            Nome = nome,
                            Indirizzo = indirizzo,
                            Email = email
                        });
                    }
                }
            }

            return aziende;
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
        static void SendEmail(string senderEmail, string senderPassword, string recipientEmail,
                     string subject, string body, string pdfPath = null)
        {
            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("", senderEmail));
                message.To.Add(new MailboxAddress("", recipientEmail));
                message.Subject = subject;
                var builder = new BodyBuilder();
                builder.TextBody = body;

                if (!string.IsNullOrEmpty(pdfPath))
                {
                    builder.Attachments.Add(pdfPath);
                }
                message.Body = builder.ToMessageBody();

                using (var client = new SmtpClient())
                {
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;

                    // Connessione con timeout esteso
                    //client.Connect("smtps.aruba.it", 465, SecureSocketOptions.SslOnConnect);
                    //client.Connect("smtp.gmail.it", 587, SecureSocketOptions.SslOnConnect);
                    client.Connect("localhost", 1025);

                    client.Authenticate(senderEmail, senderPassword);
                     
                    client.Timeout = 20000;

                    client.Send(message);
                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                WriteColoredText(ConsoleColor.Red, "\n[ERRORE DETTAGLIATO] Invio email fallito:\n");
                WriteColoredText(ConsoleColor.Red, $"Messaggio: {ex.Message}\n");
                WriteColoredText(ConsoleColor.Red, $"Stack Trace: {ex.StackTrace}\n");

                if (ex.InnerException != null)
                {
                    WriteColoredText(ConsoleColor.Red, $"Inner Exception: {ex.InnerException.Message}\n");
                }
                throw;
            }
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
        public static List<string> GetEmails(List<Azienda> aziende)
        {
            return aziende.Select(a => a.Email).ToList();
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
        public static void PrintResoconto(string senderEmail, string excelPath, string pdfPath, string subject, string body)
        {
            // RESOCONTO
            WriteColoredText(ConsoleColor.Cyan, "\n\t===== RESOCONTO =====\n");

            WriteColoredText(ConsoleColor.Green, "Email mittente: ");
            Console.WriteLine(senderEmail);

            WriteColoredText(ConsoleColor.Green, "Path Excel: ");
            Console.WriteLine(excelPath);

            WriteColoredText(ConsoleColor.Green, "Path PDF: ");
            Console.WriteLine(pdfPath);

            WriteColoredText(ConsoleColor.Green, "Oggetto email: ");
            Console.WriteLine(subject);

            WriteColoredText(ConsoleColor.Green, "Body: ");
            string shortBody = body.Length > 200 ? body.Substring(0, 200).TrimEnd() : body.TrimEnd();
            Console.WriteLine($"{shortBody}...");
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
        public static void PrintFinalReport(int emailInviate, int emailFallite)
        {
            WriteColoredText(ConsoleColor.Cyan, "\n\t===== REPORT FINALE =====\n");

            WriteColoredText(ConsoleColor.Green, $"\nEmail inviate {emailInviate}\n");

            WriteColoredText(ConsoleColor.Red, $"\nEmail fallite {emailFallite}\n");

            WriteColoredText(ConsoleColor.Green, "\nProcesso completato.\n");

            WriteColoredText(ConsoleColor.Yellow, "Premi 'E' per uscire.\n");
            WriteColoredText(ConsoleColor.Yellow, "Oppure un altro tasto per ripetere.\n");
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
        public static void WriteColoredText(ConsoleColor color, string text)
        {
            var originalColor = Console.ForegroundColor;

            try
            {
                Console.ForegroundColor = color;

                Console.Write(text);
            }
            finally
            {
                Console.ForegroundColor = originalColor;
            }
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------
    }
}
