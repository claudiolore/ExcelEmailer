using System;
using System.IO;
using OfficeOpenXml;
using System.Net.Mail;
using System.Net.Mime;
using System.Collections.Generic;
using email_sender;
using System.Net;

namespace ExcelEmailSender
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true) 
            { 
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                
                string subject = "SUPPORTO INFORMATICO PER CONTABILIZZAZIONE CALORE E MONITORAGGIO ENERGETICO";
                //string body = "È noto che i sistemi di contabilizzazione del calore per gli impianti di riscaldamento centralizzato, resi obbligatori a partire dal 2014,\r\nattualmente devono anche:\r\n essere leggibili da remoto,\r\n prevedere la fornitura agli utenti finali dei dati di consumo almeno una volta al mese.\r\n(ai sensi del D.lgs 102/2014, per maggiori dettagli si veda in allegato)\r\nRispondere a questi nuovi obblighi può essere oneroso. Non adeguare la contabilizzazione ai predetti requisiti può far perdere clienti.\r\nLa soluzione PLANERGY® che vi presentiamo è l'opportunità di innovare, senza investimenti, il servizio di contabilizzazione da voi\r\nerogato ai condomìni vostri clienti.\r\nCon la nostra soluzione, la vostra azienda potrà:\r\n fidelizzare i condomìni clienti,\r\n adeguare il servizio di contabilizzazione offerto, rispondendo agli obblighi di accesso da remoto e di fornitura dati minima\r\n(almeno una volta al mese),\r\n fornire un servizio di contabilizzazione che include un innovativo monitoraggio via web.\r\nIn allegato trova una breve descrizione dell'applicativo web PLANERGY® che la sua azienda potrà offrire ai suoi clienti.\r\nLa sua attivazione potrà avvenire con il nostro supporto informatico, basato su una Convenzione che prevede, da parte nostra:\r\n configurazione software degli apparati di raccolta e trasmissione dati,\r\n fornitura degli accessi per gli utenti finali all'applicativo web,\r\n elaborazione dati per i piani di riparto annuali (il vostro onere si ridurrà alla sola impaginazione finale personalizzata, se\r\nnecessaria, ed alla consegna all'amministratore del condominio),\r\n patto di non concorrenza (ci impegniamo a non effettuare alcuna azione commerciale rivolta ai suoi clienti in concorrenza\r\ncon la vostra azienda).\r\nSe interessato, risponda per favore a questa mail, fornendo il nominativo da contattare per un incontro di approfondimento, anche\r\nper verificare insieme quali apparati di raccolta e trasmissione dati (attualmente presenti ovvero che potreste proporre ai condomìni\r\ninteressati) siano adatti all'attivazione del nuovo servizio.\r\nIn attesa di un vostro gradito riscontro, porgiamo distinti saluti.";
                string body = @"È noto che i sistemi di contabilizzazione del calore per gli impianti di riscaldamento centralizzato, resi obbligatori a partire dal 2014, 
attualmente devono anche:

• essere leggibili da remoto,
• prevedere la fornitura agli utenti finali dei dati di consumo almeno una volta al mese.

(ai sensi del D.lgs 102/2014, per maggiori dettagli si veda in allegato)

Rispondere a questi nuovi obblighi può essere oneroso. Non adeguare la contabilizzazione ai predetti requisiti può far perdere clienti.

La soluzione PLANERGY® che vi presentiamo è l'opportunità di innovare, senza investimenti, il servizio di contabilizzazione da voi 
erogato ai condomìni vostri clienti.

Con la nostra soluzione, la vostra azienda potrà:
• fidelizzare i condomìni clienti,
• adeguare il servizio di contabilizzazione offerto, rispondendo agli obblighi di accesso da remoto e di fornitura dati minima 
  (almeno una volta al mese),
• fornire un servizio di contabilizzazione che include un innovativo monitoraggio via web.

In allegato trova una breve descrizione dell'applicativo web PLANERGY® che la sua azienda potrà offrire ai suoi clienti.

La sua attivazione potrà avvenire con il nostro supporto informatico, basato su una Convenzione che prevede, da parte nostra:
• configurazione software degli apparati di raccolta e trasmissione dati,
• fornitura degli accessi per gli utenti finali all'applicativo web,
• elaborazione dati per i piani di riparto annuali (il vostro onere si ridurrà alla sola impaginazione finale personalizzata, se
  necessaria, ed alla consegna all'amministratore del condominio),
• patto di non concorrenza (ci impegniamo a non effettuare alcuna azione commerciale rivolta ai suoi clienti in concorrenza
  con la vostra azienda).

Se interessato, risponda per favore a questa mail, fornendo il nominativo da contattare per un incontro di approfondimento, anche
per verificare insieme quali apparati di raccolta e trasmissione dati (attualmente presenti ovvero che potreste proporre ai condomìni
interessati) siano adatti all'attivazione del nuovo servizio.

In attesa di un vostro gradito riscontro, porgiamo distinti saluti.";


                string senderEmail = "claudio.lore001@gmail.com";
                string senderPassword;
                string outputPdfPath = "C:\\Users\\claud\\OneDrive\\Desktop\\planergy utili\\appoggio\\SUPPORTO INFORMATICO PLANERGY CONTABILIZZAZIONE.pdf";
                string pdfPath = "C:\\Users\\claud\\OneDrive\\Desktop\\planergy utili\\finale__SUPPORTO INFORMATICO PLANERGY CONTABILIZZAZIONE.pdf";

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
                List<Azienda> listaAziende = ReadAziendeFromExcel(excelPath);

                //LETTURA FILE E NUMERO EMAIL
                Console.WriteLine($"\nLettura del file: {excelPath}\n");
                List<string> emailAddresses = GetEmails(listaAziende);
                
                Console.WriteLine($"\nTrovate {emailAddresses.Count} email da inviare.");
                if (emailAddresses.Count == 0)
                {
                    Console.WriteLine("\nATTENZIONE!!! Nessuna email trovata nel file Excel.");
                    Console.ReadKey();
                    continue;
                }

                //INSERIMENTO OGGETTO EMAIL
                //while(true)
                //{
                //    Console.Write("\n--Inserisci l'oggetto dell'email: ");
                //    subject = Console.ReadLine();
                //    if(string.IsNullOrEmpty(subject))
                //    {
                //        Console.WriteLine("\nAttenzione!!! inserire almeno un carattere");
                //        Console.WriteLine("premi un tasto qualunque per riprovare");
                //        Console.ReadKey();
                //        continue;
                //    }
                //    break;                   
                //}

                ////INSERIMENTO CORPO EMAIL
                //while (true)
                //{
                //    Console.Write("\n--Inserisci il corpo dell'email: ");
                //    body = Console.ReadLine();
                //    if (string.IsNullOrEmpty(body))
                //    {
                //        Console.WriteLine("\nAttenzione!!! inserire almeno un carattere");
                //        Console.WriteLine("premi un tasto qualunque per riprovare");
                //        Console.ReadKey();
                //        continue;
                //    }
                //    break;
                //}

                //ALLEGARE UN PDF
                Console.Write("\nVuoi allegare un PDF? (S/N): ");
                bool attachPdf = Console.ReadLine().ToUpper() == "S";

                
                //if (attachPdf)
                //{
                //    while (true) 
                //    { 
                //        Console.Write("\nInserisci il percorso del file PDF: ");
                //        pdfPath = Console.ReadLine().Trim('"');
                //        if (!File.Exists(pdfPath))
                //        {
                //            Console.WriteLine("\nATTENZIONE!!!!! Il file pdf specificato non esiste ò il percorso è sbagliato!");
                //            Console.WriteLine("Premi qualunque tasto per riprovare");
                //            Console.ReadKey();
                //            continue;
                //        }
                //        break ;
                //    }
                //}

                //VALIDAZIONE CREDENZIALI
                while (true)
                {
                    // Configurazione email mittente
                    //Console.Write("\n--Inserisci l'email del mittente: ");
                    //senderEmail = Console.ReadLine();
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
                        string pdfDaModificare = pdfPath;

                        Azienda azienda = listaAziende.FirstOrDefault(a => a.Email.Equals(email, StringComparison.OrdinalIgnoreCase));

                        if (azienda == null)
                        {
                            Console.WriteLine($"\n[AVVISO] Nessuna azienda trovata per l'email: {email}");
                            continue;
                        }

                        string name = azienda.Nome;       // Nome dell'azienda
                        string address = azienda.Indirizzo; // Indirizzo dell'azienda

                        PdfFormFiller.FillPdf(pdfDaModificare, outputPdfPath, name, address, email);

                        SendEmail(senderEmail, senderPassword, email, subject, body,  attachPdf ? outputPdfPath : null);
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
        //------------------------------------------------------------------------------------------------------------------
        static void SendEmail(string senderEmail, string senderPassword, string recipientEmail,
                            string subject, string body, string pdfPath = null)
        {
            using(SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587))
            //using (SmtpClient smtpClient = new SmtpClient("localhost", 1025))
            {
                smtpClient.EnableSsl = true;

                //mailhog
                //smtpClient.EnableSsl = false;
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
        public static List<string> GetEmails(List<Azienda> aziende)
        {
            return aziende.Select(a => a.Email).ToList();
        }

    }
}
