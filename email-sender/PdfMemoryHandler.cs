using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Forms;
using iText.Forms.Fields;
using MailKit.Net.Smtp;
using MimeKit;
using MimeKit.Text;

namespace email_sender
{
    public static class PdfMemoryHandler
    {
        public static byte[] GeneratePdfInMemory(string inputPdfPath, string name, string address, string email)
        {
            try
            {
                if (!File.Exists(inputPdfPath))
                    throw new FileNotFoundException("Il file PDF non esiste", inputPdfPath);

                byte[] pdfBytes = File.ReadAllBytes(inputPdfPath);

                using (MemoryStream inputStream = new MemoryStream(pdfBytes))
                using (MemoryStream outputStream = new MemoryStream())
                {
                    using (PdfReader pdfReader = new PdfReader(inputStream))
                    using (PdfWriter pdfWriter = new PdfWriter(outputStream))
                    using (PdfDocument pdfDoc = new PdfDocument(pdfReader, pdfWriter))
                    {
                        PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);


                        var fields = form.GetFormFields();

                        if (fields.ContainsKey("nome"))
                            fields["nome"].SetValue(name.ToUpper());
                        if (fields.ContainsKey("indirizzo"))
                            fields["indirizzo"].SetValue(address);
                        if (fields.ContainsKey("email"))
                            fields["email"].SetValue(email.ToLower());

                        form.FlattenFields();
                    }

                    return outputStream.ToArray();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore durante la generazione del PDF in memoria: {ex.Message}");
                throw;
            }
        }

        public static void SendEmailWithPdfBytes(
            string senderEmail,
            string senderPassword,
            string recipientEmail,
            string subject,
            string body,
            byte[] pdfBytes,
            string pdfFileName)
        {
            try
            {
                var message = new MimeMessage();
                //message.From.Add(new MailboxAddress("Planergy", senderEmail));
                message.From.Add(new MailboxAddress("", senderEmail));
                message.To.Add(new MailboxAddress("", recipientEmail));
                message.Subject = subject;

                var builder = new BodyBuilder();
                builder.TextBody = body;

                // Aggiungi il PDF come allegato
                if (pdfBytes != null && pdfBytes.Length > 0)
                {
                    builder.Attachments.Add(pdfFileName, pdfBytes);
                }

                message.Body = builder.ToMessageBody();

                using (var client = new SmtpClient())
                {
                    client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    
                    client.Connect("localhost", 1025);
                    //client.Connect("smtps.aruba.it", 465, true);
                    //client.Connect("smtp.gmail.it", 587, true);

                    client.Authenticate(senderEmail, senderPassword);

                    client.Timeout = 20000;

                    client.Send(message);
                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore nell'invio email: {ex.Message}");
                throw;
            }
        }
    }
}