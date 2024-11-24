using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
    using System.IO;
    using iText.Kernel.Pdf;
    using iText.Forms;
    using iText.Forms.Fields;
    using System.Reflection.PortableExecutable;

namespace email_sender
{

    public class PdfFormFiller
    {
        public static void FillPdf(string inputPdfPath, string outputPdfPath, string name, string address, string email)
        {
            // Apri il PDF di input
            using (PdfReader pdfReader = new PdfReader(inputPdfPath))
            using (PdfWriter pdfWriter = new PdfWriter(outputPdfPath))
            using (PdfDocument pdfDoc = new PdfDocument(pdfReader, pdfWriter))
            {
                // Ottieni il modulo del PDF
                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);

                // Recupera i campi del modulo
                IDictionary<string, PdfFormField> fields = form.GetFormFields();

                // Compila i campi
                if (fields.ContainsKey("nome"))
                    fields["nome"].SetValue(name.ToUpper());

                if (fields.ContainsKey("indirizzo"))
                    fields["indirizzo"].SetValue(address.ToLower());

                if (fields.ContainsKey("email"))
                    fields["email"].SetValue(email.ToLower());

                // Rendi i campi non modificabili (opzionale)
                form.FlattenFields();

                // Chiudi il documento
            }
        }
    }

}
