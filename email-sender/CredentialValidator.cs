using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
    using System;
    using System.Text.RegularExpressions;


namespace email_sender
{
    public class CredentialValidator
    {
        private const int MinPasswordLength = 6; // Lunghezza minima della password


        public (string email, string password) ValidateCredentials()
        {
            while (true)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("\n--Inserisci l'email del mittente: ");
                Console.ResetColor();
                string senderEmail = Console.ReadLine();

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("Inserisci la password dell'email: ");
                Console.ResetColor();
                string senderPassword = Console.ReadLine();

                // Verifica se i campi sono vuoti
                if (string.IsNullOrEmpty(senderEmail) || string.IsNullOrEmpty(senderPassword))
                {
                    ShowError("⚠ ATTENZIONE! Inserire almeno un carattere.");
                    continue;
                }

                // Validazione formato email
                if (!IsValidEmail(senderEmail))
                {
                    ShowError("⚠ ATTENZIONE! L'indirizzo email non è valido.");
                    continue;
                }

                // Controllo lunghezza password
                if (senderPassword.Length < MinPasswordLength)
                {
                    ShowError($"⚠ ATTENZIONE! La password deve contenere almeno {MinPasswordLength} caratteri.");
                    continue;
                }

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("\nCredenziali validate. Inizio invio email...\n");
                Console.ResetColor();
                return (senderEmail, senderPassword);
            }
        }

        private bool IsValidEmail(string email)
        {
            var emailRegex = new Regex(@"^[^@\s]+@[^@\s]+\.[^@\s]+$");
            return emailRegex.IsMatch(email);
        }


        private void ShowError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n{message}");
            Console.WriteLine("Premi un tasto qualunque per riprovare.");
            Console.ResetColor();
            Console.ReadKey();
        }
    }

}
