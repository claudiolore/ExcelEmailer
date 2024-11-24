using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace email_sender
{
    public class Azienda
    {
        public Guid Id { get; set; }
        public string Nome { get; set; }
        public string Indirizzo { get; set; }
        public string Email { get; set; }
    }
}
