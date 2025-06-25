using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Protonmail_Selenium_Exporter
{
    internal class Correo
    {
        public string asunto { get; set; }
        public string remitente { get; set; }
        public string destinatario { get; set; }
        public string location { get; set; }
        public string size { get; set; }

        public DateTime fecha { get; set; }


        public Correo(string remitente, string destinatario, string location, string size, DateTime fecha,string asunto)
        {
            this.asunto = asunto;
            this.remitente = remitente;
            this.destinatario = destinatario;
            this.location = location;
            this.size = size;
            this.fecha = fecha;
        }
    }


}

