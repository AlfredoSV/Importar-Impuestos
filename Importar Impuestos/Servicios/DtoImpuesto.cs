using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Importar_Impuestos.App
{
    public class DtoImpuesto
    {
        public string RFC { get; set; }
        public string Fecha { get; set; }
        public string Mes { get; set; }
        public string Anio { get; set; }
        public string IVA { get; set; }
        public string ISR { get; set; }

    }
}
