using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Importar_Impuestos.App
{
    public class DtoRespuestaImpuestos
    {
        public string RFC { get; set; }
        public string Fecha { get; set; }
        public string Mes { get; set; }
        public string Anio { get; set; }
        public string Iva { get; set; }
        public string Isr { get; set; }
        public bool seImporto { get; set; }
        public string Estatus { get; set; }
       

        public DtoRespuestaImpuestos(string rfc, string fecha, string mes, string anio, string iva, string isr, bool seImporto, string estatus)
        {
            RFC = rfc;
            Fecha = fecha;
            Mes = mes;
            Anio = anio;
            Iva = iva;
            Isr = isr;
            this.seImporto = seImporto;
            Estatus = estatus;
            
        }

        public static DtoRespuestaImpuestos Create(string rfc, string fecha, string mes, string anio, string iva, string isr, bool seImporto, string estatus)
        {
            return new DtoRespuestaImpuestos( rfc,  fecha,  mes,  anio,  iva,  isr,  seImporto,  estatus);
        }
    }
}
