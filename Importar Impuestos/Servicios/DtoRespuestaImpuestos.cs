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
        public string Año { get; set; }
        public string Iva { get; set; }
        public string Isr { get; set; }
        public bool seImporto { get; set; }
        public string Estatus { get; set; }
       

        public DtoRespuestaImpuestos(string rFC, string fecha, string mes, string año, string iva, string isr, bool seImporto, string estatus)
        {
            RFC = rFC;
            Fecha = fecha;
            Mes = mes;
            Año = año;
            Iva = iva;
            Isr = isr;
            this.seImporto = seImporto;
            Estatus = estatus;
            
        }

        public static DtoRespuestaImpuestos Create(string rFC, string fecha, string mes, string año, string iva, string isr, bool seImporto, string estatus)
        {
            return new DtoRespuestaImpuestos( rFC,  fecha,  mes,  año,  iva,  isr,  seImporto,  estatus);
        }
    }
}
