using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Importar_Impuestos.App
{
    public class DtoRespuestaImportarImpuestosMes
    {
        public IEnumerable<DtoRespuestaImpuestos> impuestos { get; set; }
        public bool ExisteError { get; set; }
        public string ArchivoError { get; set; }

        public DtoRespuestaImportarImpuestosMes(IEnumerable<DtoRespuestaImpuestos> impuestos, bool existeError, string archivoError)
        {
            this.impuestos = impuestos;
            ExisteError = existeError;
            ArchivoError = archivoError;
        }

        public static DtoRespuestaImportarImpuestosMes Create(IEnumerable<DtoRespuestaImpuestos> impuestos, bool existeError, string archivoErro)
        {
            return new DtoRespuestaImportarImpuestosMes( impuestos,  existeError,  archivoErro);
        }
    }
}
