using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Importar_Impuestos.App;

namespace Importar_Impuestos.Servicios
{
    public interface IservicioExcel<DtoExtraerInfo> where DtoExtraerInfo : new()
    {
        IEnumerable<DtoImpuesto> ObtenerInformacionExcel(MemoryStream archivo,  List<string> columnasEsperadas);
        List<string> ObtenerNombresColumnas(MemoryStream archivo, List<string> columnasEsperadas);
        void OnMapearValoresAImpuestos(DtoImpuesto cargaCatalogo, string valor, int indice, List<string> nombresColumnas);
        void GenerarExcelUnaHoja(Stream rutaNueva, string[] NombreHeaders, object[,] Valores, string nombreHoja);
        string CrearHoja(SpreadsheetDocument xl, List<object[,]> Valores, List<String[]> NombreHeaders, List<String> nombreHoja, List<String[]> NombreHeadersSheet, List<CellValues[]> LCellValues, List<UInt32[]> LEstilo);
        String GenerarExcelMultiplesHojas(String nombreFinalArchivo, String rutaFinalArchivo, List<String> nombreHoja, List<object[,]> Valores, List<String[]> NombreHeadersSheet);
    }
}
