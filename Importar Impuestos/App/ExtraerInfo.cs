using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Importar_Impuestos.App
{
    public  class ExtraerInformacionExcel<DtoExtraerInfo, Entidad> : IRepositorioExcel<DtoExtraerInfo, Entidad> where DtoExtraerInfo : new()
    {
        public IEnumerable<DtoImpuesto> ObtenerInformacionSAX(MemoryStream archivo, Guid idNegocio, Guid idSucursal, List<string> columnasEsperadas)
        {
            //string ruta = @"C:\pruebaArchivos\pruebaBot.xlsx";
            List<DtoImpuesto> conjuntoCatalogo = new List<DtoImpuesto>();
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(archivo, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    //Obtenemos el workbookpart mediante el nombre de la hoja a leer
                    Worksheet sheet = workbookPart.WorksheetParts.First().Worksheet;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    //WorksheetPart worksheetPart = UtileriasOpenXML.ObtenerWorkSheetPart(spreadsheetDocument, "Catalogo");
                    //Se crea el lector de openxml que trabajara con el workpart seleccionado
                    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                    //Se leen los nombres de las columnas y se guardan en una lista
                    List<string> nombresColumnas = ObtenerNombresColumnas(archivo, columnasEsperadas);

                    while (reader.Read())
                    {
                        //Si uno de los elementos es de tipo "Row" se procede a realizar la obtención de los valores de las celdas
                        if (reader.ElementType == typeof(Row))
                        {
                            //Nos brincamos el primer renglon que es del nombre de las columnas
                            if (reader.Attributes[0].Value == "1")
                                reader.ReadNextSibling();

                            var cargaCatalogo = new DtoImpuesto();
                            List<Cell> celdas = new List<Cell>();
                            int currentCount = 0;

                            //Leemos el primer elemento hijo del Row que es la primer celda
                            reader.ReadFirstChild();

                            //Mientras el row tenga elementos hijo se realizara lo siguiente
                            do
                            {
                                //Si el tipo de elemento es "Cell"
                                if (reader.ElementType == typeof(Cell))
                                {
                                    //Se hace un cast del elemento actual a cell
                                    Cell c = (Cell)reader.LoadCurrentElement();

                                    //bloque comentado porque no lo uso ni tengo las utilerias
                                    ////Se obtiene el nombre de la columna
                                    //string columnName = UtileriasOpenXML.GetColumnName(c.CellReference);

                                    ////Se obtiene el index de la columna

                                    //int currentColumnIndex = UtileriasOpenXML.ConvertColumnNameToNumber(columnName);

                                    ////Detecta celdas faltantes y rellena el objeto con un valor vacio (Cuando la cuentaactual es menor al conteo del index actual)

                                    //for (; currentCount < currentColumnIndex; currentCount++)
                                    //{
                                    //    MapearValorACargaCatalogo(cargaCatalogo, "", currentCount, nombresColumnas);
                                    //}

                                    //Obtiene el valor de la celda si cellvalue no es nulo en caso contrario lo deja vacio
                                    string cellValue= cellValue = c.CellValue != null ? c.CellValue.InnerText : ""; ;
                                    
                                    
                                    //Si el valor de la celda se encuentra dentro de la tabla sharedstring obtiene el valor mediante la referencia en cellValue
                                    if (c.DataType != null && c.CellValue != null)
                                    {
                                        if (c.DataType == CellValues.SharedString)
                                        {
                                            cellValue = workbookPart.SharedStringTablePart.SharedStringTable.ElementAt(Int32.Parse(c.CellValue.InnerText)).InnerText;
                                        }
                                        
                                    }

                                    //Mapea el valor al objeto
                                    OnMapearValorACargaCatalogo(cargaCatalogo, cellValue.Trim(), currentCount, nombresColumnas);
                                    currentCount++;
                                }

                            }
                            while (reader.ReadNextSibling());

                            //Detecta celdas faltantes al final del renglón y rellena el ojeto con un valor vacio (Cuando la cuentaactual es menor al conteo del index actual)
                            while (currentCount < nombresColumnas.Count)
                            {
                                //MapearValorACargaCatalogo(cargaCatalogo, "", currentCount, nombresColumnas);
                                currentCount++;
                            }
                            //Se agrega el objeto al conjunto
                            conjuntoCatalogo.Add(cargaCatalogo);
                        }
                    }
                }
            }
           
            catch (Exception ex)
            {
                if (ex != null && string.IsNullOrEmpty(ex.Message) && ex.Message.Contains("was not in a correct format"))
                    throw new Exception("Imposible utilizar archivo Excel: No fue posible realizar la lectura de la hoja catálogo debido a inconsistencias de formato, favor de revisar que las columnas tengan los formatos correctos.");
                else
                    throw new Exception(string.Format("Imposible utilizar archivo excel: No fue posible realizar la lectura de la hoja catálogo debido a {0}", ex.Message));
            }
            //return OnMapDtoEntidad(conjuntoCatalogo, idNegocio, idSucursal);
            return conjuntoCatalogo;
        }

        #region "Mapeo a Entidades"
        public  void OnMapearValorACargaCatalogo(DtoImpuesto cargaCatalogo, string valor, int indice, List<string> nombresColumnas)
        {
            //cargaCatalogo = new DtoImpuesto();
            if (indice < nombresColumnas.Count)
                switch (nombresColumnas[indice])
                {
                    case "RFC":
                        cargaCatalogo.RFC = valor.Replace("\t", string.Empty ); 
                        break;
                    case "Fecha":
                        cargaCatalogo.Fecha = DateTime.FromOADate(Int32.Parse(valor)).ToString("dd/MM/yyyy");
                        break;
                    case "Mes":
                        cargaCatalogo.Mes = valor.Replace("\t", string.Empty);
                        break;
                    case "Año":
                        cargaCatalogo.Año = valor.Replace("\t", string.Empty); 
                        break;
                    case "IVA":
                        cargaCatalogo.IVA = valor.Replace(@"[^\w\s.!@$%^&*()\-\/]+", string.Empty); 
                        break;
                    case "ISR":
                        cargaCatalogo.ISR = valor.Replace("\t", string.Empty); 
                        break;
                    default:
                        break;

                }
        }


        #endregion
        private List<string> ObtenerNombresColumnas(MemoryStream archivo, List<string> columnasEsperadas)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(archivo, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                Worksheet sheet = workbookPart.WorksheetParts.First().Worksheet;


                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                var listaNombresColumnas = new List<String>();
                var listaElementos = new List<String>();
                while (reader.Read())
                {
                    text = reader.GetText();
                    if (reader.ElementType == typeof(Row))
                    {
                        var row = reader.LoadCurrentElement();
                        foreach (Cell item in row.ChildElements)
                        {
                            if (item.DataType != null && item.DataType == CellValues.SharedString)
                            {
                                SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(item.CellValue.InnerText));
                                listaNombresColumnas.Add(ssi.Text.Text);
                            }
                            else
                            {
                                listaNombresColumnas.Add(item.CellValue.InnerText);
                            }
                        }
                        break;
                    }
                }
                var columnasFaltantes = "";
                var index = 0;
                foreach (var columna in columnasEsperadas)
                {
                    if (!listaNombresColumnas.Contains(columna))
                    {
                        columnasFaltantes += $"({columna}),";
                        index++;
                    }
                }

                if (index > 0)
                {
                    if (index > 1)
                        columnasFaltantes = $"Las columnas {columnasFaltantes} no se encuentran en el archivo actual";
                    else
                        columnasFaltantes = $"La columna {columnasFaltantes} no se encuentra en el archivo actual";
                    throw new Exception( $"El formato de la plantilla es incorrecto, {columnasFaltantes}, si necesitas ayuda para elaborar tu archivo, descarga la plantilla de ejemplo.");
                }

                return listaNombresColumnas;
            }
        }
    }
}
