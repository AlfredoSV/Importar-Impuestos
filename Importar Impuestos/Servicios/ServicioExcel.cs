using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Importar_Impuestos.App;

namespace Importar_Impuestos.Servicios
{
    public class ServicioExcel : IservicioExcel<DtoImpuesto>
    {
        public IEnumerable<DtoImpuesto> ObtenerInformacionExcel(MemoryStream archivo, List<string> columnasEsperadas)
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
                                    string cellValue = cellValue = c.CellValue != null ? c.CellValue.InnerText : ""; ;


                                    //Si el valor de la celda se encuentra dentro de la tabla sharedstring obtiene el valor mediante la referencia en cellValue
                                    if (c.DataType != null && c.CellValue != null)
                                    {
                                        if (c.DataType == CellValues.SharedString)
                                        {
                                            cellValue = workbookPart.SharedStringTablePart.SharedStringTable.ElementAt(Int32.Parse(c.CellValue.InnerText)).InnerText;
                                        }

                                    }

                                    //Mapea el valor al objeto
                                    OnMapearValoresAImpuestos(cargaCatalogo, cellValue.Trim(), currentCount, nombresColumnas);
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
                    throw new Exception(string.Format("Imposible utilizar archivo excel: No fue posible realizar la lectura de la hoja impuestos debido a {0}", ex.Message));
            }
            //return OnMapDtoEntidad(conjuntoCatalogo, idNegocio, idSucursal);
            return conjuntoCatalogo;
        }

        public void OnMapearValoresAImpuestos(DtoImpuesto impuesto, string valor, int indice, List<string> nombresColumnas)
        {
            //cargaCatalogo = new DtoImpuesto();
            if (indice < nombresColumnas.Count)
                switch (nombresColumnas[indice])
                {
                    case "RFC":
                        impuesto.RFC = Regex.Replace(valor.ToString(), @"[^\w\.@-]", string.Empty);
                        break;
                    case "Fecha":
                        impuesto.Fecha = valor.ToString();//DateTime.FromOADate(Int32.Parse(valor)).ToString("dd/MM/yyyy");//.ToString("MM/dd/yyyy");
                        break;
                    case "Mes":
                        impuesto.Mes = Regex.Replace(valor.ToString(), @"[^\w\.@-]", string.Empty);
                        break;
                    case "Año":
                        impuesto.Anio = Regex.Replace(valor.ToString(), @"[^\w\.@-]", string.Empty);
                        break;
                    case "IVA":
                        impuesto.IVA = Regex.Replace(valor.ToString(), @"[^\w\.@-]", string.Empty);
                        break;
                    case "ISR":
                        impuesto.ISR = Regex.Replace(valor.ToString(), @"[^\w\.@-]", string.Empty);
                        break;
                    default:
                        break;

                }
        }

        public List<string> ObtenerNombresColumnas(MemoryStream archivo, List<string> columnasEsperadas)
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
                    throw new Exception($"El formato de la plantilla es incorrecto, {columnasFaltantes}, si necesitas ayuda para elaborar tu archivo, descarga la plantilla de ejemplo.");
                }

                return listaNombresColumnas;
            }
        }

        public  void GenerarExcelUnaHoja(Stream rutaNueva, string[] NombreHeaders, object[,] Valores, string nombreHoja)
        {
            if (rutaNueva == null) throw new ArgumentException("Imposible utilizar valor nulo: rutaNueva");
            //if (rutaNueva == "") throw new ArgumentException("Imposible utilizar valor vacio: rutaNueva");
            if (NombreHeaders == null) throw new ArgumentException("Imposible utilizar valor nulo: NombreHeaders");
            if (Valores == null) throw new ArgumentException("Imposible utilizar valor nulo: Valores");
            if (nombreHoja == null) throw new ArgumentException("Imposible utilizar valor nulo: nombreHoja");
            if (nombreHoja == "") throw new ArgumentException("Imposible utilizar valor vacio: nombreHoja");

            SpreadsheetDocument archivoNuevo = null;

            try
            {
                //se crea el archivo
                archivoNuevo = SpreadsheetDocument.Create(rutaNueva, SpreadsheetDocumentType.Workbook, true);
                //  
                archivoNuevo.AddWorkbookPart();
                WorkbookStylesPart stylesPart = archivoNuevo.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                stylesPart.Stylesheet.Save();
                Workbook libro = new Workbook();

                //hojas al libro
                Sheets hojasLibro = new Sheets();
                //hoja nueva
                Sheet hoja = new Sheet { Name = nombreHoja, SheetId = 1, Id = "idHoja1" };

                hojasLibro.Append(hoja);

                libro.Append(hojasLibro);
                archivoNuevo.WorkbookPart.Workbook = libro;

                WorksheetPart areaHoja1 = archivoNuevo.WorkbookPart.AddNewPart<WorksheetPart>("idHoja1");
                areaHoja1.Worksheet = new Worksheet();

                SheetData datos = new SheetData();

                //Headers
                Row Headers = new Row { RowIndex = 1 };
                char numeroDeFilaHeaders = '1';
                char[] letras = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

                ///Asignar los headers
                string[] tFechas = new string[] { "fecha" };
                string[] tCurrency = new string[] { "debe", "haber", "saldo", "monto", "cambio" };

                CellValues[] cellValue = new CellValues[NombreHeaders.Length];

                #region Estilos
                TableParts tp = new TableParts() { Count = 1 };
                TablePart tablePart = new TablePart() { Id = "rId1" };
                tp.Append(tablePart);

                TableDefinitionPart tableDefinitionPart1 = areaHoja1.Worksheet.WorksheetPart.AddNewPart<TableDefinitionPart>("rId1");
                Table table1 = new Table()
                {
                    Id = (UInt32Value)1U,
                    Name = "Tabla1",
                    DisplayName = "Tabla1",
                    Reference = StringValue.FromString("A1:" + letras[NombreHeaders.Length - 1].ToString() + Valores.Length),
                    TotalsRowShown = BooleanValue.FromBoolean(false)
                };

                uint cellRangeCount = (uint)NombreHeaders.Length;
                TableColumns tableColumns1 = new TableColumns() { Count = (UInt32Value)(cellRangeCount) };

                for (uint i = 0; i < NombreHeaders.Length; i++)
                    tableColumns1.Append(new TableColumn() { Id = i + 1, Name = NombreHeaders[i] });

                TableStyleInfo tsi = new TableStyleInfo() { Name = "TableStyleLight9", ShowColumnStripes = BooleanValue.FromBoolean(false), ShowRowStripes = BooleanValue.FromBoolean(true), ShowFirstColumn = BooleanValue.FromBoolean(false), ShowLastColumn = BooleanValue.FromBoolean(false) };
                table1.Append(tableColumns1);
                table1.Append(tsi);

                tableDefinitionPart1.Table = table1;
                #endregion

                for (int i = 0; i < NombreHeaders.Length; i++)
                {
                    Headers.AppendChild(new Cell()
                    {
                        CellReference = letras[i].ToString() + numeroDeFilaHeaders.ToString(),
                        CellValue = new CellValue(NombreHeaders[i]),
                        DataType = CellValues.String
                    });
                }


                datos.Append(Headers);
                ////Escribir los datos en excel

                int numFilas = Valores.Length / NombreHeaders.Length;
                for (int i = 0; i < numFilas; i++)
                {
                    string numColumna = (i + 2).ToString();
                    Row row = new Row { RowIndex = UInt32.Parse(numColumna) };
                    for (int j = 0; j < NombreHeaders.Length; j++)
                    {
                        ///creacion de celda para la fila
                        Cell celda = new Cell();
                        celda.CellReference = letras[j] + numColumna;
                        celda.CellValue = new CellValue(Valores[i, j] == DBNull.Value || Valores[i, j] == null ? string.Empty : Valores[i, j].ToString());
                        celda.DataType = CellValues.String;

                        row.AppendChild(celda);
                    }
                    datos.Append(row);
                }
                areaHoja1.Worksheet.Append(datos);
                areaHoja1.Worksheet.Append(tp);
                archivoNuevo.WorkbookPart.Workbook.Save();
                archivoNuevo.Close();
                
            }
            catch (Exception)
            {
                throw;
            }
            
        }

        public string GenerarExcelMultiplesHojas(String nombreFinalArchivo, String rutaFinalArchivo, List<String> nombreHoja, List<object[,]> Valores, List<String[]> NombreHeadersSheet)
        {
           
            string rutaCompletaExcel = Path.Combine(rutaFinalArchivo, nombreFinalArchivo);
            //se crea el archivo
            SpreadsheetDocument xl = SpreadsheetDocument.Create(rutaCompletaExcel, SpreadsheetDocumentType.Workbook, true);

            xl.AddWorkbookPart();
            WorkbookStylesPart wbsp = xl.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            OpenXmlWriter oxw = OpenXmlWriter.Create(wbsp);

            oxw.WriteStartElement(new Stylesheet());

            oxw.WriteStartElement(new Fonts() { Count = 1, KnownFonts = BooleanValue.FromBoolean(true) });
            oxw.WriteElement(new DocumentFormat.OpenXml.Spreadsheet.Font() { FontSize = new FontSize() { Val = 11 }, Color = new Color() { Theme = 1 }, FontName = new FontName() { Val = "Calibri" }, FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 }, FontScheme = new FontScheme() { Val = FontSchemeValues.Minor } });
            oxw.WriteEndElement(); //FONTS

            oxw.WriteStartElement(new Fills() { Count = 2 });
            oxw.WriteElement(new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.None } });
            oxw.WriteElement(new Fill() { PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 } });
            oxw.WriteEndElement();//FILLS

            oxw.WriteStartElement(new Borders() { Count = 1 });
            oxw.WriteElement(new Border() { StartBorder = new StartBorder(), EndBorder = new EndBorder(), TopBorder = new TopBorder(), BottomBorder = new BottomBorder(), DiagonalBorder = new DiagonalBorder() });
            oxw.WriteEndElement();//BORDERS

            oxw.WriteStartElement(new CellStyleFormats() { Count = 1 });
            oxw.WriteElement(new CellFormat() { BorderId = 0, FillId = 0, NumberFormatId = 0, FontId = 0 });
            oxw.WriteEndElement(); //CELLSTYLEFORMATS

            oxw.WriteStartElement(new CellFormats() { Count = 4 });
            oxw.WriteElement(new CellFormat() { BorderId = 0, FillId = 0, NumberFormatId = 0, FormatId = 0 });
            oxw.WriteElement(new CellFormat() { BorderId = 0, FillId = 0, NumberFormatId = 14, ApplyNumberFormat = BooleanValue.FromBoolean(true), FormatId = 0 });
            oxw.WriteElement(new CellFormat() { BorderId = 0, FillId = 0, NumberFormatId = 49, ApplyNumberFormat = BooleanValue.FromBoolean(true), FormatId = 0 });
            oxw.WriteElement(new CellFormat() { BorderId = 0, FillId = 0, NumberFormatId = 7, ApplyNumberFormat = BooleanValue.FromBoolean(true), FormatId = 0 });
            oxw.WriteEndElement(); //cellformats

            oxw.WriteStartElement(new CellStyles() { Count = 1 });
            oxw.WriteElement(new CellStyle() { BuiltinId = 0, Name = StringValue.FromString("Normal"), FormatId = 0 });
            oxw.WriteEndElement();// CELLSTYLES


            char[] letras = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            string[] tFechas = new string[] { "fecha", "vigencia" };
            string[] tCurrency = new string[] { "debe", "haber", "saldo", "monto", "cambio" };

            List<string[]> NombreHeaders = new List<string[]>();
            for (int yu = 0; yu < NombreHeadersSheet.Count(); yu++)
                NombreHeaders.Add(NombreHeadersSheet[yu]);
            List<CellValues[]> LCellValues = new List<CellValues[]>();
            List<UInt32[]> LEstilo = new List<uint[]>();


            List<DifferentialFormat> dif = new List<DifferentialFormat>();

            for (int k = 0; k < NombreHeaders.Count; k++)
            {
                CellValues[] cellValue = new CellValues[NombreHeaders[k].Length];
                UInt32[] estilo = new UInt32[NombreHeaders[k].Length];
                for (int i = 0; i < NombreHeaders[k].Length; i++)
                {
                    cellValue[i] = CellValues.String;
                    estilo[i] = 2;
                    for (int j = 0; j < tFechas.Length; j++)
                    {
                        if ((NombreHeaders[k])[i].ToLower().Contains(tFechas[j].ToLower()))
                        {
                            cellValue[i] = CellValues.Date;
                            estilo[i] = 1;
                        }
                    }

                    for (int j = 0; j < tCurrency.Length; j++)
                    {
                        if (NombreHeaders[k][i].ToLower().Contains(tCurrency[j].ToLower()))
                        {
                            cellValue[i] = CellValues.Number;
                            estilo[i] = 3;
                        }
                    }
                    if (estilo[i] == 2)
                    {
                        //MODO SAX
                        dif.Add(new DifferentialFormat() { NumberingFormat = new NumberingFormat() { NumberFormatId = 30, FormatCode = StringValue.FromString("@") } }); new DifferentialFormat() { NumberingFormat = new NumberingFormat() { NumberFormatId = 30, FormatCode = StringValue.FromString("@") } };
                    }
                    if (estilo[i] == 1)
                    {
                        //MODO SAX
                        dif.Add(new DifferentialFormat() { NumberingFormat = new NumberingFormat() { NumberFormatId = 19, FormatCode = StringValue.FromString("dd/mm/yyyy") } });
                    }
                    if (estilo[i] == 3)
                    {
                        dif.Add(new DifferentialFormat() { NumberingFormat = new NumberingFormat() { NumberFormatId = 7, FormatCode = StringValue.FromString("\"$\"#,##0.00_);(\"$\"#,##0.00)") } });
                    }
                    LCellValues.Add(cellValue);
                    LEstilo.Add(estilo);
                }
            }
            oxw.WriteStartElement(new DifferentialFormats() { Count = UInt32.Parse(dif.Count.ToString()) });

            foreach (var item in dif)
            {
                oxw.WriteElement(item);

            }

            oxw.WriteEndElement();// DIFFERENTIALFORMATS

            oxw.WriteElement(new DocumentFormat.OpenXml.Spreadsheet.TableStyles() { Count = 0, DefaultTableStyle = StringValue.FromString("TableStyleMedium2"), DefaultPivotStyle = "PivotStyleLight16" });
            oxw.WriteElement(new ExtensionList());

            oxw.WriteEndElement();//stylesheet

            oxw.Close();



            try
            {

                oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                oxw.WriteStartElement(new Workbook());
                oxw.WriteStartElement(new Sheets());

                //Se agregan tantas hojas como nombres haya en la lista de nombreHoja
                for (int mx = 1; mx <= nombreHoja.Count(); mx++)
                {
                    oxw.WriteElement(new Sheet()
                    {
                        Name = nombreHoja[mx - 1],
                        SheetId = (UInt32)mx,
                        Id = CrearHoja(xl, Valores, NombreHeaders, nombreHoja, NombreHeadersSheet, LCellValues, LEstilo)
                    });
                }
                // this is for Sheets
                oxw.WriteEndElement();
                // this is for Workbook
                oxw.WriteEndElement();
                oxw.Close();

                xl.Close();

            }
            catch (Exception ex)
            {

            }
            
            return rutaCompletaExcel;
        }


         public string CrearHoja(SpreadsheetDocument xl, List<object[,]> Valores, List<String[]> NombreHeaders, List<String> nombreHoja, List<String[]> NombreHeadersSheet, List<CellValues[]> LCellValues, List<UInt32[]> LEstilo)
        {
            
            char[] letras = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            int mx = 1;
            WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>("rId" + mx);
            List<OpenXmlAttribute> oxa;
            OpenXmlWriter oxw;
            oxw = OpenXmlWriter.Create(wsp);
            oxw.WriteStartElement(new Worksheet());

            //Agregar Valores
            int[] numFilas = new int[NombreHeaders.Count];// Valores.Length / NombreHeaders.Length;
            numFilas[0] = Valores[mx - 1].Length / NombreHeadersSheet[mx - 1].Count();
            List<SheetDimension> LDimension = new List<SheetDimension>();
            List<string> Ldimension = new List<string>();
            for (int i = 0; i < numFilas.Length; i++)
            {

                var referencia = "";
                if (NombreHeaders[mx - 1].Length > 26)
                    referencia = "A" + letras[NombreHeaders[mx - 1].Length - 1 - 26].ToString();
                else
                    referencia = letras[NombreHeaders[mx - 1].Length - 1].ToString();
                Ldimension.Add("A1:" + referencia + (numFilas[i] + 1).ToString());
                oxw.WriteElement(new SheetDimension() { Reference = StringValue.FromString(Ldimension[i]) });
            }

            oxw.WriteStartElement(new SheetData());

            //--------Partes de la tabla

            TableParts tp = new TableParts() { Count = 1 };
            TablePart tablePart = new TablePart() { Id = "rId" + mx };

            List<Object[,]> LValores = new List<object[,]>();
            LValores.Add(Valores[mx - 1]);

            ///Valores de la tabla
            char numeroDeFilaHeaders = '1';
            for (int k = 0; k < 1; k++)
            {
                Row Headers = new Row { RowIndex = 1 };
                for (int i = 0; i < NombreHeaders[mx - 1].Length; i++)
                {
                    String referencia = "";
                    if (i > 25)
                        referencia = "A" + letras[i - 26].ToString() + numeroDeFilaHeaders.ToString();
                    else
                        referencia = letras[i].ToString() + numeroDeFilaHeaders.ToString();
                    Headers.AppendChild(new Cell()
                    {
                        CellReference = referencia,
                        CellValue = new CellValue(NombreHeaders[mx - 1][i]),
                        DataType = CellValues.String
                    });
                }
                oxw.WriteElement(Headers);

                if (numFilas[k] > 0)
                {
                    for (int i = 0; i < numFilas[k]; i++)
                    {
                        string numColumna = (i + 2).ToString();
                        oxa = new List<OpenXmlAttribute>();
                        // this is the row index
                        oxa.Add(new OpenXmlAttribute("r", null, numColumna.ToString()));
                        Row row = new Row() { RowIndex = UInt32.Parse(numColumna) };
                        oxw.WriteStartElement(row, oxa);


                        for (int j = 0; j < NombreHeaders[mx - 1].Length; j++)
                        {
                            oxa = new List<OpenXmlAttribute>();

                            Cell celda = new Cell();

                            if (j > 25)
                                celda.CellReference = "A" + letras[j - 26] + numColumna;
                            else
                                celda.CellReference = letras[j] + numColumna;

                            if (LValores[k][i, j] != null)
                                celda.CellValue = new CellValue(LValores[k][i, j].ToString());

                            celda.DataType = LCellValues[mx - 1][j];
                            celda.StyleIndex = LEstilo[mx - 1][j];
                            oxw.WriteElement(celda);
                        }
                        oxw.WriteEndElement();
                    }
                }
            }

            // this is for SheetData
            oxw.WriteEndElement();
            //this is for table
            oxw.WriteStartElement(tp);
            oxw.WriteElement(tablePart);

            oxw.WriteEndElement();
            // this is for Worksheet
            oxw.WriteEndElement();
            oxw.Close();
            /////////////////////////////////definicion de la tabla 1

            TableDefinitionPart tableDefinitionPart1 = wsp.AddNewPart<TableDefinitionPart>("rId" + mx);
            OpenXmlWriter oxw2 = OpenXmlWriter.Create(tableDefinitionPart1);


            Table table1 = new Table() { Id = (UInt32)mx, Name = "Tabla1", DisplayName = "Tabla1" + mx, Reference = StringValue.FromString(Ldimension[0]), TotalsRowShown = BooleanValue.FromBoolean(false) };
            DocumentFormat.OpenXml.Spreadsheet.AutoFilter autoFilter1 = new DocumentFormat.OpenXml.Spreadsheet.AutoFilter() { Reference = StringValue.FromString(Ldimension[0]) };
            oxw2.WriteStartElement(table1);
            oxw2.WriteElement(autoFilter1);

            uint cellRangeCount = (uint)NombreHeaders[mx - 1].Length;


            TableColumns tableColumns1 = new TableColumns() { Count = (UInt32Value)(cellRangeCount) };

            oxw2.WriteStartElement(tableColumns1);
            uint tablaId = 0;
            for (uint i = 0; i < NombreHeaders[mx - 1].Length; i++)

            {
                oxw2.WriteElement(new TableColumn() { Id = i + 1, Name = NombreHeaders[mx - 1][i], DataFormatId = tablaId });
                tablaId++;

            }
            oxw2.WriteEndElement();//TableColumns

            TableStyleInfo tsi = new TableStyleInfo() { Name = "TableStyleLight9", ShowColumnStripes = BooleanValue.FromBoolean(false), ShowRowStripes = BooleanValue.FromBoolean(true), ShowFirstColumn = BooleanValue.FromBoolean(false), ShowLastColumn = BooleanValue.FromBoolean(false) };
            oxw2.WriteElement(tsi);
            oxw2.WriteEndElement();//Table
            oxw2.Close();// OpenXMLWriter tableDefinitionPart1
            
            return xl.WorkbookPart.GetIdOfPart(wsp);

        }

    }
}
