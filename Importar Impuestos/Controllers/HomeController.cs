using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Importar_Impuestos.App;
using Importar_Impuestos.Servicios;
using System.Collections;
using System.Text.RegularExpressions;

namespace Importar_Impuestos.Controllers
{
    [Route("api/[controller]")]
    public class HomeController : Controller
    {
        [HttpPost("[action]")]
        public IActionResult Index([FromForm]IFormFile archivo)
        {
            var columnasEsperadas = new List<string>()
            {
                "RFC", "Fecha","Mes","Año","IVA","ISR"
            };
            var mems = new System.IO.MemoryStream(LeerStream(archivo.OpenReadStream()));
            var stream = new System.IO.MemoryStream();

           
            List<DtoRespuestaImpuestos> impuestosRespuesta = new List<DtoRespuestaImpuestos>();
            
            //var ob = new ExtraerInformacionExcel<DtoImpuesto, Entidad>();

            //var obre = ob.ObtenerInformacionSAX(mems, Guid.Empty, Guid.Empty, columnasEsperadas);

            var obj = new ServicioExcel();
            var obre = obj.ObtenerInformacionExcel(mems, columnasEsperadas);
           
          

            //Validar información de excel
            Regex exprerfc = new Regex(@"^([A-ZÑ\x26]{3,4}([0-9]{2})(0[1-9]|1[0-2])(0[1-9]|1[0-9]|2[0-9]|3[0-1]))((-)?([A-Z\d]{3}))?$");
            var erroresEstatus = String.Empty;
            var existeRfc = true;
            DateTime fecha;
            var existenRegErrores = false;
        
            

            foreach (var impuestoRegistro in obre)
            {
                erroresEstatus = string.Empty;
                if (!exprerfc.IsMatch(impuestoRegistro.RFC)){
                    impuestosRespuesta.Add(DtoRespuestaImpuestos.Create(impuestoRegistro.RFC,impuestoRegistro.Fecha,impuestoRegistro.Mes,impuestoRegistro.Año, impuestoRegistro.IVA,impuestoRegistro.ISR,false,"Formato no valido (RFC)"));
                }
                else
                {
                    if (!existeRfc)
                    {
                        impuestosRespuesta.Add(DtoRespuestaImpuestos.Create(impuestoRegistro.RFC, impuestoRegistro.Fecha, impuestoRegistro.Mes, impuestoRegistro.Año, impuestoRegistro.IVA, impuestoRegistro.ISR, false, "RFC no encontrado"));
                        
                    }
                    else
                    {
                        if (impuestoRegistro.Mes.Equals("") || impuestoRegistro.Año.Equals("") || impuestoRegistro.IVA.Equals("") || impuestoRegistro.ISR.Equals(""))
                            erroresEstatus = "Formato no valido (Hay campos vacios en este registro) - ";
                           
                        if (!impuestoRegistro.Fecha.Equals(""))
                        {
                            if (!DateTime.TryParseExact(impuestoRegistro.Fecha, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out fecha))                            
                                erroresEstatus += "Formato no valido (La fecha no viene con el formato correcto) - ";
                             
                            else
                                erroresEstatus += fecha < DateTime.Now.AddMonths(-6) ? "Fecha no valida - " : "";
                        }

                        if (erroresEstatus.Equals(""))
                        {
                            impuestosRespuesta.Add(DtoRespuestaImpuestos.Create(impuestoRegistro.RFC, impuestoRegistro.Fecha, impuestoRegistro.Mes, impuestoRegistro.Año, impuestoRegistro.IVA, impuestoRegistro.ISR, true, "Importado correctamente"));
                        }
                        else
                        {
                            impuestosRespuesta.Add(DtoRespuestaImpuestos.Create(impuestoRegistro.RFC, impuestoRegistro.Fecha, impuestoRegistro.Mes, impuestoRegistro.Año, impuestoRegistro.IVA, impuestoRegistro.ISR, false, erroresEstatus.Trim().Trim('-').Trim()));

                        }

                    }
                    
                                      
                }
            }


            existenRegErrores = impuestosRespuesta.Count(x => x.seImporto == false) > 0;

            //Gnerar Archivo con errores
            var inputAsString = string.Empty;
            if (existenRegErrores){
                string[] listaHeaders = {
                "RFC", "Fecha","Mes","Año","IVA","ISR","Estatus"
            };

                Object[,] nuemvo = new object[impuestosRespuesta.Count(x => x.seImporto == false), 7];
                int i = 0;
                foreach (var regError in impuestosRespuesta.Where(x => x.seImporto == false).ToList())
                {
                    nuemvo[i, 0] = regError.RFC;
                    nuemvo[i, 1] = regError.Fecha;
                    nuemvo[i, 2] = regError.Mes;
                    nuemvo[i, 3] = regError.Año;
                    nuemvo[i, 4] = regError.Iva;
                    nuemvo[i, 5] = regError.Isr;
                    nuemvo[i, 6] = regError.Estatus;
                    i += 1;
                }




                obj.GenerarExcelUnaHoja(stream, listaHeaders, nuemvo, "Incidencias encontradas");

                inputAsString = Convert.ToBase64String(stream.ToArray());
            }
            
            //Generar Archivo Con errores

            DtoRespuestaImportarImpuestosMes resp = DtoRespuestaImportarImpuestosMes.Create(impuestosRespuesta,existenRegErrores,inputAsString);
            //return Json(inputAsString);
            return Json(resp);
        }

        public byte[] LeerStream(Stream stream)
        {
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}
