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

            //var ob = new ExtraerInformacionExcel<DtoImpuesto, Entidad>();

            //var obre = ob.ObtenerInformacionSAX(mems, Guid.Empty, Guid.Empty, columnasEsperadas);

            var obj = new ServicioExcel();
            var obre = obj.ObtenerInformacionExcel(mems, columnasEsperadas);
            int i;
            string[] listaHeaders = { "NumeroHabitacion", "TipoHabitacion", "EstatusHabitacion", "NumeroPersonas", "Incidencias" };

            Object[,] nuemvo = new object[1, 6];

            nuemvo[0, 0] = "GSHGDHD";
            nuemvo[0, 1] = "GSHGDHD";
            nuemvo[0, 2] = "GSHGDHD";
            nuemvo[0, 3] = "GSHGDHD";
            nuemvo[0, 4] = "GSHGDHD";
            nuemvo[0, 5] = "GSHGDHD";


            obj.GenerarExcelUnaHoja(stream, listaHeaders, nuemvo, "Incidencias encontradas");
           
            var inputAsString = Convert.ToBase64String(stream.ToArray());


            //return Json(inputAsString);
            return Json(obre);
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
