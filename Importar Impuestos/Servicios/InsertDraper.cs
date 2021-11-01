using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Dapper;
using Importar_Impuestos.App;

namespace Importar_Impuestos.Servicios
{
    public class InsertDraper
    {
        private string sqlCon = @"Data Source=DESKTOP-GUSKUDA;Initial Catalog=GESTIONHOSP;  integrated security = true";

        public void Insert(IEnumerable<DtoRespuestaImpuestos> lista)
        {
            string sqlQuery = @"INSERT INTO dbo.ImpuestosMes(rfc,fecha,anio,mes,iva,isr) VALUES (@rfc,CONVERT(datetime, @fecha, 103),@anio,@mes,@iva,@isr)";

            try
            {
                
                using (var db = new SqlConnection(sqlCon))
                {
                    db.Open();
                    using(var tran = db.BeginTransaction())
                    {
                        foreach (var item in lista)
                        {
                            db.Execute(sqlQuery, new { item.RFC, item.Fecha,item.Anio, item.Mes, item.Iva, item.Isr}, tran);
                        }
                        tran.Commit();
                    }
                 

                }


            }
            catch (Exception e)
            {
                
            }
        }
    }
}



