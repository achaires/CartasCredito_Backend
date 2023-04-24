using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Reporte
	{
		public int Id { get; set; }
		public string TipoReporte { get; set; }
		public DateTime Creado { get; set; }
		public string CreadoPorId { get; set; }
		public string CreadoPor { get; set; }
		public string Filename { get; set; }

		public static List<Reporte> Get()
		{
			List<Reporte> res = new List<Reporte>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_Reportes(out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Reporte();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.TipoReporte = row[idx].ToString(); idx++;	
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crdVal) ? crdVal : DateTime.Now; idx++;
							item.CreadoPorId = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.Filename = row[idx].ToString(); idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Reporte>();

				// Get stack trace for the exception with source file information
				var st = new StackTrace(ex, true);
				// Get the top stack frame
				var frame = st.GetFrame(0);
				// Get the line number from the stack frame
				var line = frame.GetFileLineNumber();

				var errorMsg = ex.ToString();
			}

			return res;
		}

		public static RespuestaFormato Insert(Reporte modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_Reporte(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						int id = 0;
						Int32.TryParse(row[3].ToString(), out id);

						if (id > 0)
						{
							rsp.Description = "Modelo insertado";
							rsp.Flag = true;
							rsp.DataInt = id;
						}
					}
				}
				else
				{
					rsp.Description = "Ocurrió un error";
					rsp.Errors.Add(errores);
				}
			}
			catch (Exception ex)
			{
				rsp.Errors.Add(ex.Message);
				rsp.Description = "Ocurrió un error";
			}

			return rsp;
		}
	}
}