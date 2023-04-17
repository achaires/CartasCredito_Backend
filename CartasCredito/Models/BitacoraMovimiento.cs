using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class BitacoraMovimiento
	{
		public int Id { get; set; }
		public string Titulo { get; set; }
		public string Descripcion { get; set; }
		public string CartaCreditoId { get; set; }
		public string ModeloNombre { get; set; }
		public int ModeloId { get; set; }
		public string CreadoPor { get; set; }
		public string CreadoPorId { get; set; }
		public DateTime? Creado { get; set; }

		public BitacoraMovimiento()
		{
			Id = 0;
			Titulo= string.Empty;
			Descripcion = string.Empty;
			CartaCreditoId = string.Empty;
			ModeloNombre = string.Empty;
			ModeloId = 0;
			CreadoPor = string.Empty;
			CreadoPorId = string.Empty;
			Creado = null;
		}

		public static List<BitacoraMovimiento> Get(DateTime dateStart, DateTime dateEnd)
		{
			List<BitacoraMovimiento> res = new List<BitacoraMovimiento>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_BitacoraMovimientos(dateStart, dateEnd, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new BitacoraMovimiento();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.Titulo = row[idx].ToString(); idx++;
							item.Descripcion = row[idx].ToString(); idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							item.ModeloNombre = row[idx].ToString(); idx++;
							item.ModeloId = int.TryParse(row[idx].ToString(), out int modid) ? modid : 0; idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.CreadoPorId = row[idx].ToString(); idx++;
							
							if (row[idx].ToString().Length > 0)
							{
								item.Creado = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.Creado = null;
							}

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<BitacoraMovimiento>();

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

		public static RespuestaFormato Insert(BitacoraMovimiento modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_BitacoraMovimiento(modelo, out dt, out errores))
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