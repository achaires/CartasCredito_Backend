using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Enmienda
	{
		public int RowNum { get; set; }
		public int Id { get; set; }
		public string CartaCreditoId { get; set; }
		public decimal ImporteLC { get; set; }
		public DateTime FechaLimiteEmbarque { get; set; }
		public DateTime FechaVencimiento { get; set; }
		public string DescripcionMercancia { get; set; }
		public string ConsideracionesAdicionales { get; set; }
		public string InstruccionesEspeciales { get; set; }
		public string DocumentoSwift { get; set; }
		public int Estatus { get; set; }
		public DateTime Creado { get; set; }
		public string CreadoPor { get; set; }
		public bool Activo { get; set; }
		public decimal Prev_ImporteLC { get; set; }
		public DateTime Prev_FechaLimiteEmbarque { get; set; }
		public DateTime Prev_FechaVencimiento { get; set; }
		public string Prev_DescripcionMercancia { get; set; }
		public string Prev_ConsideracionesAdicionales { get; set; }
		public string Prev_InstruccionesEspeciales { get; set; }
		public string Prev_DocumentoSwift { get; set; }

		public static RespuestaFormato Insert(Enmienda modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_Enmienda(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						int id = 0;
						Int32.TryParse(row[3].ToString(), out id);

						if (id > 0)
						{
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

		public static RespuestaFormato Update(Enmienda modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_Enmienda(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						int id = 0;
						Int32.TryParse(row[3].ToString(), out id);

						if (id > 0)
						{
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

		public static List<Enmienda> GetByCartaCreditoId(string cartaCreditoId)
		{
			List<Enmienda> res = new List<Enmienda>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_EnmiendasByCartaCreditoId(out dt, out errores, cartaCreditoId))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Enmienda();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							item.DocumentoSwift = row[idx].ToString(); idx++;
							item.ImporteLC = decimal.TryParse(row[idx].ToString(), out decimal importeLcVal) ? importeLcVal : 0; idx++;
							item.FechaLimiteEmbarque = DateTime.TryParse(row[idx].ToString(), out DateTime fleVal) ? fleVal : DateTime.Parse("2000-01-01"); idx++;
							item.FechaVencimiento = DateTime.TryParse(row[idx].ToString(), out DateTime fvVal) ? fvVal : DateTime.Parse("2000-01-01"); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.ConsideracionesAdicionales = row[idx].ToString(); idx++;
							item.InstruccionesEspeciales = row[idx].ToString(); idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int estatusVal) ? estatusVal : 0; idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime creadoVal) ? creadoVal : DateTime.Parse("2000-01-01"); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool activoVal) ? activoVal : false; idx++;
							item.Prev_DocumentoSwift = row[idx].ToString(); idx++;
							item.Prev_ImporteLC = decimal.TryParse(row[idx].ToString(), out decimal prevImporteLcVal) ? prevImporteLcVal : 0; idx++;
							item.Prev_FechaLimiteEmbarque = DateTime.TryParse(row[idx].ToString(), out DateTime prevfleVal) ? prevfleVal : DateTime.Parse("2000-01-01"); idx++;
							item.Prev_FechaVencimiento = DateTime.TryParse(row[idx].ToString(), out DateTime prevfvVal) ? prevfvVal : DateTime.Parse("2000-01-01"); idx++;
							item.Prev_DescripcionMercancia = row[idx].ToString(); idx++;
							item.Prev_ConsideracionesAdicionales = row[idx].ToString(); idx++;
							item.Prev_InstruccionesEspeciales = row[idx].ToString(); idx++;
							item.RowNum = int.TryParse(row[idx].ToString(), out int rownumval) ? rownumval : 0;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Enmienda>();

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
	}
}