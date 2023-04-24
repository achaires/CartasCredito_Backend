using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Pais
	{
		public int Id { get; set; }
		public string Nombre { get; set;}
		public string Alpha2Code { get; set;}
		public string Alpha3Code { get; set;}

		public Pais ()
		{
			Id = 0;
			Nombre = string.Empty;
			Alpha2Code = string.Empty;
			Alpha3Code = string.Empty;
		}

		public static List<Pais> Get()
		{
			List<Pais> res = new List<Pais>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_Paises(out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Pais();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.Nombre = row[idx].ToString(); idx++;
							item.Alpha2Code = row[idx].ToString(); idx++;
							item.Alpha3Code = row[idx].ToString(); idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Pais>();

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
		/*
		public static Pais GetById(int id)
		{
			Pais rsp = new Pais();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_PaisById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.Nombre = row[idx].ToString(); idx++;
						rsp.Descripcion = row[idx].ToString(); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
						rsp.CreadoPor = row[idx].ToString(); idx++;
						rsp.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crdVal) ? crdVal : DateTime.Now; idx++;

						if (row[idx].ToString().Length > 0)
						{
							rsp.Actualizado = DateTime.Parse(row[idx].ToString());
						}
						else
						{
							rsp.Actualizado = null;
						}

						idx++;

						if (row[idx].ToString().Length > 0)
						{
							rsp.Eliminado = DateTime.Parse(row[idx].ToString());
						}
						else
						{
							rsp.Eliminado = null;
						}

					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Pais();
			}

			return rsp;
		}
		*/

		public static RespuestaFormato Insert(Pais modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_Pais(modelo, out dt, out errores))
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
		/*
		public static RespuestaFormato Update(Pais modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_Pais(modelo, out dt, out errores))
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
		*/
	}
}