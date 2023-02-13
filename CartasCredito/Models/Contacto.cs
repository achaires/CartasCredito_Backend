using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Contacto
	{
		public int Id { get; set; }
		public int ModelId { get; set; }

		public string ModelNombre { get; set; }
		public string Nombre { get; set; }
		public string ApellidoPaterno { get; set; }
		public string ApellidoMaterno { get; set; }
		public string Telefono { get; set; }
		public string Celular { get; set; }
		public string Email { get; set; }
		public string Fax { get; set; }
		public string Descripcion { get; set; }
		public bool Activo { get; set; }

		public Contacto()
		{
			Id = 0;
			ModelId = 0;
			ModelNombre = "";
			Nombre = "";
			ApellidoPaterno= "";
			ApellidoMaterno = "";
			Telefono = "";
			Celular = "";
			Email = "";
			Fax = "";
			Descripcion = "";
			Activo = false;
		}

		public static List<Contacto> Get(int activo = 1)
		{
			List<Contacto> res = new List<Contacto>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_Contactos(out dt, out errores, activo))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Contacto();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.ModelId = int.Parse(row[idx].ToString()); idx++;
							item.ModelNombre = row[idx].ToString(); idx++;
							item.Nombre = row[idx].ToString(); idx++;
							item.ApellidoPaterno = row[idx].ToString(); idx++;
							item.ApellidoMaterno = row[idx].ToString(); idx++;
							item.Telefono = row[idx].ToString(); idx++;
							item.Celular = row[idx].ToString(); idx++;
							item.Email = row[idx].ToString(); idx++;
							item.Fax = row[idx].ToString(); idx++;
							item.Descripcion = row[idx].ToString(); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
							
							idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Contacto>();

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

		public static Contacto GetById(int id)
		{
			Contacto rsp = new Contacto();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_ContactoById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.ModelId = int.Parse(row[idx].ToString()); idx++;
						rsp.ModelNombre = row[idx].ToString(); idx++;
						rsp.Nombre = row[idx].ToString(); idx++;
						rsp.ApellidoPaterno = row[idx].ToString(); idx++;
						rsp.ApellidoMaterno = row[idx].ToString(); idx++;
						rsp.Telefono = row[idx].ToString(); idx++;
						rsp.Celular = row[idx].ToString(); idx++;
						rsp.Email = row[idx].ToString(); idx++;
						rsp.Fax = row[idx].ToString(); idx++;
						rsp.Descripcion = row[idx].ToString(); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
						
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Contacto();
			}

			return rsp;
		}

		public static Contacto GetByModelNombreAndId(int modelId, string modelNombre)
		{
			Contacto rsp = new Contacto();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_ContactoByModelNombreAndId(modelId, modelNombre, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.ModelId = int.Parse(row[idx].ToString()); idx++;
						rsp.ModelNombre = row[idx].ToString(); idx++;
						rsp.Nombre = row[idx].ToString(); idx++;
						rsp.ApellidoPaterno = row[idx].ToString(); idx++;
						rsp.ApellidoMaterno = row[idx].ToString(); idx++;
						rsp.Telefono = row[idx].ToString(); idx++;
						rsp.Celular = row[idx].ToString(); idx++;
						rsp.Email = row[idx].ToString(); idx++;
						rsp.Fax = row[idx].ToString(); idx++;
						rsp.Descripcion = row[idx].ToString(); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;

					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Contacto();
			}

			return rsp;
		}

		public static List<Contacto> GetManyByModelNombreAndId(int modelId, string modelNombre)
		{
			List<Contacto> res = new List<Contacto>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ContactoByModelNombreAndId(modelId, modelNombre, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Contacto();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.ModelId = int.Parse(row[idx].ToString()); idx++;
							item.ModelNombre = row[idx].ToString(); idx++;
							item.Nombre = row[idx].ToString(); idx++;
							item.ApellidoPaterno = row[idx].ToString(); idx++;
							item.ApellidoMaterno = row[idx].ToString(); idx++;
							item.Telefono = row[idx].ToString(); idx++;
							item.Celular = row[idx].ToString(); idx++;
							item.Email = row[idx].ToString(); idx++;
							item.Fax = row[idx].ToString(); idx++;
							item.Descripcion = row[idx].ToString(); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;

							idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Contacto>();

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

		public static RespuestaFormato Insert(Contacto modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_Contacto(modelo, out dt, out errores))
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

		public static RespuestaFormato Update(Contacto modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_Contacto(modelo, out dt, out errores))
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
	}
}