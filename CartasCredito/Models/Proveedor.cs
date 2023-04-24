using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Proveedor
	{
		public int Id { get; set; }
		public int EmpresaId { get; set; }
		public string Empresa { get; set; }
		public int PaisId { get; set; }
		public string Pais { get; set; }
		public string Nombre { get; set; }
		public string Descripcion { get; set; }
		public bool Activo { get; set; }
		public string CreadoPor { get; set; }
		public DateTime Creado { get; set; }
		public DateTime? Actualizado { get; set; }
		public DateTime? Eliminado { get; set; }
		public Contacto Contacto { get; set; }

		public Proveedor ()
		{
			Id = 0;
			EmpresaId = 0;
			PaisId = 0;
			Nombre = "";
			Descripcion = "";
			Activo = false;
			CreadoPor = "";
			Creado = DateTime.MinValue;
			Actualizado = null;
			Eliminado = null;
		}

		public static List<Proveedor> Get(int activo = 1)
		{
			List<Proveedor> res = new List<Proveedor>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_Proveedores(out dt, out errores, activo))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Proveedor();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.EmpresaId = int.Parse(row[idx].ToString()); idx++;
							item.PaisId = int.Parse(row[idx].ToString()); idx++;
							item.Nombre = row[idx].ToString(); idx++;
							item.Descripcion = row[idx].ToString(); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crdVal) ? crdVal : DateTime.Now; idx++;

							if (row[idx].ToString().Length > 0)
							{
								item.Actualizado = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.Actualizado = null;
							}

							idx++;

							if (row[idx].ToString().Length > 0)
							{
								item.Eliminado = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.Eliminado = null;
							}

							idx++;

							item.Empresa = row[idx].ToString(); idx++;
							item.Pais = row[idx].ToString(); idx++;

							item.Contacto = Contacto.GetByModelNombreAndId(item.Id, "Proveedor");

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Proveedor>();

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

		public static Proveedor GetById(int id)
		{
			Proveedor rsp = new Proveedor();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_ProveedorById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.EmpresaId = int.Parse(row[idx].ToString()); idx++;
						rsp.PaisId = int.Parse(row[idx].ToString()); idx++;
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

						idx++;

						rsp.Empresa = row[idx].ToString(); idx++;
						rsp.Pais = row[idx].ToString(); idx++;
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Proveedor();
			}

			return rsp;
		}

		public static RespuestaFormato Insert(Proveedor modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_Proveedor(modelo, out dt, out errores))
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

		public static RespuestaFormato Update(Proveedor modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_Proveedor(modelo, out dt, out errores))
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