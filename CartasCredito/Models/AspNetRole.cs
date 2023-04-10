using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class AspNetRole
	{
		public string Id { get; set; }
		public string Name { get; set; }
		public bool Activo { get; set; }
		public IEnumerable<AspNetPermission> Permissions { get; set; }

		public AspNetRole()
		{
			Id = String.Empty;
			Name = String.Empty;
			Activo = false;
			Permissions = new List<AspNetPermission>();
		}

		public static List<AspNetRole> Get(int activo = 1)
		{
			List<AspNetRole> res = new List<AspNetRole>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_AspNetRoles(out dt, out errores, activo))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new AspNetRole();

							item.Id = row[idx].ToString(); idx++;
							item.Name = row[idx].ToString(); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool activoVal) ? activoVal : false; idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<AspNetRole>();

				// Get stack trace for the exception with source file information
				var st = new StackTrace(ex, true);
				// Get the top stack frame
				var frame = st.GetFrame(0);
				// Get the line number from the stack frame
				var line = frame.GetFileLineNumber();

				var errorMsg = ex.ToString();
				Utility.Logger.Error("AspNetRole - " + line + " - " + errorMsg);

			}

			return res;
		}

		public static AspNetRole GetById(string id)
		{
			AspNetRole rsp = new AspNetRole();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_AspNetRoleById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = row[idx].ToString(); idx++;
						rsp.Name = row[idx].ToString(); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool activoval) ? activoval : false; idx++;
						rsp.Permissions = AspNetPermission.GetByRoleId(id);
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new AspNetRole();
				Utility.Logger.Error(ex);
			}

			return rsp;
		}

		public static RespuestaFormato Insert(AspNetRole modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_AspNetRole(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						string id = row[3].ToString();

						if (id.Length > 0)
						{
							rsp.Description = "Nuevo usuario creado correctamente";
							rsp.Flag = true;
							rsp.DataInt = 1;
							rsp.DataString = id;
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
				Utility.Logger.Error(ex);
			}

			return rsp;
		}

		public static RespuestaFormato Update(AspNetRole modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_AspNetRole(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						string id = row[3].ToString();

						if (id.Length > 0)
						{
							rsp.Description = "Nuevo usuario creado correctamente";
							rsp.Flag = true;
							rsp.DataInt = 1;
							rsp.DataString = id;
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
				Utility.Logger.Error(ex);
			}

			return rsp;
		}

		public static RespuestaFormato InsertPermission(string roleId, int permissionId)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_AspNetRolePermission(roleId, permissionId, out dt, out errores))
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

		public static RespuestaFormato DeletePermission(string roleId, int permissionId)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Del_AspNetRolePermission(roleId, permissionId, out dt, out errores))
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