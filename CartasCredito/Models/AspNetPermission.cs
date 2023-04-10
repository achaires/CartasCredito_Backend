using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class AspNetPermission
	{
		public int Id { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public bool Activo { get; set; }

		public AspNetPermission() {
			Id = 0;
			Name = string.Empty;
			Description = string.Empty;
			Activo = false;
		}

		public static List<AspNetPermission> Get(int activo = 1)
		{
			List<AspNetPermission> res = new List<AspNetPermission>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_AspNetPermissions(out dt, out errores, activo))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new AspNetPermission();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.Name = row[idx].ToString(); idx++;
							item.Description = row[idx].ToString(); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<AspNetPermission>();

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

		public static List<AspNetPermission> GetByRoleId(string id)
		{
			List<AspNetPermission> rsp = new List<AspNetPermission>();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_AspNetPermissionsByRoleId(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new AspNetPermission();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.Name = row[idx].ToString(); idx++;
							item.Description = row[idx].ToString(); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;

							rsp.Add(item);
						}
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new List<AspNetPermission>();
			}

			return rsp;
		}

		public static AspNetPermission GetById(int id)
		{
			AspNetPermission rsp = new AspNetPermission();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_AspNetPermissionById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.Name = row[idx].ToString(); idx++;
						rsp.Description = row[idx].ToString(); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new AspNetPermission();
			}

			return rsp;
		}

	}
}