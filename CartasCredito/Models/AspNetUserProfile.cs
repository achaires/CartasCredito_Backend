using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class AspNetUserProfile
	{
		public int Id { get; set; }
		public string Name { get; set; }
		public string LastName { get; set; }
		public string DisplayName { get; set; }
		public string UserId { get; set; }
		public string Notes { get; set; }

		public AspNetUserProfile() {
			Id = 0;
			Name = string.Empty;
			LastName = string.Empty;
			DisplayName = string.Empty;
			UserId = string.Empty;
			Notes = string.Empty;
		}

		public static AspNetUserProfile GetByUserId(string userId)
		{
			AspNetUserProfile rsp = new AspNetUserProfile();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_AspNetUserProfileByUserId(userId, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.TryParse(row[idx].ToString(),out int idval) ? idval : 0; idx++;
						rsp.Name = row[idx].ToString(); idx++;
						rsp.LastName = row[idx].ToString(); idx++;
						rsp.DisplayName = row[idx].ToString(); idx++;
						rsp.UserId = row[idx].ToString(); idx++;
						rsp.Notes = row[idx].ToString();
					}
				}
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex);
				Console.WriteLine(ex);
				rsp = new AspNetUserProfile();
			}

			return rsp;
		}

		public static RespuestaFormato Insert(AspNetUserProfile modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_AspNetUserProfile(modelo, out dt, out errores))
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
			}

			return rsp;
		}

		public static RespuestaFormato Update(AspNetUserProfile modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_AspNetUserProfile(modelo, out dt, out errores))
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