using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class AspNetUser
	{
		public string Id { get; set; }
		public string Email { get; set; }
		public bool EmailConfirmed { get; set; }
		public string PasswordHash { get; set; }
		public string NewPasswordPlain { get; set; }
		public string PhoneNumber { get; set; }
		public bool PhoneNumberConfirmed { get; set; }
		public bool TwoFactorEnabled { get; set; }
		public DateTime LockoutEndDateUtc { get; set; }
		public bool LockoutEnabled { get; set; }
		public int AccessFailedCount { get; set; }
		public string UserName { get; set; }
		public bool Activo { get; set; }
		//public IEnumerable<AspNetRole> Roles { get; set; }
		public AspNetRole Role { get; set; }
		public string RoleId { get; set; }
		public AspNetUserProfile Profile { get; set; }
		public IEnumerable<Empresa> Empresas { get; set; }

		public AspNetUser () {
			Id = String.Empty;
			Email = String.Empty;
			PasswordHash = String.Empty;
			PhoneNumber = String.Empty;
			PhoneNumberConfirmed = false;
			TwoFactorEnabled = false;
			LockoutEndDateUtc = DateTime.MinValue;
			LockoutEnabled = false;
			NewPasswordPlain = String.Empty;
			Activo = false;
			//Roles = new List<AspNetRole> ();
			Empresas = new List<Empresa> ();
		}

		public static List<AspNetUser> Get(int activo = 1)
		{
			List<AspNetUser> res = new List<AspNetUser>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_AspNetUsers(out dt, out errores, activo))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new AspNetUser();

							item.Id = row[idx].ToString(); idx++;
							item.Email = row[idx].ToString(); idx++;
							item.EmailConfirmed = bool.TryParse(row[idx].ToString(), out bool emailconfirmedval) ? emailconfirmedval : false; idx++;
							item.PasswordHash = row[idx].ToString(); idx++;
							item.PhoneNumber = row[idx].ToString(); idx++;
							item.PhoneNumberConfirmed = bool.TryParse(row[idx].ToString(), out bool phoneconfirmedval) ? phoneconfirmedval : false; idx++;
							item.TwoFactorEnabled = bool.TryParse(row[idx].ToString(), out bool twofactorendabledval) ? twofactorendabledval : false; idx++;
							item.LockoutEndDateUtc = DateTime.TryParse(row[idx].ToString(), out DateTime lockoutdateendval) ? lockoutdateendval : DateTime.MaxValue; idx++;
							item.LockoutEnabled = bool.TryParse(row[idx].ToString(), out bool lockoutenabledval) ? lockoutenabledval : false; idx++;
							item.AccessFailedCount = int.TryParse(row[idx].ToString(), out int accessfailcountval) ? accessfailcountval : 0; idx++;
							item.UserName = row[idx].ToString(); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool activoval) ? activoval : false; idx++;
							item.RoleId = row[idx].ToString(); idx++;
							item.Profile = AspNetUserProfile.GetByUserId(item.Id);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<AspNetUser>();

				// Get stack trace for the exception with source file information
				var st = new StackTrace(ex, true);
				// Get the top stack frame
				var frame = st.GetFrame(0);
				// Get the line number from the stack frame
				var line = frame.GetFileLineNumber();

				var errorMsg = ex.ToString();
				Utility.Logger.Error("AspNetUser - " + line + " - " + errorMsg);

			}

			return res;
		}

		public static AspNetUser GetById(string id)
		{
			AspNetUser rsp = new AspNetUser();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_AspNetUserById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = row[idx].ToString(); idx++;
						rsp.Email = row[idx].ToString(); idx++;
						rsp.EmailConfirmed = bool.TryParse(row[idx].ToString(), out bool emailconfirmedval) ? emailconfirmedval : false; idx++;
						rsp.PasswordHash = row[idx].ToString(); idx++;
						rsp.PhoneNumber = row[idx].ToString(); idx++;
						rsp.PhoneNumberConfirmed = bool.TryParse(row[idx].ToString(), out bool phoneconfirmedval) ? phoneconfirmedval : false; idx++;
						rsp.TwoFactorEnabled = bool.TryParse(row[idx].ToString(), out bool twofactorendabledval) ? twofactorendabledval : false; idx++;
						rsp.LockoutEndDateUtc = DateTime.TryParse(row[idx].ToString(), out DateTime lockoutdateendval) ? lockoutdateendval : DateTime.MaxValue; idx++;
						rsp.LockoutEnabled = bool.TryParse(row[idx].ToString(), out bool lockoutenabledval) ? lockoutenabledval : false; idx++;
						rsp.AccessFailedCount = int.TryParse(row[idx].ToString(), out int accessfailcountval) ? accessfailcountval : 0; idx++;
						rsp.UserName = row[idx].ToString(); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool activoval) ? activoval : false; idx++;
						rsp.RoleId = row[idx].ToString(); idx++;
						rsp.Profile = AspNetUserProfile.GetByUserId(rsp.Id);
						rsp.Empresas = Empresa.GetByUserId(rsp.Id);
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new AspNetUser();
			}

			return rsp;
		}

		public static RespuestaFormato Insert(AspNetUser modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_AspNetUser(modelo, out dt, out errores))
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

		public static RespuestaFormato Update(AspNetUser modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_AspNetUser(modelo, out dt, out errores))
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

		public static RespuestaFormato InsertRole(string userId, string roleId)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_AspNetUserRole(userId, roleId, out dt, out errores))
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

		public static RespuestaFormato UpdateRole(string userId, string roleId)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_AspNetUserRole(userId, roleId, out dt, out errores))
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

		public static RespuestaFormato InsertEmpresa(string userId, int empresaId)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_AspNetUserEmpresa(userId, empresaId, out dt, out errores))
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

		public static RespuestaFormato DeleteEmpresa(string userId, int empresaId)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Del_AspNetUserEmpresa(userId, empresaId, out dt, out errores))
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

		public static AspNetUser GetByUserName(string userName)
		{
			AspNetUser rsp = new AspNetUser();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_AspNetUserByUserName(out dt, out errores, userName))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = row[idx].ToString(); idx++;
						rsp.Email = row[idx].ToString(); idx++;
						rsp.EmailConfirmed = bool.TryParse(row[idx].ToString(), out bool emailconfirmedval) ? emailconfirmedval : false; idx++;
						rsp.PasswordHash = row[idx].ToString(); idx++;
						rsp.PhoneNumber = row[idx].ToString(); idx++;
						rsp.PhoneNumberConfirmed = bool.TryParse(row[idx].ToString(), out bool phoneconfirmedval) ? phoneconfirmedval : false; idx++;
						rsp.TwoFactorEnabled = bool.TryParse(row[idx].ToString(), out bool twofactorendabledval) ? twofactorendabledval : false; idx++;
						rsp.LockoutEndDateUtc = DateTime.TryParse(row[idx].ToString(), out DateTime lockoutdateendval) ? lockoutdateendval : DateTime.MaxValue; idx++;
						rsp.LockoutEnabled = bool.TryParse(row[idx].ToString(), out bool lockoutenabledval) ? lockoutenabledval : false; idx++;
						rsp.AccessFailedCount = int.TryParse(row[idx].ToString(), out int accessfailcountval) ? accessfailcountval : 0; idx++;
						rsp.UserName = row[idx].ToString(); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool activoval) ? activoval : false; idx++;
						rsp.Profile = AspNetUserProfile.GetByUserId(rsp.Id);
						rsp.Empresas = Empresa.GetByUserId(rsp.Id);

					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new AspNetUser();
			}

			return rsp;
		}
	}
}