using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Invitacion
	{
		public int Id { get; set; }
		public string Email { get; set; }
		public string Token { get; set; }
		public DateTime Creado { get; set; }
		public string CreadoPorId { get; set; }
		public string UserName { get; set; }

		public static Invitacion GetById(int id)
		{
			Invitacion rsp = new Invitacion();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_InvitacionById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.Email = row[idx].ToString(); idx++;
						rsp.Token = row[idx].ToString(); idx++;
						rsp.Creado = DateTime.Parse(row[idx].ToString());
						rsp.CreadoPorId = row[idx].ToString(); idx++;

					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Invitacion();
			}

			return rsp;
		}

		public static Invitacion GetByToken(string token)
		{
			Invitacion rsp = new Invitacion();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_InvitacionByToken(token, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.Email = row[idx].ToString(); idx++;
						rsp.Token = row[idx].ToString(); idx++;
						rsp.Creado = DateTime.Parse(row[idx].ToString());
						rsp.CreadoPorId = row[idx].ToString(); idx++;
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Invitacion();
			}

			return rsp;
		}

		public static RespuestaFormato Insert(Invitacion modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_Invitacion(modelo, out dt, out errores))
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