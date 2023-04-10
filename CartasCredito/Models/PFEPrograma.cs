using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class PFEPrograma {
		public int Id { get; set; }
		public int Anio { get; set; }
		public int Periodo { get; set; }
		public int EmpresaId { get; set; }
		public string CreadoPor { get; set; }
		public DateTime Creado { get; set; }
		public DateTime? Actualizado { get; set; }
		public DateTime? Eliminado { get; set; }
		public List<Pago> Pagos { get; set; }
		public bool Activo { get; set; }

		public PFEPrograma() {
			Anio = 0;
			Periodo = 0;
			EmpresaId = 0;
			CreadoPor = string.Empty;
			Creado = DateTime.MinValue;
			Pagos = new List<Pago>();
			Activo = false;
		}

		public static PFEPrograma Buscar(PFEPrograma model)
		{
			PFEPrograma rsp = new PFEPrograma();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_PFEProgramaBuscar(model, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.EmpresaId = int.Parse(row[idx].ToString()); idx++;
						rsp.Anio = int.Parse(row[idx].ToString()); idx++;
						rsp.Periodo = int.Parse(row[idx].ToString()); idx++;
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

						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;

						
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new PFEPrograma();
			}

			return rsp;
		}

		public static RespuestaFormato Insert(PFEPrograma modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_PFEPrograma(modelo, out dt, out errores))
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