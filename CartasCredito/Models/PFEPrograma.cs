using Antlr.Runtime.Misc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class PFEPrograma
	{
		public int Id { get; set; }
		public int Anio { get; set; }
		public int Periodo { get; set; }
		public int EmpresaId { get; set; }

		public List<Pago> Pagos { get; set; }
		public List<PFETipoCambio> TiposCambio { get; set; }

		public PFEPrograma() 
		{
			Id = 0;
			Anio = 0;
			Periodo = 0;
			EmpresaId = 0;
			Pagos = new List<Pago>();
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

		public static RespuestaFormato InsertPFEProgramaPago (int programaId, int pagoId)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_PFEPago(programaId, pagoId, out dt, out errores))
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

		public static RespuestaFormato InsertTipoCambio(PFETipoCambio model)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_PFETipoCambio(model, out dt, out errores))
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


		public static PFEPrograma Buscar(PFEPrograma model)
		{
			var rsp = new PFEPrograma();

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
						rsp.Anio = int.Parse(row[idx].ToString()); idx++;
						rsp.Periodo = int.Parse(row[idx].ToString()); idx++;
						rsp.EmpresaId = int.Parse(row[idx].ToString()); idx++;
						rsp.Pagos = PagosByPFEProgramaId(rsp.Id);
					}
				}
			} catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

		public static PFEPrograma GetById(int id)
		{
			var rsp = new PFEPrograma();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_PFEProgramaById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.Parse(row[idx].ToString()); idx++;
						rsp.Anio = int.Parse(row[idx].ToString()); idx++;
						rsp.Periodo = int.Parse(row[idx].ToString()); idx++;
						rsp.EmpresaId = int.Parse(row[idx].ToString()); idx++;
						rsp.Pagos = PagosByPFEProgramaId(rsp.Id);
						rsp.TiposCambio = PFETipoCambio.GetByProgramaId(rsp.Id);
					}
				}
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

		public static RespuestaFormato DelPagosByProgramaId(int id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Del_PFEPagosByProgramaId(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						rsp.Flag = true;
						rsp.DataString = "Registro eliminado";
					}
				}
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
				rsp.Flag = false;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}

		public static List<Pago> PagosByPFEProgramaId (int pfeProgramaId)
		{
			var rsp = new List<Pago>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_PFEPagos(pfeProgramaId, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Pago();

							item.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;
							item.MontoPago = decimal.TryParse(row[idx].ToString(), out decimal montoVal) ? montoVal : 0; idx++;
							item.MontoPagado = decimal.TryParse(row[idx].ToString(), out decimal montoPgVal) ? montoPgVal : 0; idx++;
							item.FechaVencimiento = DateTime.TryParse(row[idx].ToString(), out DateTime fvVal) ? fvVal : DateTime.Now.AddDays(1); idx++;
							item.FechaPago = DateTime.TryParse(row[idx].ToString(), out DateTime fpVal) ? fpVal : DateTime.Now.AddDays(1); idx++;
							item.RegistroPagoPor = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							item.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0; idx++;
							item.NumeroFactura = row[idx].ToString(); idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int statusVal) ? statusVal : 0; idx++;
							item.EstatusText = Pago.GetStatusText(item.Estatus);
							item.Activo = Boolean.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crVal) ? crVal : DateTime.Now.AddDays(1); idx++;
							item.CartaCredito = CartaCredito.GetById(item.CartaCreditoId);
							rsp.Add(item);
						}
					}
				}
			} catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}
	}
}