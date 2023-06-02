using CartasCredito.Models.DTOs;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Web;

namespace CartasCredito.Models
{
	public class Pago
	{
		public int Id { get; set; }
		public int? ComisionId { get; set; }
		public int? MonedaId { get; set; }
		public decimal? TipoCambio { get; set; }

		//[JsonConverter(typeof(CustomDateTimeConverter))]
		public DateTime? FechaVencimiento { get; set; }
		//[JsonConverter(typeof(CustomDateTimeConverter))]
		public DateTime? FechaPago { get; set; }
		public decimal MontoPago { get; set; }
		public decimal MontoPagado { get; set; }
		public string RegistroPagoPor { get; set; }
		public string CreadoPor { get; set; }
		public string CartaCreditoId { get; set; }
		public int NumeroPago { get; set; }
		public string NumeroFactura { get; set; }
		public int Estatus { get; set; }
		public bool Activo { get; set; }
		public DateTime Creado { get; set; }
		public DateTime? Actualizado { get; set; }
		public DateTime? Eliminado { get; set; }
		public string EstatusText { get; set; }
		public string EstatusClass { get; set; }
		public string Empresa { get; set; }
		public string Proveedor { get; set; }
		public CartaCredito CartaCredito { get; set; }
		public string Comentarios { get; set; }
		public int ProveedorId { get; set; }
		public int EmpresaId { get; set; }

		public bool PFEActivo { get; set; } // Esta flag me ayuda a saber si está seleccionado dentro de un programa PFE

		public Pago()
		{
			Id = 0;
			FechaVencimiento = null;
			FechaPago = null;
			MontoPago = 0M;
			RegistroPagoPor = "";
			CreadoPor = "";
			CartaCreditoId = "";
			NumeroPago = 0;
			NumeroFactura = "";
			Estatus = 0;
			Activo = false;
			Creado = DateTime.MinValue;
			Actualizado = null;
			Eliminado = null;
			PFEActivo = false;
		}

		public static List<Pago> Get(DateTime dateStart, DateTime dateEnd, int activo = -1)
		{
			List<Pago> res = new List<Pago>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_Pagos(out dt, out errores, dateStart, dateEnd))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Pago();

							item.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0; idx++;
							item.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;
							item.FechaVencimiento = DateTime.TryParse(row[idx].ToString(), out DateTime fvVal) ? fvVal : DateTime.Now.AddDays(1); idx++;
							item.FechaPago = DateTime.TryParse(row[idx].ToString(), out DateTime fpVal) ? fpVal : DateTime.Now.AddDays(1); idx++;
							item.MontoPago = decimal.TryParse(row[idx].ToString(), out decimal montoVal) ? montoVal : 0; idx++;
							//item.MontoPagado = decimal.TryParse(row[idx].ToString(), out decimal montoPgVal) ? montoPgVal : 0;
							idx++;
							item.RegistroPagoPor = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							//item.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0;
							idx++;
							item.NumeroFactura = row[idx].ToString(); idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int statusVal) ? statusVal : 0; idx++;
							item.EstatusText = Pago.GetStatusText(item.Estatus);
							item.Activo = Boolean.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crVal) ? crVal : DateTime.Now.AddDays(1); idx++;


							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Pago>();

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

		public static List<Pago> GetProgramados()
		{
			List<Pago> res = new List<Pago>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_PagosProgramados(out dt, out errores))
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
							item.MontoPagado = decimal.TryParse(row[idx].ToString(), out decimal montoPagadoVal) ? montoPagadoVal : 0; idx++;

							if (row[idx].ToString().Length > 0)
							{
								item.FechaVencimiento = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.FechaVencimiento = null;
							}

							idx++;

							if (row[idx].ToString().Length > 0)
							{
								item.FechaPago = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.FechaPago = null;
							}

							idx++;

							item.RegistroPagoPor = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							item.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0; idx++;	
							item.NumeroFactura = row[idx].ToString(); idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int statusVal) ? statusVal : 0; idx++;
							item.Activo = Boolean.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crVal) ? crVal : DateTime.Now.AddDays(1); idx++;

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
							item.Proveedor = row[idx].ToString(); idx++;
							item.EmpresaId = int.TryParse(row[idx].ToString(), out int empresaidval) ? empresaidval : 0; idx++;
							item.ProveedorId = int.TryParse(row[idx].ToString(), out int proveedoridval) ? proveedoridval : 0; idx++;

							item.EstatusText = Pago.GetStatusText(item.Estatus);
							item.CartaCredito = CartaCredito.GetById(item.CartaCreditoId);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Pago>();

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

		public static List<Pago> GetVencidos()
		{
			List<Pago> res = new List<Pago>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_PagosVencidos(out dt, out errores))
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
							item.MontoPagado = decimal.TryParse(row[idx].ToString(), out decimal montoPagadoVal) ? montoPagadoVal : 0; idx++;

							if (row[idx].ToString().Length > 0)
							{
								item.FechaVencimiento = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.FechaVencimiento = null;
							}

							idx++;

							if (row[idx].ToString().Length > 0)
							{
								item.FechaPago = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.FechaPago = null;
							}

							idx++;

							
							item.RegistroPagoPor = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							item.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0; idx++;
							item.NumeroFactura = row[idx].ToString(); idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int statusVal) ? statusVal : 0; idx++;
							item.Activo = Boolean.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crVal) ? crVal : DateTime.Now.AddDays(1); idx++;

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
							item.Proveedor = row[idx].ToString(); idx++;
							item.EmpresaId = int.TryParse(row[idx].ToString(), out int empresaidval) ? empresaidval : 0; idx++;
							item.ProveedorId = int.TryParse(row[idx].ToString(), out int proveedoridval) ? proveedoridval : 0; idx++;

							item.EstatusText = Pago.GetStatusText(item.Estatus);
							item.CartaCredito = CartaCredito.GetById(item.CartaCreditoId);

							res.Add(item);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Pago>();

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

		public static List<Pago> GetByCartaCreditoId(string cartaCreditoId)
		{
			List<Pago> res = new List<Pago>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_PagosByCartaCreditoId(cartaCreditoId, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new Pago();

							item.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0; idx++;
							item.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;
							
							if (row[idx].ToString().Length > 0)
							{
								item.FechaVencimiento = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.FechaVencimiento = null;
							}

							idx++;

							if (row[idx].ToString().Length > 0)
							{
								item.FechaPago = DateTime.Parse(row[idx].ToString());
							}
							else
							{
								item.FechaPago = null;
							}

							idx++;

							item.MontoPago = decimal.TryParse(row[idx].ToString(), out decimal montoVal) ? montoVal : 0; idx++;
							item.RegistroPagoPor = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							item.NumeroFactura = row[idx].ToString(); idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int statusVal) ? statusVal : 0; idx++;
							item.EstatusText = Pago.GetStatusText(item.Estatus);
							item.Activo = Boolean.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
							item.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crVal) ? crVal : DateTime.Now.AddDays(1); idx++;

							item.EstatusClass = GetStatusClass(item.Estatus);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<Pago>();

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

		public static Pago GetById(int id)
		{
			Pago rsp = new Pago();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_PagoById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{

						int idx = 0;
						var row = dt.Rows[0];

						rsp.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0; idx++;
						rsp.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;

						if (row[idx].ToString().Length > 0)
						{
							rsp.FechaVencimiento = DateTime.Parse(row[idx].ToString());
						}
						else
						{
							rsp.FechaVencimiento = null;
						}

						idx++;

						if (row[idx].ToString().Length > 0)
						{
							rsp.FechaPago = DateTime.Parse(row[idx].ToString());
						}
						else
						{
							rsp.FechaPago = null;
						}

						idx++;

						rsp.MontoPago = decimal.TryParse(row[idx].ToString(), out decimal montoVal) ? montoVal : 0; idx++;
						rsp.RegistroPagoPor = row[idx].ToString(); idx++;
						rsp.CreadoPor = row[idx].ToString(); idx++;
						rsp.CartaCreditoId = row[idx].ToString(); idx++;
						rsp.NumeroFactura = row[idx].ToString(); idx++;
						rsp.Estatus = int.TryParse(row[idx].ToString(), out int statusVal) ? statusVal : 0; idx++;
						rsp.EstatusText = Pago.GetStatusText(rsp.Estatus);
						rsp.Activo = Boolean.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
						rsp.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crVal) ? crVal : DateTime.Now.AddDays(1); idx++;

					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Pago();
			}

			return rsp;
		}

		public static Pago GetByComisionId(int id)
		{
			Pago rsp = new Pago();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_PagoByComisionId(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{

						int idx = 0;
						var row = dt.Rows[0];

						rsp.NumeroPago = int.TryParse(row[idx].ToString(), out int numPagoVal) ? numPagoVal : 0; idx++;
						rsp.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;

						if (row[idx].ToString().Length > 0)
						{
							rsp.FechaVencimiento = DateTime.Parse(row[idx].ToString());
						}
						else
						{
							rsp.FechaVencimiento = null;
						}

						idx++;

						if (row[idx].ToString().Length > 0)
						{
							rsp.FechaPago = DateTime.Parse(row[idx].ToString());
						}
						else
						{
							rsp.FechaPago = null;
						}

						idx++;

						rsp.MontoPago = decimal.TryParse(row[idx].ToString(), out decimal montoVal) ? montoVal : 0; idx++;
						rsp.RegistroPagoPor = row[idx].ToString(); idx++;
						rsp.CreadoPor = row[idx].ToString(); idx++;
						rsp.CartaCreditoId = row[idx].ToString(); idx++;
						rsp.NumeroFactura = row[idx].ToString(); idx++;
						rsp.Estatus = int.TryParse(row[idx].ToString(), out int statusVal) ? statusVal : 0; idx++;
						rsp.EstatusText = Pago.GetStatusText(rsp.Estatus);
						rsp.Activo = Boolean.TryParse(row[idx].ToString(), out bool actVal) ? actVal : false; idx++;
						rsp.Creado = DateTime.TryParse(row[idx].ToString(), out DateTime crVal) ? crVal : DateTime.Now.AddDays(1); idx++;

					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new Pago();
			}

			return rsp;
		}

		public static RespuestaFormato Insert(Pago modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_Pago(modelo, out dt, out errores))
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

		public static RespuestaFormato ComisionInsert(Pago modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_PagoComision(modelo, out dt, out errores))
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

		public static RespuestaFormato Update(Pago modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_Pago(modelo, out dt, out errores))
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

		public static string GetStatusText(int estatusInt)
		{
			string rsp = "";

			switch (estatusInt)
			{
				case 4:
					rsp = "Refinanciado";
					break;

				case 1:
					rsp = "Pendiente";
					break;

				case 3:
					rsp = "Pagado";
					break;
			}

			return rsp;
		}

		public static string GetStatusClass(int estatusInt)
		{
			string rsp = "";

			switch (estatusInt)
			{
				case 4:
					rsp = "";
					break;

				case 1:
					rsp = "bg-yellow-100";
					break;

				case 3:
					rsp = "bg-green-300";
					break;
			}

			return rsp;
		}

	}
}