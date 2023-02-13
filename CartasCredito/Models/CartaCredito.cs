using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.UI;

namespace CartasCredito.Models
{
	public class CartaCredito
	{
		public int Consecutive { get; set; }
		public string Id { get; set; }
		public string NumCartaCredito { get; set; }
		public string TipoCarta { get; set; }
		public int TipoCartaId { get; set; }
		public string TipoStandBy { get; set; }
		public string TipoActivo { get; set; }
		public string Banco { get; set; }
		public string Empresa { get; set; }
		public string Proveedor { get; set; }
		public string Moneda { get; set; }
		public int TipoActivoId { get; set; }
		public int ProyectoId { get; set; }
		public string Proyecto { get; set; }
		public int BancoId { get; set; }
		public int ProveedorId { get; set; }
		public int EmpresaId { get; set; }
		public int AgenteAduanalId { get; set; }
		public string AgenteAduanal { get; set; }
		public int MonedaId { get; set; }
		public string TipoDotacion { get; set; }
		public string Responsable { get; set; }
		public int CompradorId { get; set; }
		public string Comprador { get; set; }
		public int PorcentajeTolerancia { get; set; }
		public string NumOrdenCompra { get; set; }
		public decimal CostoApertura { get; set; }
		public decimal MontoOrdenCompra { get; set; }
		public decimal MontoOriginalLC { get; set; }
		public decimal PagosEfectuados { get; set; }
		public decimal PagosProgramados { get; set; }
		public decimal MontoDispuesto { get; set; }
		public decimal SaldoInsoluto { get; set; }
		public DateTime FechaApertura { get; set; }
		public string Incoterm { get; set; }
		public DateTime FechaLimiteEmbarque { get; set; }
		public DateTime FechaVencimiento { get; set; }
		public string EmbarquesParciales { get; set; }
		public string Transbordos { get; set; }
		public string PuntoEmbarque { get; set; }
		public string PuntoDesembarque { get; set; }
		public string DescripcionMercancia { get; set; }
		public string DescripcionCartaCredito { get; set; }
		public string InstruccionesEspeciales { get; set; }
		public string PagoCartaAceptacion { get; set; }
		public string ConsignacionMercancia { get; set; }
		public string ConsideracionesAdicionales { get; set; }
		public string ConsideracionesReclamacion { get; set; }
		public int DiasParaPresentarDocumentos { get; set; }
		public int DiasPlazoProveedor { get; set; }
		public string CondicionesPago { get; set; }
		public int NumeroPeriodos { get; set; }
		public int BancoCorresponsalId { get; set; }
		public string BancoCorresponsal { get; set; }
		public string SeguroPorCuenta { get; set; }
		public string GastosComisionesCorresponsal { get; set; }
		public string ConfirmacionBancoNotificador { get; set; }
		public string TipoEmision { get; set; }
		public string DocumentoSwift { get; set; }
		//public List<DocumentoRequerido> DocumentosRequeridos { get; set; }
		public DateTime Creado { get; set; }
		public string CreadoPor { get; set; }
		public int Estatus { get; set; }
		public bool Activo { get; set; }
		//public List<Enmienda> Enmiendas { get; set; }

		/*
		public decimal TotalPagosEfectuados()
		{
			decimal result = 0;
			List<Pago> listaPagos = new List<Pago>();
			listaPagos = Pago.GetByCartaCreditoId(Id);

			foreach (var lp in listaPagos)
			{
				result += lp.MontoPagado;
			}

			return result;
		}

		public decimal TotalPagosProgramados()
		{
			decimal result = 0;
			DateTime dateNow = DateTime.Now;
			DateTime dateLimit = new DateTime(dateNow.Year, dateNow.Month, dateNow.Day);

			List<Pago> listaPagos = new List<Pago>();
			listaPagos = Pago.GetByCartaCreditoId(Id);

			foreach (Pago lp in listaPagos)
			{
				if (lp.FechaVencimiento >= dateLimit & lp.Estatus == 1)
				{
					result += lp.MontoPago;
				}
			}

			//result -= TotalPagosEfectuados();

			return result;
		}

		public decimal TotalMontoDispuesto()
		{
			decimal result = 0;

			List<Pago> listaPagos = new List<Pago>();
			listaPagos = Pago.GetByCartaCreditoId(Id);

			foreach (var lp in listaPagos)
			{
				result += lp.MontoPago;
			}

			return result;
		}

		public decimal PagosCancelados()
		{
			return 0;
		}
		*/

		public static List<CartaCredito> Get(DateTime fechaInicio, DateTime fechaFin, int activo = -1)
		{
			List<CartaCredito> res = new List<CartaCredito>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_CartasCredito(out dt, out errores, fechaInicio, fechaFin, activo))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCredito();

							item.Consecutive = int.Parse(row[idx].ToString()); idx++;
							item.Id = row[idx].ToString(); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.Moneda = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.TipoCartaId = int.TryParse(row[idx].ToString(), out int tipoCartaVal) ? tipoCartaVal : 0; idx++;
							item.TipoCarta = item.TipoCartaId == 17 ? "Comercial" : "StandBy";
							item.TipoActivoId = int.TryParse(row[idx].ToString(), out int tactval) ? tactval : 0; idx++;
							item.BancoId = int.TryParse(row[idx].ToString(), out int bncidval) ? bncidval : 0; idx++;
							item.ProyectoId = int.TryParse(row[idx].ToString(), out int pryidval) ? pryidval : 0; idx++;
							item.ProveedorId = int.TryParse(row[idx].ToString(), out int prvidval) ? prvidval : 0; idx++;
							item.EmpresaId = int.TryParse(row[idx].ToString(), out int empidval) ? empidval : 0; idx++;
							item.AgenteAduanalId = int.TryParse(row[idx].ToString(), out int agadidval) ? agadidval : 0; idx++;
							item.MonedaId = int.TryParse(row[idx].ToString(), out int mndidval) ? mndidval : 0; idx++;
							item.TipoDotacion = row[idx].ToString(); idx++;
							item.Responsable = row[idx].ToString(); idx++;
							item.CompradorId = int.TryParse(row[idx].ToString(), out int cmprdidval) ? cmprdidval : 0; idx++;
							item.PorcentajeTolerancia = int.TryParse(row[idx].ToString(), out int prctolval) ? prctolval : 0; idx++;
							item.NumOrdenCompra = row[idx].ToString(); idx++;
							item.CostoApertura = decimal.TryParse(row[idx].ToString(), out decimal costoapval) ? costoapval : 0; idx++;
							item.MontoOrdenCompra = decimal.TryParse(row[idx].ToString(), out decimal montoordenval) ? montoordenval : 0; idx++;
							item.MontoOriginalLC = decimal.TryParse(row[idx].ToString(), out decimal montooriginalval) ? montooriginalval : 0; idx++;
							item.PagosEfectuados = decimal.TryParse(row[idx].ToString(), out decimal pagosEfectuadosVal) ? pagosEfectuadosVal : 0; idx++;
							item.PagosProgramados = decimal.TryParse(row[idx].ToString(), out decimal pagosProgramadosVal) ? pagosProgramadosVal : 0; idx++;
							item.MontoDispuesto = decimal.TryParse(row[idx].ToString(), out decimal montoDispuestoVal) ? montoDispuestoVal : 0; idx++;
							item.SaldoInsoluto = decimal.TryParse(row[idx].ToString(), out decimal saldoInsolutoVal) ? saldoInsolutoVal : 0; idx++;
							item.FechaApertura = DateTime.TryParse(row[idx].ToString(), out DateTime faval) ? faval : DateTime.Now.AddDays(365); idx++;
							item.Incoterm = row[idx].ToString(); idx++;
							item.FechaLimiteEmbarque = DateTime.TryParse(row[idx].ToString(), out DateTime flimval) ? flimval : DateTime.Now.AddDays(365); idx++;
							item.EmbarquesParciales = row[idx].ToString(); idx++;
							item.Transbordos = row[idx].ToString(); idx++;
							item.PuntoEmbarque = row[idx].ToString(); idx++;
							item.PuntoDesembarque = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.PagoCartaAceptacion = row[idx].ToString(); idx++;
							item.ConsignacionMercancia = row[idx].ToString(); idx++;
							item.ConsideracionesAdicionales = row[idx].ToString(); idx++;
							item.DiasParaPresentarDocumentos = int.TryParse(row[idx].ToString(), out int diasprsval) ? diasprsval : 0; idx++;
							item.DiasPlazoProveedor = int.TryParse(row[idx].ToString(), out int diasPlazoVal) ? diasPlazoVal : 0; idx++;
							item.CondicionesPago = row[idx].ToString(); idx++;
							item.NumeroPeriodos = int.TryParse(row[idx].ToString(), out int numeroPeriodosVal) ? numeroPeriodosVal : 0; idx++;
							item.BancoCorresponsalId = int.TryParse(row[idx].ToString(), out int bancoCoritemVal) ? bancoCoritemVal : 0; idx++;
							item.SeguroPorCuenta = row[idx].ToString(); idx++;
							item.GastosComisionesCorresponsal = row[idx].ToString(); idx++;
							item.ConfirmacionBancoNotificador = row[idx].ToString(); idx++;
							item.Estatus = int.Parse(row[idx].ToString()); idx++;
							item.TipoEmision = row[idx].ToString(); idx++;
							item.DocumentoSwift = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.Creado = DateTime.Parse(row[idx].ToString()); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool activoVal) && activoVal; idx++;
							//item.Enmiendas = Enmienda.GetByCartaCreditoId(item.Id);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				res = new List<CartaCredito>();

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


		public static List<CartaCredito> Filtrar(CartasCreditoFiltrarDTO model)
		{
			List<CartaCredito> res = new List<CartaCredito>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_CartasCreditoFiltrar(model, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCredito();

							item.Consecutive = int.Parse(row[idx].ToString()); idx++;
							item.Id = row[idx].ToString(); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.Moneda = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.TipoCartaId = int.TryParse(row[idx].ToString(), out int tipoCartaVal) ? tipoCartaVal : 0; idx++;
							item.TipoCarta = item.TipoCartaId == 17 ? "Comercial" : "StandBy";
							item.TipoActivoId = int.Parse(row[idx].ToString()); idx++;
							item.BancoId = int.Parse(row[idx].ToString()); idx++;
							item.ProyectoId = int.Parse(row[idx].ToString()); idx++;
							item.ProveedorId = int.Parse(row[idx].ToString()); idx++;
							item.EmpresaId = int.Parse(row[idx].ToString()); idx++;
							item.AgenteAduanalId = int.Parse(row[idx].ToString()); idx++;
							item.MonedaId = int.Parse(row[idx].ToString()); idx++;
							item.TipoDotacion = row[idx].ToString(); idx++;
							item.Responsable = row[idx].ToString(); idx++;
							item.CompradorId = int.Parse(row[idx].ToString()); idx++;
							item.PorcentajeTolerancia = int.Parse(row[idx].ToString()); idx++;
							item.NumOrdenCompra = row[idx].ToString(); idx++;
							item.CostoApertura = decimal.Parse(row[idx].ToString()); idx++;
							item.MontoOrdenCompra = decimal.Parse(row[idx].ToString()); idx++;
							item.MontoOriginalLC = decimal.Parse(row[idx].ToString()); idx++;
							item.PagosEfectuados = decimal.TryParse(row[idx].ToString(), out decimal pagosEfectuadosVal) ? pagosEfectuadosVal : 0; idx++;
							item.PagosProgramados = decimal.TryParse(row[idx].ToString(), out decimal pagosProgramadosVal) ? pagosProgramadosVal : 0; idx++;
							item.MontoDispuesto = decimal.TryParse(row[idx].ToString(), out decimal montoDispuestoVal) ? montoDispuestoVal : 0; idx++;
							item.SaldoInsoluto = decimal.TryParse(row[idx].ToString(), out decimal saldoInsolutoVal) ? saldoInsolutoVal : 0; idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;
							item.Incoterm = row[idx].ToString(); idx++;
							item.FechaLimiteEmbarque = DateTime.Parse(row[idx].ToString()); idx++;
							item.EmbarquesParciales = row[idx].ToString(); idx++;
							item.Transbordos = row[idx].ToString(); idx++;
							item.PuntoEmbarque = row[idx].ToString(); idx++;
							item.PuntoDesembarque = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.PagoCartaAceptacion = row[idx].ToString(); idx++;
							item.ConsignacionMercancia = row[idx].ToString(); idx++;
							item.ConsideracionesAdicionales = row[idx].ToString(); idx++;
							item.DiasParaPresentarDocumentos = int.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = int.TryParse(row[idx].ToString(), out int diasPlazoVal) ? diasPlazoVal : 0; idx++;
							item.CondicionesPago = row[idx].ToString(); idx++;
							item.NumeroPeriodos = int.TryParse(row[idx].ToString(), out int numeroPeriodosVal) ? numeroPeriodosVal : 0; idx++;
							item.BancoCorresponsalId = int.TryParse(row[idx].ToString(), out int bancoCoritemVal) ? bancoCoritemVal : 0; idx++;
							item.SeguroPorCuenta = row[idx].ToString(); idx++;
							item.GastosComisionesCorresponsal = row[idx].ToString(); idx++;
							item.ConfirmacionBancoNotificador = row[idx].ToString(); idx++;
							item.Estatus = int.Parse(row[idx].ToString()); idx++;
							item.TipoEmision = row[idx].ToString(); idx++;
							item.DocumentoSwift = row[idx].ToString(); idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.Creado = DateTime.Parse(row[idx].ToString()); idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool activoVal) && activoVal; idx++;
							//item.Enmiendas = Enmienda.GetByCartaCreditoId(item.Id);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCredito>();

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

		public static RespuestaFormato InsertStandBy(CartaCredito modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_CartaCreditoStandBy(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						string id = row[3].ToString();

						if (id.Length > 0)
						{
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

		public static RespuestaFormato Insert(CartaCredito modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_CartaCredito(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						string id = row[3].ToString();

						if (id.Length > 0)
						{
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

		public static RespuestaFormato Update(CartaCredito modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_CartaCredito(modelo, out dt, out errores))
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
							rsp.DataString = row[3].ToString();
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

		public static CartaCredito GetById(string id)
		{
			CartaCredito rsp = new CartaCredito();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_CartaCreditoById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Consecutive = int.Parse(row[idx].ToString()); idx++;
						rsp.Id = row[idx].ToString(); idx++;
						rsp.TipoActivo = row[idx].ToString(); idx++;
						rsp.Banco = row[idx].ToString(); idx++;
						rsp.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;
						rsp.NumCartaCredito = row[idx].ToString(); idx++;
						rsp.TipoCartaId = int.TryParse(row[idx].ToString(), out int tipoCartaVal) ? tipoCartaVal : 0; idx++;
						rsp.TipoCarta = rsp.TipoCartaId == 17 ? "Comercial" : "StandBy";
						rsp.TipoActivoId = int.Parse(row[idx].ToString()); idx++;
						rsp.BancoId = int.Parse(row[idx].ToString()); idx++;
						rsp.ProyectoId = int.Parse(row[idx].ToString()); idx++;
						rsp.ProveedorId = int.Parse(row[idx].ToString()); idx++;
						rsp.EmpresaId = int.Parse(row[idx].ToString()); idx++;
						rsp.AgenteAduanalId = int.Parse(row[idx].ToString()); idx++;
						rsp.MonedaId = int.Parse(row[idx].ToString()); idx++;
						rsp.TipoDotacion = row[idx].ToString(); idx++;
						rsp.Responsable = row[idx].ToString(); idx++;
						rsp.CompradorId = int.Parse(row[idx].ToString()); idx++;
						rsp.PorcentajeTolerancia = int.Parse(row[idx].ToString()); idx++;
						rsp.NumOrdenCompra = row[idx].ToString(); idx++;
						rsp.CostoApertura = decimal.Parse(row[idx].ToString()); idx++;
						rsp.MontoOrdenCompra = decimal.Parse(row[idx].ToString()); idx++;
						rsp.MontoOriginalLC = decimal.Parse(row[idx].ToString()); idx++;
						rsp.PagosEfectuados = decimal.TryParse(row[idx].ToString(), out decimal pagosEfectuadosVal) ? pagosEfectuadosVal : 0; idx++;
						rsp.PagosProgramados = decimal.TryParse(row[idx].ToString(), out decimal pagosProgramadosVal) ? pagosProgramadosVal : 0; idx++;
						rsp.MontoDispuesto = decimal.TryParse(row[idx].ToString(), out decimal montoDispuestoVal) ? montoDispuestoVal : 0; idx++;
						rsp.SaldoInsoluto = decimal.TryParse(row[idx].ToString(), out decimal saldoInsolutoVal) ? saldoInsolutoVal : 0; idx++;
						rsp.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;
						rsp.Incoterm = row[idx].ToString(); idx++;
						rsp.FechaLimiteEmbarque = DateTime.Parse(row[idx].ToString()); idx++;
						rsp.EmbarquesParciales = row[idx].ToString(); idx++;
						rsp.Transbordos = row[idx].ToString(); idx++;
						rsp.PuntoEmbarque = row[idx].ToString(); idx++;
						rsp.PuntoDesembarque = row[idx].ToString(); idx++;
						rsp.DescripcionMercancia = row[idx].ToString(); idx++;
						rsp.DescripcionCartaCredito = row[idx].ToString(); idx++;
						rsp.PagoCartaAceptacion = row[idx].ToString(); idx++;
						rsp.ConsignacionMercancia = row[idx].ToString(); idx++;
						rsp.ConsideracionesAdicionales = row[idx].ToString(); idx++;
						rsp.DiasParaPresentarDocumentos = int.Parse(row[idx].ToString()); idx++;
						rsp.DiasPlazoProveedor = int.TryParse(row[idx].ToString(), out int diasPlazoVal) ? diasPlazoVal : 0; idx++;
						rsp.CondicionesPago = row[idx].ToString(); idx++;
						rsp.NumeroPeriodos = int.TryParse(row[idx].ToString(), out int numeroPeriodosVal) ? numeroPeriodosVal : 0; idx++;
						rsp.BancoCorresponsalId = int.TryParse(row[idx].ToString(), out int bancoCorrspVal) ? bancoCorrspVal : 0; idx++;
						rsp.SeguroPorCuenta = row[idx].ToString(); idx++;
						rsp.GastosComisionesCorresponsal = row[idx].ToString(); idx++;
						rsp.ConfirmacionBancoNotificador = row[idx].ToString(); idx++;
						rsp.Estatus = int.Parse(row[idx].ToString()); idx++;
						rsp.TipoEmision = row[idx].ToString(); idx++;
						rsp.DocumentoSwift = row[idx].ToString(); idx++;
						rsp.CreadoPor = row[idx].ToString(); idx++;
						rsp.Creado = DateTime.Parse(row[idx].ToString()); idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool activoVal) && activoVal; idx++;
						rsp.Proveedor = row[idx].ToString(); idx++;
						rsp.Proyecto = row[idx].ToString(); idx++;
						rsp.AgenteAduanal = row[idx].ToString(); idx++;
						rsp.Moneda = row[idx].ToString(); idx++;
						rsp.BancoCorresponsal = row[idx].ToString(); idx++;
						rsp.Empresa = row[idx].ToString(); idx++;
						rsp.Comprador = row[idx].ToString(); idx++;
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new CartaCredito();
			}

			return rsp;
		}

		public static RespuestaFormato UpdateSwiftFile(string ccid, string numeroCartaComercial, string swiftFilename)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_CartaCreditoSwift(ccid, numeroCartaComercial, swiftFilename, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						string id = row[3].ToString();

						if (id.Length > 0)
						{
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

		public string GetStatusText()
		{
			string estatusText = String.Empty;
			switch (Estatus)
			{
				case 1:
					estatusText = "Registrada";
					break;
				case 2:
					estatusText = "Emitida";
					break;
				case 3:
					estatusText = "Enmienda Pendiente";
					break;
				case 4:
					estatusText = "Pagada";
					break;
				case 5:
					estatusText = "Cancelada";
					break;
			}

			return estatusText;
		}

		public string GetStatusClass()
		{
			string estatusText = String.Empty;
			switch (Estatus)
			{
				case 1:
					estatusText = "dark";
					break;
				case 2:
					estatusText = "primary";
					break;
				case 3:
					estatusText = "warning";
					break;
				case 4:
					estatusText = "success";
					break;
				case 5:
					estatusText = "danger";
					break;
			}

			return estatusText;
		}
	}
}