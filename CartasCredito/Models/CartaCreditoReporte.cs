using CartasCredito.Models.DTOs;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Web;
using System.Web.UI;

namespace CartasCredito.Models
{
	public class CTMObjReporte
    {
		public string key { get; set; } = "";
		public decimal value { get; set; } = 0;
    }

	public class CTMResumenCartas
    {

    }

	public class CTMComisiones
	{
		public int comision_id { get; set; } = 0;
		public decimal monto { get; set; } = 0;
		public decimal monto_convertido { get; set; } = 0;
		public int moneda_id { get; set; } = 0;
		public string moneda { get; set; } = "";
		public int empresa_id { get; set; } = 0;
		public string empresa { get; set; } = "";
		public DateTime fechaPago { get; set; } = DateTime.Parse("1969-01-01");
		public DateTime fechaCargo { get; set; } = DateTime.Parse("1969-01-01");
		public int mes { get; set; } = 0;
	}
	public class CartaCreditoReporte
	{
		public int TipoCoberturaId { get; set; }
		public string TipoCobertura { get; set; }
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
		public string TipoPago { get; set; }
		public string Responsable { get; set; }
		public int CompradorId { get; set; }
		public string Comprador { get; set; }
		public int PorcentajeTolerancia { get; set; }
		public string NumOrdenCompra { get; set; }
		public decimal CostoApertura { get; set; }
		public decimal MontoOrdenCompra { get; set; }
		public decimal MontoOriginalLC { get; set; }
		public decimal PagosEfectuados { get; set; }
		public decimal PagosEfectuadosUSD { get; set; } = 0;
		public decimal PagosProgramadosUSD { get; set; } = 0;
		public decimal PagosProgramados { get; set; }
		public decimal PagosComisionesEfectuados { get; set; }
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
		public List<Pago> Pagos { get; set; }
		public List<CartaCreditoComision> Comisiones { get; set; }
		public List<Enmienda> Enmiendas { get; set; }
		public decimal MontoDispuestoUSD { get; set; }



		//---
		public DateTime Actualizado { get; set; }
		public string Estado { get; set; } = "";
		public string Creado_str { get; set; } = "";
		public string Actualizado_str { get; set; } = "";
		public string FechaApertura_str { get; set; } = "";
		public string FechaLimiteEmbarque_str { get; set; } = "";
		public string FechaVencimiento_str { get; set; } = "";


		public string Pais { get; set; } = "";
		public decimal Refinanciado { get; set; } = 0;
		public int PaisId { get; set; } = 0;
		public int MesVencimiento { get; set; } = 0;
		public int AnhoVencimiento { get; set; } = 0;
		public int IdPago { get; set; } = 0;

		public int ComisionId { get; set; } = 0;
		public int TipoComisionId { get; set; } = 0;
		public string TipoComision { get; set; } = "";
		public decimal ComisionMonto { get; set; } = 0;
		public decimal ComisionMontoPagado { get; set; } = 0;
		public decimal ComisionMontoUSD { get; set; } = 0;
		public decimal ComisionMontoPagadoUSD { get; set; } = 0;
		public int ComisionMonedaId { get; set; } = 0;
		public string ComisionMoneda { get; set; } = "";
		public int DivisionId { get; set; } = 0;
		public string Division { get; set; } = "";
		public decimal MontoOriginalLCUSD { get; set; } = 0;
		public decimal APagar { get; set; } = 0;
		public decimal ARefinanciar { get; set; } = 0;

		#region CTM reportes
		public static List<CartaCreditoReporte> ReporteAnalisisEjecutivo(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaDivisa)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteAnalisisEjecutivo(empresaId, fechaInicio, fechaFin, fechaDivisa, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//
							item.MontoOriginalLCUSD = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosEfectuadosUSD = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramadosUSD = Decimal.Parse(row[idx].ToString()); idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteOutstanding(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaDivisa)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteOutstanding(empresaId, fechaInicio, fechaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;
							item.Refinanciado = Decimal.Parse(row[idx].ToString()); idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteAnalisisCartas(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaDivisa)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();

				var catMonedas = CartasCredito.Models.Moneda.Get();
				CartasCredito.Models.Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteAnalisisCartas(empresaId, fechaInicio, fechaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;
							item.Estado = row[idx].ToString(); idx++;

							/*item.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCartaWConversion(item.Id);
							decimal totalPagoComisionesEfectuados = 0;

							foreach (var comision in item.Comisiones)
							{
								totalPagoComisionesEfectuados += comision.MontoPagado;
							}*/

							item.PagosComisionesEfectuados = 0;// totalPagoComisionesEfectuados;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteComisiones(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();

				var catMonedas = CartasCredito.Models.Moneda.Get();
				CartasCredito.Models.Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteComisiones(empresaId, fechaInicio, fechaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;

							//idx++; //estatus vigencia
							//item.MontoOrdenCompra = Decimal.Parse(row[idx].ToString()); idx++;

							/*item.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(item.Id);
							decimal totalPagoComisionesEfectuados = 0;

							foreach (var comision in item.Comisiones)
							{
								comision.NumCartaCredito = item.NumCartaCredito;*/
								/*//----conversion de moneda-----
								TipoDeCambio tipoDeCambio = new TipoDeCambio();
								tipoDeCambio.Fecha = fechaDivisa;
								tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
								tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
								tipoDeCambio.Get();
								if (tipoDeCambio.ConversionStr == "-1")
								{
									tipoDeCambio.Conversion = Utility.GetTipoDeCambio(tipoDeCambio.MonedaOriginal, tipoDeCambio.MonedaNueva, tipoDeCambio.Fecha);
									if (tipoDeCambio.Conversion > -1)
									{
										TipoDeCambio.Insert(tipoDeCambio);
									}
								}
								else
								{
									tipoDeCambio.Conversion = Decimal.Parse(tipoDeCambio.ConversionStr);
								}

								tiposDeCambio.Add(tipoDeCambio);
								totalPagoComisionesEfectuados += (comision.MontoPagado * tipoDeCambio.Conversion);*/
								/*totalPagoComisionesEfectuados += comision.MontoPagado;
							}

							item.PagosComisionesEfectuados = totalPagoComisionesEfectuados;*/

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteComisionesMXP(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();

				var catMonedas = CartasCredito.Models.Moneda.Get();
				CartasCredito.Models.Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteComisionesMXP(empresaId, fechaInicio, fechaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.Creado = DateTime.Parse(row[idx].ToString()); idx++;

							idx++; //estatus vigencia
							idx++; //monto

							item.DivisionId = Int32.Parse(row[idx].ToString()); idx++;
							item.Division = row[idx].ToString(); idx++;
							idx++;
							idx++;
							idx++;

							item.ComisionMontoPagado = Decimal.Parse(row[idx].ToString()); idx++;
							idx++;
							item.ComisionMonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.MonedaId = item.ComisionMonedaId;

							item.MesVencimiento = item.Creado.Month;
							item.AnhoVencimiento = item.Creado.Year;

							//item.MontoOrdenCompra = Decimal.Parse(row[idx].ToString()); idx++;

							/*item.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(item.Id);
							decimal totalPagoComisionesEfectuados = 0;

							foreach (var comision in item.Comisiones)
							{
								comision.NumCartaCredito = item.NumCartaCredito;*/
							/*//----conversion de moneda-----
							TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.Get();
							if (tipoDeCambio.ConversionStr == "-1")
							{
								tipoDeCambio.Conversion = Utility.GetTipoDeCambio(tipoDeCambio.MonedaOriginal, tipoDeCambio.MonedaNueva, tipoDeCambio.Fecha);
								if (tipoDeCambio.Conversion > -1)
								{
									TipoDeCambio.Insert(tipoDeCambio);
								}
							}
							else
							{
								tipoDeCambio.Conversion = Decimal.Parse(tipoDeCambio.ConversionStr);
							}

							tiposDeCambio.Add(tipoDeCambio);
							totalPagoComisionesEfectuados += (comision.MontoPagado * tipoDeCambio.Conversion);*/
							/*totalPagoComisionesEfectuados += comision.MontoPagado;
						}

						item.PagosComisionesEfectuados = totalPagoComisionesEfectuados;*/

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteVencimientos(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaVInicio, DateTime fechVaFin)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();
			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteVencimientos(empresaId, fechaInicio, fechaFin, fechaVInicio, fechVaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;

							//idx++; //estatus vigencia
							//item.MontoOrdenCompra = Decimal.Parse(row[idx].ToString()); idx++;

							/*item.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(item.Id);
							decimal totalPagoComisionesEfectuados = 0;

							foreach (var comision in item.Comisiones)
							{
								comision.NumCartaCredito = item.NumCartaCredito;*/
							/*//----conversion de moneda-----
							TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.Get();
							if (tipoDeCambio.ConversionStr == "-1")
							{
								tipoDeCambio.Conversion = Utility.GetTipoDeCambio(tipoDeCambio.MonedaOriginal, tipoDeCambio.MonedaNueva, tipoDeCambio.Fecha);
								if (tipoDeCambio.Conversion > -1)
								{
									TipoDeCambio.Insert(tipoDeCambio);
								}
							}
							else
							{
								tipoDeCambio.Conversion = Decimal.Parse(tipoDeCambio.ConversionStr);
							}

							tiposDeCambio.Add(tipoDeCambio);
							totalPagoComisionesEfectuados += (comision.MontoPagado * tipoDeCambio.Conversion);*/
							/*totalPagoComisionesEfectuados += comision.MontoPagado;
						}

						item.PagosComisionesEfectuados = totalPagoComisionesEfectuados;*/

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReportestandBy(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaVInicio, DateTime fechVaFin)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();
			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteStandBy(empresaId, fechaInicio, fechaFin, fechaVInicio, fechVaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;

							idx++; //estatus vigencia
							item.MontoOrdenCompra = Decimal.Parse(row[idx].ToString()); idx++;
							item.CostoApertura = Decimal.Parse(row[idx].ToString()); idx++;

							item.BancoCorresponsalId = Int32.Parse(row[idx].ToString()); idx++;
							item.BancoCorresponsal = row[idx].ToString(); idx++;

							item.TipoCoberturaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCobertura = row[idx].ToString(); idx++;

							item.TipoStandBy = row[idx].ToString(); idx++;
							/*item.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(item.Id);
							decimal totalPagoComisionesEfectuados = 0;

							foreach (var comision in item.Comisiones)
							{
								comision.NumCartaCredito = item.NumCartaCredito;*/
							/*//----conversion de moneda-----
							TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.Get();
							if (tipoDeCambio.ConversionStr == "-1")
							{
								tipoDeCambio.Conversion = Utility.GetTipoDeCambio(tipoDeCambio.MonedaOriginal, tipoDeCambio.MonedaNueva, tipoDeCambio.Fecha);
								if (tipoDeCambio.Conversion > -1)
								{
									TipoDeCambio.Insert(tipoDeCambio);
								}
							}
							else
							{
								tipoDeCambio.Conversion = Decimal.Parse(tipoDeCambio.ConversionStr);
							}

							tiposDeCambio.Add(tipoDeCambio);
							totalPagoComisionesEfectuados += (comision.MontoPagado * tipoDeCambio.Conversion);*/
							/*totalPagoComisionesEfectuados += comision.MontoPagado;
						}

						item.PagosComisionesEfectuados = totalPagoComisionesEfectuados;*/

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteVencimientosPagos(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaVInicio, DateTime fechVaFin)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();
			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteVencimientosPagos(empresaId, fechaInicio, fechaFin, fechaVInicio, fechVaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.Creado = DateTime.Parse(row[idx].ToString()); idx++;
							item.IdPago = Int32.Parse(row[idx].ToString()); idx++;

							item.MesVencimiento = item.FechaVencimiento.Month;
							item.AnhoVencimiento = item.FechaVencimiento.Year;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteComisionesEstatus(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();

				var catMonedas = CartasCredito.Models.Moneda.Get();
				CartasCredito.Models.Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteComisionesEstatus(empresaId, fechaInicio, fechaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;
							idx++; //estatus vigencia
							idx++; //monto

							item.ComisionId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoComisionId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoComision = row[idx].ToString(); idx++;

							item.ComisionMonto = Decimal.Parse(row[idx].ToString()); idx++;
							item.ComisionMontoPagado = Decimal.Parse(row[idx].ToString()); idx++;
							item.ComisionMonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.ComisionMoneda = row[idx].ToString(); idx++;

							item.Estado = CartaCredito.GetStatusText(item.Estatus);
							//
							//item.MontoOrdenCompra = Decimal.Parse(row[idx].ToString()); idx++;

							/*item.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(item.Id);
							decimal totalPagoComisionesEfectuados = 0;

							foreach (var comision in item.Comisiones)
							{
								comision.NumCartaCredito = item.NumCartaCredito;*/
							/*//----conversion de moneda-----
							TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.Get();
							if (tipoDeCambio.ConversionStr == "-1")
							{
								tipoDeCambio.Conversion = Utility.GetTipoDeCambio(tipoDeCambio.MonedaOriginal, tipoDeCambio.MonedaNueva, tipoDeCambio.Fecha);
								if (tipoDeCambio.Conversion > -1)
								{
									TipoDeCambio.Insert(tipoDeCambio);
								}
							}
							else
							{
								tipoDeCambio.Conversion = Decimal.Parse(tipoDeCambio.ConversionStr);
							}

							tiposDeCambio.Add(tipoDeCambio);
							totalPagoComisionesEfectuados += (comision.MontoPagado * tipoDeCambio.Conversion);*/
							/*totalPagoComisionesEfectuados += comision.MontoPagado;
						}

						item.PagosComisionesEfectuados = totalPagoComisionesEfectuados;*/

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteLineasCredito(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			List<CartaCreditoReporte> res = new List<CartaCreditoReporte>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();

				var catMonedas = CartasCredito.Models.Moneda.Get();
				CartasCredito.Models.Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteLineasCredito(empresaId, fechaInicio, fechaFin, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;

							//item.Comisiones = CartaCreditoComision.GetByCartaCreditoId(item.Id);


							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;
							idx++; //estatus vigencia
							idx++; //monto

							item.Estado = CartaCredito.GetStatusText(item.Estatus);
							//
							//item.MontoOrdenCompra = Decimal.Parse(row[idx].ToString()); idx++;

							/*item.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(item.Id);
							decimal totalPagoComisionesEfectuados = 0;

							foreach (var comision in item.Comisiones)
							{
								comision.NumCartaCredito = item.NumCartaCredito;*/
							/*//----conversion de moneda-----
							TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.Get();
							if (tipoDeCambio.ConversionStr == "-1")
							{
								tipoDeCambio.Conversion = Utility.GetTipoDeCambio(tipoDeCambio.MonedaOriginal, tipoDeCambio.MonedaNueva, tipoDeCambio.Fecha);
								if (tipoDeCambio.Conversion > -1)
								{
									TipoDeCambio.Insert(tipoDeCambio);
								}
							}
							else
							{
								tipoDeCambio.Conversion = Decimal.Parse(tipoDeCambio.ConversionStr);
							}

							tiposDeCambio.Add(tipoDeCambio);
							totalPagoComisionesEfectuados += (comision.MontoPagado * tipoDeCambio.Conversion);*/
							/*totalPagoComisionesEfectuados += comision.MontoPagado;
						}

						item.PagosComisionesEfectuados = totalPagoComisionesEfectuados;*/

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				res = new List<CartaCreditoReporte>();

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

		public static List<CartaCreditoReporte> ReporteResumen(int empresaId, DateTime fechaInicio, DateTime fechaFin, out List<CTMComisiones> comisiones, out List<CTMComisiones> vencimientos)
		{
			List<CartaCreditoReporte> cartas = new List<CartaCreditoReporte>();
			comisiones = new List<CTMComisiones>();
			vencimientos = new List<CTMComisiones>();
			List<TipoDeCambio> tipoDeCambios = new List<TipoDeCambio>();

			try
			{
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();

				var catMonedas = CartasCredito.Models.Moneda.Get();
				CartasCredito.Models.Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var dt2 = new System.Data.DataTable();
				var dt3 = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteResumen(empresaId, fechaInicio, fechaFin, out dt, out errores, out dt2, out dt3))
				{
					if (dt.Rows.Count > 0)
					{	for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoReporte();

							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EmpresaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.BancoId = Int32.Parse(row[idx].ToString()); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.ProveedorId = Int32.Parse(row[idx].ToString()); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PaisId = Int32.Parse(row[idx].ToString()); idx++;
							item.Pais = row[idx].ToString(); idx++;
							item.TipoActivoId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoActivo = row[idx].ToString(); idx++;
							item.MonedaId = Int32.Parse(row[idx].ToString()); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							item.MontoOriginalLC = Decimal.Parse(row[idx].ToString()); idx++;
							item.Estatus = Int32.Parse(row[idx].ToString()); idx++;
							item.DiasPlazoProveedor = Int32.Parse(row[idx].ToString()); idx++;

							item.PagosEfectuados = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosProgramados = Decimal.Parse(row[idx].ToString()); idx++;

							item.TipoCartaId = Int32.Parse(row[idx].ToString()); idx++;
							item.TipoCarta = row[idx].ToString(); idx++;
							item.FechaVencimiento = DateTime.Parse(row[idx].ToString()); idx++;
							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.FechaApertura = DateTime.Parse(row[idx].ToString()); idx++;
							idx++; //estatus vigencia
							idx++; //monto
							item.TipoEmision = row[idx].ToString(); idx++;
							item.MontoOriginalLCUSD = Decimal.Parse(row[idx].ToString()); idx++;
							item.PagosEfectuadosUSD = Decimal.Parse(row[idx].ToString()); idx++;
							item.CondicionesPago = row[idx].ToString(); idx++;

							if (item.CondicionesPago == "Pago refinanciado")
							{
								item.ARefinanciar = item.MontoOriginalLCUSD - item.PagosEfectuadosUSD;
                            }
                            else
                            {
								item.APagar = item.MontoOriginalLCUSD - item.PagosEfectuadosUSD;
							}
							cartas.Add(item);
						}
					}

					if (dt2.Rows.Count > 0)
					{
						for (int i = 0; i < dt2.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt2.Rows[i];
							var item = new CTMComisiones();

							item.moneda_id = Int32.Parse(row[idx].ToString()); idx++;
							item.moneda = row[idx].ToString(); idx++;
							item.monto = Decimal.Parse(row[idx].ToString()); idx++;
							item.monto_convertido = Decimal.Parse(row[idx].ToString()); idx++;
							item.empresa_id = Int32.Parse(row[idx].ToString()); idx++;
							item.empresa = row[idx].ToString(); idx++;
							item.comision_id = Int32.Parse(row[idx].ToString()); idx++;
							item.fechaPago = DateTime.Parse(row[idx].ToString()); idx++;
							item.mes = item.fechaPago.Month;
							comisiones.Add(item);
						}
					}

					if (dt3.Rows.Count > 0)
					{
						for (int i = 0; i < dt3.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt3.Rows[i];
							var item = new CTMComisiones();

							item.moneda_id = Int32.Parse(row[idx].ToString()); idx++;
							item.moneda = row[idx].ToString(); idx++;
							item.monto = Decimal.Parse(row[idx].ToString()); idx++;
							item.monto_convertido = Decimal.Parse(row[idx].ToString()); idx++;
							item.empresa_id = Int32.Parse(row[idx].ToString()); idx++;
							item.empresa = row[idx].ToString(); idx++;
							item.comision_id = Int32.Parse(row[idx].ToString()); idx++;
							item.fechaCargo = DateTime.Parse(row[idx].ToString()); idx++;
							item.mes = item.fechaCargo.Month;
							vencimientos.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
				cartas = new List<CartaCreditoReporte>();

				// Get stack trace for the exception with source file information
				var st = new StackTrace(ex, true);
				// Get the top stack frame
				var frame = st.GetFrame(0);
				// Get the line number from the stack frame
				var line = frame.GetFileLineNumber();

				var errorMsg = ex.ToString();
			}

			return cartas;
		}
		#endregion
	}
}