using CartasCredito.Models.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class AnalisisCartas : ReporteBase
	{
		public AnalisisCartas(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Análisis de Cartas de Crédito")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();

			try
			{
				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);
				//var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
				var cartasCredito = CartaCreditoReporte.ReporteAnalisisCartas(EmpresaId, fechaInicioExact, fechaFinExact, fechaDivisa).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				ESheet.Cells["B9"].Value = "Empresa";
				ESheet.Cells["C9"].Value = "Banco";
				ESheet.Cells["D9"].Value = "Proveedor";
				ESheet.Cells["E9"].Value = "No. CC";
				ESheet.Cells["F9"].Value = "Tipo Activo";
				ESheet.Cells["G9"].Value = "Descripción";
				ESheet.Cells["H9"].Value = "Moneda";
				ESheet.Cells["I9"].Value = "Importe";
				ESheet.Cells["J9"].Value = "Pagos Efectuados";
				ESheet.Cells["K9"].Value = "Plazo Proveedor";
				ESheet.Cells["L9"].Value = "Refinanciado";
				ESheet.Cells["M9"].Value = "Saldo insoluto por embarcar (sin % tolerancia) ";
				ESheet.Cells["N9"].Value = "Saldo insoluto real + Plazo Proveedor";
				ESheet.Cells["O9"].Value = "Comisiones Pagadas";
				ESheet.Cells["P9"].Value = "Fecha Apertura";
				ESheet.Cells["Q9"].Value = "Fecha Vencimiento";
				ESheet.Cells["R9"].Value = "Días plazo proveedor despúes de B/L";

				ESheet.Cells["B9:R9"].Style.Font.Bold = true;

				ESheet.Cells["B1:R1"].Merge = true;
				ESheet.Cells["B2:R2"].Merge = true;
				ESheet.Cells["B3:R3"].Merge = true;
				ESheet.Cells["B4:R4"].Merge = true;

				ESheet.Cells["B9:R9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:R9"].Style.Fill.BackgroundColor.SetColor(1,180,198,231);

				ESheet.Cells["B9:R9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:R9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:R9"].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:R9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:R9"].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:R9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:R9"].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:R9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:R9"].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

				int fila = 10;

				var grupos = cartasCredito
					.GroupBy(carta => carta.Empresa)
					.Select(grupoEmpresa => new
					{
						grupoEmpresa.Key,
						GruposMoneda = grupoEmpresa
							.GroupBy(carta => carta.Moneda)
							.Select(grupoMoneda => new
							{
								grupoMoneda.Key,
								TotalMoneda = grupoMoneda.Sum(carta => carta.MontoOriginalLC),
								CartasDeCredito = grupoMoneda.ToList()
							}).ToList()
					}).ToList();

				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);
				TipoDeCambio tipoDeCambio = new TipoDeCambio();
				foreach (var grupoEmpresa in grupos)
				{
					ESheet.Cells[string.Format("B{0}", fila)].Value = grupoEmpresa.Key;
					ESheet.Cells["B" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

					foreach (var grupoMoneda in grupoEmpresa.GruposMoneda)
					{
						decimal gpoTotal = 0;
						decimal gpoPagosEfectuados = 0;
						decimal gpoPagosProgramados = 0;
						decimal gpoRefinanciado = 0;
						decimal gpoSaldoTolerancia = 0;
						decimal gpoSaldoProveedor = 0;
						decimal gpoComisiones = 0;

						foreach (var carta in grupoMoneda.CartasDeCredito)
						{
							TimeSpan diferencia = carta.FechaVencimiento.Subtract(carta.FechaApertura);
							int cantidadDias = (int)diferencia.TotalDays;
							decimal periodosDecimal = Convert.ToDecimal(cantidadDias) / 90M;
							var periodos = Math.Ceiling(periodosDecimal);

							//----conversion de moneda-----
							tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == carta.MonedaId).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio.Conversion = 1;
							/*tipoDeCambio.Get();
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
							}*/

							tiposDeCambio.Add(tipoDeCambio);
							decimal _MontoOriginalLC = 0;
							decimal _PagosEfectuados = 0;
							decimal _PagosProgramados = 0;
							_MontoOriginalLC = carta.MontoOriginalLC * tipoDeCambio.Conversion;
							_PagosEfectuados = carta.PagosEfectuados * tipoDeCambio.Conversion;
							_PagosProgramados = carta.PagosProgramados * tipoDeCambio.Conversion;
							//-----------------------------

							ESheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
							ESheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
							ESheet.Cells[string.Format("E{0}", fila)].Value = carta.NumCartaCredito;
							ESheet.Cells[string.Format("F{0}", fila)].Value = carta.TipoActivo;
							ESheet.Cells[string.Format("G{0}", fila)].Value = carta.DescripcionCartaCredito;
							ESheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;


							if (_MontoOriginalLC <= 0)
							{
								_MontoOriginalLC = 0;
							}
							if (_PagosEfectuados <= 0)
							{
								_PagosEfectuados = 0;
							}
							if (_PagosProgramados <= 0)
							{
								_PagosProgramados = 0;
							}


							ESheet.Cells[string.Format("I{0}", fila)].Value = _MontoOriginalLC;
							ESheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells[string.Format("J{0}", fila)].Value = _PagosEfectuados;
							ESheet.Cells[string.Format("J{0}", fila)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells[string.Format("K{0}", fila)].Value = _PagosProgramados;
							ESheet.Cells[string.Format("K{0}", fila)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells[string.Format("L{0}", fila)].Value = 0; // Refinanciado
							ESheet.Cells[string.Format("L{0}", fila)].Style.Numberformat.Format = " #,##0.00";

							/*
								* Saldo insoluto por embarcar (sin % tolerancia) = Importe  –  Pagos Efectuados – Plazo Proveedor
								* */
							var saldoInsolutoReal = _MontoOriginalLC - _PagosEfectuados - _PagosProgramados;
							if (saldoInsolutoReal <= 0)
							{
								saldoInsolutoReal = 0;
							}
							ESheet.Cells[string.Format("M{0}", fila)].Value = saldoInsolutoReal;
							ESheet.Cells[string.Format("M{0}", fila)].Style.Numberformat.Format = " #,##0.00";

							/*
								* Saldo insoluto real + Plazo Proveedor = Saldo insoluto por embarcar (sin % tolerancia) + Plazo Proveedor
								*/
							var saldoInsolutoProveedor = saldoInsolutoReal + _PagosProgramados;
							if (saldoInsolutoProveedor <= 0)
							{
								saldoInsolutoProveedor = 0;
							}
							ESheet.Cells[string.Format("N{0}", fila)].Value = saldoInsolutoProveedor;
							ESheet.Cells[string.Format("N{0}", fila)].Style.Numberformat.Format = " #,##0.00";

							/*
								* COMISIONES PAGADAS
								* Si hay comisiones pagadas en dólares y la carta es en yenes, deberá convertir los dólares a yenes para mostrarlas en el reporte
								*/

							carta.Comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(carta.Id);
							carta.PagosComisionesEfectuados = 0;
							foreach (var comision in carta.Comisiones)
							{
								if (comision.Estatus == 3)
								{
									string MonedaOriginal = catMonedas.Where(m => m.Id == comision.MonedaId).FirstOrDefault().Abbr.ToUpper();
									string MonedaNueva = catMonedas.Where(m => m.Id == carta.MonedaId).FirstOrDefault().Abbr.ToUpper();
									tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == MonedaOriginal && i.MonedaNueva == MonedaNueva).FirstOrDefault();
									decimal monto_convertido = 0M;

									if (MonedaOriginal == "JPY" && MonedaNueva == "USD")
									{
										monto_convertido = comision.MontoPagado / tipoDeCambio.Conversion;
									}
									else
									{
										monto_convertido = comision.MontoPagado * tipoDeCambio.Conversion;
									}

									carta.PagosComisionesEfectuados += monto_convertido;
								}
							}
							if (carta.PagosComisionesEfectuados <= 0)
							{
								carta.PagosComisionesEfectuados = 0;
							}
							ESheet.Cells[string.Format("O{0}", fila)].Value = carta.PagosComisionesEfectuados;
							ESheet.Cells[string.Format("O{0}", fila)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells[string.Format("P{0}", fila)].Value = carta.FechaApertura.ToString("dd-MM-yyyy");
							ESheet.Cells[string.Format("Q{0}", fila)].Value = carta.FechaVencimiento.ToString("dd-MM-yyyy");
							ESheet.Cells[string.Format("R{0}", fila)].Value = carta.DiasPlazoProveedor;

							//ESheet.Cells[string.Format("T{0}", fila)].Value = carta.Estado;


							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

							gpoTotal += _MontoOriginalLC;
							gpoPagosEfectuados += _PagosEfectuados;
							gpoPagosProgramados += _PagosProgramados;
							gpoSaldoTolerancia += saldoInsolutoReal;
							gpoSaldoProveedor += saldoInsolutoProveedor;
							gpoComisiones += carta.PagosComisionesEfectuados;
							fila++;
						}

						ESheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key + ": ";
						ESheet.Cells["I" + fila].Value = gpoTotal;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";


						ESheet.Cells["J" + fila].Value = gpoPagosEfectuados;
						ESheet.Cells["J" + fila].Style.Numberformat.Format = " #,##0.00";


						ESheet.Cells["K" + fila].Value = gpoPagosProgramados;
						ESheet.Cells["K" + fila].Style.Numberformat.Format = " #,##0.00";


						ESheet.Cells["L" + fila].Value = gpoRefinanciado;
						ESheet.Cells["L" + fila].Style.Numberformat.Format = " #,##0.00";


						ESheet.Cells["M" + fila].Value = gpoSaldoTolerancia;
						ESheet.Cells["M" + fila].Style.Numberformat.Format = " #,##0.00";

						ESheet.Cells["N" + fila].Value = gpoSaldoProveedor;
						ESheet.Cells["N" + fila].Style.Numberformat.Format = " #,##0.00";

						ESheet.Cells["O" + fila].Value = gpoComisiones;
						ESheet.Cells["O" + fila].Style.Numberformat.Format = " #,##0.00";


						ESheet.Cells["H" + fila + ":O" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["H" + fila + ":O" + fila].Style.Fill.BackgroundColor.SetColor(1, 221, 221, 221);
						ESheet.Cells["H" + fila + ":O" + fila].Style.Font.Bold = true;

						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":R" + fila].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
						fila++;
					}

					//fila++;
				}
				System.Drawing.Color marco = System.Drawing.Color.FromArgb(1, 0, 0, 0); //1,191,191,191
				ESheet.Cells["B9" + ":R9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":R9"].Style.Border.Top.Color.SetColor(marco);
				ESheet.Cells["B9" + ":B" + (fila - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + (fila - 1)].Style.Border.Left.Color.SetColor(marco);
				ESheet.Cells["R9" + ":R" + (fila - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["R9" + ":R" + (fila - 1)].Style.Border.Right.Color.SetColor(marco);
				ESheet.Cells["B" + (fila - 1) + ":R" + (fila - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (fila - 1) + ":R" + (fila - 1)].Style.Border.Bottom.Color.SetColor(marco);



				ESheet.Cells["A:AZ"].AutoFitColumns();
				ESheet.Column(2).Width = 20;
				ESheet.Column(3).Width = 15;
				ESheet.Column(4).Width = 20;
				ESheet.Column(5).Width = 15;
				ESheet.Column(8).Width = 15;
				ESheet.Column(9).Width = 15;
				ESheet.Column(10).Width = 15;
				ESheet.Column(11).Width = 15;
				ESheet.Column(12).Width = 15;
				ESheet.Column(13).Width = 15;
				ESheet.Column(14).Width = 15;
				ESheet.Column(15).Width = 15;
				ESheet.Column(19).Width = 10;

				ESheet.Column(2).Style.WrapText = true;
				ESheet.Column(3).Style.WrapText = true;

				var path = HttpContext.Current.Server.MapPath("~/Reportes/") + ReporteFilename;
				var stream = File.Create(path);
				EPackage.SaveAs(stream);
				stream.Close();

				var newReporte = new Reporte()
				{
					TipoReporte = ReporteNombre,
					CreadoPor = "Prueba Usuario",
					CreadoPorId = "7E7836AF-0F46-4F5C-944B-194ED9D87AEF",
					Filename = ReporteFilename,
				};

				Reporte.Insert(newReporte);
				
			}
			catch (Exception ex)
			{
				reporteResultado.Filename = ex.Message;
				reporteResultado.Id = 0;
			}

			return reporteResultado;
		}
	}
}