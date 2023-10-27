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
				var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);

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
				ESheet.Cells["B4:R4"].Merge = true;


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

				foreach (var grupoEmpresa in grupos)
				{
					ESheet.Cells[string.Format("B{0}", fila)].Value = grupoEmpresa.Key;

					foreach (var grupoMoneda in grupoEmpresa.GruposMoneda)
					{
						foreach (var carta in grupoMoneda.CartasDeCredito)
						{
							TimeSpan diferencia = carta.FechaVencimiento.Subtract(carta.FechaApertura);
							int cantidadDias = (int)diferencia.TotalDays;
							decimal periodosDecimal = Convert.ToDecimal(cantidadDias) / 90M;
							var periodos = Math.Ceiling(periodosDecimal);

							ESheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
							ESheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
							ESheet.Cells[string.Format("E{0}", fila)].Value = carta.NumCartaCredito;
							ESheet.Cells[string.Format("F{0}", fila)].Value = carta.TipoActivo;
							ESheet.Cells[string.Format("G{0}", fila)].Value = carta.DescripcionCartaCredito;
							ESheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;

							ESheet.Cells[string.Format("I{0}", fila)].Value = carta.MontoOriginalLC;
							ESheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							ESheet.Cells[string.Format("J{0}", fila)].Value = carta.PagosEfectuados;
							ESheet.Cells[string.Format("J{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							ESheet.Cells[string.Format("K{0}", fila)].Value = carta.PagosProgramados;
							ESheet.Cells[string.Format("K{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							ESheet.Cells[string.Format("L{0}", fila)].Value = 0; // Refinanciado

							/*
								* Saldo insoluto por embarcar (sin % tolerancia) = Importe  –  Pagos Efectuados – Plazo Proveedor
								* */
							var saldoInsolutoReal = carta.MontoOriginalLC - carta.PagosEfectuados - carta.PagosProgramados;
							ESheet.Cells[string.Format("M{0}", fila)].Value = saldoInsolutoReal;
							ESheet.Cells[string.Format("M{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							/*
								* Saldo insoluto real + Plazo Proveedor = Saldo insoluto por embarcar (sin % tolerancia) + Plazo Proveedor
								*/
							ESheet.Cells[string.Format("N{0}", fila)].Value = saldoInsolutoReal + carta.PagosProgramados;
							ESheet.Cells[string.Format("N{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							/*
								* COMISIONES PAGADAS
								* Si hay comisiones pagadas en dólares y la carta es en yenes, deberá convertir los dólares a yenes para mostrarlas en el reporte
								*/
							ESheet.Cells[string.Format("O{0}", fila)].Value = carta.PagosComisionesEfectuados;
							ESheet.Cells[string.Format("O{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							ESheet.Cells[string.Format("P{0}", fila)].Value = carta.FechaApertura.ToString("dd-MM-yyyy");
							ESheet.Cells[string.Format("Q{0}", fila)].Value = carta.FechaVencimiento.ToString("dd-MM-yyyy");
							ESheet.Cells[string.Format("R{0}", fila)].Value = carta.DiasPlazoProveedor;

							fila++;
						}
						ESheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key + ": ";
						ESheet.Cells["I" + fila].Value = grupoMoneda.TotalMoneda;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
						fila++;
					}

					fila++;
				}


				//Sheet.Cells["A:AZ"].AutoFitColumns();
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