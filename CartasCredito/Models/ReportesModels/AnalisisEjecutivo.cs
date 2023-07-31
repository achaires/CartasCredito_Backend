using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class AnalisisEjecutivo : ReporteBase
	{
		public AnalisisEjecutivo(DateTime fechaInicio, DateTime fechaFin, int empresaId) 
			: base (fechaInicio, fechaFin, empresaId, "A", "P",  "Análisis Ejecutivo de Cartas de Crédito")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();

			try
			{
				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);

				var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).Where(cc => cc.TipoCarta == "Comercial").GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
				var catMonedas = Moneda.Get();

				// Headings
				ESheet.Cells["B9"].Value = "Empresa";
				ESheet.Cells["C9"].Value = "Banco";
				ESheet.Cells["D9"].Value = "Proveedor";
				ESheet.Cells["E9"].Value = "Producto";
				ESheet.Cells["F9"].Value = "País";
				ESheet.Cells["G9"].Value = "Descripción";
				ESheet.Cells["H9"].Value = "Moneda";
				ESheet.Cells["I9"].Value = "Importe Total";
				ESheet.Cells["J9"].Value = "Importe < 50,0000";
				ESheet.Cells["K9"].Value = "50,000 < Importe > 300,000";
				ESheet.Cells["L9"].Value = "Importe > 300,000";
				ESheet.Cells["M9"].Value = "1 Periodo";
				ESheet.Cells["N9"].Value = "2 Periodo";
				ESheet.Cells["O9"].Value = "3 Periodos";
				ESheet.Cells["P9"].Value = "Días plazo proveedor después de B/L";

				ESheet.Cells["B9:P9"].Style.Font.Bold = true;

				int fila = 10;

				var proveedoresCat = Proveedor.Get(1);

				var grupos = cartasCredito
					.GroupBy(carta => carta.Empresa)
					.Select(grupoEmpresa => new
					{
						grupoEmpresa.Key,
						TotalEmpresa = grupoEmpresa.Sum(c => c.MontoOriginalLC),
						GruposMoneda = grupoEmpresa
							.GroupBy(carta => carta.Moneda)
							.Select(grupoMoneda => new
							{
								grupoMoneda.Key,
								MonedaId = grupoMoneda.First().MonedaId,
								TotalMoneda = grupoMoneda.Sum(carta => carta.MontoOriginalLC),
								CartasDeCredito = grupoMoneda.ToList()
							}).ToList()
					}).ToList();

				var granTotal = 0M;

				var divisasList = new List<int>();

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
							var proveedorObj = proveedoresCat.First(pv => pv.Id == carta.ProveedorId);

							ESheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
							ESheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
							ESheet.Cells[string.Format("E{0}", fila)].Value = carta.DescripcionMercancia;
							ESheet.Cells[string.Format("F{0}", fila)].Value = proveedorObj.Pais;
							ESheet.Cells[string.Format("G{0}", fila)].Value = carta.TipoActivo;
							ESheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;

							ESheet.Cells[string.Format("I{0}", fila)].Value = carta.MontoOriginalLC;
							ESheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							ESheet.Cells[string.Format("J{0}", fila)].Value = carta.MontoOriginalLC < 50000 ? "Sí" : "No";
							ESheet.Cells[string.Format("K{0}", fila)].Value = carta.MontoOriginalLC > 50000 && carta.MontoOriginalLC < 300000 ? "Sí" : "No";
							ESheet.Cells[string.Format("L{0}", fila)].Value = carta.MontoOriginalLC > 300000 ? "Sí" : "No";
							ESheet.Cells[string.Format("M{0}", fila)].Value = periodos <= 1 ? "Sí" : "No";
							ESheet.Cells[string.Format("N{0}", fila)].Value = periodos == 2 ? "Sí" : "No";
							ESheet.Cells[string.Format("O{0}", fila)].Value = periodos > 2 ? "Sí" : "No";
							ESheet.Cells[string.Format("P{0}", fila)].Value = carta.DiasPlazoProveedor;

							fila++;
						}
						ESheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key;
						ESheet.Cells["I" + fila].Value = grupoMoneda.TotalMoneda;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						// Calcula y agrega fila de conversión a dólares
						fila++;
						var totalMonedaEnUsd = 0M; // ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);
						divisasList.Add(grupoMoneda.MonedaId);

						ESheet.Cells["H" + fila].Value = "Total USD";
						ESheet.Cells["I" + fila].Value = totalMonedaEnUsd;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						fila++;
						fila++;
					}

					granTotal += grupoEmpresa.TotalEmpresa;

					fila++;
					fila++;
				}

				ESheet.Cells["H" + fila].Value = "GRAN TOTAL:";
				ESheet.Cells["I" + fila].Value = granTotal;
				ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

				fila++;
				fila++;

				var sumaProcentajes = 0M;

				foreach (var grupoEmpresa in grupos)
				{
					var porcentajeEmpresa = Math.Round(Math.Round(grupoEmpresa.TotalEmpresa, 4) / granTotal, 4) * 100;

					sumaProcentajes += porcentajeEmpresa;

					ESheet.Cells["H" + fila].Value = grupoEmpresa.Key;
					ESheet.Cells["I" + fila].Value = porcentajeEmpresa + "%";

					fila++;
				}


				ESheet.Cells["A:AZ"].AutoFitColumns();

				ESheet.Column(5).Width = 25;
				ESheet.Column(4).Width = 25;
				ESheet.Column(8).Width = 25;
				ESheet.Column(11).Width = 25;
				ESheet.Column(16).Width = 25;

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
			} catch (Exception ex)
			{
				reporteResultado.Filename = ex.Message;
				reporteResultado.Id = 0;
			}

			return reporteResultado;
		}
	}
}