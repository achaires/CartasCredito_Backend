using CartasCredito.Models.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class ComisionesCartasPorEstatus : ReporteBase
	{
		public ComisionesCartasPorEstatus(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Comisiones de Cartas de Crédito por Estatus")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();

			try
			{
				ESheet.Cells.Style.Font.Size = 10;
				ESheet.Cells["B4:G4"].Style.Font.Bold = true;

				ESheet.Cells["B1:G1"].Style.Font.Size = 22;
				ESheet.Cells["B1:G1"].Style.Font.Bold = true;
				ESheet.Cells["B1:G1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				ESheet.Cells["B2:G2"].Style.Font.Size = 16;
				ESheet.Cells["B2:G2"].Style.Font.Bold = true;
				ESheet.Cells["B2:G2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B2"].Value = ReporteFilename;

				ESheet.Cells["B4:G4"].Style.Font.Size = 16;
				ESheet.Cells["B4:G4"].Style.Font.Bold = false;
				ESheet.Cells["B4:G4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B4"].Value = "Periodo " + FechaInicio.ToString("yyyy-MM-dd") + " - " + FechaFin.ToString("yyyy-MM-dd");

				ESheet.Cells["B9"].Value = "Comisión";
				ESheet.Cells["C9"].Value = "Estatus Carta";
				ESheet.Cells["D9"].Value = "Moneda";
				ESheet.Cells["E9"].Value = "Monto Programado";
				ESheet.Cells["F9"].Value = "Monto Pagado";
				ESheet.Cells["G9"].Value = "Monto Pagado en USD";

				ESheet.Cells["B9:G9"].Style.Font.Bold = true;

				ESheet.Cells["B1:G1"].Merge = true;
				ESheet.Cells["B2:G2"].Merge = true;
				ESheet.Cells["B4:G4"].Merge = true;


				// Fila inicial en excel
				int fila = 10;

				// Consulta de cartas
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = FechaInicio,
					FechaFin = FechaFin,
				};

				var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.FechaVencimiento);
				var comisionesDeTodasLasCartas = new List<CartaCreditoComision>();

				foreach (var cc in cartasCredito)
				{
					var ccComisiones = CartaCreditoComision.GetByCartaCreditoId(cc.Id);

					comisionesDeTodasLasCartas.AddRange(ccComisiones);
				}

				// Consulta de tipos de comisión
				var tiposComisiones = TipoComision.Get(1);

				var agrupadoPorTipoComision = tiposComisiones
					.GroupJoin(comisionesDeTodasLasCartas, tipoComision => tipoComision.Id, comision => comision.ComisionId,
						(tipoComision, comisionesDeTipo) => new {
							TipoComision = tipoComision,
							Comisiones = comisionesDeTipo.GroupBy(comision => comision.EstatusCartaId)
						});

				foreach (var tipoYComisionesPorEstatus in agrupadoPorTipoComision)
				{
					ESheet.Cells[string.Format("B{0}", fila)].Value = tipoYComisionesPorEstatus.TipoComision.Nombre;

					foreach (var comisionPorEstatus in tipoYComisionesPorEstatus.Comisiones)
					{
						//Console.WriteLine($"Estatus: {comisionPorEstatus.Key}");
						ESheet.Cells[string.Format("C{0}", fila)].Value = CartaCredito.GetStatusText(comisionPorEstatus.Key);

						foreach (var comision in comisionPorEstatus)
						{
							ESheet.Cells[string.Format("D{0}", fila)].Value = comision.Moneda;
							ESheet.Cells[string.Format("E{0}", fila)].Value = comision.Monto;
							ESheet.Cells[string.Format("F{0}", fila)].Value = comision.MontoPagado;
							ESheet.Cells[string.Format("G{0}", fila)].Value = "MONTO USD";
						}

						fila++;
					}
				}


				ESheet.Cells["A:AZ"].AutoFitColumns();
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