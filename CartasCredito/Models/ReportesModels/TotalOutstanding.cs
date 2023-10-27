using CartasCredito.Models.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class TotalOutstanding : ReporteBase
	{
		public TotalOutstanding(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Total Outstanding")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();

			try
			{
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = FechaInicio,
					FechaFin = FechaFin,
				};

				var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.EmpresaId);
				var tipoActivoGroup = cartasCredito.GroupBy(carta => carta.TipoActivoId).Select(gpoTipoActivo => new
				{
					gpoTipoActivo.Key,
					CartasDeCredito = gpoTipoActivo.ToList()
				});

				ESheet.Cells.Style.Font.Size = 10;
				ESheet.Cells["B4:H4"].Style.Font.Bold = true;

				ESheet.Cells["B1:H1"].Style.Font.Size = 22;
				ESheet.Cells["B1:H1"].Style.Font.Bold = true;
				ESheet.Cells["B1:H1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				ESheet.Cells["B2:H2"].Style.Font.Size = 16;
				ESheet.Cells["B2:H2"].Style.Font.Bold = true;
				ESheet.Cells["B2:H2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B2"].Value = ReporteFilename;

				ESheet.Cells["B4:H4"].Style.Font.Size = 16;
				ESheet.Cells["B4:H4"].Style.Font.Bold = false;
				ESheet.Cells["B4:H4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B4"].Value = "Periodo " + FechaInicio.ToString("yyyy-MM-dd") + " - " + FechaFin.ToString("yyyy-MM-dd");

				ESheet.Cells["B9"].Value = "Tipo Activo";
				ESheet.Cells["C9"].Value = "Empresa";
				ESheet.Cells["D9"].Value = "Monto Original";
				ESheet.Cells["E9"].Value = "Pagos Efectuados";
				ESheet.Cells["F9"].Value = "Plazo Proveedor";
				ESheet.Cells["G9"].Value = "Refinanciado";
				ESheet.Cells["H9"].Value = "No Embarcado";
				ESheet.Cells["J9"].Value = "Total Outstanding";

				ESheet.Cells["B9:H9"].Style.Font.Bold = true;

				ESheet.Cells["B1:H1"].Merge = true;
				ESheet.Cells["B2:H2"].Merge = true;
				ESheet.Cells["B4:H4"].Merge = true;


				int fila = 10;

				var gruposPorTipoActivo = cartasCredito.GroupBy(c => c.TipoActivo)
									  .Select(g => new {
										  TipoActivo = g.Key,
										  TotalMontoOriginalLC = g.Sum(c => c.MontoOriginalLC),
										  TotalPagosEfectuados = g.Sum(c => c.PagosEfectuados),
										  TotalPlazoProveedor = g.Sum(c => c.PagosProgramados),
										  TotalNoEmbarcado = g.Sum(c => (c.MontoOriginalLC - c.PagosEfectuados)),
										  GrupoEmpresas = g.GroupBy(c => c.Empresa)
											.Select(ge => new {
												Empresa = ge.Key,
												TotalEmpresa = ge.Sum(c => c.MontoOriginalLC),
												CartasCredito = ge
											}).ToList()
									  })
									  .ToArray();

				foreach (var grupo in gruposPorTipoActivo)
				{
					ESheet.Cells[string.Format("B{0}", fila)].Value = grupo.TipoActivo;

					foreach (var gpoEmpresa in grupo.GrupoEmpresas)
					{
						ESheet.Cells[string.Format("C{0}", fila)].Value = gpoEmpresa.Empresa;
						ESheet.Cells[string.Format("D{0}", fila)].Value = gpoEmpresa.TotalEmpresa;
						ESheet.Cells[string.Format("E{0}", fila)].Value = 0;
						ESheet.Cells[string.Format("F{0}", fila)].Value = 0;
						ESheet.Cells[string.Format("G{0}", fila)].Value = 0;
						ESheet.Cells[string.Format("H{0}", fila)].Value = 0;

						fila++;
					}

					ESheet.Cells[string.Format("C{0}", fila)].Value = "Total por Activo";
					ESheet.Cells[string.Format("D{0}", fila)].Value = grupo.TotalMontoOriginalLC;
					ESheet.Cells[string.Format("E{0}", fila)].Value = grupo.TotalPagosEfectuados;
					ESheet.Cells[string.Format("F{0}", fila)].Value = grupo.TotalPlazoProveedor;
					ESheet.Cells[string.Format("G{0}", fila)].Value = 0;
					ESheet.Cells[string.Format("H{0}", fila)].Value = grupo.TotalNoEmbarcado;
					ESheet.Cells[string.Format("J{0}", fila)].Value = grupo.TotalMontoOriginalLC - grupo.TotalPagosEfectuados;

					fila++;

				}

				ESheet.Cells["A:AZ"].AutoFitColumns();
				ESheet.Column(5).Width = 50;
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