using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class LineasDeCreditoDisponibles : ReporteBase
	{
		public LineasDeCreditoDisponibles(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Líneas de Crédito Disponibles")
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

				ESheet.Cells["B9"].Value = "Empresa";

				var lineasBancos = LineaDeCredito.Get(1);
				var filaInicio = 10;
				var columnaInicio = 2;

				ESheet.Cells[filaInicio, columnaInicio].Value = "Banco / Empresa";

				// Recorrer las empresas y agregarlas como filas
				int filaActual = filaInicio + 1;
				var empresas = lineasBancos.Select(lc => lc.Empresa).Distinct();
				foreach (var empresa in empresas)
				{
					ESheet.Cells[filaActual, columnaInicio].Value = empresa;
					filaActual++;
				}

				// Recorrer los bancos y agregarlos como columnas
				int columnaActual = columnaInicio + 1;
				var bancos = lineasBancos.Select(lc => lc.Banco).Distinct();
				foreach (var banco in bancos)
				{
					ESheet.Cells[filaInicio, columnaActual].Value = banco;
					columnaActual++;
				}

				// Recorrer las líneas de crédito y agregar el monto en la coordenada correspondiente
				foreach (var lineaCredito in lineasBancos)
				{
					// Encontrar la fila y columna correspondiente a la línea de crédito
					int fila = filaInicio + empresas.ToList().IndexOf(lineaCredito.Empresa) + 1;
					int columna = columnaInicio + bancos.ToList().IndexOf(lineaCredito.Banco) + 1;

					// Agregar el monto en la coordenada correspondiente
					ESheet.Cells[fila, columna].Value = lineaCredito.Monto;
				}


				ESheet.Cells["B1:G1"].Merge = true;
				ESheet.Cells["B2:G2"].Merge = true;
				ESheet.Cells["B4:G4"].Merge = true;


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