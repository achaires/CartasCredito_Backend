using CartasCredito.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace CartasCredito.Controllers
{
	public class GeneradorReportes
	{
		private int TipoReporteId { get; }
		private DateTime FechaInicio { get; }
		private DateTime FechaFin { get; }
		private int EmpresaId { get; }
		private DateTime FechaDivisas { get; }
		private string FileName { get; }
		private int ColumnStart { get; }
		private int ColumnEnd { get; }
		private int RowStart { get; }
		private string NombreReporte { get; }

		public GeneradorReportes(int tipoReporteId, DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisas)
		{
			TipoReporteId = tipoReporteId;
			FechaInicio = fechaInicio;
			FechaFin = fechaFin;
			EmpresaId = empresaId;
			FechaDivisas = fechaDivisas;

			// Generar el nombre del archivo
			FileName = $"reporte_{tipoReporteId}_{fechaInicio.ToString("yyyyMMdd")}_{fechaFin.ToString("yyyyMMdd")}.xlsx";

			// Setear las variables de columna y fila
			switch (tipoReporteId)
			{
				case 1: // Reporte tipo A
					ColumnStart = 1;
					ColumnEnd = 5;
					RowStart = 2;
					break;
				case 2: // Reporte tipo B
					ColumnStart = 1;
					ColumnEnd = 6;
					RowStart = 3;
					break;
				default:
					throw new ArgumentException("Tipo de reporte inválido.");
			}
		}

		public void GenerateHeader(ExcelWorksheet worksheet, string nombreReporte)
		{
			// Escribir el nombre del reporte
			worksheet.Cells[1, ColumnStart].Value = nombreReporte;

			// Escribir la fecha de inicio y fin
			worksheet.Cells[2, ColumnStart].Value = "Fecha de inicio:";
			worksheet.Cells[2, ColumnStart + 1].Value = FechaInicio.ToString("dd/MM/yyyy");

			worksheet.Cells[3, ColumnStart].Value = "Fecha de fin:";
			worksheet.Cells[3, ColumnStart + 1].Value = FechaFin.ToString("dd/MM/yyyy");

			// Escribir la fecha de generación
			worksheet.Cells[4, ColumnStart].Value = "Fecha de generación:";
			worksheet.Cells[4, ColumnStart + 1].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
		}

		public void WriteRowsAndColumns(ExcelWorksheet worksheet)
		{
			// Escribir las filas y columnas
			for (int i = RowStart; i <= RowStart + 10; i++)
			{
				for (int j = ColumnStart; j <= ColumnEnd; j++)
				{
					worksheet.Cells[i, j].Value = $"Celda {i}, {j}";
				}
			}
		}

		public void StylizeWorkSheet(ExcelWorksheet worksheet)
		{
			worksheet.Cells["A:AZ"].AutoFitColumns();
		}

		public void SaveWorkbook(ExcelPackage package)
		{
			// Guardar el archivo
			// var filePath = Path.Combine(@"C:\Reports\", FileName);
			// package.SaveAs(new FileInfo(filePath));

			var path = HttpContext.Current.Server.MapPath("~/Reportes/") + FileName;
			var stream = File.Create(path);
			package.SaveAs(stream);
			stream.Close();

			var newReporte = new Reporte()
			{
				TipoReporte = NombreReporte,
				CreadoPor = "Prueba Usuario",
				CreadoPorId = "7E7836AF-0F46-4F5C-944B-194ED9D87AEF",
				Filename = FileName,
			};

			Reporte.Insert(newReporte);
		}

	}
}