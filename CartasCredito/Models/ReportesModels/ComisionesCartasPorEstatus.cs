using CartasCredito.Models.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Drawing;
using System.Drawing.Imaging;
using File = System.IO.File;

namespace CartasCredito.Models.ReportesModels
{
	public class ComisionesCartasPorEstatus : ReporteBase
	{
		public ComisionesCartasPorEstatus(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Reporte de Comisiones de Cartas de Crédito por Estatus")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();

			try
			{
				/*ESheet.Cells.Style.Font.Size = 10;
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
				ESheet.Cells["B4"].Value = "Periodo " + FechaInicio.ToString("yyyy-MM-dd") + " - " + FechaFin.ToString("yyyy-MM-dd");*/

				ESheet.Cells["B9"].Value = "Comisión";
				ESheet.Cells["C9"].Value = "Estatus Carta";
				ESheet.Cells["D9"].Value = "Moneda";
				ESheet.Cells["E9"].Value = "Monto Programado";
				ESheet.Cells["F9"].Value = "Monto Pagado";
				ESheet.Cells["G9"].Value = "Monto Pagado en USD";

				ESheet.Cells["B9:G9"].Style.Font.Bold = true;

				ESheet.Cells["B1:G1"].Merge = true;
				ESheet.Cells["B2:G2"].Merge = true;
				ESheet.Cells["B3:G3"].Merge = true;
				ESheet.Cells["B4:G4"].Merge = true;

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(),".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20,450);

				// Fila inicial en excel
				int fila = 10;

				// Consulta de cartas
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = FechaInicio,
					FechaFin = FechaFin,
				};

				//var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.FechaVencimiento);
				//var cartasCredito = CartaCredito.Filtrar(ccFiltro).Where(cc => cc.Consecutive>1075 && cc.Consecutive<1085).OrderBy(cc => cc.FechaVencimiento);
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
					.GroupJoin(comisionesDeTodasLasCartas, tipoComision => tipoComision.Id, comision => comision.NumeroComision,
						(tipoComision, comisionesDeTipo) => new {
							TipoComision = tipoComision,
							Comisiones = comisionesDeTipo.GroupBy(comision => comision.EstatusCartaId)
						});

				foreach (var tipoYComisionesPorEstatus in agrupadoPorTipoComision)
				{
					var totalComisionProgramado = 0M;
					var totalComisionPagado = 0M;
					var hayComisiones = 0;
					foreach (var comisionGroup in tipoYComisionesPorEstatus.Comisiones)
					{
						hayComisiones++;
					}
					if(hayComisiones>0){
						ESheet.Cells[string.Format("B{0}", fila)].Value = tipoYComisionesPorEstatus.TipoComision.Nombre;
					}

					foreach (var comisionPorEstatus in tipoYComisionesPorEstatus.Comisiones)
					{
						totalComisionProgramado = 0M;
						totalComisionPagado = 0M;
						//Console.WriteLine($"Estatus: {comisionPorEstatus.Key}");
						ESheet.Cells[string.Format("C{0}", fila)].Value = CartaCredito.GetStatusText(comisionPorEstatus.Key);

						foreach (var comision in comisionPorEstatus)
						{
							ESheet.Cells[string.Format("D{0}", fila)].Value = comision.Moneda;
							ESheet.Cells[string.Format("E{0}", fila)].Value = comision.Monto;
							ESheet.Cells[string.Format("E{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";
							ESheet.Cells[string.Format("F{0}", fila)].Value = comision.MontoPagado;
							ESheet.Cells[string.Format("F{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";
							ESheet.Cells[string.Format("G{0}", fila)].Value = "MONTO USD";
							fila++;
							totalComisionProgramado += comision.Monto;
							totalComisionPagado += comision.MontoPagado;
						}

						if(hayComisiones>0){ //si tipo comision tiene comisiones>0
							ESheet.Cells[string.Format("D{0}", fila)].Value = "Total";
							ESheet.Cells[string.Format("E{0}", fila)].Value = totalComisionProgramado;
							ESheet.Cells[string.Format("E{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";
							ESheet.Cells[string.Format("F{0}", fila)].Value = totalComisionPagado;
							ESheet.Cells[string.Format("F{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";
							fila++;
							fila++;
						}	
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