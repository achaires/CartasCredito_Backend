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
	public class ComisionesPorTipoComision : ReporteBase
	{
		public ComisionesPorTipoComision(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa)
			: base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Reporte de Comisiones por Tipo de Comisión (USD)")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();

			try
			{
				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);
				var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento).ToList();

				var empresas = Empresa.Get(1);

				if (EmpresaId > 0)
				{
					empresas = empresas.FindAll(e => e.Id == EmpresaId);
				}

				ESheet.Cells["B9"].Value = "Empresa";
				ESheet.Cells["C9"].Value = "Comisión";
				ESheet.Cells["D9"].Value = "Número de Carta";
				ESheet.Cells["E9"].Value = "Moneda Original";
				ESheet.Cells["F9"].Value = "Monto Programado";
				ESheet.Cells["G9"].Value = "Monto Pagado";
				ESheet.Cells["H9"].Value = "Estatus Carta";

				ESheet.Cells["B9:H9"].Style.Font.Bold = true;

				ESheet.Cells["B1:H1"].Merge = true;
				ESheet.Cells["B2:H2"].Merge = true;
				ESheet.Cells["B3:H3"].Merge = true;
				ESheet.Cells["B4:H4"].Merge = true;

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(),".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20,400);

				int row = 10;
				var granTotalProgramado = 0M;
				var granTotalPagado = 0M;

				foreach (var empresa in empresas)
				{
					var empresaCartas = cartasCredito.FindAll(cc => cc.EmpresaId == empresa.Id);
					var empresaCartasComisiones = new List<CartaCreditoComision>();

					if (empresaCartas.Count < 1)
					{
						continue;
					}

					foreach (var cartaCredito in empresaCartas)
					{
						var cartaComisiones = CartaCreditoComision.GetByCartaCreditoId(cartaCredito.Id);
						
						if ( cartaComisiones.Count > 0 )
						{
							empresaCartasComisiones.AddRange(cartaComisiones);
						}
					}
					
					var groupedComisiones = empresaCartasComisiones.OrderBy(ecc => ecc.Comision).GroupBy(ecc => ecc.ComisionId);

					//validar que la empresa tenga comisiones>0
					var hayComisiones = 0;
					foreach (var comisionGroup in groupedComisiones)
					{
						foreach (var comision in comisionGroup)
						{
							if(comision.MontoPagado>0M){ hayComisiones++; }
						}
					}
					//si la empresa tiene comisiones>0
					if(hayComisiones>0){ ESheet.Cells[string.Format("B{0}", row)].Value = empresa.Nombre; }
					var totalEmpresaProgramado = 0M;
					var totalEmpresaPagado = 0M;
					var valComision="uno";


					foreach (var comisionGroup in groupedComisiones)
					{
						//var rowOrigin = row;

						foreach (var comision in comisionGroup)
						{
							if(comision.MontoPagado>0M){
								if(valComision!=comision.Comision){
									valComision=comision.Comision;
									ESheet.Cells[string.Format("C{0}", row)].Value = comision.Comision;
								}
								ESheet.Cells[string.Format("D{0}", row)].Value = comision.NumCartaCredito;
								ESheet.Cells[string.Format("E{0}", row)].Value = comision.Moneda;

								ESheet.Cells[string.Format("F{0}", row)].Value = comision.Monto - comision.MontoPagado;
								ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

								ESheet.Cells[string.Format("G{0}", row)].Value = comision.MontoPagado;
								ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

								ESheet.Cells[string.Format("H{0}", row)].Value = comision.EstatusCartaText;

								row++;

								totalEmpresaProgramado += (comision.Monto - comision.MontoPagado);
								totalEmpresaPagado += comision.MontoPagado;
							}
						}

						//var rowFinal = row - 1;
						//row++;

						//Sheet.Cells[string.Format("C{0}:C{1}",rowOrigin,rowFinal)].Merge = true;
					}

					granTotalProgramado += totalEmpresaProgramado;
					granTotalPagado += totalEmpresaPagado;

					if(hayComisiones>0){ //si la empresa tiene comisiones>0
					ESheet.Cells[string.Format("C{0}", row)].Value = "Total";
					ESheet.Cells[string.Format("F{0}", row)].Value = totalEmpresaProgramado;
					ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					ESheet.Cells[string.Format("G{0}", row)].Value = totalEmpresaPagado;
					ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					row++;
					}

				}

				ESheet.Cells[string.Format("C{0}", row)].Value = "Gran Total";
				ESheet.Cells[string.Format("F{0}", row)].Value = granTotalProgramado;
				ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
				ESheet.Cells[string.Format("G{0}", row)].Value = granTotalPagado;
				ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

				/*
				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(), ".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}

				var sheetLogo = Sheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(50, 50);
				*/

				ESheet.Cells["A:AZ"].AutoFitColumns();
				ESheet.Column(3).Width = 30;

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