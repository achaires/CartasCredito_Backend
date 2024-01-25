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
	public class StandBy : ReporteBase
	{
		public DateTime FechaVencimientoInicio { get; set; } = DateTime.Parse("1969-01-01");
		public DateTime FechaVencimientoFin { get; set; } = DateTime.Parse("1969-01-01");
		//DateTime.Parse("1969-01-01");
		//, DateTime fechaVencimientoInicio, DateTime fechaVencimientoFin
		public StandBy(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa, "A","O", "Reporte de Cartas de Crédito Stand By")//, fechaVencimientoInicio, fechaVencimientoFin)
		{
		}

		public override Reporte Generar()
		{
			System.Drawing.Color _BORDE = System.Drawing.Color.FromArgb(1, 191, 191, 191);
			System.Drawing.Color _MARCO = System.Drawing.Color.FromArgb(1, 0, 0, 0); //1,191,191,191
			var reporteRsp = new Reporte();
			try
			{

				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);
				//var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).Where(cc => cc.TipoCarta == "StandBy").GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento).ToList();
				var cartasCredito = CartaCreditoReporte.ReportestandBy(EmpresaId, fechaInicioExact, fechaFinExact, this.FechaVencimientoInicio, this.FechaVencimientoFin).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento).ToList();

				//
				/*if(this.FechaVencimientoInicio.Year != 1969 && this.FechaVencimientoFin.Year != 1969)
				{
					DateTime fvi = this.FechaVencimientoInicio.Date;
					TimeSpan ts = new TimeSpan(0, 0, 0);
					fvi = fvi.Date + ts;
					DateTime fvf = this.FechaVencimientoFin.Date;
					ts = new TimeSpan(23, 59, 59);
					fvf = fvf.Date + ts;


					cartasCredito = cartasCredito.Where(i => i.FechaVencimiento >= fvi &&
					i.FechaVencimiento <= fvf
					).ToList();
                }*/
				//


				var empresas = Empresa.Get(1);

				if (EmpresaId > 0)
				{
					empresas = empresas.FindAll(e => e.Id == EmpresaId);
				}

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);


				if (FechaVencimientoFin.Year == 2099)
				{
					ESheet.Cells["B4"].Value = ESheet.Cells["B4"].Value + " | Periodo de Vencimiento: A partir de " + FechaVencimientoInicio.ToString("dd-MM-yyyy");
				}
				else
				{
					ESheet.Cells["B4"].Value = ESheet.Cells["B4"].Value + " | Periodo de Vencimiento: Del " + FechaVencimientoInicio.ToString("dd-MM-yyyy") + " Al " + FechaVencimientoFin.ToString("dd-MM-yyyy");
				}
				//ESheet.Cells["B4"].Value += ESheet.Cells["B4"].Value + " | Periodo de Vencimiento: Del " + FechaVencimientoInicio.ToString("dd-MM-yyyy") + " Al " + FechaVencimientoFin.ToString("dd-MM-yyyy");

				ESheet.Cells["B9:O9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:O9"].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);
				ESheet.Cells["B9:O9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:O9"].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:O9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:O9"].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:O9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:O9"].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:O9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:O9"].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

				ESheet.Cells["B9"].Value = "Número de Carta";
				ESheet.Cells["C9"].Value = "Banco Emisor";
				ESheet.Cells["D9"].Value = "Banco Confirmador";
				ESheet.Cells["E9"].Value = "Referencia Banco Confirmador";
				ESheet.Cells["F9"].Value = "Empresa";
				ESheet.Cells["G9"].Value = "Proveedor";
				ESheet.Cells["H9"].Value = "Monto Apertura";
				ESheet.Cells["I9"].Value = "Fecha Apertura";
				ESheet.Cells["J9"].Value = "Fecha Vencimiento";
				ESheet.Cells["K9"].Value = "Tipo de Crédito";
				ESheet.Cells["L9"].Value = "Tipo de Cobertura";
				ESheet.Cells["M9"].Value = "Carta A";
				ESheet.Cells["N9"].Value = "Moneda";
				ESheet.Cells["O9"].Value = "Estatus Carta";

				ESheet.Cells["B9:O9"].Style.Font.Bold = true;

				ESheet.Cells["B1:O1"].Merge = true;
				ESheet.Cells["B2:O2"].Merge = true;
				ESheet.Cells["B3:O3"].Merge = true;
				ESheet.Cells["B4:O4"].Merge = true;

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(),".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20,650);

				int row = 10;
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				TipoDeCambio tipoDeCambio = new TipoDeCambio();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);

				foreach (var empresa in empresas)
				{
					var empresaCartas = cartasCredito.FindAll(cc => cc.EmpresaId == empresa.Id);

					if (empresaCartas.Count < 1)
					{
						continue;
					}


					foreach (var cartaCredito in empresaCartas)
					{
						var rowOrigin = row;
						//var bancoConfirmador=Banco.GetById(cartaCredito.BancoCorresponsalId);

						ESheet.Cells[string.Format("B{0}", row)].Value = cartaCredito.NumCartaCredito;
						ESheet.Cells[string.Format("C{0}", row)].Value = cartaCredito.Banco;
						ESheet.Cells[string.Format("D{0}", row)].Value = cartaCredito.BancoCorresponsal;
						//ESheet.Cells[string.Format("D{0}", row)].Value = bancoConfirmador.Nombre;
						//ESheet.Cells[string.Format("E{0}", row)].Value = "Referencia Banco Confirmador";
						ESheet.Cells[string.Format("E{0}", row)].Value = cartaCredito.NumCartaCredito;
						ESheet.Cells[string.Format("F{0}", row)].Value = cartaCredito.Empresa;
						ESheet.Cells[string.Format("G{0}", row)].Value = cartaCredito.Proveedor;
						ESheet.Cells[string.Format("H{0}", row)].Value = cartaCredito.MontoOriginalLC;
						ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells[string.Format("I{0}", row)].Value = cartaCredito.FechaApertura.ToString("dd-MM-yyyy");
						ESheet.Cells[string.Format("J{0}", row)].Value = cartaCredito.FechaVencimiento.ToString("dd-MM-yyyy");
						//ESheet.Cells[string.Format("K{0}", row)].Value = cartaCredito.TipoActivo;
						ESheet.Cells[string.Format("K{0}", row)].Value = cartaCredito.TipoCarta;
						ESheet.Cells[string.Format("L{0}", row)].Value = cartaCredito.TipoCobertura;// "Tipo de Cobertura";
																			//ESheet.Cells[string.Format("M{0}", row)].Value = "Carta A";
						ESheet.Cells[string.Format("M{0}", row)].Value = cartaCredito.TipoStandBy;// "A favor";
						ESheet.Cells[string.Format("N{0}", row)].Value = cartaCredito.Moneda;
						ESheet.Cells[string.Format("O{0}", row)].Value = CartaCredito.GetStatusText(cartaCredito.Estatus);
					
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":O" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

						row++;
					}

					row++;
				}

				ESheet.Cells["B9" + ":O9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":O9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B9" + ":B" + (row - 2)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + (row - 2)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["O9" + ":O" + (row - 2)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["O9" + ":O" + (row - 2)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 2) + ":O" + (row - 2)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 2) + ":O" + (row - 2)].Style.Border.Bottom.Color.SetColor(_MARCO);

				ESheet.Cells["A:AZ"].AutoFitColumns();
				var path = HttpContext.Current.Server.MapPath("~/Reportes/") + ReporteFilename;
				var stream = File.Create(path);
				EPackage.SaveAs(stream);
				stream.Close();

				var newReporte = new Reporte()
				{
					TipoReporte = "Reporte de Cartas de Crédito Stand By",
					CreadoPor = "Prueba Usuario",
					CreadoPorId = "7E7836AF-0F46-4F5C-944B-194ED9D87AEF",
					Filename = ReporteFilename,
				};

				Reporte.Insert(newReporte);
			} catch (Exception ex)
			{
				reporteRsp.Filename = ex.Message;
				reporteRsp.Id = 0;
			}

			return reporteRsp;
		}
	}
}