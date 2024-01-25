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

	public class OutstandingGrupoActivo
    {
		public string TipoActivo { get; set; } = "";
		public decimal TotalMontoOriginalLC { get; set; } = 0;
		public decimal TotalPagosEfectuados { get; set; } = 0;
		public decimal TotalPlazoProveedor { get; set; } = 0;
		public decimal TotalRefinanciado { get; set; } = 0;
		public decimal TotalNoEmbarcado { get; set; } = 0;
		public decimal TotalOutstanding { get; set; } = 0;
		public List<OutstandingGrupoEmpresas> GrupoEmpresas { get; set; } = new List<OutstandingGrupoEmpresas>();
    }

	public class LineaCreditoGpo
	{
		public string Empresa { get; set; } = "";
		public int EmpresaId { get; set; } = 0;
		public string Banco { get; set; } = "";
		public int BancoId { get; set; } = 0;
		public decimal TotalEmpresaBanco { get; set; } = 0;
		public decimal TotalMontoDispuesto { get; set; } = 0;
	}

	public class OutstandingGrupoEmpresas
	{
		public string Empresa { get; set; } = "";
		public decimal TotalEmpresa { get; set; } = 0;
		public decimal TotalMontoOriginalLC { get; set; } = 0;
		public decimal TotalPagosEfectuados { get; set; } = 0;
		public decimal TotalPlazoProveedor { get; set; } = 0;
		public decimal TotalRefinanciado { get; set; } = 0;
		public decimal TotalNoEmbarcado { get; set; } = 0;
		public decimal TotalOutstanding { get; set; } = 0;
		public List<CartaCreditoReporte> CartasCredito { get; set; } = new List<CartaCreditoReporte>();
	}

	public class ComisionesPorTipoComision : ReporteBase
	{
		public ComisionesPorTipoComision(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa)
			: base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Reporte de Comisiones por Tipo de Comisión (USD)")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();
			System.Drawing.Color _BORDE = System.Drawing.Color.FromArgb(1, 191, 191, 191);
			System.Drawing.Color _MARCO = System.Drawing.Color.FromArgb(1, 0, 0, 0); //1,191,191,191

			try
			{
				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);
				//var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento).ToList();
				var cartasCredito = CartaCreditoReporte.ReporteComisiones(EmpresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento).ToList();
				
				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				var empresas = Empresa.Get(1);

				if (EmpresaId > 0)
				{
					empresas = empresas.FindAll(e => e.Id == EmpresaId);
				}

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				ESheet.Cells["B9"].Value = "Empresa";
				ESheet.Cells["C9"].Value = "Comisión";
				ESheet.Cells["D9"].Value = "Número de Carta";
				ESheet.Cells["E9"].Value = "Moneda Original";
				ESheet.Cells["F9"].Value = "Monto Programado";
				ESheet.Cells["G9"].Value = "Monto Pagado (USD)";
				ESheet.Cells["H9"].Value = "Estatus Carta";

				ESheet.Cells["B9:H9"].Style.Font.Bold = true;


				ESheet.Cells["B9:H9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:H9"].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);
				ESheet.Cells["B9:H9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:H9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:H9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:H9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

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
				var granTotalPagadoUSD = 0M;

				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				TipoDeCambio tipoDeCambio = new TipoDeCambio();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);
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
						var cartaComisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(cartaCredito.Id);
						
						if ( cartaComisiones.Count > 0 )
						{
							empresaCartasComisiones.AddRange(cartaComisiones);
						}
					}
					
					var groupedComisiones = empresaCartasComisiones.OrderBy(ecc => ecc.Comision).ThenBy(ecc => ecc.Moneda).GroupBy(ecc => ecc.Comision);

					//validar que la empresa tenga comisiones>0
					var hayComisiones = 0;
					foreach (var comisionGroup in groupedComisiones)
					{
						//hayComisiones++;
						foreach (var comision in comisionGroup)
						{
							if(comision.MontoPagado>0M){ hayComisiones++; }
						}
					}
					//si la empresa tiene comisiones>0
					if(hayComisiones>0){ ESheet.Cells[
						string.Format("B{0}", row)].Value = empresa.Nombre;

						ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
					}
					var totalEmpresaProgramado = 0M;
					var totalEmpresaPagado = 0M;
					var totalEmpresaPagadoUSD = 0M;
                    var valComisionGroup = "unog";
                    var valComision="uno";


					foreach (var comisionGroup in groupedComisiones)
					{
                        //var rowOrigin = row;
						totalEmpresaProgramado = 0M;
						totalEmpresaPagado = 0M;
						totalEmpresaPagadoUSD = 0M;
						foreach (var comision in comisionGroup)
						{
							if(comision.MontoPagado>0M){
								if(valComision!=comision.Comision){
									valComision=comision.Comision;
									ESheet.Cells[string.Format("C{0}", row)].Value = comision.Comision;
									ESheet.Cells["C" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["C" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
									ESheet.Cells["C" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["C" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
									ESheet.Cells["C" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["C" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
									ESheet.Cells["C" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["C" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
								}
								ESheet.Cells[string.Format("D{0}", row)].Value = comision.NumCartaCredito;
								ESheet.Cells[string.Format("E{0}", row)].Value = comision.Moneda;

								ESheet.Cells[string.Format("F{0}", row)].Value = comision.MontoPagado;//comision.Monto - comision.MontoPagado;
								ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = " #,##0.00";

								ESheet.Cells[string.Format("G{0}", row)].Value = comision.MontoPagado;
								ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";

								ESheet.Cells[string.Format("H{0}", row)].Value = comision.EstatusCartaText;
								//----conversion de moneda-----

								string MonedaOriginal = catMonedas.Where(m => m.Id == comision.MonedaId).FirstOrDefault().Abbr.ToUpper();
								string MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();

								/*TipoDeCambio tipoDeCambio = new TipoDeCambio();
								tipoDeCambio.Fecha = fechaDivisa;
								tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == comision.MonedaId).FirstOrDefault().Abbr.ToUpper();
								tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
								
								List<TipoDeCambio> busqueda = tiposDeCambio.Where(t => t.MonedaOriginal == tipoDeCambio.MonedaOriginal &&
								t.MonedaNueva == tipoDeCambio.MonedaNueva &&
								t.Fecha == tipoDeCambio.Fecha).ToList();
								if (busqueda.Count > 0)
								{
									tipoDeCambio = busqueda[0];
								}
								else
								{
									tipoDeCambio.Get();
								}
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
								}
								tiposDeCambio.Add(tipoDeCambio);*/
								tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == MonedaOriginal && i.MonedaNueva == MonedaNueva).FirstOrDefault();

								if (tipoDeCambio.MonedaOriginal == "JPY")
								{
									comision.MontoPagado_USD = comision.MontoPagado / tipoDeCambio.Conversion;
								}
								else
								{
									comision.MontoPagado_USD = comision.MontoPagado * tipoDeCambio.Conversion;
								}
								//comision.MontoPagado_USD = comision.MontoPagado * tipoDeCambio.Conversion;

								ESheet.Cells[string.Format("G{0}", row)].Value = comision.MontoPagado_USD;
								ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";
								//-----------------------------

								ESheet.Cells["D" + row + ":H" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                ESheet.Cells["D" + row + ":H" + row].Style.Fill.BackgroundColor.SetColor(1, 255, 255, 255);

                                ESheet.Cells["D" + row + ":H" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["D" + row + ":H" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
								ESheet.Cells["D" + row + ":H" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["D" + row + ":H" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
								ESheet.Cells["D" + row + ":H" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["D" + row + ":H" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
								ESheet.Cells["D" + row + ":H" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["D" + row + ":H" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
								row++;

								totalEmpresaProgramado += (comision.Monto - comision.MontoPagado);
								totalEmpresaPagado += comision.MontoPagado;
								totalEmpresaPagadoUSD += comision.MontoPagado_USD;
							}

							
						}

						if(valComisionGroup != comisionGroup.Key && totalEmpresaPagado>0){ //si cambia la comison de comisionGroup
							valComisionGroup = comisionGroup.Key;
							ESheet.Cells[string.Format("E{0}", row)].Value = "Total";
							ESheet.Cells[string.Format("F{0}", row)].Value = totalEmpresaPagado;
							ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = " #,##0.00";
							ESheet.Cells[string.Format("G{0}", row)].Value = totalEmpresaPagado;
							ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";
							ESheet.Cells[string.Format("G{0}", row)].Value = totalEmpresaPagadoUSD;
							ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells["E" + row + ":G" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["E" + row + ":G" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
							ESheet.Cells["E" + row + ":G" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["E" + row + ":G" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
							ESheet.Cells["E" + row + ":G" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["E" + row + ":G" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
							ESheet.Cells["E" + row + ":G" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["E" + row + ":G" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
							ESheet.Cells["E" + row + ":G" + row].Style.Font.Bold = true;

							row++;
						}

						//var rowFinal = row - 1;
						//row++;

						//Sheet.Cells[string.Format("C{0}:C{1}",rowOrigin,rowFinal)].Merge = true;
						granTotalProgramado += totalEmpresaProgramado;
						granTotalPagado += totalEmpresaPagado;
						granTotalPagadoUSD += totalEmpresaPagadoUSD;
					}
				}

				ESheet.Cells[string.Format("E{0}", row)].Value = "Gran Total";
				ESheet.Cells[string.Format("F{0}", row)].Value = granTotalPagado;
				ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[string.Format("F{0}", row)].Style.Font.Bold = true;
				ESheet.Cells[string.Format("G{0}", row)].Value = granTotalPagado;
				ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[string.Format("G{0}", row)].Style.Font.Bold = true;
				ESheet.Cells[string.Format("G{0}", row)].Value = granTotalPagadoUSD;
				ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["E" + row + ":G" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["E" + row + ":G" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":H" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row + ":H" + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				row++;
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


				ESheet.Cells["B9" + ":H9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":H9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["H9" + ":H" + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["H9" + ":H" + (row - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 1) + ":H" + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 1) + ":H" + (row - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);


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