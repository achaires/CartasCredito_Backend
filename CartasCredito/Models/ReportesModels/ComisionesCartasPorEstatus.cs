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

		public Reporte GenerarV1()
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
				var cartasCredito = CartaCreditoReporte.ReporteComisionesEstatus(EmpresaId, fechaInicioExact, fechaFinExact).OrderBy(cc => cc.FechaVencimiento).ToList();

				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				var empresas = Empresa.Get(1);

				if (EmpresaId > 0)
				{
					empresas = empresas.FindAll(e => e.Id == EmpresaId);
				}

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				ESheet.Cells["B9"].Value = "Comisión";
				ESheet.Cells["C9"].Value = "Estatus Carta";
				ESheet.Cells["D9"].Value = "Moneda";
				ESheet.Cells["E9"].Value = "Monto programado";
				ESheet.Cells["F9"].Value = "Monto pagado";
				ESheet.Cells["G9"].Value = "Monto pagado en USD";
				ESheet.Cells["H9"].Value = "";

				ESheet.Cells["B9:H9"].Style.Font.Bold = true;


				ESheet.Cells["B9:G9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:G9"].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);
				ESheet.Cells["B9:G9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:G9"].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:G9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:G9"].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:G9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:G9"].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:G9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:G9"].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

				ESheet.Cells["B1:G1"].Merge = true;
				ESheet.Cells["B2:G2"].Merge = true;
				ESheet.Cells["B3:G3"].Merge = true;
				ESheet.Cells["B4:G4"].Merge = true;

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(), ".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20, 400);

				int row = 10;

				decimal totalMontoProgramado = 0m;
				decimal totalMontoProgramadoUSD = 0m;
				decimal totalMontoPagado = 0m;
				decimal totalMontoPagadoUSD = 0m;

				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);
				TipoDeCambio tipoDeCambio = new TipoDeCambio();

				List<string> gpoComisiones = new List<string>();
				gpoComisiones = cartasCredito.Select(i => i.TipoComision).Distinct().ToList();
				gpoComisiones.Sort();

				foreach (string comision in gpoComisiones)
				{
					int filaInicioComision = row;
					int filaFinComision = row;

					decimal COM_totalMontoProgramado = 0m;
					decimal COM_totalMontoProgramadoUSD = 0m;
					decimal COM_totalMontoPagado = 0m;
					decimal COM_totalMontoPagadoUSD = 0m;

					ESheet.Cells[string.Format("B{0}", row)].Value = comision.ToUpper();

					ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					List<CartaCreditoReporte> cartasTipoComision = new List<CartaCreditoReporte>();
					cartasTipoComision = cartasCredito.Where(j => j.TipoComision == comision).ToList();

					List<string> gpoEstatus = new List<string>();
					gpoEstatus = cartasTipoComision.Select(i => i.Estado).Distinct().ToList();
					gpoEstatus.Sort();
					foreach (string estatus in gpoEstatus)
					{
						int filaInicioEstatus = row;
						int filaFinEstatus = row;

						ESheet.Cells[string.Format("C{0}", row)].Value = estatus.ToUpper();

						ESheet.Cells["C" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["C" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["C" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["C" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

						List<CartaCreditoReporte> cartasEstado = new List<CartaCreditoReporte>();
						cartasEstado = cartasTipoComision.Where(k=> k.Estado == estatus).ToList();


						List<string> gpoMoneda = new List<string>();
						gpoMoneda = cartasEstado.Select(i => i.ComisionMoneda).Distinct().ToList();
						gpoMoneda.Sort();

						foreach (string moneda in gpoMoneda)
						{
							var mainMoenda = catMonedas.Where(m => m.Nombre.ToUpper() == moneda.ToUpper()).FirstOrDefault();
							ESheet.Cells[string.Format("D{0}", row)].Value = mainMoenda.Abbr.ToUpper();

							List<CartaCreditoReporte> cartasMoneda = new List<CartaCreditoReporte>();
							cartasMoneda = cartasEstado.Where(l => l.ComisionMoneda == moneda).ToList();

							decimal MontoProgramado = 0M;
							decimal MontoProgramadoUSD = 0M;
							decimal MontoPagado = 0M;
							decimal MontoPagadoUSD = 0M;

							//----conversion de moneda-----
							/*TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Nombre.ToUpper() == moneda.ToUpper()).FirstOrDefault().Abbr.ToUpper();
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

							string MonedaOriginal = catMonedas.Where(m => m.Nombre.ToUpper() == moneda.ToUpper()).FirstOrDefault().Abbr.ToUpper();
							string MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == MonedaOriginal && i.MonedaNueva == MonedaNueva).FirstOrDefault();

							//-----------------------------

							foreach (CartaCreditoReporte item in cartasMoneda)
                            {
								if (tipoDeCambio.MonedaOriginal == "JPY")
								{
									item.ComisionMontoUSD = item.ComisionMonto / tipoDeCambio.Conversion;
									item.ComisionMontoPagadoUSD = item.ComisionMontoPagado / tipoDeCambio.Conversion;
								}
								else
								{
									item.ComisionMontoUSD = item.ComisionMonto * tipoDeCambio.Conversion;
									item.ComisionMontoPagadoUSD = item.ComisionMontoPagado * tipoDeCambio.Conversion;
								}
							}

							MontoProgramado = cartasMoneda.Sum(m => m.ComisionMonto);
							MontoProgramadoUSD = cartasMoneda.Sum(m => m.ComisionMontoUSD);
							MontoPagado = cartasMoneda.Sum(m => m.ComisionMontoPagado);
							MontoPagadoUSD = cartasMoneda.Sum(m => m.ComisionMontoPagadoUSD);

							ESheet.Cells[string.Format("E{0}", row)].Value = MontoProgramado;
							ESheet.Cells[string.Format("E{0}", row)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells[string.Format("F{0}", row)].Value = MontoPagado;
							ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells[string.Format("G{0}", row)].Value = MontoPagadoUSD;
							ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";

							ESheet.Cells["D" + row + ":G" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":G" + row].Style.Border.Left.Color.SetColor(_BORDE);
							ESheet.Cells["D" + row + ":G" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":G" + row].Style.Border.Top.Color.SetColor(_BORDE);
							ESheet.Cells["D" + row + ":G" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":G" + row].Style.Border.Right.Color.SetColor(_BORDE);
							ESheet.Cells["D" + row + ":G" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":G" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

							COM_totalMontoProgramado += MontoProgramado;
							COM_totalMontoProgramadoUSD += MontoProgramadoUSD;
							COM_totalMontoPagado += MontoPagado;
							COM_totalMontoPagadoUSD += MontoPagadoUSD;

							row++;
						}


						filaFinEstatus = row - 1;
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Merge = true;
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + filaInicioEstatus + ":C" + filaFinEstatus].Style.Border.Bottom.Color.SetColor(_BORDE);
					}

					
					filaFinComision = row;
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Merge = true;
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;

					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + filaInicioComision + ":B" + filaFinComision].Style.Border.Bottom.Color.SetColor(_BORDE);

					ESheet.Cells[string.Format("D{0}", row)].Value = "Total en USD";

					ESheet.Cells[string.Format("E{0}", row)].Value = COM_totalMontoProgramadoUSD;
					ESheet.Cells[string.Format("E{0}", row)].Style.Numberformat.Format = " #,##0.00";

					ESheet.Cells[string.Format("F{0}", row)].Value = COM_totalMontoPagado;
					ESheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = " #,##0.00";

					ESheet.Cells[string.Format("G{0}", row)].Value = COM_totalMontoPagadoUSD;
					ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";

					ESheet.Cells["D" + row + ":G" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["D" + row + ":G" + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["D" + row + ":G" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["D" + row + ":G" + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["D" + row + ":G" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["D" + row + ":G" + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["D" + row + ":G" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["D" + row + ":G" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					ESheet.Cells["C" + row + ":G" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["C" + row + ":G" + row].Style.Fill.BackgroundColor.SetColor(1, 139, 236, 255);
					ESheet.Cells["D" + row + ":G" + row].Style.Font.Bold = true;

					totalMontoProgramado += COM_totalMontoProgramado;
					totalMontoProgramadoUSD += COM_totalMontoProgramadoUSD;
					totalMontoPagado += COM_totalMontoPagado;
					totalMontoPagadoUSD += COM_totalMontoPagadoUSD;


					row++;
				}

				ESheet.Cells[string.Format("B{0}", row)].Value = "Total en USD";
				ESheet.Cells[string.Format("G{0}", row)].Value = totalMontoPagadoUSD;
				ESheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = " #,##0.00";

				ESheet.Cells["B" + row + ":G" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				ESheet.Cells["B" + row + ":G" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row + ":G" + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);
				ESheet.Cells["B" + row + ":G" + row].Style.Font.Bold = true;
				

				ESheet.Cells["B9" + ":G9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":G9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B9" + ":B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + row].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["G9" + ":G" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["G9" + ":G" + row].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + row + ":G" + row].Style.Border.Bottom.Color.SetColor(_MARCO);

				ESheet.Cells["A:AZ"].AutoFitColumns();
				ESheet.Column(2).Width = 45;

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