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
	public class Vencimientos : ReporteBase
	{
		public DateTime FechaVencimientoInicio { get; set; } = DateTime.Parse("1969-01-01");
		public DateTime FechaVencimientoFin { get; set; } = DateTime.Parse("1969-01-01");
		public Vencimientos(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa, "A", "K", "Reporte de Vencimientos de Cartas de Crédito")
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
					FechaVencimientoInicio = this.FechaVencimientoInicio,
					FechaVencimientoFin = this.FechaVencimientoFin,
					TipoCarta = "0"
				};

				//var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.FechaVencimiento);
				//var cartasCredito = CartaCredito.Filtrar(ccFiltro).Where(cc => cc.Consecutive<10).OrderBy(cc => cc.Moneda).ThenBy(cc => cc.FechaVencimiento);
				//List<CartaCredito> cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.Moneda).ThenBy(cc => cc.FechaVencimiento).ToList();
				//List<CartaCredito> cartasCredito = CartaCredito.FiltrarReporte(ccFiltro).OrderBy(cc => cc.Moneda).ThenBy(cc => cc.FechaVencimiento).ToList();
				List<CartaCreditoReporte> cartasCredito = CartaCreditoReporte.ReporteVencimientos(this.EmpresaId, FechaInicio, FechaFin, this.FechaVencimientoInicio, this.FechaVencimientoFin).OrderBy(cc => cc.Moneda).ThenBy(cc => cc.FechaVencimiento).ToList();
				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				/*if (this.FechaVencimientoInicio.Year != 1969 && this.FechaVencimientoFin.Year != 1969)
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

				var monedas = Moneda.Get(1);

				/*ESheet.Cells.Style.Font.Size = 10;
				ESheet.Cells["B4:K4"].Style.Font.Bold = true;

				ESheet.Cells["B1:K1"].Style.Font.Size = 22;
				ESheet.Cells["B1:K1"].Style.Font.Bold = true;
				ESheet.Cells["B1:K1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				ESheet.Cells["B2:K2"].Style.Font.Size = 16;
				ESheet.Cells["B2:K2"].Style.Font.Bold = true;
				ESheet.Cells["B2:K2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B2"].Value = ReporteFilename;

				ESheet.Cells["B4:K4"].Style.Font.Size = 16;
				ESheet.Cells["B4:K4"].Style.Font.Bold = false;
				ESheet.Cells["B4:K4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B4"].Value = "Periodo " + FechaInicio.ToString("yyyy-MM-dd") + " - " + FechaFin.ToString("yyyy-MM-dd");*/



				System.Drawing.Color _BORDE = System.Drawing.Color.FromArgb(1, 191, 191, 191);
				System.Drawing.Color _MARCO = System.Drawing.Color.FromArgb(1, 0, 0, 0); //1,191,191,191

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				ESheet.Cells["B4"].Value += ESheet.Cells["B4"].Value + " | Periodo de Vencimiento: Del " + FechaVencimientoInicio.ToString("dd-MM-yyyy") + " Al " + FechaVencimientoFin.ToString("dd-MM-yyyy");

				ESheet.Cells["B9"].Value = "Moneda";
				ESheet.Cells["C9"].Value = "Fecha Vencimiento";
				ESheet.Cells["D9"].Value = "Empresa";
				ESheet.Cells["E9"].Value = "Número de Carta";
				ESheet.Cells["F9"].Value = "Estatus Carta";
				ESheet.Cells["G9"].Value = "Proveedor";
				ESheet.Cells["H9"].Value = "Banco";
				ESheet.Cells["I9"].Value = "Monto Pago";
				ESheet.Cells["J9"].Value = "Comisión Banco Corresponsal (USD)";
				ESheet.Cells["K9"].Value = "Comisión de Aceptación (USD)";

				ESheet.Cells["B9:K9"].Style.Font.Bold = true;


				ESheet.Cells["B9:K9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:K9"].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells["B9:K9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:K9"].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:K9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:K9"].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:K9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:K9"].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:K9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:K9"].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

				ESheet.Cells["B1:K1"].Merge = true;
				ESheet.Cells["B2:K2"].Merge = true;
				ESheet.Cells["B3:K3"].Merge = true;
				ESheet.Cells["B4:K4"].Merge = true;

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(),".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20,600);

				int row = 10;

				/*var tipoMoneda="uno";
				foreach (var cc in cartasCredito)
				{
					var rowOrigin = row;

					var comisionBancoCorresponsal = EncontrarComisionEnCarta("COM. BANCO CORRESPONSAL", cc.Id);
					var comisionAceptacion = EncontrarComisionEnCarta("COMISION DE ACEPTACION", cc.Id);

					if(tipoMoneda!=cc.Moneda){
						ESheet.Cells[string.Format("B{0}", row)].Value = cc.Moneda;
						tipoMoneda=cc.Moneda;
						var totalMonedaPagado=0M;
					}
					ESheet.Cells[string.Format("C{0}", row)].Value = cc.FechaVencimiento.ToString("dd-MM-yyyy");
					ESheet.Cells[string.Format("D{0}", row)].Value = cc.Empresa;
					ESheet.Cells[string.Format("E{0}", row)].Value = cc.NumCartaCredito;
					ESheet.Cells[string.Format("F{0}", row)].Value = CartaCredito.GetStatusText(cc.Estatus);
					ESheet.Cells[string.Format("G{0}", row)].Value = cc.Proveedor;
					ESheet.Cells[string.Format("H{0}", row)].Value = cc.Banco;
					ESheet.Cells[string.Format("I{0}", row)].Value = cc.MontoOriginalLC;
					ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

					if (comisionBancoCorresponsal != null && comisionBancoCorresponsal.Id > 0)
					{
						ESheet.Cells[string.Format("J{0}", row)].Value = comisionBancoCorresponsal.MontoPagado;
						ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					}

					if (comisionAceptacion != null && comisionAceptacion.Id > 0)
					{
						ESheet.Cells[string.Format("K{0}", row)].Value = comisionAceptacion.MontoPagado;
						ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					}
					
					row++;
				}*/

				var grupos = cartasCredito
								.GroupBy(carta => carta.Moneda)
								.Select(grupoMoneda => new
								{
									grupoMoneda.Key,
									MonedaId = grupoMoneda.First().MonedaId,
									TotalMoneda = grupoMoneda.Sum(c => c.MontoOriginalLC),
									CartasDeCredito = grupoMoneda.ToList()
								}).ToList();

				decimal grantotal = 0M;
				decimal grantotalUSD = 0M;
				foreach (var grupoMoneda in grupos)
				{
					var rowOrigin = row;

					decimal _comisionBancoCorresponsalMoneda = 0;
					decimal _comisionAceptacionMoneda = 0;

					ESheet.Cells[string.Format("B{0}", row)].Value = grupoMoneda.Key;
					ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

					foreach (var carta in grupoMoneda.CartasDeCredito)
					{
						//var comisionBancoCorresponsal = EncontrarComisionEnCarta("COM. BANCO CORRESPONSAL", carta.Id);
						//var comisionAceptacion = EncontrarComisionEnCarta("COMISION DE ACEPTACION", carta.Id);

						List<CartaCreditoComision> comisiones = CartaCreditoComision.GetByCartaCreditoId(carta.Id);
						decimal _comisionBancoCorresponsal = 0;
						decimal _comisionAceptacion = 0;
						decimal _MontoOriginalUSD = 0M;
						//----conversion de moneda-----
						TipoDeCambio tipoDeCambio = new TipoDeCambio();
						foreach (CartaCreditoComision comision in comisiones)
                        {
							if (comision.Comision == "COM. BANCO CORRESPONSAL" ||
								comision.Comision == "COMISION DE ACEPTACION") {
								tipoDeCambio = new TipoDeCambio();
								tipoDeCambio.Fecha = fechaDivisa;
								tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == comision.MonedaId).FirstOrDefault().Abbr.ToUpper();
								tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
								tipoDeCambio.Get();
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

								decimal _monto = 0;
								_monto = comision.MontoPagado * tipoDeCambio.Conversion;
								if (comision.Comision == "COM. BANCO CORRESPONSAL")
								{
									_comisionBancoCorresponsal += _monto;
								}
								if (comision.Comision == "COMISION DE ACEPTACION")
								{
									_comisionAceptacion += _monto;
								}
							}
						}

						tipoDeCambio = new TipoDeCambio();
						tipoDeCambio.Fecha = fechaDivisa;
						tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == grupoMoneda.MonedaId).FirstOrDefault().Abbr.ToUpper();
						tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
						tipoDeCambio.Get();
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
							_MontoOriginalUSD = carta.MontoOriginalLC * tipoDeCambio.Conversion;
						}
						//-----------------------------

						/*_comisionBancoCorresponsal = comisiones.Where(c => c.Comision == "COM. BANCO CORRESPONSAL").Sum(c => c.MontoPagado);
						comisionAceptacion = comisiones.Where(c => c.Comision == "COMISION DE ACEPTACION").Sum(c => c.MontoPagado);*/

						ESheet.Cells[string.Format("C{0}", row)].Value = carta.FechaVencimiento.ToString("dd-MM-yyyy");
						ESheet.Cells[string.Format("D{0}", row)].Value = carta.Empresa;
						ESheet.Cells[string.Format("E{0}", row)].Value = carta.NumCartaCredito;
						ESheet.Cells[string.Format("F{0}", row)].Value = CartaCredito.GetStatusText(carta.Estatus);
						ESheet.Cells[string.Format("G{0}", row)].Value = carta.Proveedor;
						ESheet.Cells[string.Format("H{0}", row)].Value = carta.Banco;
						ESheet.Cells[string.Format("I{0}", row)].Value = carta.MontoOriginalLC;
						ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
						//if (comisionBancoCorresponsal != null && comisionBancoCorresponsal.Id > 0)
						if (_comisionBancoCorresponsal > 0)
						{
							ESheet.Cells[string.Format("J{0}", row)].Value = _comisionBancoCorresponsal;
							ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
						}
						else{
							ESheet.Cells[string.Format("J{0}", row)].Value = "-";
						}

						//if (comisionAceptacion != null && comisionAceptacion.Id > 0)
						if (_comisionAceptacion > 0)
						{
							ESheet.Cells[string.Format("K{0}", row)].Value = _comisionAceptacion;
							ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";
						}
						else{
							ESheet.Cells[string.Format("K{0}", row)].Value = "-";
						}


						ESheet.Cells[string.Format("L{0}", row)].Value = _MontoOriginalUSD;
						ESheet.Cells[string.Format("L{0}", row)].Style.Numberformat.Format = " #,##0.00";

						_comisionBancoCorresponsalMoneda += _comisionBancoCorresponsal;
						_comisionAceptacionMoneda += _comisionAceptacion;
						grantotalUSD += _MontoOriginalUSD;

						ESheet.Cells["B" + row + ":K" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":K" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row + ":K" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":K" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row + ":K" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":K" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row + ":K" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":K" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
						row++;
					}
					ESheet.Cells[string.Format("H{0}", row)].Value = "Total";
					ESheet.Cells[string.Format("I{0}", row)].Value = grupoMoneda.TotalMoneda;
					ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[string.Format("J{0}", row)].Value = _comisionBancoCorresponsalMoneda;
					ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[string.Format("K{0}", row)].Value = _comisionAceptacionMoneda;
					ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";

					grantotal += grupoMoneda.TotalMoneda;

					ESheet.Cells["B" + row + ":K" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["B" + row + ":K" + row].Style.Fill.BackgroundColor.SetColor(1, 217, 225, 242);

					ESheet.Cells["B" + row + ":K" + row].Style.Font.Bold = true;
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":K" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);


					row++;

				}

				ESheet.Cells[string.Format("H{0}", row)].Value = "Gran Total";
				ESheet.Cells[string.Format("I{0}", row)].Value = grantotal;
				ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";

				ESheet.Cells[string.Format("L{0}", row)].Value = grantotalUSD;
				ESheet.Cells[string.Format("L{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["B" + row + ":K" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row + ":K" + row].Style.Fill.BackgroundColor.SetColor(1, 217, 225, 242);
				ESheet.Cells["B" + row + ":K" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":K" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

				row++;

				ESheet.Cells["B9" + ":K9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":K9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["K9" + ":K" + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["K9" + ":K" + (row - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 1) + ":K" + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 1) + ":K" + (row - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);

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

		private CartaCreditoComision EncontrarComisionEnCarta(string tipoComisionNombre, string cartaCreditoId)
		{
			var rsp = new CartaCreditoComision();

			try
			{
				var tiposCartaComision = TipoComision.Get(1);
				var tipoCartaComision = tiposCartaComision.Find(tc => tc.Nombre.Trim().ToLower() == tipoComisionNombre.Trim().ToLower());
				var cartaComisiones = CartaCreditoComision.GetByCartaCreditoId(cartaCreditoId);
				if (cartaComisiones.Count > 0)
				{
					rsp = cartaComisiones.Find(cartacom => cartacom.ComisionId == tipoCartaComision.Id);
				}

			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

	}
}