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
	public class VencimientosPagos : ReporteBase
	{
		public DateTime FechaVencimientoInicio { get; set; } = DateTime.Parse("1969-01-01");
		public DateTime FechaVencimientoFin { get; set; } = DateTime.Parse("1969-01-01");
		public VencimientosPagos(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa, "A", "K", "Reporte de Vencimientos de Cartas de Crédito")
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

				List<CartaCreditoReporte> data = CartaCreditoReporte.ReporteVencimientosPagos(this.EmpresaId, FechaInicio, FechaFin, this.FechaVencimientoInicio, this.FechaVencimientoFin).OrderBy(cc => cc.Moneda).ThenBy(cc => cc.FechaVencimiento).ToList();
				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();


				var monedas = Moneda.Get(1);



				System.Drawing.Color _BORDE = System.Drawing.Color.FromArgb(1, 191, 191, 191);
				System.Drawing.Color _MARCO = System.Drawing.Color.FromArgb(1, 0, 0, 0); //1,191,191,191

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				if(FechaVencimientoFin.Year == 2099)
                {
					ESheet.Cells["B4"].Value = ESheet.Cells["B4"].Value + " | Periodo de Vencimiento: A partir de " + FechaVencimientoInicio.ToString("dd-MM-yyyy");
				}
                else
                {
					ESheet.Cells["B4"].Value = ESheet.Cells["B4"].Value + " | Periodo de Vencimiento: Del " + FechaVencimientoInicio.ToString("dd-MM-yyyy") + " Al " + FechaVencimientoFin.ToString("dd-MM-yyyy");
				}

				

				ESheet.Cells["B9"].Value = "Moneda";
				ESheet.Cells["C9"].Value = "Fecha Vencimiento";
				ESheet.Cells["D9"].Value = "Empresa";
				ESheet.Cells["E9"].Value = "Número de Carta";
				ESheet.Cells["F9"].Value = "Estatus Carta";
				ESheet.Cells["G9"].Value = "Proveedor";
				ESheet.Cells["H9"].Value = "Banco";
				ESheet.Cells["I9"].Value = "Monto Pago";
				ESheet.Cells["J9"].Value = "Monto Pago (USD)";
				ESheet.Cells["K9"].Value = "Comisión Banco Corresponsal (USD)";
				ESheet.Cells["L9"].Value = "Comisión de Aceptación (USD)";
				/*ESheet.Cells["J9"].Value = "Comisión Banco Corresponsal (USD)";
				ESheet.Cells["K9"].Value = "Comisión de Aceptación (USD)";*/

				ESheet.Cells["B9:L9"].Style.Font.Bold = true;


				ESheet.Cells["B9:L9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:L9"].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells["B9:L9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:L9"].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:L9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:L9"].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:L9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:L9"].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:L9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:L9"].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

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


				decimal _comisionBancoCorresponsalTotal = 0M;
				decimal _comisionAceptacionTotal = 0M;
				decimal _montoTotal = 0M;
				decimal _montoTotalUSD = 0M;

				List<int> gpoMonedas = new List<int>();
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);
				TipoDeCambio tipoDeCambio = new TipoDeCambio();

				gpoMonedas = data.Select(i => i.MonedaId).Distinct().ToList();
				gpoMonedas.Sort();
				foreach(int gpoMoneda in gpoMonedas) //por cada tipo de monada
                {
					Moneda moneda = monedas.Where(m => m.Id == gpoMoneda).FirstOrDefault();
					if(moneda != null)
					{
						decimal _comisionBancoCorresponsalMoneda = 0M;
						decimal _comisionAceptacionMoneda = 0M;
						decimal _montoMoneda = 0M;
						decimal _montoMonedaUSD = 0M;

						ESheet.Cells[string.Format("B{0}", row)].Value = moneda.Nombre;
						ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
						ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

						List<CartaCreditoReporte> pagosMoneda = new List<CartaCreditoReporte>(); //pagos de esas monedas
						pagosMoneda = data.Where(d => d.MonedaId == gpoMoneda).ToList();

						List<int> gpoAnhos = new List<int>();
						gpoAnhos = pagosMoneda.Select(i => i.AnhoVencimiento).Distinct().ToList();
						gpoAnhos.Sort();

						foreach(int anho in gpoAnhos) //por cada anho de esas cartas de esa moneda
                        {
							List<CartaCreditoReporte> pagosAnho = new List<CartaCreditoReporte>(); //pagos de ese anho con esa moneda
							pagosAnho = pagosMoneda.Where(d => d.AnhoVencimiento == anho).OrderBy(d=>d.FechaVencimiento).ToList();

							List<int> gpoMes = new List<int>();
							gpoMes = pagosAnho.Select(i => i.MesVencimiento).Distinct().ToList();
							gpoMes.Sort();

							decimal _comisionBancoCorresponsalAnho = 0M;
							decimal _comisionAceptacionAnho = 0M;
							decimal _montoAnho = 0M;
							decimal _montoAnhoUSD = 0M;

							foreach (int mes in gpoMes) //por cada anho de esas cartas de esa moneda en ese mes
							{
								List<CartaCreditoReporte> pagosMes = new List<CartaCreditoReporte>(); //pagos de ese anho con esa moneda en ese mes
								pagosMes = pagosAnho.Where(d => d.AnhoVencimiento == anho && d.MesVencimiento == mes).OrderBy(d => d.FechaVencimiento).ToList();


								decimal _comisionBancoCorresponsalMes = 0M;
								decimal _comisionAceptacionMes = 0M;
								decimal _montoMes = 0M;
								decimal _montoMesUSD = 0M;

								foreach (CartaCreditoReporte pago in pagosMes) //por cada pago
								{
									List<CartaCreditoComision> comisiones = CartaCreditoComision.GetByCartaCreditoFromCarta(pago.Id);
									if(pago.NumCartaCredito == "DC MXH231824")
                                    {
										var op = 0;
                                    }
									decimal _comisionBancoCorresponsal = 0;
									decimal _comisionAceptacion = 0;
									decimal _comisionBancoCorresponsalUSD = 0;
									decimal _comisionAceptacionUSD = 0;
									decimal _MontoOriginal = pago.MontoOriginalLC;
									decimal _MontoOriginalUSD = pago.MontoOriginalLC;

									//----conversion monto original
									/*TipoDeCambio tipoDeCambio = new TipoDeCambio();
									tipoDeCambio.Fecha = fechaDivisa;
									tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == pago.MonedaId).FirstOrDefault().Abbr.ToUpper();
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
									string MonedaOriginal = catMonedas.Where(m => m.Id == pago.MonedaId).FirstOrDefault().Abbr.ToUpper();
									string MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
									tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == MonedaOriginal && i.MonedaNueva == MonedaNueva).FirstOrDefault();
									if (tipoDeCambio.MonedaOriginal == "JPY")
									{
										_MontoOriginalUSD = _MontoOriginal / tipoDeCambio.Conversion;
									}
									else
									{
										_MontoOriginalUSD = _MontoOriginal * tipoDeCambio.Conversion;
									}

									//----conversion monto original


									//----conversion de moneda-----
									foreach (CartaCreditoComision comision in comisiones)
									{
										if (comision.Comision == "COM. BANCO CORRESPONSAL" ||
											comision.Comision == "COMISION DE ACEPTACION")
										{
											decimal _monto = 0;
											_monto = comision.Monto;
											decimal _montoUSD = comision.Monto;

											/*tipoDeCambio = new TipoDeCambio();
											tipoDeCambio.Fecha = fechaDivisa;
											tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == comision.MonedaId).FirstOrDefault().Abbr.ToUpper();
											tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();

											busqueda = tiposDeCambio.Where(t => t.MonedaOriginal == tipoDeCambio.MonedaOriginal &&
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

											MonedaOriginal = catMonedas.Where(m => m.Id == comision.MonedaId).FirstOrDefault().Abbr.ToUpper();
											MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
											tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == MonedaOriginal && i.MonedaNueva == MonedaNueva).FirstOrDefault();

											if (tipoDeCambio.MonedaOriginal == "JPY")
											{
												_montoUSD = comision.Monto / tipoDeCambio.Conversion;
											}
											else
											{
												_montoUSD = comision.Monto * tipoDeCambio.Conversion;
											}
											/*
											1	COMISION POR APERTURA
											2	COMISION POR DISPOSICION/NEGOC
											3	COM. BANCO CORRESPONSAL
											4	COM. POR PAGO DIFERIDO
											5	COM. POR MODIFICACION
											6	COMISION POR CANCELACION
											7	COMISION DE ACEPTACION
											8	COM. BANCO CORRESPONSAL POR ENMIENDA
											9	COM. DE CONFIRMACION
											10	COM. BANCO CORRESPONSAL POR APERTURA
											11	COM. BANCO CORRESPONSAL POR MODIFICACION
											12	COM. BANCO CORRESPONSAL CONFIRMACION
											 */
											if (comision.Comision == "COM. BANCO CORRESPONSAL" || comision.Comision == "COM. BANCO CORRESPONSAL POR ENMIENDA" || comision.Comision == "COM. BANCO CORRESPONSAL POR APERTURA" ||
												comision.Comision == "COM. BANCO CORRESPONSAL POR MODIFICACION" || comision.Comision == "COM. BANCO CORRESPONSAL CONFIRMACION")
                                            {
												_comisionBancoCorresponsal += _montoUSD;
											}
											if (comision.Comision == "COMISION DE ACEPTACION")
											{
												_comisionAceptacion += _montoUSD;
											}
										}
									}
									//-----------------------------

									ESheet.Cells[string.Format("C{0}", row)].Value = pago.FechaVencimiento.ToString("dd-MM-yyyy");
									ESheet.Cells[string.Format("D{0}", row)].Value = pago.Empresa;
									ESheet.Cells[string.Format("E{0}", row)].Value = pago.NumCartaCredito;
									ESheet.Cells[string.Format("F{0}", row)].Value = CartaCredito.GetStatusText(pago.Estatus);
									ESheet.Cells[string.Format("G{0}", row)].Value = pago.Proveedor;
									ESheet.Cells[string.Format("H{0}", row)].Value = pago.Banco;
									ESheet.Cells[string.Format("I{0}", row)].Value = _MontoOriginal;
									ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
									ESheet.Cells[string.Format("J{0}", row)].Value = _MontoOriginalUSD;
									ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";


									if (_comisionBancoCorresponsal > 0)
									{
										ESheet.Cells[string.Format("K{0}", row)].Value = _comisionBancoCorresponsal;
										ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";
									}
									else
									{
										ESheet.Cells[string.Format("K{0}", row)].Value = "-";
									}

									if (_comisionAceptacion > 0)
									{
										ESheet.Cells[string.Format("L{0}", row)].Value = _comisionAceptacion;
										ESheet.Cells[string.Format("L{0}", row)].Style.Numberformat.Format = " #,##0.00";
									}
									else
									{
										ESheet.Cells[string.Format("L{0}", row)].Value = "-";
									}


									_comisionBancoCorresponsalMes += _comisionBancoCorresponsal;
									_comisionAceptacionMes += _comisionAceptacion;
									_montoMes += _MontoOriginal;
									_montoMesUSD += _MontoOriginalUSD;

									ESheet.Cells["B" + row + ":L" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["B" + row + ":L" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
									ESheet.Cells["B" + row + ":L" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["B" + row + ":L" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
									ESheet.Cells["B" + row + ":L" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["B" + row + ":L" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
									ESheet.Cells["B" + row + ":L" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["B" + row + ":L" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);
									ESheet.Cells["B" + row + ":L" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
									row++;
								}

								ESheet.Cells[string.Format("H{0}", row)].Value = "Total";
								ESheet.Cells[string.Format("I{0}", row)].Value = _montoMes;
								ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
								ESheet.Cells[string.Format("J{0}", row)].Value = _montoMesUSD;
								ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
								ESheet.Cells[string.Format("K{0}", row)].Value = _comisionBancoCorresponsalMes;
								ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";
								ESheet.Cells[string.Format("L{0}", row)].Value = _comisionAceptacionMes;
								ESheet.Cells[string.Format("L{0}", row)].Style.Numberformat.Format = " #,##0.00";

								_comisionBancoCorresponsalAnho += _comisionBancoCorresponsalMes;
								_comisionAceptacionAnho += _comisionAceptacionMes;
								_montoAnho += _montoMes;
								_montoAnhoUSD += _montoMesUSD;


								ESheet.Cells["H" + row + ":L" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
								ESheet.Cells["H" + row + ":L" + row].Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);


								ESheet.Cells["B" + row + ":L" + row].Style.Font.Bold = true;
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["B" + row + ":L" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);


								row++;
							}


							_comisionBancoCorresponsalMoneda += _comisionBancoCorresponsalAnho;
							_comisionAceptacionMoneda += _comisionAceptacionAnho;
							_montoMoneda += _montoAnho;
							_montoMonedaUSD += _montoAnhoUSD;
						}


						ESheet.Cells[string.Format("H{0}", row)].Value = "Total " + moneda.Nombre;
						ESheet.Cells[string.Format("I{0}", row)].Value = _montoMoneda;
						ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells[string.Format("J{0}", row)].Value = _montoMonedaUSD;
						ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells[string.Format("K{0}", row)].Value = _comisionBancoCorresponsalMoneda;
						ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells[string.Format("L{0}", row)].Value = _comisionAceptacionMoneda;
						ESheet.Cells[string.Format("L{0}", row)].Style.Numberformat.Format = " #,##0.00";

						ESheet.Cells["B" + row + ":L" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":L" + row].Style.Fill.BackgroundColor.SetColor(1, 217, 225, 242);
						ESheet.Cells["B" + row + ":L" + row].Style.Font.Bold = true;

						_comisionBancoCorresponsalTotal += _comisionBancoCorresponsalMoneda;
						_comisionAceptacionTotal += _comisionAceptacionMoneda;
						_montoTotal += _montoMoneda;
						_montoTotalUSD += _montoMonedaUSD;

						row++;
					}
                }

				ESheet.Cells[string.Format("H{0}", row)].Value = "Gran Total";
				ESheet.Cells[string.Format("I{0}", row)].Value = _montoTotal;
				ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[string.Format("J{0}", row)].Value = _montoTotalUSD;
				ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[string.Format("K{0}", row)].Value = _comisionBancoCorresponsalTotal;
				ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[string.Format("L{0}", row)].Value = _comisionAceptacionTotal;
				ESheet.Cells[string.Format("L{0}", row)].Style.Numberformat.Format = " #,##0.00";


				ESheet.Cells["B" + row + ":L" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row + ":L" + row].Style.Fill.BackgroundColor.SetColor(1, 217, 225, 242);
				ESheet.Cells["B" + row + ":L" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":L" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":L" + row].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B" + row + ":L" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":L" + row].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B" + row + ":L" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":L" + row].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B" + row + ":L"+ row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":L" + row].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

				row++;

				ESheet.Cells["B9" + ":L9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":L9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["L9" + ":L" + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["L9" + ":L" + (row - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 1) + ":L" + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 1) + ":L" + (row - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);

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