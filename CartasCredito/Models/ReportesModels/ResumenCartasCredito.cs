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
using System.Globalization;

namespace CartasCredito.Models.ReportesModels
{
	public class ResumenCartasCredito : ReporteBase
	{
		public DateTime FechaVencimientoInicio { get; set; } = DateTime.Parse("1969-01-01");
		public DateTime FechaVencimientoFin { get; set; } = DateTime.Parse("1969-01-01");
		public ResumenCartasCredito(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa)
			: base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Resumen de carta de creditos")
		{
		}

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();
			System.Drawing.Color _BORDE = System.Drawing.Color.FromArgb(1, 191, 191, 191);
			System.Drawing.Color _MARCO = System.Drawing.Color.FromArgb(1, 0, 0, 0); //1,191,191,191

			try
			{
				DateTime fechaBase = DateTime.Parse("2023-10-31");
				DateTime comisionesFin = new DateTime(fechaBase.Year, fechaBase.Month, DateTime.DaysInMonth(fechaBase.Year, fechaBase.Month));
				fechaBase = fechaBase.AddMonths(-1);
				DateTime comisionesInicio = new DateTime(fechaBase.Year, fechaBase.Month, 1);
				fechaBase = fechaBase.AddMonths(2);
				DateTime vencimientoMes1 = new DateTime(fechaBase.Year, fechaBase.Month, 1);
				fechaBase = fechaBase.AddMonths(1);
				DateTime vencimientoMes2 = new DateTime(fechaBase.Year, fechaBase.Month, 1);
				fechaBase = fechaBase.AddMonths(1);
				DateTime vencimientoMes3 = new DateTime(fechaBase.Year, fechaBase.Month, 1);




				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);
				//var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento).ToList();
				List<CTMComisiones> comisiones = new List<CTMComisiones>();
				List<CTMComisiones> vencimientos = new List<CTMComisiones>();
				var cartasCredito = CartaCreditoReporte.ReporteResumen(EmpresaId, fechaInicioExact, fechaFinExact, out comisiones, out vencimientos).ToList();
				
				var catMonedas = Moneda.Get();
				var catBancos = Banco.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();
				Moneda mndMxp = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "mxp").FirstOrDefault();

				var lista_empresas = Empresa.Get();

				if (EmpresaId > 0)
				{
					lista_empresas = lista_empresas.FindAll(e => e.Id == EmpresaId);
				}

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(), ".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20, 500);

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				TipoDeCambio tipoDeCambio = new TipoDeCambio();
				int row = 9;

				int columnaInicio = 0;
				int columnaFin = 0;
				int filaInicio = 0;
				int filaFin = 0;
				string columnaStr = "B";

				#region primera seccion
				columnaInicio = 2;
				columnaFin = 12;
				filaInicio = row;

				int comisionesMesInicio = comisionesInicio.Month;
				int comisionesMesFin = comisionesFin.Month;

				ESheet.Cells["B" + row].Value = "";
				ESheet.Cells["C" + row].Value = "Comisiones MXP";
				ESheet.Cells["C" + row + ":D" + row].Merge = true;
				ESheet.Cells["C" + row + ":D" + row].Style.Font.Bold = true;
				ESheet.Cells["E" + row].Value = "Vencimientos";
				ESheet.Cells["E" + row + ":J" + row].Merge = true;
				ESheet.Cells["E" + row + ":J" + row].Style.Font.Bold = true;
				ESheet.Cells["K" + row].Value = "Refinanciamientos USD";
				ESheet.Cells["K" + row + ":L" + row].Merge = true;
				ESheet.Cells["K" + row + ":L" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":L" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				row++;

				ESheet.Cells["B" + row].Value = "";
				ESheet.Cells["C" + row].Value = comisionesInicio.ToString("MMMM", new CultureInfo("es-MX")).Replace(".", "").ToUpper();
				ESheet.Cells["C" + row].Style.Font.Bold = true;
				ESheet.Cells["D" + row].Value = comisionesFin.ToString("MMMM", new CultureInfo("es-MX")).Replace(".", "").ToUpper();
				ESheet.Cells["D" + row].Style.Font.Bold = true;
				ESheet.Cells["E" + row].Value = vencimientoMes1.ToString("MMMM", new CultureInfo("es-MX")).Replace(".", "").ToUpper();
				ESheet.Cells["E" + row + ":F" + row].Merge = true;
				ESheet.Cells["E" + row + ":F" + row].Style.Font.Bold = true;
				ESheet.Cells["G" + row].Value = vencimientoMes2.ToString("MMMM", new CultureInfo("es-MX")).Replace(".", "").ToUpper();
				ESheet.Cells["G" + row + ":H" + row].Merge = true;
				ESheet.Cells["G" + row + ":H" + row].Style.Font.Bold = true;
				ESheet.Cells["I" + row].Value = vencimientoMes3.ToString("MMMM", new CultureInfo("es-MX")).Replace(".", "").ToUpper();
				ESheet.Cells["I" + row + ":J" + row].Merge = true;
				ESheet.Cells["I" + row + ":J" + row].Style.Font.Bold = true;
				ESheet.Cells["K" + row].Value = "Nuevos";
				ESheet.Cells["K" + row].Style.Font.Bold = true;
				ESheet.Cells["L" + row].Value = "Pagados";
				ESheet.Cells["L" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":L" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				row++;

				ESheet.Cells["B" + row].Value = "";
				ESheet.Cells["C" + row].Value = "Comisiones MXP";
				ESheet.Cells["C" + row].Style.Font.Bold = true;
				ESheet.Cells["D" + row].Value = "Comisiones MXP";
				ESheet.Cells["D" + row].Style.Font.Bold = true;
				ESheet.Cells["E" + row].Value = "A Pagar";
				ESheet.Cells["E" + row].Style.Font.Bold = true;
				ESheet.Cells["F" + row].Value = "A Refinanciar";
				ESheet.Cells["F" + row].Style.Font.Bold = true;
				ESheet.Cells["G" + row].Value = "A Pagar";
				ESheet.Cells["G" + row].Style.Font.Bold = true;
				ESheet.Cells["H" + row].Value = "A Refinanciar";
				ESheet.Cells["H" + row].Style.Font.Bold = true;
				ESheet.Cells["I" + row].Value = "A Pagar";
				ESheet.Cells["I" + row].Style.Font.Bold = true;
				ESheet.Cells["J" + row].Value = "A Refinanciar";
				ESheet.Cells["J" + row].Style.Font.Bold = true;
				ESheet.Cells["K" + row].Value = "---";
				ESheet.Cells["L" + row].Value = "---";
				ESheet.Cells["B" + row + ":L" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				row++;

				List<int> tbl1_empresas = comisiones.Select(i => i.empresa_id).Distinct().ToList();
				decimal totalMesAnterior = 0M;
				decimal totalMesActual = 0M;
				decimal totalVencimiento_apartir_1 = 0M;
				decimal totalVencimiento_apartir_2 = 0M;
				decimal totalVencimiento_apartir_3 = 0M;
				decimal totalVencimiento_arefinanciar_1 = 0M;
				decimal totalVencimiento_arefinanciar_2 = 0M;
				decimal totalVencimiento_arefinanciar_3 = 0M;

				foreach (int empresa in tbl1_empresas)
                {
					Empresa mainEmpresa = new Empresa();
					if(lista_empresas.Where(i => i.Id == empresa).Count() > 0)
                    {
						mainEmpresa = lista_empresas.Where(i => i.Id == empresa).FirstOrDefault();

						ESheet.Cells["B" + row].Value = mainEmpresa.Nombre.ToUpper();

						decimal montoMesAnterior = comisiones.Where(i => i.mes == comisionesMesInicio && i.empresa_id == empresa).Sum(j => j.monto_convertido);
						decimal montoMesActual = comisiones.Where(i => i.mes == comisionesMesFin && i.empresa_id == empresa).Sum(j => j.monto_convertido);

						var cartasMes1 = cartasCredito.Where(i => i.FechaVencimiento.Month == vencimientoMes1.Month && i.FechaVencimiento.Year == vencimientoMes1.Year && i.EmpresaId == empresa).ToList();
						var cartasMes2 = cartasCredito.Where(i => i.FechaVencimiento.Month == vencimientoMes2.Month && i.FechaVencimiento.Year == vencimientoMes2.Year && i.EmpresaId == empresa).ToList();
						var cartasMes3 = cartasCredito.Where(i => i.FechaVencimiento.Month == vencimientoMes3.Month && i.FechaVencimiento.Year == vencimientoMes3.Year && i.EmpresaId == empresa).ToList();


						decimal vencimiento_apartir_1 = vencimientos.Where(i => i.mes == vencimientoMes1.Month && i.empresa_id == empresa).Sum(j => j.monto_convertido); //cartasMes1.Sum(i => i.APagar);
						decimal vencimiento_apartir_2 = vencimientos.Where(i => i.mes == vencimientoMes2.Month && i.empresa_id == empresa).Sum(j => j.monto_convertido); //;
						decimal vencimiento_apartir_3 = vencimientos.Where(i => i.mes == vencimientoMes3.Month && i.empresa_id == empresa).Sum(j => j.monto_convertido); //;
						decimal vencimiento_arefinanciar_1 = 0;//cartasMes1.Sum(i => i.ARefinanciar);
						decimal vencimiento_arefinanciar_2 = 0;//cartasMes2.Sum(i => i.ARefinanciar);
						decimal vencimiento_arefinanciar_3 = 0;//cartasMes3.Sum(i => i.ARefinanciar);

						ESheet.Cells["C" + row].Value = montoMesAnterior;
						ESheet.Cells["D" + row].Value = montoMesActual;

						ESheet.Cells["E" + row].Value = vencimiento_apartir_1;
						ESheet.Cells["F" + row].Value = vencimiento_arefinanciar_1;
						ESheet.Cells["G" + row].Value = vencimiento_apartir_2;
						ESheet.Cells["H" + row].Value = vencimiento_arefinanciar_2;
						ESheet.Cells["I" + row].Value = vencimiento_apartir_3;
						ESheet.Cells["J" + row].Value = vencimiento_arefinanciar_3;

						ESheet.Cells["C" + row + ":J" + row].Style.Numberformat.Format = " #,##0.00";

						totalMesAnterior += montoMesAnterior;
						totalMesActual += montoMesActual;


						totalVencimiento_apartir_1 += vencimiento_apartir_1;
						totalVencimiento_apartir_2 += vencimiento_apartir_2;
						totalVencimiento_apartir_3 += vencimiento_apartir_3;
						totalVencimiento_arefinanciar_1 += vencimiento_arefinanciar_1;
						totalVencimiento_arefinanciar_2 += vencimiento_arefinanciar_2;
						totalVencimiento_arefinanciar_3 += vencimiento_arefinanciar_3;
					}
                    else
                    {
						//no dibujar empresa
                    }
					row++;
                }


				ESheet.Cells["B" + row].Value = "TOTAL";
				ESheet.Cells["C" + row].Value = totalMesAnterior;
				ESheet.Cells["D" + row].Value = totalMesActual;

				ESheet.Cells["E" + row].Value = totalVencimiento_apartir_1;
				ESheet.Cells["F" + row].Value = totalVencimiento_arefinanciar_1;
				ESheet.Cells["G" + row].Value = totalVencimiento_apartir_2;
				ESheet.Cells["H" + row].Value = totalVencimiento_arefinanciar_2;
				ESheet.Cells["I" + row].Value = totalVencimiento_apartir_3;
				ESheet.Cells["J" + row].Value = totalVencimiento_arefinanciar_3;

				ESheet.Cells["C" + row + ":J" + row].Style.Numberformat.Format = " #,##0.00";

				ESheet.Cells["B" + row + ":L" + row].Style.Font.Bold = true;

				filaFin = row;

				columnaStr = Utility.ExcelColumnFromNumber(columnaFin);

				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + filaInicio + ":" + columnaStr + filaFin].Style.Border.Bottom.Color.SetColor(_BORDE);


				row++;
				row++;
				#endregion

				#region segunda seccion
				decimal granTotalMonto = 0M;
				decimal granTotalPagosEfectuados = 0M;


				decimal granTotalMontoLC = 0M;
				decimal granTotalMontoPF = 0M;
				decimal granTotalPagosEfectuadosLC = 0M;
				decimal granTotalPagosEfectuadosPF = 0M;

				decimal totalMontoLC = 0M;
				decimal totalMontoPF = 0M;
				decimal totalPagosEfectuadosLC = 0M;
				decimal totalPagosEfectuadosPF = 0M;
				List<int> tbl2_bancos = cartasCredito.Select(i => i.BancoId).Distinct().ToList();
				foreach (int banco in tbl2_bancos)
				{
					totalMontoLC = 0M;
					totalMontoPF = 0M;
					totalPagosEfectuadosLC = 0M;
					totalPagosEfectuadosPF = 0M;

					Banco mainBanco = new Banco();
					mainBanco = catBancos.Where(i => i.Id == banco).FirstOrDefault();
					ESheet.Cells["B" + row].Value = "Cartas de crédito de " + mainBanco.Nombre.ToUpper();
					ESheet.Cells["B" + row + ":I" + row].Merge = true;
					ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
					ESheet.Cells["B" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["B" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);
					ESheet.Cells["B" + row + ":I" + row].Style.Font.Color.SetColor(1, 0, 0, 255);
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
					row++;


					List<CartaCreditoReporte> listaLC = cartasCredito.Where(i => i.TipoEmision == "Líneas de crédito" && i.BancoId == banco).ToList();
					if(listaLC.Count > 0)
					{
						ESheet.Cells["B" + row].Value = "Líneas de crédito";
						ESheet.Cells["B" + row + ":I" + row].Merge = true;
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Color.SetColor(1, 0, 0, 255);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
						row++;

						ESheet.Cells["B" + row].Value = "";
						ESheet.Cells["C" + row].Value = "Monto USD";
						ESheet.Cells["D" + row].Value = "Referencia";
						ESheet.Cells["E" + row].Value = "Apertura";
						ESheet.Cells["F" + row].Value = "Vence";
						ESheet.Cells["G" + row].Value = "Tipo";
						ESheet.Cells["H" + row].Value = "Beneficiario";
						ESheet.Cells["I" + row].Value = "Banco";
						/*ESheet.Cells["j" + row].Value = "monto original";
						ESheet.Cells["k" + row].Value = "moneda";
						ESheet.Cells["l" + row].Value = "monto efectuados";
						ESheet.Cells["m" + row].Value = "pagos efectuados";
						ESheet.Cells["n" + row].Value = "pagos efectuados usd";*/
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
						row++;

						foreach (var item in listaLC)
						{
							ESheet.Cells["B" + row].Value = item.Empresa;
							ESheet.Cells["C" + row].Value = item.MontoOriginalLCUSD;
							ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
							ESheet.Cells["D" + row].Value = item.NumCartaCredito;
							ESheet.Cells["E" + row].Value = item.FechaApertura.ToString("dd-MM-yyyy");
							ESheet.Cells["F" + row].Value = item.FechaVencimiento.ToString("dd-MM-yyyy");
							ESheet.Cells["G" + row].Value = item.TipoCarta;
							ESheet.Cells["H" + row].Value = item.Proveedor;
							ESheet.Cells["I" + row].Value = item.Banco;

							//
							decimal montoEfectuado = 0M;
							if (item.PagosEfectuadosUSD <= 0)
							{
								tipoDeCambio.Fecha = fechaDivisa;
								tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
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
								tiposDeCambio.Add(tipoDeCambio);

								if (tipoDeCambio.MonedaOriginal == "JPY")
								{
									montoEfectuado = item.PagosEfectuados / tipoDeCambio.Conversion;
								}
								else
								{
									montoEfectuado = item.PagosEfectuados * tipoDeCambio.Conversion;
								}
								totalPagosEfectuadosLC += montoEfectuado;
							}
                            else
                            {
								totalPagosEfectuadosLC += item.PagosEfectuadosUSD;
							}
							/*ESheet.Cells["J" + row].Value = item.MontoOriginalLC;
							ESheet.Cells["K" + row].Value = item.Moneda;
							ESheet.Cells["L" + row].Value = montoEfectuado;
							ESheet.Cells["m" + row].Value = item.PagosEfectuados;
							ESheet.Cells["n" + row].Value = item.PagosEfectuadosUSD;*/

							//

							totalMontoLC += item.MontoOriginalLCUSD;
							row++;
						}
						ESheet.Cells["B" + row].Value = "TOTAL USD";
						ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
						ESheet.Cells["C" + row].Value = totalMontoLC;
						ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells["B" + row + ":C" + row].Style.Font.Bold = true;
						row++;
						ESheet.Cells["B" + row].Value = "PAGOS EFECTUADOS";
						ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
						ESheet.Cells["C" + row].Value = totalPagosEfectuadosLC;
						ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells["B" + row + ":C" + row].Style.Font.Bold = true;
						row++;
						ESheet.Cells["B" + row].Value = "TOTAL";
						ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
						ESheet.Cells["C" + row].Value = totalMontoLC - totalPagosEfectuadosLC;
						ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);
						row++;
					}

					List<CartaCreditoReporte> listaPF = cartasCredito.Where(i => i.TipoEmision == "Provisión de fondo" && i.BancoId == banco).ToList();
					if (listaPF.Count > 0)
					{
						ESheet.Cells["B" + row].Value = "Provisión de fondo";
						ESheet.Cells["B" + row + ":I" + row].Merge = true;
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Color.SetColor(1, 0, 0, 255);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
						row++;

						ESheet.Cells["B" + row].Value = "";
						ESheet.Cells["C" + row].Value = "Monto USD";
						ESheet.Cells["D" + row].Value = "Referencia";
						ESheet.Cells["E" + row].Value = "Apertura";
						ESheet.Cells["F" + row].Value = "Vence";
						ESheet.Cells["G" + row].Value = "Tipo";
						ESheet.Cells["H" + row].Value = "Beneficiario";
						ESheet.Cells["I" + row].Value = "Banco";
						ESheet.Cells["j" + row].Value = "monto original";
						ESheet.Cells["k" + row].Value = "moneda";
						ESheet.Cells["l" + row].Value = "monto efectuados";
						ESheet.Cells["m" + row].Value = "pagos efectuados";
						ESheet.Cells["n" + row].Value = "pagos efectuados usd";
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
						row++;

						foreach (var item in listaPF)
						{
							ESheet.Cells["B" + row].Value = item.Empresa;
							ESheet.Cells["C" + row].Value = item.MontoOriginalLCUSD;
							ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
							ESheet.Cells["D" + row].Value = item.NumCartaCredito;
							ESheet.Cells["E" + row].Value = item.FechaApertura.ToString("dd-MM-yyyy");
							ESheet.Cells["F" + row].Value = item.FechaVencimiento.ToString("dd-MM-yyyy");
							ESheet.Cells["G" + row].Value = item.TipoCarta;
							ESheet.Cells["H" + row].Value = item.Proveedor;
							ESheet.Cells["I" + row].Value = item.Banco;

							//
							//
							decimal montoEfectuado = 0M;
							if (item.PagosEfectuadosUSD <= 0)
							{
								tipoDeCambio.Fecha = fechaDivisa;
								tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
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
								tiposDeCambio.Add(tipoDeCambio);

								if (tipoDeCambio.MonedaOriginal == "JPY")
								{
									montoEfectuado = item.PagosEfectuados / tipoDeCambio.Conversion;
								}
								else
								{
									montoEfectuado = item.PagosEfectuados * tipoDeCambio.Conversion;
								}
								totalPagosEfectuadosPF += montoEfectuado;
							}
							else
							{
								totalPagosEfectuadosPF += item.PagosEfectuadosUSD;
							}
							ESheet.Cells["J" + row].Value = item.MontoOriginalLC;
							ESheet.Cells["K" + row].Value = item.Moneda;
							ESheet.Cells["L" + row].Value = montoEfectuado;
							ESheet.Cells["m" + row].Value = item.PagosEfectuados;
							ESheet.Cells["n" + row].Value = item.PagosEfectuadosUSD;
							//

							totalMontoPF += item.MontoOriginalLCUSD;
							row++;
						}
						ESheet.Cells["B" + row].Value = "TOTAL USD";
						ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
						ESheet.Cells["C" + row].Value = totalMontoPF;
						ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells["B" + row + ":C" + row].Style.Font.Bold = true;
						row++;
						ESheet.Cells["B" + row].Value = "PAGOS EFECTUADOS";
						ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
						ESheet.Cells["C" + row].Value = totalPagosEfectuadosPF;
						ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells["B" + row + ":C" + row].Style.Font.Bold = true;
						row++;
						ESheet.Cells["B" + row].Value = "TOTAL";
						ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
						ESheet.Cells["C" + row].Value = totalMontoPF - totalPagosEfectuadosPF;
						ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);
						row++;


						granTotalPagosEfectuados += totalPagosEfectuadosPF;
					}

					granTotalMontoLC += totalMontoLC;
					granTotalMontoPF += totalMontoPF;
					granTotalPagosEfectuadosLC += totalPagosEfectuadosLC;
					granTotalPagosEfectuadosPF += totalPagosEfectuadosPF;
					row++;
				}

				granTotalMonto = granTotalMontoLC + granTotalMontoPF;
				granTotalPagosEfectuados = granTotalPagosEfectuadosLC + granTotalPagosEfectuadosPF;

				ESheet.Cells["B" + row].Value = "GRAN TOTAL USD";
				ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["C" + row].Value = granTotalMonto;
				ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["B" + row + ":C" + row].Style.Font.Bold = true;
				row++;
				ESheet.Cells["B" + row].Value = "PAGOS EFECTUADOS";
				ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["C" + row].Value = granTotalPagosEfectuados;
				ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["B" + row + ":C" + row].Style.Font.Bold = true;
				row++;
				ESheet.Cells["B" + row].Value = "GRAN TOTAL";
				ESheet.Cells["B" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["C" + row].Value = granTotalMonto - granTotalPagosEfectuados;
				ESheet.Cells["C" + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 200, 200, 200);
				row++;
				row++;
				#endregion

				#region tercera seccion

				decimal resumenTotalLC = granTotalMontoLC - granTotalPagosEfectuadosLC;
				decimal resumenTotalDepositos = granTotalMontoPF - granTotalPagosEfectuadosPF;
				decimal resumenTotalMontoLC = granTotalMonto;
				decimal resumenTotalPF = granTotalPagosEfectuados;
				ESheet.Cells["B" + row].Value = "Resumen por emisión";
				ESheet.Cells["B" + row + ":C" + row].Merge = true;
				ESheet.Cells["D" + row].Value = "Sin quitarle pagos todas sin corte por emisión";
				ESheet.Cells["D" + row + ":E" + row].Merge = true;

				ESheet.Cells["B" + row + ":E" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":E" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row + ":E" + row].Style.Fill.BackgroundColor.SetColor(1, 255, 255, 102);

				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
				row++;

				ESheet.Cells["B" + row].Value = "Total LC con línea de crédito";
				ESheet.Cells["C" + row].Value = resumenTotalLC;
				ESheet.Cells["D" + row].Value = "Total LC con línea de crédito";
				ESheet.Cells["E" + row].Value = granTotalMonto;
				ESheet.Cells["B" + row + ":E" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":E" + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
				row++;

				ESheet.Cells["B" + row].Value = "Total depósitos en garantía";
				ESheet.Cells["C" + row].Value = resumenTotalDepositos;
				ESheet.Cells["D" + row].Value = "Menos pagos efectuados";
				ESheet.Cells["E" + row].Value = granTotalPagosEfectuados;
				ESheet.Cells["B" + row + ":E" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":E" + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
				row++;

				ESheet.Cells["B" + row].Value = "Total L/C";
				ESheet.Cells["C" + row].Value = resumenTotalLC + resumenTotalDepositos;
				ESheet.Cells["D" + row].Value = "Total L/C";
				ESheet.Cells["E" + row].Value = granTotalMonto - granTotalPagosEfectuados;
				ESheet.Cells["B" + row + ":E" + row].Style.Font.Bold = true;
				ESheet.Cells["B" + row + ":E" + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":E" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
				row++;


				#endregion







				ESheet.Cells["A:AZ"].AutoFitColumns();
				//ESheet.Column(3).Width = 30;

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