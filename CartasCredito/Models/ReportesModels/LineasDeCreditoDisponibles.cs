using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
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

		public Reporte GenerarV1()
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
				var cartasCredito = CartaCreditoReporte.ReporteLineasCredito(EmpresaId, fechaInicioExact, fechaFinExact).OrderBy(cc => cc.FechaVencimiento).ToList();
				var lineasCredito = LineaDeCredito.Get(1);


				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				var empresas = Empresa.Get();
				var bancos = Banco.Get();

				if (EmpresaId > 0)
				{
					empresas = empresas.FindAll(e => e.Id == EmpresaId);
				}

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(), ".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20, 500);

				/*ESheet.Cells["B9"].Value = "Comisión";
				ESheet.Cells["C9"].Value = "Estatus Carta";
				ESheet.Cells["D9"].Value = "Moneda";
				ESheet.Cells["E9"].Value = "Monto programado";
				ESheet.Cells["F9"].Value = "Monto pagado";
				ESheet.Cells["G9"].Value = "Monto pagado en USD";
				ESheet.Cells["H9"].Value = "";

				ESheet.Cells["B9:H9"].Style.Font.Bold = true;*/
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);
				TipoDeCambio tipoDeCambio = new TipoDeCambio();
				int columna = 3;
				int row = 10;
				int numFila = 0;
				string ultimaColumna = "B";
				List<string> lista_empresas = new List<string>();
				List<decimal> lista_empresas_banco_totales = new List<decimal>();
				List<string> lista_bancos = new List<string>();
				List<string> lista_bancos_cartas = new List<string>();

				//------------

				List<string> lista_empresas_pre = new List<string>();
				List<string> lista_bancos_pre = new List<string>();

				lista_empresas_pre = cartasCredito.Select(i => i.Empresa).Distinct().ToList();
				lista_bancos_pre = lineasCredito.Select(i => i.Banco).Distinct().ToList();



				//------------


				lista_empresas = cartasCredito.Select(i => i.Empresa).Distinct().ToList();
				lista_bancos = lineasCredito.Select(i => i.Banco).Distinct().ToList();
				lista_bancos_cartas = cartasCredito.Where(a => a.Banco != "").Select(i => i.Banco).Distinct().ToList();
				lista_empresas.Sort();
				lista_bancos.Sort();

				List<LineaCreditoGpo> totales = new List<LineaCreditoGpo>();


				#region tabla de lineas de credito

				string headerTotal = "B" + row;

				columna = 3;
				foreach (string banco in lista_bancos)
				{
					columna++;
				}
				string columnaStr = Utility.ExcelColumnFromNumber(columna);
				headerTotal = "B" + row + ":" + columnaStr + row;

				ESheet.Cells[headerTotal].Merge = true;
				ESheet.Cells[headerTotal].Value = "Líneas autorizadas";
				ESheet.Cells[headerTotal].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells[headerTotal].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells[headerTotal].Style.Fill.BackgroundColor.SetColor(1, 191, 191, 189);

				row++;

				columna = 3;
				ESheet.Cells[string.Format("B{0}", row)].Value = "EMPRESA";

				ESheet.Cells["B" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);


				columna = 3;
				foreach (string banco in lista_bancos)
				{
					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = banco.ToUpper();

					ESheet.Cells[columnaStr + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells[columnaStr + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					columna++;
				}

				columnaStr = Utility.ExcelColumnFromNumber(columna);
				ESheet.Cells[columnaStr + row].Value = "TOTAL";
				ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

				ESheet.Cells[columnaStr + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells[columnaStr + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				columna++;

				ultimaColumna = Utility.ExcelColumnFromNumber(columna - 1);

				row++;

				numFila = 0;
				decimal autorizada_fila_total_total = 0M;
				foreach (string empresa in lista_empresas)
				{

					ESheet.Cells[string.Format("B{0}", row)].Value = empresa;
					columna = 3;

					decimal autorizada_fila_total = 0M;
					foreach (string banco in lista_bancos)
					{
						decimal monto = 0M;
						List<LineaDeCredito> lineas = new List<LineaDeCredito>();
						lineas = lineasCredito.Where(j => j.Empresa == empresa &&
						j.Banco == banco).ToList();

						int bancoId = 0;
						int empresaId = 0;

						bancoId = bancos.Where(b => b.Nombre == banco).First().Id;
						empresaId = empresas.Where(b => b.Nombre == empresa).First().Id;

						foreach (var linea in lineas)
						{
							bancoId = linea.BancoId;
							empresaId = linea.EmpresaId;
							//----conversion de moneda-----
							/*TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
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
							*/
							tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == "USD" && i.MonedaNueva == "USD").FirstOrDefault();

							if (tipoDeCambio.MonedaOriginal == "JPY")
							{
								linea.MontoUSD = linea.Monto / tipoDeCambio.Conversion;
							}
							else
							{
								linea.MontoUSD = linea.Monto * tipoDeCambio.Conversion;
							}
							//-----------------------------
						}

						monto = lineas.Sum(k => k.MontoUSD);

						totales.Add(new LineaCreditoGpo()
						{
							Empresa = empresa,
							EmpresaId = empresaId,
							Banco = banco,
							BancoId = bancoId,
							TotalEmpresaBanco = monto
						});

						columnaStr = Utility.ExcelColumnFromNumber(columna);

						autorizada_fila_total += monto;

						ESheet.Cells[columnaStr + row].Value = monto;
						ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";

						ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

						columna++;
					}

					columnaStr = Utility.ExcelColumnFromNumber(columna);
					autorizada_fila_total_total += autorizada_fila_total;
					ESheet.Cells[columnaStr + row].Value = autorizada_fila_total;
					ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);


					ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);


					if (numFila % 2 > 0)
					{
						ESheet.Cells["B" + row + ":" + ultimaColumna + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":" + ultimaColumna + row].Style.Fill.BackgroundColor.SetColor(1, 226, 248, 254);
					}

					row++;

					numFila++;
				}


				columna = 3;
				ESheet.Cells[string.Format("B{0}", row)].Value = "Total";
				ESheet.Cells["B" + row].Style.Font.Bold = true;
				foreach (string banco in lista_bancos)
				{
					decimal montoLineaCredito = 0M;

					List<LineaDeCredito> lineas = new List<LineaDeCredito>();
					lineas = lineasCredito.Where(j => j.Banco == banco).ToList();

					if (lineas.Count > 0)
					{
						montoLineaCredito = lineas.Sum(m => m.MontoUSD);
					}
					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = montoLineaCredito;
					ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					columna++;
				}

				columnaStr = Utility.ExcelColumnFromNumber(columna);
				ESheet.Cells[columnaStr + row].Value = autorizada_fila_total_total;
				ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

				ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				row++;


				ESheet.Cells["B10" + ":" + ultimaColumna + "10"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B10" + ":" + ultimaColumna + "10"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B10" + ":B" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B10" + ":B" + (row - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells[ultimaColumna + "10" + ":" + ultimaColumna + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells[ultimaColumna + "10" + ":" + ultimaColumna + (row - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 1) + ":" + ultimaColumna + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 1) + ":" + ultimaColumna + (row - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);
				row++;
				#endregion
				//-------
				//-------
				#region tabla monto dispuesto
				int filaInicioMonto = row;
				columna = 3;
				foreach (string banco in lista_bancos_cartas)
				{
					columna++;
				}
				columnaStr = Utility.ExcelColumnFromNumber(columna);
				headerTotal = "B" + row + ":" + columnaStr + row;
				ESheet.Cells[headerTotal].Merge = true;
				ESheet.Cells[headerTotal].Value = "Monto dispuesto";
				ESheet.Cells[headerTotal].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells[headerTotal].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells[headerTotal].Style.Fill.BackgroundColor.SetColor(1, 191, 191, 189);

				row++;


				columna = 3;
				ESheet.Cells[string.Format("B{0}", row)].Value = "EMPRESA";

				ESheet.Cells["B" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
				foreach (string banco in lista_bancos_cartas)
				{
					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = banco.ToUpper();


					ESheet.Cells[columnaStr + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells[columnaStr + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);


					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					columna++;
				}

				columnaStr = Utility.ExcelColumnFromNumber(columna);
				ESheet.Cells[columnaStr + row].Value = "TOTAL";
				ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

				ESheet.Cells[columnaStr + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells[columnaStr + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				columna++;
				ultimaColumna = Utility.ExcelColumnFromNumber(columna - 1);

				row++;

				numFila = 0;
				decimal monto_fila_total_total = 0M;
				foreach (string empresa in lista_empresas)
                {
					ESheet.Cells[string.Format("B{0}", row)].Value = empresa;
					columna = 3;
					decimal monto_fila_total = 0M;
					foreach (string banco in lista_bancos_cartas)
					{
						int bancoId = 0;
						int empresaId = 0;

						decimal montoEmpresaBanco = 0M;
						List<CartaCreditoReporte> cartas = new List<CartaCreditoReporte>();
						cartas = cartasCredito.Where(j => j.Empresa == empresa &&
						j.Banco == banco).ToList();


						bancoId = bancos.Where(b => b.Nombre == banco).First().Id;
						empresaId = empresas.Where(b => b.Nombre == empresa).First().Id;

						foreach (var carta in cartas)
						{
							bancoId = carta.BancoId;
							empresaId = carta.EmpresaId;
							//----conversion de moneda-----
							/*TipoDeCambio tipoDeCambio = new TipoDeCambio();
							tipoDeCambio.Fecha = fechaDivisa;
							tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == carta.MonedaId).FirstOrDefault().Abbr.ToUpper();
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



							carta.MontoDispuesto = carta.PagosEfectuados + carta.PagosProgramados;
							*/
							string MonedaOriginal = catMonedas.Where(m => m.Id == carta.MonedaId).FirstOrDefault().Abbr.ToUpper();
							string MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
							tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == MonedaOriginal && i.MonedaNueva == MonedaNueva).FirstOrDefault();
							if (tipoDeCambio.MonedaOriginal == "JPY")
							{
								carta.MontoDispuestoUSD = carta.MontoDispuesto / tipoDeCambio.Conversion;
							}
							else
							{
								carta.MontoDispuestoUSD = carta.MontoDispuesto * tipoDeCambio.Conversion;
							}
							//-----------------------------
						}

						montoEmpresaBanco = cartas.Sum(k => k.MontoDispuestoUSD);

						//
						LineaCreditoGpo existe = new LineaCreditoGpo();


						if (empresa == "Draxton México, S. de R.L. de C.V." &&
							banco == "HSBC")
						{
							var o = 0;
						}

						existe = totales.Where(z => z.EmpresaId == empresaId && z.BancoId == bancoId).FirstOrDefault();

						if(existe != null)
                        {
							totales.Where(z => z.EmpresaId == empresaId && z.BancoId == bancoId).FirstOrDefault().TotalMontoDispuesto = montoEmpresaBanco;
						}
						//

						columnaStr = Utility.ExcelColumnFromNumber(columna);

						monto_fila_total += montoEmpresaBanco;
						ESheet.Cells[columnaStr + row].Value = montoEmpresaBanco;
						ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";

						ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

						columna++;
					}

					monto_fila_total_total += monto_fila_total;

					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = monto_fila_total;
					ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					if (numFila % 2 > 0)
					{
						ESheet.Cells["B" + row + ":" + ultimaColumna + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":" + ultimaColumna + row].Style.Fill.BackgroundColor.SetColor(1, 226, 248, 254);
					}

					row++;

					numFila++;
				}

				columna = 3;
				ESheet.Cells[string.Format("B{0}", row)].Value = "Total";
				ESheet.Cells["B" + row].Style.Font.Bold = true;
				foreach (string banco in lista_bancos_cartas)
				{
					decimal montoEmpresaBanco = 0M;

					List<CartaCreditoReporte> cartas = cartasCredito.Where(l => l.Banco == banco).ToList();

					if (cartas.Count > 0)
					{
						montoEmpresaBanco = cartas.Sum(m => m.MontoDispuestoUSD);
					}
					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = montoEmpresaBanco;
					ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					columna++;
				}


				columnaStr = Utility.ExcelColumnFromNumber(columna);
				ESheet.Cells[columnaStr + row].Value = monto_fila_total_total;
				ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

				ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				row++;


				ESheet.Cells["B" + filaInicioMonto + ":" + ultimaColumna + filaInicioMonto].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + filaInicioMonto + ":" + ultimaColumna + filaInicioMonto].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B" + filaInicioMonto + ":B" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + filaInicioMonto + ":B" + (row - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells[ultimaColumna + filaInicioMonto + ":" + ultimaColumna + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells[ultimaColumna + filaInicioMonto + ":" + ultimaColumna + (row - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 1) + ":" + ultimaColumna + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 1) + ":" + ultimaColumna + (row - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);
				row++;

				#endregion
				//-------
				//-------
				#region tabla de lineas de credito totales

				headerTotal = "B" + row;
				filaInicioMonto = row;
				columna = 3;
				foreach (string banco in lista_bancos)
				{
					columna++;
				}
				columnaStr = Utility.ExcelColumnFromNumber(columna);
				headerTotal = "B" + row + ":" + columnaStr + row;

				ESheet.Cells[headerTotal].Merge = true;
				ESheet.Cells[headerTotal].Value = "Total disponible";
				ESheet.Cells[headerTotal].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells[headerTotal].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells[headerTotal].Style.Fill.BackgroundColor.SetColor(1, 191, 191, 189);

				row++;

				columna = 3;
				ESheet.Cells[string.Format("B{0}", row)].Value = "EMPRESA";

				ESheet.Cells["B" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);


				columna = 3;
				foreach (string banco in lista_bancos)
				{
					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = banco.ToUpper();

					ESheet.Cells[columnaStr + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells[columnaStr + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					columna++;
				}

				columnaStr = Utility.ExcelColumnFromNumber(columna);
				ESheet.Cells[columnaStr + row].Value = "TOTAL";
				ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

				ESheet.Cells[columnaStr + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells[columnaStr + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);

				ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);
				columna++;
				ultimaColumna = Utility.ExcelColumnFromNumber(columna - 1);

				row++;

				numFila = 0;
				decimal total_fila_total_total = 0M;
				foreach (string empresa in lista_empresas)
				{

					ESheet.Cells[string.Format("B{0}", row)].Value = empresa;
					columna = 3;
					decimal total_fila_total = 0M;
					foreach (string banco in lista_bancos)
					{

						if(empresa == "Draxton México, S. de R.L. de C.V." &&
							banco == "HSBC")
                        {
							var o = 0;
                        }
						decimal _montoDispuesto = 0M;
						decimal _lineaDeCredito = 0M;

						LineaCreditoGpo item = new LineaCreditoGpo();
						item = totales.Where(z => z.Empresa == empresa && z.Banco == banco).FirstOrDefault();

						if(item != null)
                        {
							_montoDispuesto = item.TotalMontoDispuesto;
							_lineaDeCredito = item.TotalEmpresaBanco;
                        }

						decimal monto = _lineaDeCredito - _montoDispuesto;
						total_fila_total += monto;
						columnaStr = Utility.ExcelColumnFromNumber(columna);
						ESheet.Cells[columnaStr + row].Value = monto;
						ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";

						ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

						columna++;
					}

					total_fila_total_total += total_fila_total;

					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = total_fila_total;
					ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);


					if (numFila % 2 > 0)
					{
						ESheet.Cells["B" + row + ":" + ultimaColumna + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["B" + row + ":" + ultimaColumna + row].Style.Fill.BackgroundColor.SetColor(1, 226, 248, 254);
					}

					row++;

					numFila++;
				}


				columna = 3;
				ESheet.Cells[string.Format("B{0}", row)].Value = "Total";
				ESheet.Cells["B" + row].Style.Font.Bold = true;
				foreach (string banco in lista_bancos)
				{
					decimal montoLineaCredito = 0M;

					List<LineaDeCredito> lineas = new List<LineaDeCredito>();
					lineas = lineasCredito.Where(j => j.Banco == banco).ToList();

					if (lineas.Count > 0)
					{
						decimal _montoDispuesto = 0M;
						decimal _lineaDeCredito = 0M;

						LineaCreditoGpo item = new LineaCreditoGpo();
						_montoDispuesto = totales.Where(z => z.Banco == banco).Sum(y => y.TotalMontoDispuesto);
						_lineaDeCredito = totales.Where(z => z.Banco == banco).Sum(y => y.TotalEmpresaBanco);

						montoLineaCredito = _lineaDeCredito - _montoDispuesto;
					}
					columnaStr = Utility.ExcelColumnFromNumber(columna);
					ESheet.Cells[columnaStr + row].Value = montoLineaCredito;
					ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

					ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

					columna++;
				}

				columnaStr = Utility.ExcelColumnFromNumber(columna);
				ESheet.Cells[columnaStr + row].Value = total_fila_total_total;
				ESheet.Cells[columnaStr + row].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[columnaStr + row].Style.Font.Bold = true;

				ESheet.Cells[columnaStr + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells[columnaStr + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				ESheet.Cells["B" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

				row++;


				ESheet.Cells["B" + filaInicioMonto + ":" + ultimaColumna + filaInicioMonto].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + filaInicioMonto + ":" + ultimaColumna + filaInicioMonto].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B" + filaInicioMonto + ":B" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + filaInicioMonto + ":B" + (row - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells[ultimaColumna + filaInicioMonto + ":" + ultimaColumna + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells[ultimaColumna + filaInicioMonto + ":" + ultimaColumna + (row - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 1) + ":" + ultimaColumna + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 1) + ":" + ultimaColumna + (row - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);
				row++;
				#endregion
				


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