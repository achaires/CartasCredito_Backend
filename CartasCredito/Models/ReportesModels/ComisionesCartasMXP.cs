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
	public class ComisionesCartas : ReporteBase
	{
		public DateTime FechaVencimientoInicio { get; set; } = DateTime.Parse("1969-01-01");
		public DateTime FechaVencimientoFin { get; set; } = DateTime.Parse("1969-01-01");
		public ComisionesCartas(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa)
			: base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Reporte de Comisiones de Cartas de Crédito (MXP)")
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
				var cartasCredito = CartaCreditoReporte.ReporteComisionesMXP(EmpresaId, fechaInicioExact, fechaFinExact).ToList();
				
				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();
				Moneda mndMxp = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "mxp").FirstOrDefault();

				var lista_empresas = Empresa.Get(1);

				if (EmpresaId > 0)
				{
					lista_empresas = lista_empresas.FindAll(e => e.Id == EmpresaId);
				}

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				ESheet.Cells["B9"].Value = "Año";
				ESheet.Cells["C9"].Value = "Mes";
				ESheet.Cells["D9"].Value = "Division";
				ESheet.Cells["E9"].Value = "Empresa";
				ESheet.Cells["F9"].Value = "Proveedor";
				ESheet.Cells["G9"].Value = "Moneda";
				ESheet.Cells["H9"].Value = "Monto programado";
				ESheet.Cells["I9"].Value = "Monto pagado (MXP)";
				/*ESheet.Cells["J9"].Value = "Monto pagado";
				ESheet.Cells["K9"].Value = "Monto pagado (MXP)";*/

				ESheet.Cells["B9:I9"].Style.Font.Bold = true;


				ESheet.Cells["B9:I9"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B9:I9"].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);
				ESheet.Cells["B9:I9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:I9"].Style.Border.Left.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:I9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:I9"].Style.Border.Top.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:I9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:I9"].Style.Border.Right.Color.SetColor(1, 191, 191, 191);
				ESheet.Cells["B9:I9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:I9"].Style.Border.Bottom.Color.SetColor(1, 191, 191, 191);

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
				sheetLogo.SetPosition(20,500);

				int row = 10;
				var granTotalProgramado = 0M;
				var granTotalPagado = 0M;
				var granTotalPagadoUSD = 0M;

				int idx_inicio_proveedor = row;
				int idx_fin_proveedor = row;
				int idx_inicio_empresa = row;
				int idx_fin_empresa = row;
				int idx_inicio_division = row;
				int idx_fin_division = row;
				int idx_inicio_mes = row;
				int idx_fin_mes = row;
				int idx_inicio_anho = row;
				int idx_fin_anho = row;

				decimal anho_monto_programado = 0;
				decimal anho_monto_programado_mx = 0;
				decimal anho_monto_pagado = 0;
				decimal anho_monto_pagado_mx = 0;

				decimal mes_monto_programado = 0;
				decimal mes_monto_programado_mx = 0;
				decimal mes_monto_pagado = 0;
				decimal mes_monto_pagado_mx = 0;

				decimal division_monto_programado = 0;
				decimal division_monto_programado_mx = 0;
				decimal division_monto_pagado = 0;
				decimal division_monto_pagado_mx = 0;

				decimal empresa_monto_programado = 0;
				decimal empresa_monto_programado_mx = 0;
				decimal empresa_monto_pagado = 0;
				decimal empresa_monto_pagado_mx = 0;

				decimal proveedor_monto_programado = 0;
				decimal proveedor_monto_programado_mx = 0;
				decimal proveedor_monto_pagado = 0;
				decimal proveedor_monto_pagado_mx = 0;

				///--------------
				///
				var lista_divisiones = Division.Get();
				var lista_proveedores = Proveedor.Get();
				var lista_empresa = Empresa.Get();

				List<int> anhos = new List<int>();
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				anhos = cartasCredito.Select(i => i.AnhoVencimiento).Distinct().ToList();
				anhos.Sort();

				foreach(int anho in anhos)
				{
					ESheet.Cells[string.Format("B{0}", row)].Value = anho;

					List<CartaCreditoReporte> cartasAnho = new List<CartaCreditoReporte>();
					cartasAnho = cartasCredito.Where(j => j.AnhoVencimiento == anho).ToList();

					List<int> meses = new List<int>();
					meses = cartasAnho.Select(i => i.MesVencimiento).Distinct().ToList();
					meses.Sort();

					idx_inicio_anho = row;
					foreach (int mes in meses)
					{
						ESheet.Cells[string.Format("C{0}", row)].Value = mes;


						List<CartaCreditoReporte> cartasMes = new List<CartaCreditoReporte>();
						cartasMes = cartasAnho.Where(j => j.MesVencimiento == mes).ToList();

						List<int> divisiones = new List<int>();
						divisiones = cartasMes.Select(j => j.DivisionId).Distinct().ToList();
						divisiones.Sort();


						idx_inicio_mes = row;

						foreach (int division in divisiones)
						{
							Division mainDivision = new Division();
							mainDivision = lista_divisiones.Where(k => k.Id == division).FirstOrDefault();
							if(mainDivision == null)
                            {
								mainDivision = new Division();
								mainDivision.Nombre = "División con Id: " + division;
                            }
							ESheet.Cells[string.Format("D{0}", row)].Value = mainDivision.Nombre;


							List<CartaCreditoReporte> cartasDivison = new List<CartaCreditoReporte>();
							cartasDivison = cartasMes.Where(j => j.DivisionId == division).ToList();

							List<int> empresas = new List<int>();
							empresas = cartasDivison.Select(i => i.EmpresaId).Distinct().ToList();
							empresas.Sort();


							idx_inicio_division = row;

							foreach (int empresa in empresas)
                            {
								Empresa mainEmpresa = new Empresa();
								mainEmpresa = lista_empresas.Where(k => k.Id == empresa).FirstOrDefault();
								if (mainEmpresa == null)
								{
									mainEmpresa = new Empresa();
									mainEmpresa.Nombre = "Empresa con Id: " + empresa;
								}
								ESheet.Cells[string.Format("E{0}", row)].Value = mainEmpresa.Nombre;


								List<CartaCreditoReporte> cartasEmpresa = new List<CartaCreditoReporte>();
								cartasEmpresa = cartasDivison.Where(j => j.DivisionId == division).ToList();

								List<int> proveedores = new List<int>();
								proveedores = cartasEmpresa.Select(i => i.ProveedorId).Distinct().ToList();
								proveedores.Sort();


								empresa_monto_programado = 0;
								empresa_monto_programado_mx = 0;
								empresa_monto_pagado = 0;
								empresa_monto_pagado_mx = 0;

								idx_inicio_empresa = row;

								foreach (int proveedor in proveedores)
								{
									Proveedor mainProveedor = new Proveedor();
									mainProveedor = lista_proveedores.Where(l => l.Id == proveedor).FirstOrDefault();
									if (mainProveedor == null)
									{
										mainProveedor = new Proveedor();
										mainProveedor.Nombre = "Proveedor con Id: " + proveedor;
									}
									ESheet.Cells[string.Format("F{0}", row)].Value = mainProveedor.Nombre;


									List<CartaCreditoReporte> cartasProveedor = new List<CartaCreditoReporte>();
									cartasProveedor = cartasEmpresa.Where(j => j.ProveedorId == proveedor).ToList();

									List<int> monedas = new List<int>();
									monedas = cartasProveedor.Select(i => i.MonedaId).Distinct().ToList();
									monedas.Sort();

									idx_inicio_proveedor = row;

									proveedor_monto_programado = 0;
									proveedor_monto_programado_mx = 0;
									proveedor_monto_pagado = 0;
									proveedor_monto_pagado_mx = 0;
									foreach (int moneda in monedas)
									{
										Moneda mainMoneda = new Moneda();
										mainMoneda = catMonedas.Where(l => l.Id == moneda).FirstOrDefault();
										if (mainMoneda == null)
										{
											mainMoneda = new Moneda();
											mainMoneda.Abbr = "MONEDA con Id: " + moneda;
										}
										ESheet.Cells[string.Format("G{0}", row)].Value = mainMoneda.Abbr;


										List<CartaCreditoReporte> cartasMoneda = new List<CartaCreditoReporte>();
										cartasMoneda = cartasProveedor.Where(j => j.MonedaId == moneda).ToList();


                                        #region conversion comision
                                        decimal monto_programado = 0;
										decimal monto_pagado = 0;
										decimal monto_programado_MXP = 0;
										decimal monto_pagado_MXP = 0;

										monto_programado = cartasMoneda.Sum(m => m.ComisionMontoPagado);
										monto_pagado = cartasMoneda.Sum(m => m.ComisionMontoPagado);

										TipoDeCambio tipoDeCambio = new TipoDeCambio();
										tipoDeCambio.Fecha = fechaDivisa;
										tipoDeCambio.MonedaOriginal = mainMoneda.Abbr;
										tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndMxp.Id).FirstOrDefault().Abbr.ToUpper();

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
										}
										else
										{
										}
										monto_programado_MXP = monto_programado * tipoDeCambio.Conversion;
										monto_pagado_MXP = monto_pagado * tipoDeCambio.Conversion;
										#endregion

										proveedor_monto_pagado += monto_pagado;
										proveedor_monto_pagado_mx += monto_pagado_MXP;
										proveedor_monto_programado += monto_programado;
										proveedor_monto_programado_mx += monto_programado_MXP;

                                        ESheet.Cells[string.Format("H{0}", row)].Value = monto_programado;
										ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
										ESheet.Cells[string.Format("I{0}", row)].Value = monto_programado_MXP;
										ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
										/*ESheet.Cells[string.Format("J{0}", row)].Value = monto_programado_MXP;
										ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
										ESheet.Cells[string.Format("K{0}", row)].Value = monto_pagado_MXP;
										ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";*/
										//comision.MontoPagado_USD = comision.MontoPagado * tipoDeCambio.Conversion;

										ESheet.Cells["G" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
										ESheet.Cells["G" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
										ESheet.Cells["G" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
										ESheet.Cells["G" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
										ESheet.Cells["G" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
										ESheet.Cells["G" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
										ESheet.Cells["G" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
										ESheet.Cells["G" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

										//------
										row++;
									}

									idx_fin_proveedor = row;

									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Merge = true;
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Left.Color.SetColor(_BORDE);
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Top.Color.SetColor(_BORDE);
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Right.Color.SetColor(_BORDE);
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + idx_inicio_proveedor + ":F" + idx_fin_proveedor].Style.Border.Bottom.Color.SetColor(_BORDE);

									ESheet.Cells[string.Format("G{0}", row)].Value = "Total Montos";
									ESheet.Cells["G" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
									//ESheet.Cells["F" + row + ":G" + row].Merge = true;
									ESheet.Cells[string.Format("H{0}", row)].Value = proveedor_monto_programado;
									ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
									ESheet.Cells[string.Format("I{0}", row)].Value = proveedor_monto_programado_mx;
									ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
									/*ESheet.Cells[string.Format("J{0}", row)].Value = proveedor_monto_programado_mx;
									ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
									ESheet.Cells[string.Format("K{0}", row)].Value = proveedor_monto_pagado_mx;
									ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";*/

									ESheet.Cells["F" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
									ESheet.Cells["F" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
									ESheet.Cells["F" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
									ESheet.Cells["F" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
									ESheet.Cells["F" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);

									ESheet.Cells["F" + row + ":I" + row].Style.Font.Bold = true;


									ESheet.Cells["G" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
									ESheet.Cells["G" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 221,235,247);

									row++;
								}

								idx_fin_empresa = row;
								
								empresa_monto_pagado += proveedor_monto_pagado;
								empresa_monto_pagado_mx += proveedor_monto_pagado_mx;
								empresa_monto_programado += proveedor_monto_programado;
								empresa_monto_programado_mx += proveedor_monto_programado_mx;
								
								ESheet.Cells[string.Format("F{0}", row)].Value = "Total Proveedores";
								ESheet.Cells[string.Format("H{0}", row)].Value = empresa_monto_programado;
								ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
								ESheet.Cells[string.Format("I{0}", row)].Value = empresa_monto_programado_mx;
								ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
								/*ESheet.Cells[string.Format("J{0}", row)].Value = empresa_monto_programado_mx;
								ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
								ESheet.Cells[string.Format("K{0}", row)].Value = empresa_monto_pagado_mx;
								ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";*/

								ESheet.Cells["F" + row + ":G" + row].Merge = true;
								ESheet.Cells["F" + row + ":G" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
								ESheet.Cells["E" + row + ":I" + row].Style.Font.Bold = true;

								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Merge = true;
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Left.Color.SetColor(_BORDE);
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Top.Color.SetColor(_BORDE);
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Right.Color.SetColor(_BORDE);
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
								ESheet.Cells["E" + idx_inicio_empresa + ":E" + idx_fin_empresa].Style.Border.Bottom.Color.SetColor(_BORDE);

								ESheet.Cells["F" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
								ESheet.Cells["F" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 189, 215, 238);

								row++;
							}

							idx_fin_division = row;
							
							division_monto_pagado += empresa_monto_pagado;
							division_monto_pagado_mx += empresa_monto_pagado_mx;
							division_monto_programado += empresa_monto_programado;
							division_monto_programado_mx += empresa_monto_programado_mx;
							
							ESheet.Cells[string.Format("E{0}", row)].Value = "Total Empresa";
							ESheet.Cells[string.Format("H{0}", row)].Value = division_monto_programado;
							ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
							ESheet.Cells[string.Format("I{0}", row)].Value = division_monto_programado_mx;
							ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
							/*ESheet.Cells[string.Format("J{0}", row)].Value = division_monto_programado_mx;
							ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
							ESheet.Cells[string.Format("K{0}", row)].Value = division_monto_pagado_mx;
							ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";*/

							ESheet.Cells["E" + row + ":G" + row].Merge = true;
							ESheet.Cells["E" + row + ":G" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
							ESheet.Cells["D" + row + ":I" + row].Style.Font.Bold = true;

							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Merge = true;
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Left.Color.SetColor(_BORDE);
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Top.Color.SetColor(_BORDE);
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Right.Color.SetColor(_BORDE);
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
							ESheet.Cells["D" + idx_inicio_division + ":D" + idx_fin_division].Style.Border.Bottom.Color.SetColor(_BORDE);

							ESheet.Cells["E" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
							ESheet.Cells["E" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 180, 198, 231);
							row++;
						}

						idx_fin_mes = row;

						mes_monto_pagado += division_monto_pagado;
						mes_monto_pagado_mx += division_monto_pagado_mx;
						mes_monto_programado += division_monto_programado;
						mes_monto_programado_mx += division_monto_programado_mx;

						ESheet.Cells[string.Format("D{0}", row)].Value = "Total División";
						ESheet.Cells[string.Format("H{0}", row)].Value = mes_monto_programado;
						ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells[string.Format("I{0}", row)].Value = mes_monto_programado_mx;
						ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
						/*ESheet.Cells[string.Format("J{0}", row)].Value = mes_monto_programado_mx;
						ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells[string.Format("K{0}", row)].Value = mes_monto_pagado_mx;
						ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";*/

						ESheet.Cells["D" + row + ":G" + row].Merge = true;
						ESheet.Cells["D" + row + ":G" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
						ESheet.Cells["C" + row + ":I" + row].Style.Font.Bold = true;

						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Merge = true;
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + idx_inicio_mes + ":C" + idx_fin_mes].Style.Border.Bottom.Color.SetColor(_BORDE);

						ESheet.Cells["D" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["D" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 155, 194, 230);

						row++;
					}

					idx_fin_anho = row;

					anho_monto_pagado += mes_monto_pagado;
					anho_monto_pagado_mx += mes_monto_pagado_mx;
					anho_monto_programado += mes_monto_programado;
					anho_monto_programado_mx += mes_monto_programado_mx;

					ESheet.Cells[string.Format("C{0}", row)].Value = "Total Mes";
					ESheet.Cells[string.Format("H{0}", row)].Value = anho_monto_programado;
					ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[string.Format("I{0}", row)].Value = anho_monto_programado_mx;
					ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
					/*ESheet.Cells[string.Format("J{0}", row)].Value = anho_monto_programado_mx;
					ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
					ESheet.Cells[string.Format("K{0}", row)].Value = anho_monto_pagado_mx;
					ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";*/

					ESheet.Cells["C" + row + ":G" + row].Merge = true;
					ESheet.Cells["C" + row + ":G" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
					ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;

					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Merge = true;
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + idx_inicio_anho + ":B" + idx_fin_anho].Style.Border.Bottom.Color.SetColor(_BORDE);

					ESheet.Cells["C" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["C" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 142, 169, 219);

					row++;
				}
				///--------------
				ESheet.Cells[string.Format("B{0}", row)].Value = "Total Año";
				ESheet.Cells[string.Format("H{0}", row)].Value = anho_monto_programado;
				ESheet.Cells[string.Format("H{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[string.Format("I{0}", row)].Value = anho_monto_programado_mx;
				ESheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = " #,##0.00";
				/*ESheet.Cells[string.Format("J{0}", row)].Value = anho_monto_programado_mx;
				ESheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells[string.Format("K{0}", row)].Value = anho_monto_pagado_mx;
				ESheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = " #,##0.00";*/

				ESheet.Cells["B" + row + ":G" + row].Merge = true;
				ESheet.Cells["B" + row + ":G" + row].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + row + ":I" + row].Style.Border.Bottom.Color.SetColor(_BORDE);
				ESheet.Cells["B" + row + ":I" + row].Style.Font.Bold = true;



				ESheet.Cells["B" + row + ":I" + row].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + row + ":I" + row].Style.Fill.BackgroundColor.SetColor(1, 47, 117, 181);

				row++;

				row++;
				/*
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
				*/
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

				/*
				ESheet.Cells["B9" + ":I9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":I9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + (row - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["K9" + ":I" + (row - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["K9" + ":I" + (row - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (row - 1) + ":I" + (row - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (row - 1) + ":I" + (row - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);
				*/

				ESheet.Cells["A:AZ"].AutoFitColumns();
				ESheet.Column(2).Width = 10;
				ESheet.Column(3).Width = 10;
				ESheet.Column(4).Width = 20;
				ESheet.Column(5).Width = 35;
				ESheet.Column(6).Width = 40;
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