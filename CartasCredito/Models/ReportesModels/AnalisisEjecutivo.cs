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
	public class AnalisisEjecutivo : ReporteBase
	{
		public AnalisisEjecutivo(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) 
			: base (fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "P",  "Reporte de Análisis Ejecutivo de Cartas de Crédito")
		{
		}

		public Reporte GenerarV1()
		{
			var reporteResultado = new Reporte();

			try
			{
				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);

				var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).Where(cc => cc.TipoCarta == "Comercial").GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
				
				var catMonedas = Moneda.Get();

				// Headings
				ESheet.Cells["B9"].Value = "Empresa";
				ESheet.Cells["C9"].Value = "Banco";
				ESheet.Cells["D9"].Value = "Proveedor";
				ESheet.Cells["E9"].Value = "Producto";
				ESheet.Cells["F9"].Value = "País";
				ESheet.Cells["G9"].Value = "Descripción";
				ESheet.Cells["H9"].Value = "Moneda";
				ESheet.Cells["I9"].Value = "Importe Total";
				ESheet.Cells["J9"].Value = "Importe < 50,0000";
				ESheet.Cells["K9"].Value = "50,000 < Importe > 300,000";
				ESheet.Cells["L9"].Value = "Importe > 300,000";
				ESheet.Cells["M9"].Value = "1 Periodo";
				ESheet.Cells["N9"].Value = "2 Periodo";
				ESheet.Cells["O9"].Value = "3 Periodos";
				ESheet.Cells["P9"].Value = "Días plazo proveedor después de B/L";
				ESheet.Cells["B9:P9"].Style.Font.Bold = true;

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(),".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20,200);

				int fila = 10;

				/*var proveedoresCat = Proveedor.Get(1);

				var grupos = cartasCredito
					.GroupBy(carta => carta.Empresa)
					.Select(grupoEmpresa => new
					{
						grupoEmpresa.Key,
						TotalEmpresa = grupoEmpresa.Sum(c => c.MontoOriginalLC),
						GruposMoneda = grupoEmpresa
							.GroupBy(carta => new {carta.Moneda, carta.Banco, carta.Proveedor})
							.Select(grupoMoneda => new
							{
								grupoMoneda.Key,
								MonedaId = grupoMoneda.First().MonedaId,
								TotalMoneda = grupoMoneda.Sum(carta => carta.MontoOriginalLC),
								CartasDeCredito = grupoMoneda.ToList()
							}).ToList()
					}).ToList();

				var granTotal = 0M;
				var moneda1="";
				var moneda2="";
				var moneda3="";
				var moneda4="";
				var moneda5="";
				var sumaMoneda1 = 0M;
				var sumaMoneda2 = 0M;
				var sumaMoneda3 = 0M;
				var sumaMoneda4 = 0M;
				var sumaMoneda5 = 0M;

				var divisasList = new List<int>();

				foreach (var grupoEmpresa in grupos)
				{
					ESheet.Cells[string.Format("B{0}", fila)].Value = grupoEmpresa.Key;
					var sumaTipoMoneda = 0M;
					var tipoMoneda="";
					var monedaId = 0;
					foreach (var grupoMoneda in grupoEmpresa.GruposMoneda)
					{
						foreach (var carta in grupoMoneda.CartasDeCredito)
						{
							TimeSpan diferencia = carta.FechaVencimiento.Subtract(carta.FechaApertura);
							int cantidadDias = (int)diferencia.TotalDays;
							decimal periodosDecimal = Convert.ToDecimal(cantidadDias) / 90M;
							var periodos = Math.Ceiling(periodosDecimal);
							var proveedorObj = proveedoresCat.First(pv => pv.Id == carta.ProveedorId);

							ESheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
							ESheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
							ESheet.Cells[string.Format("E{0}", fila)].Value = carta.DescripcionMercancia;
							ESheet.Cells[string.Format("F{0}", fila)].Value = proveedorObj.Pais;
							ESheet.Cells[string.Format("G{0}", fila)].Value = carta.TipoActivo;
							ESheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;

							ESheet.Cells[string.Format("I{0}", fila)].Value = carta.MontoOriginalLC;
							ESheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							ESheet.Cells[string.Format("J{0}", fila)].Value = carta.MontoOriginalLC < 50000 ? "1" : "0";
							ESheet.Cells[string.Format("K{0}", fila)].Value = carta.MontoOriginalLC > 50000 && carta.MontoOriginalLC < 300000 ? "1" : "0";
							ESheet.Cells[string.Format("L{0}", fila)].Value = carta.MontoOriginalLC > 300000 ? "1" : "0";
							ESheet.Cells[string.Format("M{0}", fila)].Value = periodos <= 1 ? "1" : "0";
							ESheet.Cells[string.Format("N{0}", fila)].Value = periodos == 2 ? "1" : "0";
							ESheet.Cells[string.Format("O{0}", fila)].Value = periodos > 2 ? "1" : "0";
							ESheet.Cells[string.Format("P{0}", fila)].Value = carta.DiasPlazoProveedor;

							//fila++;
						}

						//ESheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key.Moneda;
						//ESheet.Cells["I" + fila].Value = grupoMoneda.TotalMoneda;
						//ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						// Calcula y agrega fila de conversión a dólares
						//fila++;
						//var totalMonedaEnUsd = 0M; // ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);
						//divisasList.Add(grupoMoneda.MonedaId);

						//ESheet.Cells["H" + fila].Value = "Total USD";
						//ESheet.Cells["I" + fila].Value = totalMonedaEnUsd;
						//ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						//suma totales de la moneda, descripcion de moneda y id moneda
						monedaId=grupoMoneda.MonedaId;
						tipoMoneda=grupoMoneda.Key.Moneda;
						sumaTipoMoneda += grupoMoneda.TotalMoneda;

						// suma los totales por moneda
						if(grupoMoneda.MonedaId==1){
							moneda1=grupoMoneda.Key.Moneda;
							sumaMoneda1 += grupoMoneda.TotalMoneda;
						}else if(grupoMoneda.MonedaId==2){
							moneda2=grupoMoneda.Key.Moneda;
							sumaMoneda2 += grupoMoneda.TotalMoneda;
						}else if(grupoMoneda.MonedaId==3){
							moneda3=grupoMoneda.Key.Moneda;
							sumaMoneda3 += grupoMoneda.TotalMoneda;
						}else if(grupoMoneda.MonedaId==4){
							moneda4=grupoMoneda.Key.Moneda;
							sumaMoneda4 += grupoMoneda.TotalMoneda;
						}else if(grupoMoneda.MonedaId==5){
							moneda5=grupoMoneda.Key.Moneda;
							sumaMoneda5 += grupoMoneda.TotalMoneda;
						}

						//fila++;
						fila++;
					}

					ESheet.Cells["H" + fila].Value = "Total " + tipoMoneda;
					ESheet.Cells["I" + fila].Value = sumaTipoMoneda;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

					fila++;
					// Calcula y agrega fila de conversión a dólares
					//var totalMonedaEnUsd = 0M; // ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);
					var totalMonedaEnUsd = ConversionUSD(monedaId, sumaTipoMoneda, fechaDivisa);
					divisasList.Add(monedaId);
					ESheet.Cells["H" + fila].Value = "Total USD";
					ESheet.Cells["I" + fila].Value = totalMonedaEnUsd;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

					granTotal += grupoEmpresa.TotalEmpresa;

					fila++;
					fila++;
				}

				ESheet.Cells["H" + fila].Value = "GRAN TOTAL:";
				ESheet.Cells["I" + fila].Value = granTotal;
				ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

				if(moneda1 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda1;
					ESheet.Cells["I" + fila].Value = sumaMoneda1;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda2 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda2;
					ESheet.Cells["I" + fila].Value = sumaMoneda2;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda3 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda3;
					ESheet.Cells["I" + fila].Value = sumaMoneda3;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda4 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda4;
					ESheet.Cells["I" + fila].Value = sumaMoneda4;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda5 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda5;
					ESheet.Cells["I" + fila].Value = sumaMoneda5;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}

				fila++;
				fila++;

				var sumaProcentajes = 0M;

				foreach (var grupoEmpresa in grupos)
				{
					var porcentajeEmpresa = Math.Round(Math.Round(grupoEmpresa.TotalEmpresa, 4) / granTotal, 4) * 100;

					sumaProcentajes += porcentajeEmpresa;

					ESheet.Cells["H" + fila].Value = grupoEmpresa.Key;
					ESheet.Cells["I" + fila].Value = porcentajeEmpresa + "%";

					fila++;
				}*/

				var proveedoresCat = Proveedor.Get(1);

				/*var grupos = cartasCredito
					.GroupBy(carta => carta.Empresa)
					.Select(grupoEmpresa => new
					{
						grupoEmpresa.Key,
						TotalEmpresa = grupoEmpresa.Sum(c => c.MontoOriginalLC),
						GruposBanco = grupoEmpresa
							.GroupBy(carta => new {carta.Banco, carta.Proveedor})
							.Select(grupoBanco => new 
							{
								grupoBanco.Key,
								BancoId = grupoBanco.First().BancoId,
								GruposMoneda = grupoBanco
									.GroupBy(carta => carta.Moneda)
									.Select(grupoMoneda => new
									{
										grupoMoneda.Key,
										MonedaId = grupoMoneda.First().MonedaId,
										TotalMoneda = grupoMoneda.Sum(carta => carta.MontoOriginalLC),
										CartasDeCredito = grupoMoneda.ToList()
									}).ToList()
							}).ToList()
					}).ToList();*/
				var grupos = cartasCredito
					.GroupBy(carta => carta.Empresa)
					.Select(grupoEmpresa => new
					{
						grupoEmpresa.Key,
						TotalEmpresa = grupoEmpresa.Sum(c => c.MontoOriginalLC),
						GruposMoneda = grupoEmpresa
							.GroupBy(carta => carta.Moneda)
							.Select(grupoMoneda => new
							{
								grupoMoneda.Key,
								MonedaId = grupoMoneda.First().MonedaId,
								GruposBanco=grupoMoneda
									.GroupBy(carta => new {carta.Banco, carta.Proveedor })
									.Select(grupoBanco => new 
									{
										grupoBanco.Key,
										BancoId=grupoBanco.First().BancoId,
										TotalMoneda = grupoBanco.Sum(carta => carta.MontoOriginalLC),
										CartasDeCredito = grupoBanco.ToList()
									}).ToList()
							}).ToList()
					}).ToList();

				var granTotal = 0M;
				var moneda1="";
				var moneda2="";
				var moneda3="";
				var moneda4="";
				var moneda5="";
				var tipoCambioMoneda1 = 0M;
				var tipoCambioMoneda2 = 0M;
				var tipoCambioMoneda3 = 0M;
				var tipoCambioMoneda4 = 0M;
				var tipoCambioMoneda5 = 0M;
				var sumaTipoMoneda = 0M;
				var tipoMoneda="";
				var monedaId = 0;

				var divisasList = new List<int>();

				foreach (var grupoEmpresa in grupos)
				{
					ESheet.Cells[string.Format("B{0}", fila)].Value = grupoEmpresa.Key;
					
					foreach (var grupoMoneda in grupoEmpresa.GruposMoneda)
					{
						sumaTipoMoneda=0M;
						tipoMoneda="";
						monedaId=0;
						foreach(var grupoBanco in grupoMoneda.GruposBanco)
						{
							foreach (var carta in grupoBanco.CartasDeCredito)
							{
								TimeSpan diferencia = carta.FechaVencimiento.Subtract(carta.FechaApertura);
								int cantidadDias = (int)diferencia.TotalDays;
								decimal periodosDecimal = Convert.ToDecimal(cantidadDias) / 90M;
								var periodos = Math.Ceiling(periodosDecimal);
								var proveedorObj = proveedoresCat.First(pv => pv.Id == carta.ProveedorId);

								ESheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
								ESheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
								ESheet.Cells[string.Format("E{0}", fila)].Value = carta.DescripcionMercancia;
								ESheet.Cells[string.Format("F{0}", fila)].Value = proveedorObj.Pais;
								ESheet.Cells[string.Format("G{0}", fila)].Value = carta.TipoActivo;
								ESheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;

								//ESheet.Cells[string.Format("I{0}", fila)].Value = carta.MontoOriginalLC;
								ESheet.Cells[string.Format("I{0}", fila)].Value = grupoBanco.TotalMoneda;
								ESheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

								ESheet.Cells[string.Format("J{0}", fila)].Value = carta.MontoOriginalLC < 50000 ? "1" : "0";
								ESheet.Cells[string.Format("K{0}", fila)].Value = carta.MontoOriginalLC > 50000 && carta.MontoOriginalLC < 300000 ? "1" : "0";
								ESheet.Cells[string.Format("L{0}", fila)].Value = carta.MontoOriginalLC > 300000 ? "1" : "0";
								ESheet.Cells[string.Format("M{0}", fila)].Value = periodos <= 1 ? "1" : "0";
								ESheet.Cells[string.Format("N{0}", fila)].Value = periodos == 2 ? "1" : "0";
								ESheet.Cells[string.Format("O{0}", fila)].Value = periodos > 2 ? "1" : "0";
								ESheet.Cells[string.Format("P{0}", fila)].Value = carta.DiasPlazoProveedor;

								//fila++;
							}

							//ESheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key.Moneda;
							//ESheet.Cells["I" + fila].Value = grupoMoneda.TotalMoneda;
							//ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

							// Calcula y agrega fila de conversión a dólares
							//fila++;
							//var totalMonedaEnUsd = 0M; // ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);
							//divisasList.Add(grupoMoneda.MonedaId);

							//ESheet.Cells["H" + fila].Value = "Total USD";
							//ESheet.Cells["I" + fila].Value = totalMonedaEnUsd;
							//ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

							
							//fila++;
							fila++;
							//suma totales de la moneda, descripcion de moneda y id moneda
							sumaTipoMoneda += grupoBanco.TotalMoneda;
						}
						//suma totales de la moneda, descripcion de moneda y id moneda
						monedaId=grupoMoneda.MonedaId;
						//tipoMoneda=grupoMoneda.Key.Moneda;
						tipoMoneda=grupoMoneda.Key;
						//sumaTipoMoneda += grupoMoneda.TotalMoneda;
						
						ESheet.Cells["H" + fila].Value = "Total " + tipoMoneda;
						ESheet.Cells["I" + fila].Value = sumaTipoMoneda;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						fila++;
						// Calcula y agrega fila de conversión a dólares
						//var totalMonedaEnUsd = 0M; // ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);
						var totalMonedaEnUsd = ConversionUSD(monedaId, sumaTipoMoneda, fechaDivisa);
						divisasList.Add(monedaId);
						ESheet.Cells["H" + fila].Value = "Total USD";
						ESheet.Cells["H" + fila].Style.Font.Bold=true;
						ESheet.Cells["I" + fila].Value = totalMonedaEnUsd;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
						ESheet.Cells["I" + fila].Style.Font.Bold=true;
						fila++;
						fila++;

						// suma los totales por moneda
						if(grupoMoneda.MonedaId==1){
							//moneda1=grupoMoneda.Key.Moneda;
							moneda1=grupoMoneda.Key;
							tipoCambioMoneda1 = 0;
						}else if(grupoMoneda.MonedaId==2){
							//moneda2=grupoMoneda.Key.Moneda;
							moneda2=grupoMoneda.Key;
							tipoCambioMoneda2 = 0;
						}else if(grupoMoneda.MonedaId==3){
							//moneda3=grupoMoneda.Key.Moneda;
							moneda3=grupoMoneda.Key;
							tipoCambioMoneda3 = 0;
						}else if(grupoMoneda.MonedaId==4){
							//moneda4=grupoMoneda.Key.Moneda;
							moneda4=grupoMoneda.Key;
							tipoCambioMoneda4 = 0;
						}else if(grupoMoneda.MonedaId==5){
							//moneda5=grupoMoneda.Key.Moneda;
							moneda5=grupoMoneda.Key;
							tipoCambioMoneda5 = 0;
						}
					}
					/*ESheet.Cells["H" + fila].Value = "Total " + tipoMoneda;
					ESheet.Cells["I" + fila].Value = sumaTipoMoneda;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";*/

					

					granTotal += grupoEmpresa.TotalEmpresa;
				}

				ESheet.Cells["H" + fila].Value = "GRAN TOTAL:";
				ESheet.Cells["I" + fila].Value = granTotal;
				ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

				if(moneda1 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda1;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda1;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda2 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda2;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda2;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda3 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda3;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda3;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda4 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda4;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda4;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}
				if(moneda5 != ""){
					fila++;
					ESheet.Cells["H" + fila].Value = moneda5;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda5;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";
				}

				fila++;
				fila++;

				var sumaProcentajes = 0M;

				foreach (var grupoEmpresa in grupos)
				{
					var porcentajeEmpresa = Math.Round(Math.Round(grupoEmpresa.TotalEmpresa, 4) / granTotal, 4) * 100;

					sumaProcentajes += porcentajeEmpresa;

					ESheet.Cells["H" + fila].Value = grupoEmpresa.Key;
					ESheet.Cells["I" + fila].Value = porcentajeEmpresa + "%";

					fila++;
				}


				ESheet.Cells["A:AZ"].AutoFitColumns();

				ESheet.Column(5).Width = 25;
				ESheet.Column(4).Width = 25;
				ESheet.Column(8).Width = 25;
				ESheet.Column(11).Width = 25;
				ESheet.Column(16).Width = 25;

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

		public override Reporte Generar()
		{
			var reporteResultado = new Reporte();
			//fechaDivisa = fechaTipoDeCambio;
			try
			{
				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);

				//var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).Where(cc => cc.TipoCarta == "Comercial").GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);

				var cartasCredito = CartaCreditoReporte.ReporteAnalisisEjecutivo(EmpresaId, fechaInicioExact, fechaFinExact, fechaDivisa).Where(cc => cc.TipoCarta == "Comercial").GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
				// Headings
				ESheet.Cells["B9"].Value = "Empresa";
				ESheet.Cells["C9"].Value = "Banco";
				ESheet.Cells["D9"].Value = "Proveedor";
				ESheet.Cells["E9"].Value = "Producto";
				ESheet.Cells["F9"].Value = "País";
				ESheet.Cells["G9"].Value = "Descripción";
				ESheet.Cells["H9"].Value = "Moneda";
				ESheet.Cells["I9"].Value = "Importe Total";
				ESheet.Cells["J9"].Value = "Importe < 50,0000";
				ESheet.Cells["K9"].Value = "50,000 < Importe > 300,000";
				ESheet.Cells["L9"].Value = "Importe > 300,000";
				ESheet.Cells["M9"].Value = "1 Periodo";
				ESheet.Cells["N9"].Value = "2 Periodo";
				ESheet.Cells["O9"].Value = "3 Periodos";
				ESheet.Cells["P9"].Value = "Días plazo proveedor después de B/L";
				ESheet.Cells["B9:P9"].Style.Font.Bold = true;

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(), ".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20, 320);

				int fila = 10;

				var proveedoresCat = Proveedor.Get(1);

				var grupos = cartasCredito
					.GroupBy(carta => carta.Empresa)
					.Select(grupoEmpresa => new
					{
						grupoEmpresa.Key,
						TotalEmpresa = grupoEmpresa.Sum(c => c.MontoOriginalLC),
						TotalEmpresaUSD = grupoEmpresa.Sum(c => c.MontoOriginalLCUSD),
						TotalEmpresa_TipoCambio = (decimal)0,
						GruposMoneda = grupoEmpresa
							.GroupBy(carta => carta.Moneda)
							.Select(grupoMoneda => new
							{
								grupoMoneda.Key,
								MonedaId = grupoMoneda.First().MonedaId,
								GruposBanco = grupoMoneda
									.GroupBy(carta => new { carta.Banco, carta.Proveedor })
									.Select(grupoBanco => new
									{
										grupoBanco.Key,
										BancoId = grupoBanco.First().BancoId,
										TotalMoneda = grupoBanco.Sum(carta => carta.MontoOriginalLC),
										TotalMonedaUSD = grupoBanco.Sum(carta => carta.MontoOriginalLCUSD),
										CartasDeCredito = grupoBanco.ToList()
									}).ToList()
							}).ToList()
					}).ToList();


				var granTotal = 0M;
				var moneda1 = "";
				var moneda2 = "";
				var moneda3 = "";
				var moneda4 = "";
				var moneda5 = "";
				var tipoCambioMoneda1 = 0M;
				var tipoCambioMoneda2 = 0M;
				var tipoCambioMoneda3 = 0M;
				var tipoCambioMoneda4 = 0M;
				var tipoCambioMoneda5 = 0M;
				var sumaTipoMoneda = 0M;
				var sumaTipoMonedaUSD = 0M;


				var sumaImporteMenor = 0M;
				var sumaImporteIgual = 0M;
				var sumaImporteMayor = 0M;
				var suma1Periodo = 0M;
				var suma2Periodo = 0M;
				var suma3Periodo = 0M;


				var tipoMoneda = "";
				var monedaId = 0;

				var divisasList = new List<int>();
				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);


				List<CTMObjReporte> empresas = new List<CTMObjReporte>();

				TipoDeCambio tipoDeCambio = new TipoDeCambio();
				tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == "USD" && i.MonedaNueva == "MXP").FirstOrDefault();
				/*tipoDeCambio.Fecha = fechaDivisa;
				tipoDeCambio.MonedaOriginal = "USD";
				tipoDeCambio.MonedaNueva = "MXP";
				tipoDeCambio.Get();*/

				/*if (tipoDeCambio.ConversionStr == "-1")
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
				}*/

				decimal dolarAPesos = tipoDeCambio.Conversion;

				fila++;
				fila++;

				foreach (var grupoEmpresa in grupos)
				{
					ESheet.Cells[string.Format("B{0}", fila)].Value = grupoEmpresa.Key;
					decimal totales = 0;
					foreach (var grupoMoneda in grupoEmpresa.GruposMoneda)
					{
						sumaTipoMoneda = 0M;
						sumaTipoMonedaUSD = 0M;
						sumaImporteMenor = 0M;
						sumaImporteIgual = 0M;
						sumaImporteMayor = 0M;
						suma1Periodo = 0M;
						suma2Periodo = 0M;
						suma3Periodo = 0M;
						tipoMoneda = "";
						monedaId = 0;
						foreach (var grupoBanco in grupoMoneda.GruposBanco)
						{
							int importeMenor = 0;
							int importeIgual = 0;
							int importeMayor = 0;
							int periodo1 = 0;
							int periodo2 = 0;
							int periodo3 = 0;

							foreach (var carta in grupoBanco.CartasDeCredito)
							{
								TimeSpan diferencia = carta.FechaVencimiento.Subtract(carta.FechaApertura);
								int cantidadDias = (int)diferencia.TotalDays;
								decimal periodosDecimal = Convert.ToDecimal(cantidadDias) / 90M;
								var periodos = Math.Ceiling(periodosDecimal);
								var proveedorObj = proveedoresCat.First(pv => pv.Id == carta.ProveedorId);

								ESheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
								ESheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
								ESheet.Cells[string.Format("E{0}", fila)].Value = carta.DescripcionMercancia;
								ESheet.Cells[string.Format("F{0}", fila)].Value = proveedorObj.Pais;
								ESheet.Cells[string.Format("G{0}", fila)].Value = carta.TipoActivo;
								ESheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;


								//ESheet.Cells[string.Format("I{0}", fila)].Value = carta.MontoOriginalLC;
								ESheet.Cells[string.Format("I{0}", fila)].Value = grupoBanco.TotalMoneda;
								ESheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = " #,##0.00";


								//ESheet.Cells[string.Format("R{0}", fila)].Value = grupoBanco.TotalMonedaUSD;
								//ESheet.Cells[string.Format("R{0}", fila)].Style.Numberformat.Format = " #,##0.00";

								//---
								if (carta.MontoOriginalLC < 50000)
								{
									importeMenor += 1;
								}
								else if (carta.MontoOriginalLC > 50000 && carta.MontoOriginalLC < 300000)
								{
									importeIgual += 1;
								}
								else if (carta.MontoOriginalLC > 300000)
								{
									importeMayor += 1;
								}

								if (periodos <= 1)
								{
									periodo1 += 1;
								}
								else if (periodos == 2)
								{
									periodo2 += 1;
								}
								else if (periodos > 2)
								{
									periodo3 += 1;
								}
								//---
								/*
								ESheet.Cells[string.Format("J{0}", fila)].Value = carta.MontoOriginalLC < 50000 ? "1" : "0";
								ESheet.Cells[string.Format("K{0}", fila)].Value = carta.MontoOriginalLC > 50000 && carta.MontoOriginalLC < 300000 ? "1" : "0";
								ESheet.Cells[string.Format("L{0}", fila)].Value = carta.MontoOriginalLC > 300000 ? "1" : "0";
								*/
								ESheet.Cells[string.Format("J{0}", fila)].Value = importeMenor;
								ESheet.Cells[string.Format("K{0}", fila)].Value = importeIgual;
								ESheet.Cells[string.Format("L{0}", fila)].Value = importeMayor;

								/*
								ESheet.Cells[string.Format("M{0}", fila)].Value = periodos <= 1 ? "1" : "0";
								ESheet.Cells[string.Format("N{0}", fila)].Value = periodos == 2 ? "1" : "0";
								ESheet.Cells[string.Format("O{0}", fila)].Value = periodos > 2 ? "1" : "0";
								*/

								ESheet.Cells[string.Format("M{0}", fila)].Value = periodo1;
								ESheet.Cells[string.Format("N{0}", fila)].Value = periodo2;
								ESheet.Cells[string.Format("O{0}", fila)].Value = periodo3;


								ESheet.Cells[string.Format("P{0}", fila)].Value = carta.DiasPlazoProveedor;

								//fila++;

								/*if (carta.MontoOriginalLC < 50000)
								{
									sumaImporteMenor += 1;
								}
								else if (carta.MontoOriginalLC == 50000)
								{
									sumaImporteIgual += 1;
								}
								else if (carta.MontoOriginalLC > 50000)
								{
									sumaImporteMayor += 1;
								}

								if (periodos <= 1) {
									suma1Periodo += 1;
								}
								else if (periodos == 2) {
									suma2Periodo += 1;
								}
								else if (periodos > 2) {
									suma3Periodo += 1;
								}*/
							}
							fila++;
							//suma totales de la moneda, descripcion de moneda y id moneda
							sumaTipoMoneda += grupoBanco.TotalMoneda;
							sumaTipoMonedaUSD += grupoBanco.TotalMonedaUSD;
							sumaImporteMenor += importeMenor;
							sumaImporteIgual += importeIgual;
							sumaImporteMayor += importeMayor;

							suma1Periodo = periodo1;
							suma2Periodo = periodo2;
							suma3Periodo = periodo3;
						}
						//suma totales de la moneda, descripcion de moneda y id moneda
						monedaId = grupoMoneda.MonedaId;
						//tipoMoneda=grupoMoneda.Key.Moneda;
						tipoMoneda = grupoMoneda.Key;
						//sumaTipoMoneda += grupoMoneda.TotalMoneda;

						ESheet.Cells["H" + fila].Value = "Total " + tipoMoneda;
						ESheet.Cells["I" + fila].Value = sumaTipoMoneda;
						//ESheet.Cells["R" + fila].Value = sumaTipoMonedaUSD;

						ESheet.Cells["J" + fila].Value = sumaImporteMenor;
						ESheet.Cells["K" + fila].Value = sumaImporteIgual;
						ESheet.Cells["L" + fila].Value = sumaImporteMayor;

						ESheet.Cells["M" + fila].Value = suma1Periodo;
						ESheet.Cells["N" + fila].Value = suma2Periodo;
						ESheet.Cells["O" + fila].Value = suma3Periodo;


						ESheet.Cells["I" + fila +":O" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["I" + fila + ":O" + fila].Style.Fill.BackgroundColor.SetColor(1,102,255,51);
						ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";

						fila++;
						// Calcula y agrega fila de conversión a dólares
						//var totalMonedaEnUsd = 0M; // ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);



						//----conversion de moneda-----
						/*tipoDeCambio = new TipoDeCambio();
						tipoDeCambio.Fecha = fechaDivisa;
						tipoDeCambio.MonedaOriginal = catMonedas.Where(m => m.Id == monedaId).FirstOrDefault().Abbr.ToUpper();
						tipoDeCambio.MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();

						List<TipoDeCambio> busqueda = tiposDeCambio.Where(t => t.MonedaOriginal == tipoDeCambio.MonedaOriginal &&
						t.MonedaNueva == tipoDeCambio.MonedaNueva &&
						t.Fecha == tipoDeCambio.Fecha).ToList();
						if(busqueda.Count > 0)
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
						//-----------------------------

						//var totalMonedaEnUsd = ConversionUSD(monedaId, sumaTipoMoneda, fechaDivisa);

						string abbr = catMonedas.Where(m => m.Id == monedaId).FirstOrDefault().Abbr.ToUpper();
						tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == abbr && i.MonedaNueva == "USD").FirstOrDefault();
						decimal sumaConvertido = 0;
						if(tipoDeCambio.MonedaOriginal == "JPY")
                        {
							sumaConvertido = sumaTipoMoneda / tipoDeCambio.Conversion;
                        }
                        else
						{
							sumaConvertido = sumaTipoMoneda * tipoDeCambio.Conversion;
						}

						var totalMonedaEnUsd = sumaConvertido;
						divisasList.Add(monedaId);
						ESheet.Cells["H" + fila].Value = "Total USD";
						ESheet.Cells["H" + fila].Style.Font.Bold = true;
						ESheet.Cells["I" + fila].Value = totalMonedaEnUsd;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";
						ESheet.Cells["I" + fila].Style.Font.Bold = true;
						ESheet.Cells["I" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["I" + fila].Style.Fill.BackgroundColor.SetColor(1, 0, 176, 240);
						fila++;
						fila++;

						// suma los totales por moneda
						if (grupoMoneda.MonedaId == 1)
						{
							//moneda1=grupoMoneda.Key.Moneda;
							moneda1 = grupoMoneda.Key;
							tipoCambioMoneda1 = dolarAPesos; // tipoDeCambio.Conversion;// 0;
						}
						else if (grupoMoneda.MonedaId == 2)
						{
							//moneda2=grupoMoneda.Key.Moneda;
							moneda2 = grupoMoneda.Key;
							tipoCambioMoneda2 = tipoDeCambio.Conversion;//0;
						}
						else if (grupoMoneda.MonedaId == 3)
						{
							//moneda3=grupoMoneda.Key.Moneda;
							moneda3 = grupoMoneda.Key;
							tipoCambioMoneda3 = tipoDeCambio.Conversion;//0;
						}
						else if (grupoMoneda.MonedaId == 4)
						{
							//moneda4=grupoMoneda.Key.Moneda;
							moneda4 = grupoMoneda.Key;
							tipoCambioMoneda4 = tipoDeCambio.Conversion;//0;
						}
						else if (grupoMoneda.MonedaId == 5)
						{
							//moneda5=grupoMoneda.Key.Moneda;
							moneda5 = grupoMoneda.Key;
							tipoCambioMoneda5 = tipoDeCambio.Conversion;//0;
						}

						totales += totalMonedaEnUsd;
					}

					granTotal += totales;
					empresas.Add(new CTMObjReporte()
					{
						key = grupoEmpresa.Key,
						value = totales
					});
					//granTotal += grupoEmpresa.TotalEmpresa;
				}

				ESheet.Cells["H" + fila].Value = "GRAN TOTAL:";
				ESheet.Cells["I" + fila].Value = granTotal;
				ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";
				ESheet.Cells["I" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["I" + fila].Style.Fill.BackgroundColor.SetColor(1, 0, 176, 240);

				ESheet.Cells["H" + fila + ":I" + fila].Style.Font.Bold = true;

				fila++;

				if (moneda1 != "")
				{
					fila++;
					ESheet.Cells["H" + fila].Value = moneda1;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda1;
					ESheet.Cells["H" + fila].Style.Font.Bold = true;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";
				}
				if (moneda2 != "")
				{
					fila++;
					ESheet.Cells["H" + fila].Value = moneda2;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda2;
					ESheet.Cells["H" + fila].Style.Font.Bold = true;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";
				}
				if (moneda3 != "")
				{
					fila++;
					ESheet.Cells["H" + fila].Value = moneda3;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda3;
					ESheet.Cells["H" + fila].Style.Font.Bold = true;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";
				}
				if (moneda4 != "")
				{
					fila++;
					ESheet.Cells["H" + fila].Value = moneda4;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda4;
					ESheet.Cells["H" + fila].Style.Font.Bold = true;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";
				}
				if (moneda5 != "")
				{
					fila++;
					ESheet.Cells["H" + fila].Value = moneda5;
					ESheet.Cells["I" + fila].Value = tipoCambioMoneda5;
					ESheet.Cells["H" + fila].Style.Font.Bold = true;
					ESheet.Cells["I" + fila].Style.Numberformat.Format = " #,##0.00";
				}

				fila++;
				fila++;

				var sumaProcentajes = 0M;

				int idxFila = 0;
				foreach (var grupoEmpresa in empresas)
				{
					var porcentajeEmpresa = Math.Round(Math.Round(grupoEmpresa.value, 4) / granTotal, 4) * 100;

					sumaProcentajes += porcentajeEmpresa;

					ESheet.Cells["H" + fila].Value = grupoEmpresa.key;
					ESheet.Cells["I" + fila].Value = porcentajeEmpresa + "%";

					if ((idxFila % 2) == 0)
					{
						ESheet.Cells["G" + fila + ":J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["G" + fila + ":J" + fila].Style.Fill.BackgroundColor.SetColor(1, 183, 222, 232);
					}
					else
					{
						ESheet.Cells["G" + fila + ":J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						ESheet.Cells["G" + fila + ":J" + fila].Style.Fill.BackgroundColor.SetColor(1, 146, 205, 220);
					}
					ESheet.Cells["G" + fila + ":J" + fila].Style.Font.Bold = true;

					fila++;
					idxFila++;
				}
				ESheet.Cells["H" + fila].Value = "TOTAL";
				ESheet.Cells["I" + fila].Value = sumaProcentajes + "%";
				ESheet.Cells["G" + fila + ":J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["G" + fila + ":J" + fila].Style.Fill.BackgroundColor.SetColor(1, 49, 134, 135);
				ESheet.Cells["G" + fila + ":J" + fila].Style.Font.Bold = true;

				ESheet.Cells["I1:" + "I" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["J1:" + "J" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["K1:" + "K" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["L1:" + "L" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["M1:" + "M" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["N1:" + "N" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["O1:" + "O" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["P1:" + "P" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
				ESheet.Cells["A:AZ"].AutoFitColumns();

				ESheet.Cells["A:AZ"].AutoFitColumns();
				/*ESheet.Column(5).Width = 25;
				ESheet.Column(4).Width = 25;
				ESheet.Column(8).Width = 25;
				ESheet.Column(11).Width = 25;
				ESheet.Column(16).Width = 25;*/

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
				reporteResultado.Id = 1;
			}
			catch (Exception ex)
			{
				reporteResultado.Filename = ex.Message;
				reporteResultado.Id = 0;
			}

			return reporteResultado;
		}
		private decimal ConversionUSD(int monedaId, decimal valorIn, DateTime fecha)
		{
			var valorOut = 0M;
			try
			{
				var mndUsd = Moneda.Get(1).First(m => m.Abbr.Trim().ToLower() == "usd");
				var rateEx = GetRateEx(monedaId, mndUsd.Id, fecha);

				valorOut = valorIn * rateEx;
			} catch ( Exception ex )
			{
				valorOut = 0M;
			}

			return valorOut;
		}

		private decimal GetRateEx(int monedaIdIn, int monedaIdOut, DateTime fecha)
		{
			var rateEx = 1M;

			
			try
			{
				var monedaInDb = Moneda.GetById(monedaIdIn);
				var monedaOutDb = Moneda.GetById(monedaIdOut);

				var clnt = new ConversionMonedaService.BPELToolsClient();
				var req = new ConversionMonedaService.processRequest();
				var res = new ConversionMonedaService.processResponse();

				/*var timeoutSpan = new TimeSpan(0, 0, 1);
				clnt.Endpoint.Binding.CloseTimeout = timeoutSpan;
				clnt.Endpoint.Binding.OpenTimeout = timeoutSpan;
				clnt.Endpoint.Binding.ReceiveTimeout = timeoutSpan;
				clnt.Endpoint.Binding.SendTimeout = timeoutSpan;*/

				req.process = new ConversionMonedaService.process();
				req.process.P_USER_CONVERSION_TYPE = "Financiero Venta";
				req.process.P_CONVERSION_DATESpecified = true;
				req.process.P_CONVERSION_DATE = fecha;
				req.process.P_FROM_CURRENCY = monedaInDb.Abbr.Trim();
				req.process.P_TO_CURRENCY = monedaOutDb.Abbr.Trim();

				res = clnt.process(req.process);

				if (res.X_CONVERSION_RATE != null && res.X_MNS_ERROR == null)
				{
					rateEx = res.X_CONVERSION_RATE.Value;
				}

				Utility.Logger.Info("Conversion Rate " + res.X_CONVERSION_RATE.Value.ToString());
			} catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
				rateEx = 1M;
			}
			
			return rateEx;
		}

	}
}