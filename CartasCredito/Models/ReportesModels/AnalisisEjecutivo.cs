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
			: base (fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "P",  "Reporte Análisis Ejecutivo de Cartas de Crédito")
		{
		}

		public override Reporte Generar()
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

				var proveedoresCat = Proveedor.Get(1);

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
						/*ESheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key.Moneda;
						ESheet.Cells["I" + fila].Value = grupoMoneda.TotalMoneda;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						// Calcula y agrega fila de conversión a dólares
						fila++;
						var totalMonedaEnUsd = 0M; // ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);
						divisasList.Add(grupoMoneda.MonedaId);

						ESheet.Cells["H" + fila].Value = "Total USD";
						ESheet.Cells["I" + fila].Value = totalMonedaEnUsd;
						ESheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";*/

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

				var timeoutSpan = new TimeSpan(0, 0, 1);
				clnt.Endpoint.Binding.CloseTimeout = timeoutSpan;
				clnt.Endpoint.Binding.OpenTimeout = timeoutSpan;
				clnt.Endpoint.Binding.ReceiveTimeout = timeoutSpan;
				clnt.Endpoint.Binding.SendTimeout = timeoutSpan;

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