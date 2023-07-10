using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;
using OfficeOpenXml;
using System.Configuration;
using System.Text.Json;
using System.IO;
using System.Web;
using Microsoft.AspNet.Identity;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Text.RegularExpressions;
using OfficeOpenXml.Drawing;
using System.Drawing;
using static System.Web.Razor.Parser.SyntaxConstants;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using System.Web.Services.Description;
using System.Web.WebPages;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using System.CodeDom;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class ReportesController : ApiController
	{
		private ExcelPackage Ep;
		private ExcelWorksheet Sheet;
		private string InitialColumn;
		private string InitialRow;
		private string EndColumn;
		private string Filename;

		public ReportesController() 
		{
			Ep = new ExcelPackage();
			Sheet = Ep.Workbook.Worksheets.Add("Reporte");
		}

		[Route("api/reportes/historial")]
		[HttpPost]
		public List<Reporte> Generar()
		{
			var rsp = new List<Reporte>();

			try
			{
				rsp = Reporte.Get();
			}
			catch (Exception ex)
			{
				//
			}

			return rsp;
		}

		[Route("api/reportes/generar")]
		[HttpPost]
		public RespuestaFormato Generar (SolicitudReporteDTO solicitudReporte)
		{
			var rsp = new RespuestaFormato ();

			try
			{
				rsp.DataInt = 1;
				rsp.Flag = true;
				
				switch ( solicitudReporte.TipoReporteId )
				{
					case 1:
						rsp.DataString = ReporteAnalisisEjecutivoCartas(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.FechaDivisas);
						break;
					/*
					case 2:
						rsp.DataString = ReporteComisiones(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);
						break;
					case 3:
						rsp.DataString = ReporteCartasStandBy(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);
						break;
					case 4:
						rsp.DataString = ReporteVencimientos(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);
						break;
					case 5:
						rsp.DataString = ReporteComisionesPorEstatus(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);
						break;
					case 6:
						rsp.DataString = ReporteLineasDeCreditoDisponibles(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);
						break;
					case 7:
						rsp.DataString = ReporteTotalOutstanding(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);
						break;
					case 8:
						rsp.DataString = ReporteAnalisisCartas(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);
						break;
					*/
				}

				if ( rsp.DataString == "error" )
				{
					rsp.DataInt = -1;
					rsp.Flag = false;
					rsp.DataString = "Ocurrió un error al asignar valores";
				}
				
			}
			catch (Exception ex)
			{
				rsp.DataInt = -1;
				rsp.Flag = false;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}

		private string BuildReporteFilename(string repNombre)
		{
			var curDate = DateTime.Now;
			var filename = repNombre.Trim().Replace(' ','-').ToLower() + curDate.Year.ToString() + curDate.Month.ToString() + curDate.Day.ToString() + curDate.Hour.ToString() + curDate.Minute.ToString() + curDate.Second.ToString() + ".xlsx";

			return filename;
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

		private string ReporteAnalisisEjecutivoCartas(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaDivisa)
		{
			var reporteNombre = "Análisis Ejecutivo de Cartas de Crédito";
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = BuildReporteFilename(reporteNombre);

			try
			{

				var fechaInicioExact = new DateTime(fechaInicio.Year, fechaInicio.Month, fechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(fechaFin.Year, fechaFin.Month, fechaFin.Day, 23, 59, 59);
				
				var cartasCredito = CartaCredito.Reporte(empresaId, fechaInicioExact, fechaFinExact).Where(cc => cc.TipoCarta == "Comercial").GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
				var catMonedas = Moneda.Get();

				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:P4"].Style.Font.Bold = true;

				Sheet.Cells["B1:P1"].Style.Font.Size = 22;
				Sheet.Cells["B1:P1"].Style.Font.Bold = true;
				Sheet.Cells["B1:P1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:P2"].Style.Font.Size = 16;
				Sheet.Cells["B2:P2"].Style.Font.Bold = true;
				Sheet.Cells["B2:P2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = reporteNombre;

				Sheet.Cells["B4:P4"].Style.Font.Size = 16;
				Sheet.Cells["B4:P4"].Style.Font.Bold = false;
				Sheet.Cells["B4:P4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Empresa";
				Sheet.Cells["C9"].Value = "Banco";
				Sheet.Cells["D9"].Value = "Proveedor";
				Sheet.Cells["E9"].Value = "Producto";
				Sheet.Cells["F9"].Value = "País";
				Sheet.Cells["G9"].Value = "Descripción";
				Sheet.Cells["H9"].Value = "Moneda";
				Sheet.Cells["I9"].Value = "Importe Total";
				Sheet.Cells["J9"].Value = "Importe < 50,0000";
				Sheet.Cells["K9"].Value = "50,000 < Importe > 300,000";
				Sheet.Cells["L9"].Value = "Importe > 300,000";
				Sheet.Cells["M9"].Value = "1 Periodo";
				Sheet.Cells["N9"].Value = "2 Periodo";
				Sheet.Cells["O9"].Value = "3 Periodos";
				Sheet.Cells["P9"].Value = "Días plazo proveedor después de B/L";

				Sheet.Cells["B9:P9"].Style.Font.Bold = true;

				Sheet.Cells["B1:P1"].Merge = true;
				Sheet.Cells["B2:P2"].Merge = true;
				Sheet.Cells["B4:P4"].Merge = true;

				
				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(),".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}

				var sheetLogo = Sheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(50,50);
				
				int fila = 10;

				var proveedoresCat = Proveedor.Get(1);

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
								TotalMoneda = grupoMoneda.Sum(carta => carta.MontoOriginalLC),
								CartasDeCredito = grupoMoneda.ToList()
							}).ToList()
					}).ToList();

				var granTotal = 0M;

				var divisasList = new List<int>();
				
				foreach (var grupoEmpresa in grupos)
				{
					Sheet.Cells[string.Format("B{0}", fila)].Value = grupoEmpresa.Key;

					foreach (var grupoMoneda in grupoEmpresa.GruposMoneda)
					{
						foreach (var carta in grupoMoneda.CartasDeCredito)
						{
							TimeSpan diferencia = carta.FechaVencimiento.Subtract(carta.FechaApertura);
							int cantidadDias = (int)diferencia.TotalDays;
							decimal periodosDecimal = Convert.ToDecimal(cantidadDias) / 90M;
							var periodos = Math.Ceiling(periodosDecimal);
							var proveedorObj = proveedoresCat.First(pv => pv.Id == carta.ProveedorId);

							Sheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
							Sheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
							Sheet.Cells[string.Format("E{0}", fila)].Value = carta.DescripcionMercancia;
							Sheet.Cells[string.Format("F{0}", fila)].Value = proveedorObj.Pais;
							Sheet.Cells[string.Format("G{0}", fila)].Value = carta.TipoActivo;
							Sheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;
							
							Sheet.Cells[string.Format("I{0}", fila)].Value = carta.MontoOriginalLC;
							Sheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							Sheet.Cells[string.Format("J{0}", fila)].Value = carta.MontoOriginalLC < 50000 ? "Sí" : "No";
							Sheet.Cells[string.Format("K{0}", fila)].Value = carta.MontoOriginalLC > 50000 && carta.MontoOriginalLC < 300000 ? "Sí" : "No";
							Sheet.Cells[string.Format("L{0}", fila)].Value = carta.MontoOriginalLC > 300000 ? "Sí" : "No";
							Sheet.Cells[string.Format("M{0}", fila)].Value = periodos <= 1 ? "Sí" : "No";
							Sheet.Cells[string.Format("N{0}", fila)].Value = periodos == 2 ? "Sí" : "No";
							Sheet.Cells[string.Format("O{0}", fila)].Value = periodos > 2 ? "Sí" : "No";
							Sheet.Cells[string.Format("P{0}", fila)].Value = carta.DiasPlazoProveedor;

							fila++;
						}
						Sheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key;
						Sheet.Cells["I" + fila].Value = grupoMoneda.TotalMoneda;
						Sheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						// Calcula y agrega fila de conversión a dólares
						fila++;
						var totalMonedaEnUsd = ConversionUSD(grupoMoneda.MonedaId, grupoMoneda.TotalMoneda, fechaDivisa);
						divisasList.Add(grupoMoneda.MonedaId);

						Sheet.Cells["H" + fila].Value = "Total USD";
						Sheet.Cells["I" + fila].Value = totalMonedaEnUsd;
						Sheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

						fila++;
						fila++;
					}

					granTotal += grupoEmpresa.TotalEmpresa;

					fila++;
					fila++;
				}

				Sheet.Cells["H" + fila].Value = "GRAN TOTAL:";
				Sheet.Cells["I" + fila].Value = granTotal;
				Sheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

				fila++;
				fila++;

				var sumaProcentajes = 0M;
				
				foreach (var grupoEmpresa in grupos)
				{
					var porcentajeEmpresa = Math.Round(Math.Round(grupoEmpresa.TotalEmpresa, 4) / granTotal,4) * 100;
		
					sumaProcentajes += porcentajeEmpresa;

					Sheet.Cells["H" + fila].Value = grupoEmpresa.Key;
					Sheet.Cells["I" + fila].Value =  porcentajeEmpresa + "%";

					fila++;
				}


				Sheet.Cells["A:AZ"].AutoFitColumns();
				
				Sheet.Column(5).Width = 25;
				Sheet.Column(4).Width = 25;
				Sheet.Column(8).Width = 25;
				Sheet.Column(11).Width = 25;
				Sheet.Column(16).Width = 25;

				var path = HttpContext.Current.Server.MapPath("~/Reportes/") + filename;
				var stream = File.Create(path);
				Ep.SaveAs(stream);
				stream.Close();

				var newReporte = new Reporte()
				{
					TipoReporte = reporteNombre,
					CreadoPor = "Prueba Usuario",
					CreadoPorId = "7E7836AF-0F46-4F5C-944B-194ED9D87AEF",
					Filename = filename,
				};

				Reporte.Insert(newReporte);
			}
			catch (Exception ex)
			{
				rsp = "error";
			}

			return rsp;
		}

	}
}