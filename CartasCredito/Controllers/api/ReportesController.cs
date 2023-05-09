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

		public void BuildHeader()
		{

		}

		private string ReporteTotalOutstanding(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var reporteNombre = "Resumen de Cartas de Crédito para Dirección";
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = BuildReporteFilename(reporteNombre);

			try
			{
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = fechaInicio,
					FechaFin = fechaFin,
				};

				var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.EmpresaId);
				var tipoActivoGroup = cartasCredito.GroupBy(carta => carta.TipoActivoId).Select(gpoTipoActivo => new
				{
					gpoTipoActivo.Key,
					CartasDeCredito = gpoTipoActivo.ToList()
				});

				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:H4"].Style.Font.Bold = true;

				Sheet.Cells["B1:H1"].Style.Font.Size = 22;
				Sheet.Cells["B1:H1"].Style.Font.Bold = true;
				Sheet.Cells["B1:H1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:H2"].Style.Font.Size = 16;
				Sheet.Cells["B2:H2"].Style.Font.Bold = true;
				Sheet.Cells["B2:H2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = reporteNombre;

				Sheet.Cells["B4:H4"].Style.Font.Size = 16;
				Sheet.Cells["B4:H4"].Style.Font.Bold = false;
				Sheet.Cells["B4:H4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Tipo Activo";
				Sheet.Cells["C9"].Value = "Empresa";
				Sheet.Cells["D9"].Value = "Monto Original";
				Sheet.Cells["E9"].Value = "Pagos Efectuados";
				Sheet.Cells["F9"].Value = "Plazo Proveedor";
				Sheet.Cells["G9"].Value = "Refinanciado";
				Sheet.Cells["H9"].Value = "No Embarcado";
				Sheet.Cells["J9"].Value = "Total Outstanding";

				Sheet.Cells["B9:H9"].Style.Font.Bold = true;

				Sheet.Cells["B1:H1"].Merge = true;
				Sheet.Cells["B2:H2"].Merge = true;
				Sheet.Cells["B4:H4"].Merge = true;


				int fila = 10;

				var gruposPorTipoActivo = cartasCredito.GroupBy(c => c.TipoActivo)
									  .Select(g => new {
										  TipoActivo = g.Key,
										  TotalMontoOriginalLC = g.Sum(c => c.MontoOriginalLC),
										  TotalPagosEfectuados = g.Sum(c => c.PagosEfectuados),
										  TotalPlazoProveedor = g.Sum(c => c.PagosProgramados),
										  TotalNoEmbarcado = g.Sum(c => (c.MontoOriginalLC - c.PagosEfectuados)),
										  GrupoEmpresas = g.GroupBy(c => c.Empresa)
											.Select(ge => new {
												Empresa = ge.Key,
												TotalEmpresa = ge.Sum(c => c.MontoOriginalLC),
												CartasCredito = ge
											}).ToList()
									  })
									  .ToArray();

				foreach (var grupo in gruposPorTipoActivo)
				{
					Sheet.Cells[string.Format("B{0}", fila)].Value = grupo.TipoActivo;

					foreach ( var gpoEmpresa in grupo.GrupoEmpresas )
					{
						Sheet.Cells[string.Format("C{0}", fila)].Value = gpoEmpresa.Empresa;
						Sheet.Cells[string.Format("D{0}", fila)].Value = gpoEmpresa.TotalEmpresa;
						Sheet.Cells[string.Format("E{0}", fila)].Value = 0;
						Sheet.Cells[string.Format("F{0}", fila)].Value = 0;
						Sheet.Cells[string.Format("G{0}", fila)].Value = 0;
						Sheet.Cells[string.Format("H{0}", fila)].Value = 0;

						fila++;
					}

					Sheet.Cells[string.Format("C{0}", fila)].Value = "Total por Activo";
					Sheet.Cells[string.Format("D{0}", fila)].Value = grupo.TotalMontoOriginalLC;
					Sheet.Cells[string.Format("E{0}", fila)].Value = grupo.TotalPagosEfectuados;
					Sheet.Cells[string.Format("F{0}", fila)].Value = grupo.TotalPlazoProveedor;
					Sheet.Cells[string.Format("G{0}", fila)].Value = 0;
					Sheet.Cells[string.Format("H{0}", fila)].Value = grupo.TotalNoEmbarcado;
					Sheet.Cells[string.Format("J{0}", fila)].Value = grupo.TotalMontoOriginalLC - grupo.TotalPagosEfectuados;

					fila++;
					
				}

				Sheet.Cells["A:AZ"].AutoFitColumns();
				Sheet.Column(5).Width = 50;
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

		private decimal GetRateEx(int monedaIdIn, int monedaIdOut, DateTime fecha)
		{
			var rateEx = 1M;

			/*
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
				req.process.P_FROM_CURRENCY = monedaInDb.Abbr;
				req.process.P_TO_CURRENCY = monedaOutDb.Abbr;

				res = clnt.process(req.process);

				if (res.X_CONVERSION_RATE != null && res.X_MNS_ERROR == null)
				{
					rateEx = res.X_CONVERSION_RATE.Value;
				}
			} catch (Exception ex)
			{
				rateEx = 1M;
			}
			*/
			
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
				

				var cartasCredito = CartaCredito.Reporte(empresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
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

				/*
				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(),".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}

				var sheetLogo = Sheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(50,50);
				*/
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

						granTotal += totalMonedaEnUsd;

						fila++;
						fila++;
					}
					fila++;
					fila++;
				}

				Sheet.Cells["H" + fila].Value = "GRAN TOTAL:";
				Sheet.Cells["I" + fila].Value = granTotal;
				Sheet.Cells["I" + fila].Style.Numberformat.Format = "$ #,##0.00";

				fila++;
				fila++;

				foreach (var grupoEmpresa in grupos)
				{
					Sheet.Cells["H" + fila].Value = grupoEmpresa.Key;
					Sheet.Cells["I" + fila].Value = Math.Round(grupoEmpresa.TotalEmpresa / granTotal * 100).ToString() + "%"; ;

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

		private string ReporteAnalisisCartas(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var reporteNombre = "Análisis de Cartas de Crédito";
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = BuildReporteFilename(reporteNombre);

			try
			{
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = fechaInicio,
					FechaFin = fechaFin,
				};

				//var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.FechaVencimiento);
				var fechaInicioExact = new DateTime(fechaInicio.Year, fechaInicio.Month, fechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(fechaFin.Year, fechaFin.Month, fechaFin.Day, 23, 59, 59);
				var cartasCredito = CartaCredito.Reporte(empresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);

				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:S4"].Style.Font.Bold = true;

				Sheet.Cells["B1:S1"].Style.Font.Size = 22;
				Sheet.Cells["B1:S1"].Style.Font.Bold = true;
				Sheet.Cells["B1:S1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:S2"].Style.Font.Size = 16;
				Sheet.Cells["B2:S2"].Style.Font.Bold = true;
				Sheet.Cells["B2:S2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = reporteNombre;

				Sheet.Cells["B4:S4"].Style.Font.Size = 16;
				Sheet.Cells["B4:S4"].Style.Font.Bold = false;
				Sheet.Cells["B4:S4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Empresa";
				Sheet.Cells["C9"].Value = "Banco";
				Sheet.Cells["D9"].Value = "Proveedor";
				Sheet.Cells["E9"].Value = "No. CC";
				Sheet.Cells["F9"].Value = "Tipo Activo";
				Sheet.Cells["G9"].Value = "Descripción";
				Sheet.Cells["H9"].Value = "Moneda";
				Sheet.Cells["I9"].Value = "Importe";
				Sheet.Cells["J9"].Value = "Pagos Efectuados";
				Sheet.Cells["K9"].Value = "Plazo Proveedor";
				Sheet.Cells["L9"].Value = "Refinanciado";
				Sheet.Cells["M9"].Value = "Saldo Insoluto por Embarcar";
				Sheet.Cells["N9"].Value = "Saldo insoluto por embarcar (sin % tolerancia) ";
				Sheet.Cells["O9"].Value = "Saldo insoluto real + Plazo Proveedor";
				Sheet.Cells["P9"].Value = "Comisiones Pagadas";
				Sheet.Cells["Q9"].Value = "Fecha Apertura";
				Sheet.Cells["R9"].Value = "Fecha Vencimiento";
				Sheet.Cells["S9"].Value = "Días plazo proveedor despúes de B/L";

				Sheet.Cells["B9:S9"].Style.Font.Bold = true;

				Sheet.Cells["B1:S1"].Merge = true;
				Sheet.Cells["B2:S2"].Merge = true;
				Sheet.Cells["B4:S4"].Merge = true;


				int fila = 10;

				var grupos = cartasCredito
					.GroupBy(carta => carta.Empresa)
					.Select(grupoEmpresa => new
					{
						grupoEmpresa.Key,
						GruposMoneda = grupoEmpresa
							.GroupBy(carta => carta.Moneda)
							.Select(grupoMoneda => new
							{
								grupoMoneda.Key,
								TotalMoneda = grupoMoneda.Sum(carta => carta.MontoOriginalLC),
								CartasDeCredito = grupoMoneda.ToList()
							}).ToList()
					}).ToList();

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

							Sheet.Cells[string.Format("C{0}", fila)].Value = carta.Banco;
							Sheet.Cells[string.Format("D{0}", fila)].Value = carta.Proveedor;
							Sheet.Cells[string.Format("E{0}", fila)].Value = carta.NumCartaCredito;
							Sheet.Cells[string.Format("F{0}", fila)].Value = carta.TipoActivo;
							Sheet.Cells[string.Format("G{0}", fila)].Value = carta.DescripcionCartaCredito;
							Sheet.Cells[string.Format("H{0}", fila)].Value = carta.Moneda;
							
							Sheet.Cells[string.Format("I{0}", fila)].Value = carta.MontoOriginalLC;
							Sheet.Cells[string.Format("I{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							Sheet.Cells[string.Format("J{0}", fila)].Value = carta.PagosEfectuados;
							Sheet.Cells[string.Format("J{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							Sheet.Cells[string.Format("K{0}", fila)].Value = carta.PagosProgramados;
							Sheet.Cells[string.Format("K{0}", fila)].Style.Numberformat.Format = "$ #,##0.00";

							Sheet.Cells[string.Format("L{0}", fila)].Value = 0;
							Sheet.Cells[string.Format("M{0}", fila)].Value = 0;
							Sheet.Cells[string.Format("N{0}", fila)].Value = 0;
							Sheet.Cells[string.Format("O{0}", fila)].Value = 0;
							Sheet.Cells[string.Format("P{0}", fila)].Value = 0;
							Sheet.Cells[string.Format("Q{0}", fila)].Value = carta.FechaApertura.ToString("dd-MM-yyyy");
							Sheet.Cells[string.Format("R{0}", fila)].Value = carta.FechaVencimiento.ToString("dd-MM-yyyy");
							Sheet.Cells[string.Format("S{0}", fila)].Value = carta.DiasPlazoProveedor;

							fila++;
						}
						Sheet.Cells["H" + fila].Value = "Total " + grupoMoneda.Key + ": ";
						Sheet.Cells["I" + fila].Value = grupoMoneda.TotalMoneda;
						fila++;
					}

					fila++;
				}


				//Sheet.Cells["A:AZ"].AutoFitColumns();
				Sheet.Column(2).Width = 20;
				Sheet.Column(3).Width = 15;
				Sheet.Column(4).Width = 20;
				Sheet.Column(5).Width = 15;
				Sheet.Column(8).Width = 15;
				Sheet.Column(9).Width = 15;
				Sheet.Column(19).Width = 10;

				Sheet.Column(2).Style.WrapText = true;
				Sheet.Column(3).Style.WrapText = true;

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

		private string ReporteLineasDeCreditoDisponibles(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var reporteNombre = "Reporte de Líneas de Cartas de Crédito Disponibles";
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = BuildReporteFilename(reporteNombre);

			try
			{
				/** Preparación de EXCEL */
				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:G4"].Style.Font.Bold = true;

				Sheet.Cells["B1:G1"].Style.Font.Size = 22;
				Sheet.Cells["B1:G1"].Style.Font.Bold = true;
				Sheet.Cells["B1:G1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:G2"].Style.Font.Size = 16;
				Sheet.Cells["B2:G2"].Style.Font.Bold = true;
				Sheet.Cells["B2:G2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = reporteNombre;

				Sheet.Cells["B4:G4"].Style.Font.Size = 16;
				Sheet.Cells["B4:G4"].Style.Font.Bold = false;
				Sheet.Cells["B4:G4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Empresa";

				// Traemos las líneas de crédito para headers
				// var bancos = Banco.Get(1);
				// var empresas = Empresa.Get(1);
				var lineasBancos = LineaDeCredito.Get(1);
				var filaInicio = 10;
				var columnaInicio = 2;

				Sheet.Cells[filaInicio, columnaInicio].Value = "Banco / Empresa";

				// Recorrer las empresas y agregarlas como filas
				int filaActual = filaInicio + 1;
				var empresas = lineasBancos.Select(lc => lc.Empresa).Distinct();
				foreach (var empresa in empresas)
				{
					Sheet.Cells[filaActual, columnaInicio].Value = empresa;
					filaActual++;
				}

				// Recorrer los bancos y agregarlos como columnas
				int columnaActual = columnaInicio + 1;
				var bancos = lineasBancos.Select(lc => lc.Banco).Distinct();
				foreach (var banco in bancos)
				{
					Sheet.Cells[filaInicio, columnaActual].Value = banco;
					columnaActual++;
				}

				// Recorrer las líneas de crédito y agregar el monto en la coordenada correspondiente
				foreach (var lineaCredito in lineasBancos)
				{
					// Encontrar la fila y columna correspondiente a la línea de crédito
					int fila = filaInicio + empresas.ToList().IndexOf(lineaCredito.Empresa) + 1;
					int columna = columnaInicio + bancos.ToList().IndexOf(lineaCredito.Banco) + 1;

					// Agregar el monto en la coordenada correspondiente
					Sheet.Cells[fila, columna].Value = lineaCredito.Monto;
				}

				// Sheet.Cells["B9:G9"].Style.Font.Bold = true;

				Sheet.Cells["B1:G1"].Merge = true;
				Sheet.Cells["B2:G2"].Merge = true;
				Sheet.Cells["B4:G4"].Merge = true;

				/*


				// Fila inicial en excel
				int fila = 10;

				// Consulta de cartas
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = fechaInicio,
					FechaFin = fechaFin,
				};

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
					.GroupJoin(comisionesDeTodasLasCartas, tipoComision => tipoComision.Id, comision => comision.ComisionId,
						(tipoComision, comisionesDeTipo) => new {
							TipoComision = tipoComision,
							Comisiones = comisionesDeTipo.GroupBy(comision => comision.EstatusCartaId)
						});

				foreach (var tipoYComisionesPorEstatus in agrupadoPorTipoComision)
				{
					Sheet.Cells[string.Format("B{0}", fila)].Value = tipoYComisionesPorEstatus.TipoComision.Nombre;

					foreach (var comisionPorEstatus in tipoYComisionesPorEstatus.Comisiones)
					{
						//Console.WriteLine($"Estatus: {comisionPorEstatus.Key}");
						Sheet.Cells[string.Format("C{0}", fila)].Value = CartaCredito.GetStatusText(comisionPorEstatus.Key);

						foreach (var comision in comisionPorEstatus)
						{
							Sheet.Cells[string.Format("D{0}", fila)].Value = comision.Moneda;
							Sheet.Cells[string.Format("E{0}", fila)].Value = comision.Monto;
							Sheet.Cells[string.Format("F{0}", fila)].Value = comision.MontoPagado;
							Sheet.Cells[string.Format("G{0}", fila)].Value = "MONTO USD";
						}

						fila++;
					}
				}
				*/

				Sheet.Cells["A:AZ"].AutoFitColumns();
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

		private string ReporteComisionesPorEstatus(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var reporteNombre = "Reporte de Comisiones de Cartas de Crédito por Estatus";
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = BuildReporteFilename(reporteNombre);

			try
			{
				/** Preparación de EXCEL */
				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:G4"].Style.Font.Bold = true;

				Sheet.Cells["B1:G1"].Style.Font.Size = 22;
				Sheet.Cells["B1:G1"].Style.Font.Bold = true;
				Sheet.Cells["B1:G1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:G2"].Style.Font.Size = 16;
				Sheet.Cells["B2:G2"].Style.Font.Bold = true;
				Sheet.Cells["B2:G2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = reporteNombre;

				Sheet.Cells["B4:G4"].Style.Font.Size = 16;
				Sheet.Cells["B4:G4"].Style.Font.Bold = false;
				Sheet.Cells["B4:G4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Comisión";
				Sheet.Cells["C9"].Value = "Estatus Carta";
				Sheet.Cells["D9"].Value = "Moneda";
				Sheet.Cells["E9"].Value = "Monto Programado";
				Sheet.Cells["F9"].Value = "Monto Pagado";
				Sheet.Cells["G9"].Value = "Monto Pagado en USD";

				Sheet.Cells["B9:G9"].Style.Font.Bold = true;

				Sheet.Cells["B1:G1"].Merge = true;
				Sheet.Cells["B2:G2"].Merge = true;
				Sheet.Cells["B4:G4"].Merge = true;


				// Fila inicial en excel
				int fila = 10;

				// Consulta de cartas
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = fechaInicio,
					FechaFin = fechaFin,
				};

				var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.FechaVencimiento);
				var comisionesDeTodasLasCartas = new List<CartaCreditoComision>();

				foreach ( var cc in cartasCredito )
				{
					var ccComisiones = CartaCreditoComision.GetByCartaCreditoId(cc.Id);

					comisionesDeTodasLasCartas.AddRange(ccComisiones);
				}

				// Consulta de tipos de comisión
				var tiposComisiones = TipoComision.Get(1);

				var agrupadoPorTipoComision = tiposComisiones
					.GroupJoin(comisionesDeTodasLasCartas, tipoComision => tipoComision.Id, comision => comision.ComisionId,
						(tipoComision, comisionesDeTipo) => new {
							TipoComision = tipoComision,
							Comisiones = comisionesDeTipo.GroupBy(comision => comision.EstatusCartaId)
						});

				foreach (var tipoYComisionesPorEstatus in agrupadoPorTipoComision)
				{
					Sheet.Cells[string.Format("B{0}", fila)].Value = tipoYComisionesPorEstatus.TipoComision.Nombre;

					foreach (var comisionPorEstatus in tipoYComisionesPorEstatus.Comisiones)
					{
						//Console.WriteLine($"Estatus: {comisionPorEstatus.Key}");
						Sheet.Cells[string.Format("C{0}", fila)].Value = CartaCredito.GetStatusText(comisionPorEstatus.Key);

						foreach (var comision in comisionPorEstatus)
						{
							Sheet.Cells[string.Format("D{0}", fila)].Value = comision.Moneda;
							Sheet.Cells[string.Format("E{0}", fila)].Value = comision.Monto;
							Sheet.Cells[string.Format("F{0}", fila)].Value = comision.MontoPagado;
							Sheet.Cells[string.Format("G{0}", fila)].Value = "MONTO USD";
						}

						fila++;
					}
				}


				Sheet.Cells["A:AZ"].AutoFitColumns();
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

		private string ReporteVencimientos(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var reporteNombre = "Reporte de Vencimientos de Cartas de Crédito";
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = BuildReporteFilename(reporteNombre);

			try
			{
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = fechaInicio,
					FechaFin = fechaFin,
				};

				var cartasCredito = CartaCredito.Filtrar(ccFiltro).OrderBy(cc => cc.FechaVencimiento);
				var monedas = Moneda.Get(1);

				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:K4"].Style.Font.Bold = true;

				Sheet.Cells["B1:K1"].Style.Font.Size = 22;
				Sheet.Cells["B1:K1"].Style.Font.Bold = true;
				Sheet.Cells["B1:K1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:K2"].Style.Font.Size = 16;
				Sheet.Cells["B2:K2"].Style.Font.Bold = true;
				Sheet.Cells["B2:K2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = reporteNombre;

				Sheet.Cells["B4:K4"].Style.Font.Size = 16;
				Sheet.Cells["B4:K4"].Style.Font.Bold = false;
				Sheet.Cells["B4:K4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Moneda";
				Sheet.Cells["C9"].Value = "Fecha Vencimiento";
				Sheet.Cells["D9"].Value = "Empresa";
				Sheet.Cells["E9"].Value = "Número de Carta";
				Sheet.Cells["F9"].Value = "Estatus Carta";
				Sheet.Cells["G9"].Value = "Proveedor";
				Sheet.Cells["H9"].Value = "Banco";
				Sheet.Cells["I9"].Value = "Monto Pago";
				Sheet.Cells["J9"].Value = "Comisión Banco Corresponsal (USD)";
				Sheet.Cells["K9"].Value = "Comisión de Aceptación (USD)";

				Sheet.Cells["B9:K9"].Style.Font.Bold = true;

				Sheet.Cells["B1:K1"].Merge = true;
				Sheet.Cells["B2:K2"].Merge = true;
				Sheet.Cells["B4:K4"].Merge = true;


				int row = 10;

				foreach (var cc in cartasCredito)
				{
					var rowOrigin = row;

					var comisionBancoCorresponsal = EncontrarComisionEnCarta("COM. BANCO CORRESPONSAL",cc.Id);
					var comisionAceptacion = EncontrarComisionEnCarta("COMISION DE ACEPTACION", cc.Id);

					Sheet.Cells[string.Format("B{0}", row)].Value = cc.Moneda;
					Sheet.Cells[string.Format("C{0}", row)].Value = cc.FechaVencimiento.ToString("dd-MM-yyyy");
					Sheet.Cells[string.Format("D{0}", row)].Value = cc.Empresa;
					Sheet.Cells[string.Format("E{0}", row)].Value = cc.NumCartaCredito;
					Sheet.Cells[string.Format("F{0}", row)].Value = CartaCredito.GetStatusText(cc.Estatus);
					Sheet.Cells[string.Format("G{0}", row)].Value = cc.Proveedor;
					Sheet.Cells[string.Format("H{0}", row)].Value = cc.Banco;
					Sheet.Cells[string.Format("I{0}", row)].Value = cc.MontoOriginalLC;
					Sheet.Cells[string.Format("I{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

					if ( comisionBancoCorresponsal != null && comisionBancoCorresponsal.Id > 0 )
					{
						Sheet.Cells[string.Format("J{0}", row)].Value = comisionBancoCorresponsal.MontoPagado;
						Sheet.Cells[string.Format("J{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					}
					
					if ( comisionAceptacion !=  null && comisionAceptacion.Id > 0 )
					{
						Sheet.Cells[string.Format("K{0}", row)].Value = comisionAceptacion.MontoPagado;
						Sheet.Cells[string.Format("K{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					}

					row++;
				}

				Sheet.Cells["A:AZ"].AutoFitColumns();
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

		private CartaCreditoComision EncontrarComisionEnCarta (string tipoComisionNombre, string cartaCreditoId)
		{
			var rsp = new CartaCreditoComision();

			try
			{
				var tiposCartaComision = TipoComision.Get(1);
				var tipoCartaComision = tiposCartaComision.Find(tc => tc.Nombre.Trim().ToLower() == tipoComisionNombre.Trim().ToLower());
				var cartaComisiones = CartaCreditoComision.GetByCartaCreditoId(cartaCreditoId);
				rsp = cartaComisiones.Find(cartacom => cartacom.ComisionId == tipoCartaComision.Id);

			} catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}


		private string ReporteCartasStandBy(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = "Reporte-CartasCreditoStandBy" + curDate.Year.ToString() + curDate.Month.ToString() + curDate.Day.ToString() + curDate.Hour.ToString() + curDate.Minute.ToString() + curDate.Second.ToString() + ".xlsx";

			try
			{
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = fechaInicio,
					FechaFin = fechaFin,
					TipoCartaId = 18
				};

				var cartasCredito = CartaCredito.Filtrar(ccFiltro);

				var empresas = Empresa.Get(1);

				if (empresaId > 0)
				{
					empresas = empresas.FindAll(e => e.Id == empresaId);
				}

				// Formar Excel
				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:O4"].Style.Font.Bold = true;

				Sheet.Cells["B1:O1"].Style.Font.Size = 22;
				Sheet.Cells["B1:O1"].Style.Font.Bold = true;
				Sheet.Cells["B1:O1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:O2"].Style.Font.Size = 16;
				Sheet.Cells["B2:O2"].Style.Font.Bold = true;
				Sheet.Cells["B2:O2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = "Reporte de Cartas de Crédito Stand By";

				Sheet.Cells["B4:O4"].Style.Font.Size = 16;
				Sheet.Cells["B4:O4"].Style.Font.Bold = false;
				Sheet.Cells["B4:O4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Número de Carta";
				Sheet.Cells["C9"].Value = "Banco Emisor";
				Sheet.Cells["D9"].Value = "Banco Confirmador";
				Sheet.Cells["E9"].Value = "Referencia Banco Confirmador";
				Sheet.Cells["F9"].Value = "Empresa";
				Sheet.Cells["G9"].Value = "Proveedor";
				Sheet.Cells["H9"].Value = "Monto Apertura";
				Sheet.Cells["I9"].Value = "Fecha Apertura";
				Sheet.Cells["J9"].Value = "Fecha Vencimiento";
				Sheet.Cells["K9"].Value = "Tipo de Crédito";
				Sheet.Cells["L9"].Value = "Tipo de Cobertura";
				Sheet.Cells["M9"].Value = "Carta A";
				Sheet.Cells["N9"].Value = "Moneda";
				Sheet.Cells["O9"].Value = "Estatus Carta";

				Sheet.Cells["B9:O9"].Style.Font.Bold = true;

				Sheet.Cells["B1:O1"].Merge = true;
				Sheet.Cells["B2:O2"].Merge = true;
				Sheet.Cells["B4:O4"].Merge = true;


				int row = 10;

				foreach (var empresa in empresas)
				{
					var empresaCartas = cartasCredito.FindAll(cc => cc.EmpresaId == empresa.Id);

					if (empresaCartas.Count < 1)
					{
						continue;
					}


					foreach (var cartaCredito in empresaCartas)
					{
						var rowOrigin = row;

						Sheet.Cells[string.Format("B{0}",row)].Value = cartaCredito.NumCartaCredito;
						Sheet.Cells[string.Format("C{0}",row)].Value = cartaCredito.Banco;
						Sheet.Cells[string.Format("D{0}",row)].Value = cartaCredito.BancoCorresponsal;
						Sheet.Cells[string.Format("E{0}",row)].Value = "Referencia Banco Confirmador";
						Sheet.Cells[string.Format("F{0}",row)].Value = cartaCredito.Empresa;
						Sheet.Cells[string.Format("G{0}",row)].Value = cartaCredito.Proveedor;
						Sheet.Cells[string.Format("H{0}",row)].Value = cartaCredito.CostoApertura;
						Sheet.Cells[string.Format("I{0}",row)].Value = cartaCredito.FechaApertura;
						Sheet.Cells[string.Format("J{0}",row)].Value = cartaCredito.FechaVencimiento;
						Sheet.Cells[string.Format("K{0}",row)].Value = cartaCredito.TipoActivo;
						Sheet.Cells[string.Format("L{0}",row)].Value = "Tipo de Cobertura";
						Sheet.Cells[string.Format("M{0}",row)].Value = "Carta A";
						Sheet.Cells[string.Format("N{0}",row)].Value = cartaCredito.Moneda;
						Sheet.Cells[string.Format("O{0}",row)].Value = CartaCredito.GetStatusText(cartaCredito.Estatus);

						row++;
					}

					row++;
				}

				Sheet.Cells["A:AZ"].AutoFitColumns();
				var path = HttpContext.Current.Server.MapPath("~/Reportes/") + filename;
				var stream = File.Create(path);
				Ep.SaveAs(stream);
				stream.Close();

				var newReporte = new Reporte()
				{
					TipoReporte = "Reporte de Cartas de Crédito Stand By",
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


		private string ReporteComisiones (int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var rsp = "";
			var curDate = DateTime.Now;
			var filename = "Reporte-ComisionesTipoComision" + curDate.Year.ToString() + curDate.Month.ToString() + curDate.Day.ToString() + curDate.Hour.ToString() + curDate.Minute.ToString() + curDate.Second.ToString() + ".xlsx";

			try
			{
				var ccFiltro = new CartasCreditoFiltrarDTO()
				{
					FechaInicio = fechaInicio,
					FechaFin = fechaFin,
				};
				var cartasCredito = CartaCredito.Filtrar(ccFiltro);

				var empresas = Empresa.Get(1);

				if ( empresaId > 0 )
				{
					empresas = empresas.FindAll(e => e.Id == empresaId);
				}

				// Formar Excel
				ExcelPackage Ep = new ExcelPackage();
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");

				Sheet.Cells.Style.Font.Size = 10;
				Sheet.Cells["B4:H4"].Style.Font.Bold = true;
				
				Sheet.Cells["B1:H1"].Style.Font.Size = 22;
				Sheet.Cells["B1:H1"].Style.Font.Bold = true;
				Sheet.Cells["B1:H1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				Sheet.Cells["B2:H2"].Style.Font.Size = 16;
				Sheet.Cells["B2:H2"].Style.Font.Bold = true;
				Sheet.Cells["B2:H2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B2"].Value = "REPORTE DE COMISIONES DE CARTAS DE CRÉDITO POR TIPO DE COMISIÓN";

				Sheet.Cells["B4:H4"].Style.Font.Size = 16;
				Sheet.Cells["B4:H4"].Style.Font.Bold = false;
				Sheet.Cells["B4:H4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				Sheet.Cells["B4"].Value = "Periodo " + fechaInicio.ToString("yyyy-MM-dd") + " - " + fechaFin.ToString("yyyy-MM-dd");

				Sheet.Cells["B9"].Value = "Empresa";
				Sheet.Cells["C9"].Value = "Comisión";
				Sheet.Cells["D9"].Value = "Número de Carta";
				Sheet.Cells["E9"].Value = "Moneda Original";
				Sheet.Cells["F9"].Value = "Monto Programado";
				Sheet.Cells["G9"].Value = "Monto Pagado";
				Sheet.Cells["H9"].Value = "Estatus Carta";

				Sheet.Cells["B9:H9"].Style.Font.Bold = true;

				Sheet.Cells["B1:H1"].Merge = true;
				Sheet.Cells["B2:H2"].Merge = true;
				Sheet.Cells["B4:H4"].Merge = true;


				int row = 10;
				var granTotalProgramado = 0M;
				var granTotalPagado = 0M;

				foreach (var empresa in empresas)
				{
					var empresaCartas = cartasCredito.FindAll(cc => cc.EmpresaId == empresa.Id);
					var empresaCartasComisiones = new List<CartaCreditoComision>();

					if ( empresaCartas.Count < 1 )
					{
						continue;
					}
					
					foreach ( var cartaCredito in empresaCartas)
					{
						var cartaComisiones = CartaCreditoComision.GetByCartaCreditoId(cartaCredito.Id);
						empresaCartasComisiones.AddRange(cartaComisiones);
					}

					var groupedComisiones = empresaCartasComisiones.GroupBy(ecc => ecc.ComisionId);

					Sheet.Cells[string.Format("B{0}", row)].Value = empresa.Nombre;
					var totalEmpresaProgramado = 0M;
					var totalEmpresaPagado = 0M;

					foreach ( var comisionGroup in groupedComisiones)
					{
						//var rowOrigin = row;
						
						foreach ( var comision in comisionGroup )
						{
							Sheet.Cells[string.Format("C{0}", row)].Value = comision.Comision;
							Sheet.Cells[string.Format("D{0}", row)].Value = comision.NumCartaCredito;
							Sheet.Cells[string.Format("E{0}", row)].Value = comision.Moneda;
							
							Sheet.Cells[string.Format("F{0}", row)].Value = comision.Monto - comision.MontoPagado;
							Sheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

							Sheet.Cells[string.Format("G{0}", row)].Value = comision.MontoPagado;
							Sheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

							Sheet.Cells[string.Format("H{0}", row)].Value = comision.EstatusCartaText;

							row++;
							
							totalEmpresaProgramado += (comision.Monto - comision.MontoPagado);
							totalEmpresaPagado += comision.MontoPagado;
						}

						//var rowFinal = row - 1;
						row++;

						//Sheet.Cells[string.Format("C{0}:C{1}",rowOrigin,rowFinal)].Merge = true;
					}

					granTotalProgramado += totalEmpresaProgramado;
					granTotalPagado += totalEmpresaPagado;

					Sheet.Cells[string.Format("C{0}", row)].Value = "Total";
					Sheet.Cells[string.Format("F{0}", row)].Value = totalEmpresaProgramado;
					Sheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					Sheet.Cells[string.Format("G{0}", row)].Value = totalEmpresaPagado;
					Sheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
					row++;

				}

				Sheet.Cells[string.Format("C{0}", row)].Value = "Gran Total";
				Sheet.Cells[string.Format("F{0}", row)].Value = granTotalProgramado;
				Sheet.Cells[string.Format("F{0}", row)].Style.Numberformat.Format = "$ #,##0.00";
				Sheet.Cells[string.Format("G{0}", row)].Value = granTotalPagado;
				Sheet.Cells[string.Format("G{0}", row)].Style.Numberformat.Format = "$ #,##0.00";

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

				Sheet.Cells["A:AZ"].AutoFitColumns();
				Sheet.Column(3).Width = 30;
				var path = HttpContext.Current.Server.MapPath("~/Reportes/") + filename;
				var stream = File.Create(path);
				Ep.SaveAs(stream);
				stream.Close();

				var newReporte = new Reporte()
				{
					TipoReporte = "Reporte Comisiones por Tipo de Comisión",
					CreadoPor = "Prueba Usuario",
					CreadoPorId = "7E7836AF-0F46-4F5C-944B-194ED9D87AEF",
					Filename = filename,
				};

				Reporte.Insert(newReporte);
			} catch ( Exception ex )
			{
				rsp = "error";
			}

			return rsp;
		}

		private void Ejemplo ()
		{
			/*
			var users = AspNetUser.Get(1);

			ExcelPackage Ep = new ExcelPackage();
			ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Reporte");
			Sheet.Cells["A1"].Value = "Name";
			Sheet.Cells["B1"].Value = "Department";
			Sheet.Cells["C1"].Value = "Address";
			Sheet.Cells["D1"].Value = "City";
			Sheet.Cells["E1"].Value = "Country";
			int row = 2;
			foreach (var item in users)
			{

				Sheet.Cells[string.Format("A{0}", row)].Value = item.UserName;
				Sheet.Cells[string.Format("B{0}", row)].Value = item.Email;
				Sheet.Cells[string.Format("C{0}", row)].Value = item.Profile.Name;
				Sheet.Cells[string.Format("D{0}", row)].Value = item.Profile.LastName;
				row++;
			}


			Sheet.Cells["A:AZ"].AutoFitColumns();
			var path = HttpContext.Current.Server.MapPath("~/Reportes/") + filename;
			var stream = File.Create(path);
			Ep.SaveAs(stream);
			stream.Close();
			*/
		}
	}
}