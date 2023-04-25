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

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class ReportesController : ApiController
	{
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

		private string ReporteComisionesPorEstatus(int empresaId, DateTime fechaInicio, DateTime fechaFin)
		{
			var reporteNombre = "Reporte de Comisiones de Cartas de Crédito por Estatus";
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
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
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
				Sheet.Cells["J9"].Value = "Comisión Banco Corresponsal";
				Sheet.Cells["K9"].Value = "Comisión Banco Aceptación";

				Sheet.Cells["B9:K9"].Style.Font.Bold = true;

				Sheet.Cells["B1:K1"].Merge = true;
				Sheet.Cells["B2:K2"].Merge = true;
				Sheet.Cells["B4:K4"].Merge = true;


				int row = 10;

				foreach (var cc in cartasCredito)
				{
					var rowOrigin = row;

					Sheet.Cells[string.Format("B{0}", row)].Value = cc.Moneda;
					Sheet.Cells[string.Format("C{0}", row)].Value = cc.FechaVencimiento.ToString("yyyy-MM-dd");
					Sheet.Cells[string.Format("D{0}", row)].Value = cc.Empresa;
					Sheet.Cells[string.Format("E{0}", row)].Value = cc.NumCartaCredito;
					Sheet.Cells[string.Format("F{0}", row)].Value = CartaCredito.GetStatusText(cc.Estatus);
					Sheet.Cells[string.Format("G{0}", row)].Value = cc.Proveedor;
					Sheet.Cells[string.Format("H{0}", row)].Value = cc.Banco;
					Sheet.Cells[string.Format("I{0}", row)].Value = cc.MontoOriginalLC;
					Sheet.Cells[string.Format("J{0}", row)].Value = "";
					Sheet.Cells[string.Format("K{0}", row)].Value = "";

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
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
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
				Sheet.Cells["J9"].Value = "Comisión Banco Corresponsal";
				Sheet.Cells["K9"].Value = "Comisión Banco Aceptación";

				Sheet.Cells["B9:K9"].Style.Font.Bold = true;

				Sheet.Cells["B1:K1"].Merge = true;
				Sheet.Cells["B2:K2"].Merge = true;
				Sheet.Cells["B4:K4"].Merge = true;


				int row = 10;

				foreach (var cc in cartasCredito)
				{
					var rowOrigin = row;

					Sheet.Cells[string.Format("B{0}", row)].Value = cc.Moneda;
					Sheet.Cells[string.Format("C{0}", row)].Value = cc.FechaVencimiento.ToString("yyyy-MM-dd");
					Sheet.Cells[string.Format("D{0}", row)].Value = cc.Empresa;
					Sheet.Cells[string.Format("E{0}", row)].Value = cc.NumCartaCredito;
					Sheet.Cells[string.Format("F{0}", row)].Value = CartaCredito.GetStatusText(cc.Estatus);
					Sheet.Cells[string.Format("G{0}", row)].Value = cc.Proveedor;
					Sheet.Cells[string.Format("H{0}", row)].Value = cc.Banco;
					Sheet.Cells[string.Format("I{0}", row)].Value = cc.MontoOriginalLC;
					Sheet.Cells[string.Format("J{0}", row)].Value = "";
					Sheet.Cells[string.Format("K{0}", row)].Value = "";

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
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
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
				ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
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

				foreach (var empresa in empresas)
				{
					var empresaCartas = cartasCredito.FindAll(cc => cc.EmpresaId == empresa.Id);
					var empresaCartasComisiones = new List<CartaCreditoComision>();

					if ( empresaCartas.Count < 1 )
					{
						continue;
					}

					Sheet.Cells[string.Format("B{0}", row)].Value = empresa.Nombre;

					foreach ( var cartaCredito in empresaCartas)
					{
						var cartaComisiones = CartaCreditoComision.GetByCartaCreditoId(cartaCredito.Id);
						empresaCartasComisiones.AddRange(cartaComisiones);
					}

					var groupedComisiones = empresaCartasComisiones.GroupBy(ecc => ecc.ComisionId);

					foreach ( var comisionGroup in groupedComisiones)
					{
						var rowOrigin = row;
						var comisionTotal = 0M;
						foreach ( var comision in comisionGroup )
						{
							Sheet.Cells[string.Format("C{0}", row)].Value = comision.ComisionId+ " " +comision.Comision;
							Sheet.Cells[string.Format("D{0}", row)].Value = comision.NumCartaCredito;
							Sheet.Cells[string.Format("E{0}", row)].Value = comision.Moneda;
							Sheet.Cells[string.Format("F{0}", row)].Value = comision.Monto;
							Sheet.Cells[string.Format("G{0}", row)].Value = comision.MontoPagado;
							Sheet.Cells[string.Format("H{0}", row)].Value = comision.EstatusCartaText;

							row++;
							comisionTotal += comision.MontoPagado;
						}

						Sheet.Cells[string.Format("C{0}", row)].Value = "Total";
						Sheet.Cells[string.Format("G{0}", row)].Value = comisionTotal;

						var rowFinal = row - 1;
						row++;

						Sheet.Cells[string.Format("C{0}:C{1}",rowOrigin,rowFinal)].Merge = true;
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
			ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
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