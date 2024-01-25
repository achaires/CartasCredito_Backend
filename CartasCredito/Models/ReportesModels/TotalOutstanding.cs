using CartasCredito.Models.DTOs;
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
	public class TotalOutstanding : ReporteBase
	{
		public TotalOutstanding(DateTime fechaInicio, DateTime fechaFin, int empresaId, DateTime fechaDivisa) : base(fechaInicio, fechaFin, empresaId, fechaDivisa,"A", "R", "Total Outstanding")
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
				};

				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);

				
				var cartasCredito = CartaCreditoReporte.ReporteOutstanding(EmpresaId, fechaInicioExact, fechaFinExact, fechaDivisa).OrderBy(cc => cc.FechaVencimiento).ToList();
				//var cartasCredito = CartaCredito.FiltrarReporte(ccFiltro).OrderBy(cc => cc.EmpresaId);
				/*var tipoActivoGroup = cartasCredito.GroupBy(carta => carta.TipoActivoId).Select(gpoTipoActivo => new
				{
					gpoTipoActivo.Key,
					CartasDeCredito = gpoTipoActivo.ToList()
				});*/

				ESheet.Cells["A:AZ"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["A:AZ"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

				ESheet.Cells.Style.Font.Size = 10;
				ESheet.Cells["B4:H4"].Style.Font.Bold = true;

				ESheet.Cells["B1:H1"].Style.Font.Size = 22;
				ESheet.Cells["B1:H1"].Style.Font.Bold = true;
				ESheet.Cells["B1:H1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

				ESheet.Cells["B2:H2"].Style.Font.Size = 16;
				ESheet.Cells["B2:H2"].Style.Font.Bold = true;
				ESheet.Cells["B2:H2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B2"].Value = "RESUMEN DE CARTAS DE CRÉDITO PARA DIRECCIÓN";

				ESheet.Cells["B4:H4"].Style.Font.Size = 16;
				ESheet.Cells["B4:H4"].Style.Font.Bold = false;
				ESheet.Cells["B4:H4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
				ESheet.Cells["B4"].Value = "Periodo " + FechaInicio.ToString("yyyy-MM-dd") + " - " + FechaFin.ToString("yyyy-MM-dd");

				/*ESheet.Cells["B9"].Value = "Tipo Activo";
				ESheet.Cells["C9"].Value = "Empresa";
				ESheet.Cells["D9"].Value = "Monto Original";
				ESheet.Cells["E9"].Value = "Pagos Efectuados";
				ESheet.Cells["F9"].Value = "Plazo Proveedor";
				ESheet.Cells["G9"].Value = "Refinanciado";
				ESheet.Cells["H9"].Value = "No Embarcado";
				ESheet.Cells["J9"].Value = "Total Outstanding";*/

				ESheet.Cells["B9:J9"].Style.Font.Bold = true;

				ESheet.Cells["B1:H1"].Merge = true;
				ESheet.Cells["B2:H2"].Merge = true;
				ESheet.Cells["B3:H3"].Merge = true;
				ESheet.Cells["B4:H4"].Merge = true;

				System.Drawing.Color _BORDE = System.Drawing.Color.FromArgb(1, 191, 191, 191);
				System.Drawing.Color _MARCO = System.Drawing.Color.FromArgb(1, 0, 0, 0); //1,191,191,191

				ESheet.Cells["B9:H9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B9:H9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B9:H9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B9:H9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B9:H9"].Style.Border.Bottom.Color.SetColor(_BORDE);

				var imagen = Image.FromFile(HttpContext.Current.Server.MapPath(@"~/assets/GIS_BN.jpg"));
				var imagenTempFile = new FileInfo(Path.ChangeExtension(Path.GetTempFileName(), ".jpg"));
				using (var imgStream = new FileStream(imagenTempFile.FullName, FileMode.Create))
				{
					imagen.Save(imgStream, ImageFormat.Jpeg);
				}
				var sheetLogo = ESheet.Drawings.AddPicture("GIS_BN.jpg", imagenTempFile);
				sheetLogo.SetPosition(20, 320);

				/*ESheet.Cells["J9"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J9"].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["J9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J9"].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["J9"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J9"].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["J9"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J9"].Style.Border.Bottom.Color.SetColor(_BORDE);*/

				int fila = 9;


				decimal totalMonto = 0M;
				decimal totalEfectuados = 0M;
				decimal totalPlazo = 0M;
				decimal totalRefinanciado = 0M;
				decimal totalNoEmbarcado = 0M;
				decimal totalOutstanding = 0M;


				decimal ptotalMonto = 0M;
				decimal ptotalEfectuados = 0M;
				decimal ptotalPlazo = 0M;
				decimal ptotalRefinanciado = 0M;
				decimal ptotalNoEmbarcado = 0M;
				decimal ptotalOutstanding = 0M;

				List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
				tiposDeCambio = TipoDeCambio.TiposDeCambioAlDia(fechaDivisa);
				TipoDeCambio tipoDeCambio = new TipoDeCambio();
				var catMonedas = Moneda.Get();
				Moneda mndUsd = catMonedas.Where(m => m.Abbr.Trim().ToLower() == "usd").FirstOrDefault();

				foreach(var item in cartasCredito)
                {
					//----conversion de moneda-----
					/*TipoDeCambio tipoDeCambio = new TipoDeCambio();
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
					*/
					string MonedaOriginal = catMonedas.Where(m => m.Id == item.MonedaId).FirstOrDefault().Abbr.ToUpper();
					string MonedaNueva = catMonedas.Where(m => m.Id == mndUsd.Id).FirstOrDefault().Abbr.ToUpper();
					tipoDeCambio = tiposDeCambio.Where(i => i.MonedaOriginal == MonedaOriginal && i.MonedaNueva == MonedaNueva).FirstOrDefault();
					if (tipoDeCambio.MonedaOriginal == "JPY")
					{
						item.MontoOriginalLCUSD = item.MontoOriginalLC / tipoDeCambio.Conversion;
						item.PagosEfectuadosUSD = item.PagosEfectuados / tipoDeCambio.Conversion;
						item.PagosProgramadosUSD = item.PagosProgramados / tipoDeCambio.Conversion;
						item.Refinanciado = item.Refinanciado / tipoDeCambio.Conversion;
					}
					else
					{
						item.MontoOriginalLCUSD = item.MontoOriginalLC * tipoDeCambio.Conversion;
						item.PagosEfectuadosUSD = item.PagosEfectuados * tipoDeCambio.Conversion;
						item.PagosProgramadosUSD = item.PagosProgramados * tipoDeCambio.Conversion;
						item.Refinanciado = item.Refinanciado * tipoDeCambio.Conversion;
					}
					//-----------------------------

					item.MontoOriginalLC = item.MontoOriginalLCUSD;
					item.PagosEfectuados = item.PagosEfectuadosUSD;
					item.PagosProgramados = item.PagosProgramadosUSD;
					item.Refinanciado = item.Refinanciado;
				}

				//
				List<OutstandingGrupoActivo> gruposPorTipoActivo = new List<OutstandingGrupoActivo>();
				List<string> distinctActivo = cartasCredito.Select(i => i.TipoActivo).Distinct().ToList();

				foreach(string activo in distinctActivo)
                {
					OutstandingGrupoActivo gpoActivo = new OutstandingGrupoActivo();
					gpoActivo.TipoActivo = activo;

					List<CartaCreditoReporte> cartasActivo = new List<CartaCreditoReporte>();
					cartasActivo = cartasCredito.Where(j => j.TipoActivo == activo).ToList();

					List<string> distinctEmpresa = cartasActivo.Select(i => i.Empresa).Distinct().ToList();

					foreach (string empresa in distinctEmpresa)
					{
						OutstandingGrupoEmpresas gpoEmpresa = new OutstandingGrupoEmpresas();
						gpoEmpresa.Empresa = empresa;

						List<CartaCreditoReporte> cartasEmpresa = new List<CartaCreditoReporte>();
						cartasEmpresa = cartasActivo.Where(k => k.TipoActivo == activo && k.Empresa == empresa).ToList();

						gpoEmpresa.TotalEmpresa = cartasEmpresa.Sum(c => c.MontoOriginalLC);
						gpoEmpresa.TotalMontoOriginalLC = cartasEmpresa.Sum(c => c.MontoOriginalLCUSD);
						gpoEmpresa.TotalPagosEfectuados = cartasEmpresa.Sum(c => c.PagosEfectuados);
						gpoEmpresa.TotalPlazoProveedor = cartasEmpresa.Sum(c => c.PagosProgramados);
						gpoEmpresa.TotalRefinanciado = cartasEmpresa.Sum(c => c.Refinanciado);
						//
						decimal NoEmbarcado_gpo = cartasEmpresa.Sum(c => (c.MontoOriginalLC - (c.PagosEfectuados + c.PagosProgramados)));
						decimal NoEmbarcado = gpoEmpresa.TotalMontoOriginalLC - (gpoEmpresa.TotalPagosEfectuados + gpoEmpresa.TotalPlazoProveedor);

						//decimal Outstanding_gpo = cartasEmpresa.Sum(c => (c.MontoOriginalLC - c.PagosEfectuados));
						decimal Outstanding = NoEmbarcado + gpoEmpresa.TotalPlazoProveedor;
						//
						gpoEmpresa.TotalNoEmbarcado = NoEmbarcado;
						gpoEmpresa.TotalOutstanding = Outstanding;
						gpoEmpresa.CartasCredito = cartasEmpresa;

						gpoActivo.TotalMontoOriginalLC += gpoEmpresa.TotalMontoOriginalLC;
						gpoActivo.TotalPagosEfectuados += gpoEmpresa.TotalPagosEfectuados;
						gpoActivo.TotalPlazoProveedor += gpoEmpresa.TotalPlazoProveedor;
						gpoActivo.TotalRefinanciado += gpoEmpresa.TotalRefinanciado;
						gpoActivo.TotalNoEmbarcado += gpoEmpresa.TotalNoEmbarcado;
						gpoActivo.TotalOutstanding += gpoEmpresa.TotalOutstanding;

						gpoActivo.GrupoEmpresas.Add(gpoEmpresa);
					}

					gruposPorTipoActivo.Add(gpoActivo);
				}

				//


				/*var gruposPorTipoActivo = cartasCredito.GroupBy(c => c.TipoActivo)
									  .Select(g => new {
										  TipoActivo = g.Key,
										  TotalMontoOriginalLC = g.Sum(c => c.MontoOriginalLC),
										  TotalPagosEfectuados = g.Sum(c => c.PagosEfectuados),
										  TotalPlazoProveedor = g.Sum(c => c.PagosProgramados),
										  TotalRefinanciado = g.Sum(c => c.Refinanciado),
										  TotalNoEmbarcado = g.Sum(c => (c.MontoOriginalLC - c.PagosEfectuados)),
										  GrupoEmpresas = g.GroupBy(c => c.Empresa)
											.Select(ge => new {
												Empresa = ge.Key,
												TotalEmpresa = ge.Sum(c => c.MontoOriginalLC),
												CartasCredito = ge,
												TotalMontoOriginalLC = ge.Sum(c => c.MontoOriginalLC),
												TotalPagosEfectuados = ge.Sum(c => c.PagosEfectuados),
												TotalPlazoProveedor = ge.Sum(c => c.PagosProgramados),
												TotalRefinanciado = ge.Sum(c => c.Refinanciado),
												TotalNoEmbarcado = ge.Sum(c => (c.MontoOriginalLC - c.PagosEfectuados)),
											}).ToList()
									  })
									  .ToArray();*/

				foreach (var grupo in gruposPorTipoActivo)
				{

					decimal ptotalMontoA = 0M;
					decimal ptotalEfectuadosA = 0M;
					decimal ptotalPlazoA = 0M;
					decimal ptotalRefinanciadoA = 0M;
					decimal ptotalNoEmbarcadoA = 0M;
					decimal ptotalOutstandingA = 0M;

					ESheet.Cells[string.Format("B{0}", fila)].Value = grupo.TipoActivo;

					//ESheet.Cells["B9"].Value = "Tipo Activo";
					ESheet.Cells["C" + fila].Value = "Empresa";
					ESheet.Cells["D" + fila].Value = "Monto Original";
					ESheet.Cells["E" + fila].Value = "Pagos Efectuados";
					ESheet.Cells["F" + fila].Value = "Plazo Proveedor";
					ESheet.Cells["G" + fila].Value = "Refinanciado";
					ESheet.Cells["H" + fila].Value = "No Embarcado";
					ESheet.Cells["J" + fila].Value = "Total Outstanding";


					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Bottom.Color.SetColor(_BORDE);


					ESheet.Cells["J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["J" + fila].Style.Fill.BackgroundColor.SetColor(1, 0, 204, 255);


					ESheet.Cells["B" + fila + ":J" + fila].Style.Font.Bold = true;
					fila++;

					int filaInicioActivo = fila;

					foreach (var gpoEmpresa in grupo.GrupoEmpresas)
					{

						ESheet.Cells[string.Format("C{0}", fila)].Value = gpoEmpresa.Empresa;
						ESheet.Cells[string.Format("D{0}", fila)].Value = gpoEmpresa.TotalEmpresa;
						ESheet.Cells[string.Format("E{0}", fila)].Value = gpoEmpresa.TotalPagosEfectuados;
						ESheet.Cells[string.Format("F{0}", fila)].Value = gpoEmpresa.TotalPlazoProveedor;
						ESheet.Cells[string.Format("G{0}", fila)].Value = gpoEmpresa.TotalRefinanciado;
						ESheet.Cells[string.Format("H{0}", fila)].Value = gpoEmpresa.TotalNoEmbarcado;

						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Color.SetColor(_BORDE);
						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Color.SetColor(_BORDE);
						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Color.SetColor(_BORDE);
						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
						ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Color.SetColor(_BORDE);
						ESheet.Cells["C" + fila + ":H" + fila].Style.Numberformat.Format = " #,##0.00";



						fila++;
					}

					ESheet.Cells[string.Format("C{0}", fila)].Value = "Total por Activo";
					ESheet.Cells[string.Format("D{0}", fila)].Value = grupo.TotalMontoOriginalLC;
					ESheet.Cells[string.Format("E{0}", fila)].Value = grupo.TotalPagosEfectuados;
					ESheet.Cells[string.Format("F{0}", fila)].Value = grupo.TotalPlazoProveedor;
					ESheet.Cells[string.Format("G{0}", fila)].Value = grupo.TotalRefinanciado;
					ESheet.Cells[string.Format("H{0}", fila)].Value = grupo.TotalNoEmbarcado;
					ESheet.Cells[string.Format("J{0}", fila)].Value = grupo.TotalOutstanding;// grupo.TotalMontoOriginalLC - grupo.TotalPagosEfectuados;

					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Color.SetColor(_BORDE);

					ESheet.Cells["C" + fila + ":H" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Fill.BackgroundColor.SetColor(1, 206, 254, 228);

					ESheet.Cells["C" + fila + ":H" + fila].Style.Font.Bold = true;

					/*ESheet.Cells["B" + filaInicioActivo + ":B" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["B" + filaInicioActivo + ":B" + fila].Style.Fill.BackgroundColor.SetColor(1,0,204,255);*/

					ESheet.Cells["J" + filaInicioActivo + ":J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["J" + filaInicioActivo + ":J" + fila].Style.Fill.BackgroundColor.SetColor(1, 0, 204, 255);

					ESheet.Cells["J" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Left.Color.SetColor(1,0,0,0);
					ESheet.Cells["J" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Top.Color.SetColor(1, 0, 0, 0);
					ESheet.Cells["J" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Right.Color.SetColor(1, 0, 0, 0);
					ESheet.Cells["J" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Bottom.Color.SetColor(1, 0, 0, 0);

					ESheet.Cells["J" + fila].Style.Font.Bold = true;

					ESheet.Cells["C" + fila + ":H" + fila].Style.Numberformat.Format = " #,##0.00";

					ESheet.Cells["J" + fila].Style.Numberformat.Format = " #,##0.00";
					fila++;

					//porcentajes


					ptotalEfectuadosA = Math.Round(Math.Round(grupo.TotalPagosEfectuados, 4) / grupo.TotalMontoOriginalLC, 4) * 100;
					ptotalPlazoA = Math.Round(Math.Round(grupo.TotalPlazoProveedor, 4) / grupo.TotalMontoOriginalLC, 4) * 100;
					ptotalRefinanciadoA = Math.Round(Math.Round(grupo.TotalRefinanciado, 4) / grupo.TotalMontoOriginalLC, 4) * 100;
					ptotalNoEmbarcadoA = Math.Round(Math.Round(grupo.TotalNoEmbarcado, 4) / grupo.TotalMontoOriginalLC, 4) * 100;
					ptotalOutstandingA = ptotalPlazoA + ptotalNoEmbarcadoA;

					ESheet.Cells[string.Format("C{0}", fila)].Value = "";
					ESheet.Cells[string.Format("D{0}", fila)].Value = 100;
					ESheet.Cells[string.Format("E{0}", fila)].Value = ptotalEfectuadosA + "%";
					ESheet.Cells[string.Format("F{0}", fila)].Value = ptotalPlazoA + "%";
					ESheet.Cells[string.Format("G{0}", fila)].Value = ptotalRefinanciadoA + "%";
					ESheet.Cells[string.Format("H{0}", fila)].Value = ptotalNoEmbarcadoA + "%";
					ESheet.Cells[string.Format("J{0}", fila)].Value = ptotalOutstandingA + "%";// grupo.TotalMontoOriginalLC - grupo.TotalPagosEfectuados;

					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Color.SetColor(_BORDE);
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Color.SetColor(_BORDE);
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Color.SetColor(_BORDE);
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Color.SetColor(_BORDE);

					/*ESheet.Cells["C" + fila + ":H" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["C" + fila + ":H" + fila].Style.Fill.BackgroundColor.SetColor(1, 206, 254, 228);*/

					ESheet.Cells["C" + fila + ":H" + fila].Style.Font.Bold = true;

					/*ESheet.Cells["B" + filaInicioActivo + ":B" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["B" + filaInicioActivo + ":B" + fila].Style.Fill.BackgroundColor.SetColor(1,0,204,255);*/

					ESheet.Cells["J" + filaInicioActivo + ":J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
					ESheet.Cells["J" + filaInicioActivo + ":J" + fila].Style.Fill.BackgroundColor.SetColor(1, 0, 204, 255);

					ESheet.Cells["J" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Left.Color.SetColor(1, 0, 0, 0);
					ESheet.Cells["J" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Top.Color.SetColor(1, 0, 0, 0);
					ESheet.Cells["J" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Right.Color.SetColor(1, 0, 0, 0);
					ESheet.Cells["J" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
					ESheet.Cells["J" + fila].Style.Border.Bottom.Color.SetColor(1, 0, 0, 0);

					ESheet.Cells["J" + fila].Style.Font.Bold = true;

					ESheet.Cells["C" + fila + ":J" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

					/*ESheet.Cells["C" + fila + ":H" + fila].Style.Numberformat.Format = " #,##0.00 %";

					ESheet.Cells["J" + fila].Style.Numberformat.Format = " #,##0.00 %";*/
					fila++;
					//porcentajes


					totalMonto += grupo.TotalMontoOriginalLC;
					totalEfectuados += grupo.TotalPagosEfectuados;
					totalPlazo += grupo.TotalPlazoProveedor;
					totalRefinanciado += grupo.TotalRefinanciado;
					totalNoEmbarcado += grupo.TotalNoEmbarcado;
					totalOutstanding += grupo.TotalOutstanding;// grupo.TotalMontoOriginalLC;
				}


				ESheet.Cells[string.Format("C{0}", fila)].Value = "Sumas totales";
				ESheet.Cells[string.Format("D{0}", fila)].Value = totalMonto;
				ESheet.Cells[string.Format("E{0}", fila)].Value = totalEfectuados;
				ESheet.Cells[string.Format("F{0}", fila)].Value = totalPlazo;
				ESheet.Cells[string.Format("G{0}", fila)].Value = totalRefinanciado;
				ESheet.Cells[string.Format("H{0}", fila)].Value = totalNoEmbarcado;
				ESheet.Cells[string.Format("J{0}", fila)].Value = totalOutstanding;

				ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ESheet.Cells["B" + fila + ":H" + fila].Style.Border.Bottom.Color.SetColor(_BORDE);


                ESheet.Cells["B" + fila + ":H" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + fila + ":H" + fila].Style.Fill.BackgroundColor.SetColor(1, 161, 239, 237);


                ESheet.Cells["J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["J" + fila].Style.Fill.BackgroundColor.SetColor(1, 0, 204, 255);

				fila++;

				//porcentajes


				ptotalEfectuados = Math.Round(Math.Round(totalEfectuados, 4) / totalMonto, 4) * 100;
				ptotalPlazo = Math.Round(Math.Round(totalPlazo, 4) / totalMonto, 4) * 100;
				ptotalRefinanciado = Math.Round(Math.Round(totalRefinanciado, 4) / totalMonto, 4) * 100;
				ptotalNoEmbarcado = Math.Round(Math.Round(totalNoEmbarcado, 4) / totalMonto, 4) * 100;
				ptotalOutstanding = ptotalPlazo + ptotalNoEmbarcado;

				ESheet.Cells[string.Format("C{0}", fila)].Value = "";
				ESheet.Cells[string.Format("D{0}", fila)].Value = 100;
				ESheet.Cells[string.Format("E{0}", fila)].Value = ptotalEfectuados + "%";
				ESheet.Cells[string.Format("F{0}", fila)].Value = ptotalPlazo + "%";
				ESheet.Cells[string.Format("G{0}", fila)].Value = ptotalRefinanciado + "%";
				ESheet.Cells[string.Format("H{0}", fila)].Value = ptotalNoEmbarcado + "%";
				ESheet.Cells[string.Format("J{0}", fila)].Value = ptotalOutstanding + "%";// grupo.TotalMontoOriginalLC - grupo.TotalPagosEfectuados;

				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Left.Color.SetColor(_BORDE);
				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Top.Color.SetColor(_BORDE);
				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Right.Color.SetColor(_BORDE);
				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["C" + fila + ":H" + fila].Style.Border.Bottom.Color.SetColor(_BORDE);

				/*ESheet.Cells["C" + fila + ":H" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["C" + fila + ":H" + fila].Style.Fill.BackgroundColor.SetColor(1, 206, 254, 228);*/

				ESheet.Cells["C" + fila + ":H" + fila].Style.Font.Bold = true;

				/*ESheet.Cells["B" + filaInicioActivo + ":B" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["B" + filaInicioActivo + ":B" + fila].Style.Fill.BackgroundColor.SetColor(1,0,204,255);*/

				ESheet.Cells["J" + fila + ":J" + fila].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ESheet.Cells["J" + fila + ":J" + fila].Style.Fill.BackgroundColor.SetColor(1, 0, 204, 255);

				ESheet.Cells["J" + fila].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J" + fila].Style.Border.Left.Color.SetColor(1, 0, 0, 0);
				ESheet.Cells["J" + fila].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J" + fila].Style.Border.Top.Color.SetColor(1, 0, 0, 0);
				ESheet.Cells["J" + fila].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J" + fila].Style.Border.Right.Color.SetColor(1, 0, 0, 0);
				ESheet.Cells["J" + fila].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
				ESheet.Cells["J" + fila].Style.Border.Bottom.Color.SetColor(1, 0, 0, 0);

				ESheet.Cells["J" + fila].Style.Font.Bold = true;
				ESheet.Cells["C" + fila + ":J" + fila].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

				/*ESheet.Cells["C" + fila + ":H" + fila].Style.Numberformat.Format = " #,##0.00";

				ESheet.Cells["J" + fila].Style.Numberformat.Format = " #,##0.00";*/
				fila++;
				//porcentajes

				ESheet.Cells["B9" + ":H9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":H9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["B9" + ":B" + (fila - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B9" + ":B" + (fila - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["H9" + ":H" + (fila - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["H9" + ":H" + (fila - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["B" + (fila - 1) + ":H" + (fila - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["B" + (fila - 1) + ":H" + (fila - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);


				ESheet.Cells["J9"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["J9"].Style.Border.Top.Color.SetColor(_MARCO);
				ESheet.Cells["J9" + ":J" + (fila - 1)].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["J9" + ":J" + (fila - 1)].Style.Border.Left.Color.SetColor(_MARCO);
				ESheet.Cells["J9" + ":J" + (fila - 1)].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["J9" + ":J" + (fila - 1)].Style.Border.Right.Color.SetColor(_MARCO);
				ESheet.Cells["J" + (fila - 1)].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
				ESheet.Cells["J" + (fila - 1)].Style.Border.Bottom.Color.SetColor(_MARCO);


				ESheet.Cells["A:AZ"].AutoFitColumns();
				ESheet.Column(5).Width = 50;
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