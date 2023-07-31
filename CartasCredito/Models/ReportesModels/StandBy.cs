using CartasCredito.Models.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class StandBy : ReporteBase
	{
		public StandBy(DateTime fechaInicio, DateTime fechaFin, int empresaId) : base(fechaInicio, fechaFin, empresaId, "A","O", "Reporte de Cartas de Crédito Stand By")
		{
		}

		public override Reporte Generar()
		{
			var reporteRsp = new Reporte();
			try
			{
				var fechaInicioExact = new DateTime(FechaInicio.Year, FechaInicio.Month, FechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(FechaFin.Year, FechaFin.Month, FechaFin.Day, 23, 59, 59);
				var cartasCredito = CartaCredito.Reporte(EmpresaId, fechaInicioExact, fechaFinExact).Where(cc => cc.TipoCarta == "StandBy").GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento).ToList();

				var empresas = Empresa.Get(1);

				if (EmpresaId > 0)
				{
					empresas = empresas.FindAll(e => e.Id == EmpresaId);
				}

				ESheet.Cells["B9"].Value = "Número de Carta";
				ESheet.Cells["C9"].Value = "Banco Emisor";
				ESheet.Cells["D9"].Value = "Banco Confirmador";
				ESheet.Cells["E9"].Value = "Referencia Banco Confirmador";
				ESheet.Cells["F9"].Value = "Empresa";
				ESheet.Cells["G9"].Value = "Proveedor";
				ESheet.Cells["H9"].Value = "Monto Apertura";
				ESheet.Cells["I9"].Value = "Fecha Apertura";
				ESheet.Cells["J9"].Value = "Fecha Vencimiento";
				ESheet.Cells["K9"].Value = "Tipo de Crédito";
				ESheet.Cells["L9"].Value = "Tipo de Cobertura";
				ESheet.Cells["M9"].Value = "Carta A";
				ESheet.Cells["N9"].Value = "Moneda";
				ESheet.Cells["O9"].Value = "Estatus Carta";

				ESheet.Cells["B9:O9"].Style.Font.Bold = true;

				ESheet.Cells["B1:O1"].Merge = true;
				ESheet.Cells["B2:O2"].Merge = true;
				ESheet.Cells["B4:O4"].Merge = true;


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

						ESheet.Cells[string.Format("B{0}", row)].Value = cartaCredito.NumCartaCredito;
						ESheet.Cells[string.Format("C{0}", row)].Value = cartaCredito.Banco;
						ESheet.Cells[string.Format("D{0}", row)].Value = cartaCredito.BancoCorresponsal;
						ESheet.Cells[string.Format("E{0}", row)].Value = "Referencia Banco Confirmador";
						ESheet.Cells[string.Format("F{0}", row)].Value = cartaCredito.Empresa;
						ESheet.Cells[string.Format("G{0}", row)].Value = cartaCredito.Proveedor;
						ESheet.Cells[string.Format("H{0}", row)].Value = cartaCredito.CostoApertura;
						ESheet.Cells[string.Format("I{0}", row)].Value = cartaCredito.FechaApertura;
						ESheet.Cells[string.Format("J{0}", row)].Value = cartaCredito.FechaVencimiento;
						ESheet.Cells[string.Format("K{0}", row)].Value = cartaCredito.TipoActivo;
						ESheet.Cells[string.Format("L{0}", row)].Value = "Tipo de Cobertura";
						ESheet.Cells[string.Format("M{0}", row)].Value = "Carta A";
						ESheet.Cells[string.Format("N{0}", row)].Value = cartaCredito.Moneda;
						ESheet.Cells[string.Format("O{0}", row)].Value = CartaCredito.GetStatusText(cartaCredito.Estatus);

						row++;
					}

					row++;
				}

				ESheet.Cells["A:AZ"].AutoFitColumns();
				var path = HttpContext.Current.Server.MapPath("~/Reportes/") + ReporteFilename;
				var stream = File.Create(path);
				EPackage.SaveAs(stream);
				stream.Close();

				var newReporte = new Reporte()
				{
					TipoReporte = "Reporte de Cartas de Crédito Stand By",
					CreadoPor = "Prueba Usuario",
					CreadoPorId = "7E7836AF-0F46-4F5C-944B-194ED9D87AEF",
					Filename = ReporteFilename,
				};

				Reporte.Insert(newReporte);
			} catch (Exception ex)
			{
				reporteRsp.Filename = ex.Message;
				reporteRsp.Id = 0;
			}

			return reporteRsp;
		}
	}
}