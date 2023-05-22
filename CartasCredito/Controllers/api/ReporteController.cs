using CartasCredito.Models.DTOs;
using CartasCredito.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace CartasCredito.Controllers.api
{
    public class ReporteController : ApiController
    {
		private ExcelPackage Ep;
		private ExcelWorksheet Sheet;
		private int InitialColumn = 0;
		private int InitialRow = 0;
		private int EndColumn = 0;
		private string Filename;
		public ReporteController()
		{
			Ep = new ExcelPackage();
			Sheet = Ep.Workbook.Worksheets.Add("Reporte");
		}

		[Route("api/reporte/generar")]
		[HttpPost]
		public RespuestaFormato Generar(SolicitudReporteDTO solicitudReporte)
		{
			var rsp = new RespuestaFormato();

			try
			{
				// rsp.DataInt = 1;
				// rsp.Flag = true;

				switch (solicitudReporte.TipoReporteId)
				{
					case 1:
						rsp = ReporteAnalisisEjecutivoCartas(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.FechaDivisas);
						break;
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

		#region ReporteAnalisisEjecutivo
		private RespuestaFormato ReporteAnalisisEjecutivoCartas(int empresaId, DateTime fechaInicio, DateTime fechaFin, DateTime fechaDivisas)
		{
			Filename = SetFileName("Análisis Ejecutivo de Cartas de Crédito");
			
			var rsp = new RespuestaFormato();
			var curDate = DateTime.Now;

			try
			{
				var fechaInicioExact = new DateTime(fechaInicio.Year, fechaInicio.Month, fechaInicio.Day, 0, 0, 0);
				var fechaFinExact = new DateTime(fechaFin.Year, fechaFin.Month, fechaFin.Day, 23, 59, 59);

				var cartasCredito = CartaCredito.Reporte(empresaId, fechaInicioExact, fechaFinExact).GroupBy(cc => cc.NumCartaCredito).Select(cg => cg.First()).OrderBy(cc => cc.FechaVencimiento);
				var catMonedas = Moneda.Get();

				rsp.DataString = Filename;
				rsp.DataInt = 1;
				rsp.Flag = true;
			} catch (Exception ex)
			{
				rsp.DataInt = -1;
				rsp.Flag = false;
				rsp.DataString = "Ocurrió un error al asignar valores";
			}

			return rsp;
		}
		#endregion

		#region CommonContent
		private string SetFileName(string repNombre)
		{
			var curDate = DateTime.Now;
			var filename = repNombre.Trim().Replace(' ', '-').ToLower() + curDate.Year.ToString() + curDate.Month.ToString() + curDate.Day.ToString() + curDate.Hour.ToString() + curDate.Minute.ToString() + curDate.Second.ToString() + ".xlsx";

			return filename;
		}
		#endregion
	}
}
