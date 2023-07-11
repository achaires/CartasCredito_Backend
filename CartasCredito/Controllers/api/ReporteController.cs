using CartasCredito.Models.DTOs;
using CartasCredito.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CartasCredito.Interfaces;
using CartasCredito.Models.ReportesModels;

namespace CartasCredito.Controllers.api
{
    public class ReporteController : ApiController
    {
		[Route("api/reportes/generar")]
		[HttpPost]
		public RespuestaFormato Generar(SolicitudReporteDTO solicitudReporte)
		{
			var rsp = new RespuestaFormato();
			var reporteResultado = new Reporte();

			var reporteGeneradores = new List<IGeneradorReporte>()
			{
				new AnalisisEjecutivo(),
				new StandBy(),
				new Vencimientos(),
				new ComisionesCartasPorEstatus(),
				new LineasDeCreditoDisponibles(),
				new TotalOutstanding(),
				new AnalisisCartas(),
			};

			var reporteGenerador = reporteGeneradores.FirstOrDefault(gen => gen.Verificar(solicitudReporte.TipoReporteId));

			if ( reporteGenerador == null )
			{
				rsp.Flag = false;
				rsp.Description = "Error";
				rsp.Errors.Add("ID de reporte no registrado");
			}

			try
			{
				reporteResultado = reporteGenerador.Generar(solicitudReporte.EmpresaId, solicitudReporte.FechaInicio, solicitudReporte.FechaFin);

				rsp.Flag = true;
				rsp.Content.Add(reporteResultado);
			} catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.Description = "Error";
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}
	}
}
