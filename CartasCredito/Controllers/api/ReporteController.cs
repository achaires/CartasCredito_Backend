using CartasCredito.Models.DTOs;
using CartasCredito.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using CartasCredito.Models.ReportesModels;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class ReporteController : ApiController
    {
		[Route("api/reportes2/generar")]
		[HttpPost]
		public RespuestaFormato Generar(SolicitudReporteDTO solicitudReporte)
		{
			var rsp = new RespuestaFormato();

			if ( solicitudReporte.TipoReporteId == 1 )
			{
				var reporteRes = new AnalisisEjecutivo(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 2)
			{
				var reporteRes = new ComisionesPorTipoComision(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 3)
			{
				var reporteRes = new StandBy(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 4)
			{
				var reporteRes = new Vencimientos(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 5)
			{
				var reporteRes = new ComisionesCartasPorEstatus(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 6)
			{
				var reporteRes = new LineasDeCreditoDisponibles(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 7)
			{
				var reporteRes = new TotalOutstanding(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 8)
			{
				var reporteRes = new AnalisisCartas(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}

			return rsp;
		}
	}
}
