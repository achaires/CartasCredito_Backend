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
			var hoy = DateTime.Now;
			var rsp = new RespuestaFormato();
			var resultReporte = new Reporte();
			int diasEnMes = 0;

			//---
			if (solicitudReporte.FechaDivisas.Year == 1969 || solicitudReporte.FechaDivisas.Year == 1)
			{
				solicitudReporte.FechaDivisas = new DateTime(hoy.Year, hoy.Month, hoy.Day, 0, 0, 0);
			}

			var tiposCambio = TipoDeCambio.TiposDeCambioAlDia(solicitudReporte.FechaDivisas);
			if(tiposCambio.Count == 0)
            {
				rsp.Flag = false;
				rsp.Description = "Error, no se pudo obtener el tipo de cambio, intentalo mas tarde.";
				rsp.Errors.Add("Error, no se pudo obtener la conversion de tipo de cambio, intentalo mas tarde.");
				return rsp;
            }

			//---

			if ( solicitudReporte.TipoReporteId == 1 )
			{
				if(solicitudReporte.FechaInicio.Year == 1969)
                {
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, 1, 0, 0, 0);
					solicitudReporte.FechaFin = hoy;
                }
				var reporteRes = new AnalisisEjecutivo(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				resultReporte = reporteRes.Generar();
			}


			if (solicitudReporte.TipoReporteId == 8)
			{
				if (solicitudReporte.FechaInicio.Year == 1969)
				{
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, 1, 0, 0, 0);
					solicitudReporte.FechaFin = hoy;
				}
				var reporteRes = new AnalisisCartas(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				resultReporte = reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 2)
			{
				if (solicitudReporte.FechaInicio.Year == 1969)
				{
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, 1, 0, 0, 0);
					diasEnMes = DateTime.DaysInMonth(hoy.Year, hoy.Month);
					solicitudReporte.FechaFin = new DateTime(hoy.Year, hoy.Month, diasEnMes, 0, 0, 0);
				}
				var reporteRes = new ComisionesPorTipoComision(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				resultReporte = reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 3)
			{
				if (solicitudReporte.FechaVencimientoInicio.Year == 1969)
				{
					solicitudReporte.FechaVencimientoInicio = new DateTime(hoy.Year, hoy.Month, hoy.Day, 0, 0, 0);
					solicitudReporte.FechaFin = new DateTime(2099, 1, 1, 0, 0, 0);
				}
				var reporteRes = new StandBy(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.FechaVencimientoInicio = solicitudReporte.FechaVencimientoInicio;
				reporteRes.FechaVencimientoFin = solicitudReporte.FechaVencimientoFin;
				resultReporte = reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 4)
			{
				if (solicitudReporte.FechaVencimientoInicio.Year == 1969)
				{
					solicitudReporte.FechaVencimientoInicio = new DateTime(hoy.Year, hoy.Month, hoy.Day, 0, 0, 0);
					solicitudReporte.FechaFin = new DateTime(2099, 1, 1, 0, 0, 0);
				}
				var reporteRes = new VencimientosPagos(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.FechaVencimientoInicio = solicitudReporte.FechaVencimientoInicio;
				reporteRes.FechaVencimientoFin = solicitudReporte.FechaVencimientoFin;
				resultReporte = reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 5)
			{
				if (solicitudReporte.FechaInicio.Year == 1969)
				{
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, 1, 0, 0, 0);
					diasEnMes = DateTime.DaysInMonth(hoy.Year, hoy.Month);
					solicitudReporte.FechaFin = new DateTime(hoy.Year, hoy.Month, diasEnMes, 0, 0, 0);
				}
				var reporteRes = new ComisionesCartasPorEstatus(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				resultReporte = reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 6)
			{
				if (solicitudReporte.FechaInicio.Year == 1969)
				{
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, 1, 0, 0, 0);
					diasEnMes = DateTime.DaysInMonth(hoy.Year, hoy.Month);
					solicitudReporte.FechaFin = new DateTime(hoy.Year, hoy.Month, diasEnMes, 0, 0, 0);
				}
				var reporteRes = new LineasDeCreditoDisponibles(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				resultReporte = reporteRes.Generar();
			}

			if (solicitudReporte.TipoReporteId == 7)
			{
				if (solicitudReporte.FechaInicio.Year == 1969)
				{
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, 1, 0, 0, 0);
					solicitudReporte.FechaFin = hoy;
				}
				var reporteRes = new TotalOutstanding(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.Generar();
			}


			if (solicitudReporte.TipoReporteId == 9)
			{
				if (solicitudReporte.FechaInicio.Year == 1969)
				{
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, hoy.Day, 0, 0, 0);
					solicitudReporte.FechaFin = new DateTime(2099, 1, 1, 0, 0, 0);
				}
				var reporteRes = new ComisionesCartas(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.FechaVencimientoInicio = solicitudReporte.FechaVencimientoInicio;
				reporteRes.FechaVencimientoFin = solicitudReporte.FechaVencimientoFin;
				resultReporte = reporteRes.Generar();
			}


			if (solicitudReporte.TipoReporteId == 10)
			{
				if (solicitudReporte.FechaInicio.Year == 1969)
				{
					solicitudReporte.FechaInicio = new DateTime(hoy.Year, hoy.Month, hoy.Day, 0, 0, 0);
					solicitudReporte.FechaFin = new DateTime(2099, 1, 1, 0, 0, 0);
				}
				var reporteRes = new ResumenCartasCredito(solicitudReporte.FechaInicio, solicitudReporte.FechaFin, solicitudReporte.EmpresaId, solicitudReporte.FechaDivisas);
				reporteRes.FechaVencimientoInicio = solicitudReporte.FechaVencimientoInicio;
				reporteRes.FechaVencimientoFin = solicitudReporte.FechaVencimientoFin;
				resultReporte = reporteRes.Generar();
			}

			if(resultReporte.Id == 0)
			{
				rsp.Flag = false;
				rsp.Description = "Error";
				rsp.Errors.Add(resultReporte.Filename);
			}
			return rsp;
		}
	}
}
