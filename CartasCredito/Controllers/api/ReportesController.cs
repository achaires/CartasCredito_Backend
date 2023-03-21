using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class ReportesController : ApiController
	{
		[Route("api/reportes/analisiscartas")]
		[HttpPost]
		public RespuestaFormato AnalisisCartas (SolicitudReporteDTO solicitudReporteDTO)
		{
			var rsp = new RespuestaFormato ();

			try
			{
				var cartasReporte = Reporte.GetReporteAnalisisCartasCredito(solicitudReporteDTO.FechaInicio, solicitudReporteDTO.FechaFin, solicitudReporteDTO.EmpresaId);
				rsp.DataInt = solicitudReporteDTO.TipoReporteId;

				cartasReporte.ForEach(x => rsp.Content.Add(x));
			} catch (Exception ex)
			{
				rsp.DataInt = -1;
				rsp.Flag = false;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}
	}
}