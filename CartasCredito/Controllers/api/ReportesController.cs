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
		[Route("api/reportes/historial")]
		[HttpGet]
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
		public RespuestaFormato Generar (SolicitudReporteDTO solicitudReporteDTO)
		{
			var rsp = new RespuestaFormato ();

			try
			{
				
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