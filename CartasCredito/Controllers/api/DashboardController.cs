using CartasCredito.Models;
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
	public class DashboardController : ApiController
	{
		[HttpGet]
		[Route("api/dashboard/pagosprogramados")]
		public IEnumerable<Pago> PagosProgramados ()
		{
			List<Pago> rsp;

			try
			{
				rsp = Pago.GetProgramados();
			} catch (Exception ex)
			{
				rsp = new List<Pago>();
			}

			return rsp;
		}

		[HttpGet]
		[Route("api/dashboard/pagosvencidos")]
		public IEnumerable<Pago> PagosVencidos()
		{
			var rsp = new List<Pago>();
			try
			{
				rsp = Pago.GetVencidos();
			}
			catch (Exception ex)
			{
				rsp = new List<Pago>();
			}

			return rsp;
		}
	}
}