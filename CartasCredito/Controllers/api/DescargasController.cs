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
	public class DescargasController : ApiController
	{
		[Route("api/descargas/swift/{id}")]
		[HttpGet]
		public RespuestaFormato Swift (string id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				rsp.DataString = CartaCredito.GetSwiftByCartaCreditoId(id);
				rsp.Flag = true;
			}
			catch (Exception ex)
			{
				rsp.Flag = false;
			}

			return rsp;
		}
	}
}