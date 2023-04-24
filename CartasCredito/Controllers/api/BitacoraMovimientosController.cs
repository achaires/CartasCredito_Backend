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
	public class BitacoraMovimientosController : ApiController
	{
		[HttpPost]
		public IEnumerable<BitacoraMovimiento> Filtrar([FromBody] BitacoraMovimientosFiltrarDTO filtros)
		{
			//var filtros = new BitacoraMovimientosFiltrarDTO();
			var dateStartOrigin = new DateTime(1970,1,1,0,0,0);
			var dateEndOrigin = new DateTime(1970,1,1,0,0,0);
			var dateStart = dateStartOrigin.AddSeconds(filtros.DateStart);
			var dateEnd = dateEndOrigin.AddSeconds(filtros.DateEnd);

			return BitacoraMovimiento.Get(dateStart, dateEnd);
		}
	}
}