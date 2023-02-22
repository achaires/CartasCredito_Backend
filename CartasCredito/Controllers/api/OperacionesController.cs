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
	public class OperacionesController : ApiController
	{
		[HttpPost]
		public IEnumerable<CartaCredito> Filtrar([FromBody] CartasCreditoFiltrarDTO filtros)
		{
			return CartaCredito.Filtrar(filtros);
		}

		[Route("api/operaciones/cambiarestatus/{id}")]
		[HttpPost]
		public RespuestaFormato CambiarEstatus(string id, [FromBody] CartaCredito modelo)
		{
			var cc = CartaCredito.GetById(id);
			var rsp = CartaCredito.UpdateStatus(id, modelo.Estatus);

			return rsp;
		}
	}
}