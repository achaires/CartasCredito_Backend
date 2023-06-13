using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	[Authorize]
	public class PagosComisionesController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<string> Get()
		{
			return new string[] { "value1", "value2" };
		}

		// GET api/<controller>/5
		public string Get(int id)
		{
			return "value";
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] PagoComisionInsertDTO model)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var nuevoPago = new Pago();
				nuevoPago.CartaCreditoId = model.CartaCreditoId;
				nuevoPago.FechaPago = model.FechaPago;
				nuevoPago.MontoPago = model.MontoPago;
				nuevoPago.ComisionId = model.ComisionId;
				nuevoPago.TipoCambio = model.TipoCambio;
				nuevoPago.MonedaId = model.MonedaId;
				nuevoPago.Estatus = 3;
				nuevoPago.Comentarios = model.Comentarios;

				rsp = Pago.ComisionInsert(nuevoPago);
			}
			catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.DataString = "";
				rsp.Errors.Add(ex.Message);
				rsp.DataInt = 0;
			}

			return rsp;
		}

		// PUT api/<controller>/5
		public void Put(int id, [FromBody] string value)
		{
		}

		// DELETE api/<controller>/5
		public void Delete(int id)
		{
		}
	}
}