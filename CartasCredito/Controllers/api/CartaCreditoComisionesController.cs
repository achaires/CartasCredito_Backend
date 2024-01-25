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
	public class CartaCreditoComisionesController : ApiController
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
		public RespuestaFormato Post([FromBody] CartaCreditoComisionInsertDTO model)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{

				var cc = CartaCredito.GetById(model.CartaCreditoId);
				var nuevoReg = new CartaCreditoComision();
				nuevoReg.CartaCreditoId = model.CartaCreditoId;
				nuevoReg.ComisionId = model.ComisionId;
				nuevoReg.FechaCargo = model.FechaCargo;
				nuevoReg.MonedaId = cc.MonedaId;
				nuevoReg.Monto = model.Monto;
				nuevoReg.PagoId = model.NumReferencia;
				nuevoReg.CreadoPor = usr.Id;

				rsp = CartaCreditoComision.Insert(nuevoReg);

				var bm = new BitacoraMovimiento();
				bm.Titulo = "Comisión";
				bm.CreadoPorId = usr.Id;
				bm.Descripcion = "Se ha agregado una nueva comisión a la carta de crédito";
				bm.CartaCreditoId = model.CartaCreditoId;
				bm.ModeloNombre = "CartaCreditoComision";
				bm.ModeloId = rsp.DataInt;

				BitacoraMovimiento.Insert(bm);
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