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
	public class CartasCreditoController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<CartaCredito> Get(CartasCreditoFiltrarDTO model)
		{
			return CartaCredito.Get(model.FechaInicio, model.FechaFin, 1);
		}

		// GET api/<controller>/5
		public CartaCredito Get(string id)
		{
			return CartaCredito.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] CartaCredito modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				modelo.CreadoPor = usr;
				rsp = CartaCredito.Insert(modelo);
			}
			catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}

		// PUT api/<controller>/5
		public RespuestaFormato Put(string id, [FromBody] CartaCredito modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var m = CartaCredito.GetById(id);

				/*
				m.EmpresaId = modelo.EmpresaId;
				m.Nombre = modelo.Nombre;
				m.Descripcion = modelo.Descripcion;
				m.Activo = m.Activo;
				*/

				rsp = CartaCredito.Update(m);
			}
			catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}

		// DELETE api/<controller>/5
		public RespuestaFormato Delete(string id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var modelo = CartaCredito.GetById(id);
				modelo.Activo = modelo.Activo ? false : true;

				rsp = CartaCredito.Update(modelo);
			}
			catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.DataString = ex.Message;
			}

			return rsp;
		}
	}
}