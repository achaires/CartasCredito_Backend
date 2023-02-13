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
	public class TiposCoberturaController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<TipoCobertura> Get()
		{
			return TipoCobertura.Get(0);
		}

		// GET api/<controller>/5
		public TipoCobertura Get(int id)
		{
			return TipoCobertura.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] TipoCobertura modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var m = new TipoCobertura()
				{
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					CreadoPor = usr
				};

				rsp = TipoCobertura.Insert(m);
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
		public RespuestaFormato Put(int id, [FromBody] TipoCobertura modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var m = TipoCobertura.GetById(id);

				m.Nombre = modelo.Nombre;
				m.Descripcion = modelo.Descripcion;
				m.Activo = m.Activo;

				rsp = TipoCobertura.Update(m);
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
		public RespuestaFormato Delete(int id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var m = TipoCobertura.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = TipoCobertura.Update(m);
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