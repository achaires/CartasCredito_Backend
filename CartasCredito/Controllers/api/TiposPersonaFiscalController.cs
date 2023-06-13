using CartasCredito.Models;
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
	public class TiposPersonaFiscalController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<TipoPersonaFiscal> Get()
		{
			return TipoPersonaFiscal.Get(0);
		}

		// GET api/<controller>/5
		public TipoPersonaFiscal Get(int id)
		{
			return TipoPersonaFiscal.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] TipoPersonaFiscal modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = new TipoPersonaFiscal()
				{
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					CreadoPor = usr.Id
				};

				rsp = TipoPersonaFiscal.Insert(m);
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
		public RespuestaFormato Put(int id, [FromBody] TipoPersonaFiscal modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = TipoPersonaFiscal.GetById(id);

				m.Nombre = modelo.Nombre;
				m.Descripcion = modelo.Descripcion;
				m.Activo = m.Activo;

				rsp = TipoPersonaFiscal.Update(m);
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
				var m = TipoPersonaFiscal.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = TipoPersonaFiscal.Update(m);
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