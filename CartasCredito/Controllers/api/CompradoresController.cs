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
	public class CompradoresController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Comprador> Get()
		{
			return Comprador.Get(0);
		}

		// GET api/<controller>/5
		public Comprador Get(int id)
		{
			return Comprador.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Comprador modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = new Comprador()
				{
					EmpresaId = modelo.EmpresaId,
					TipoPersonaFiscalId = modelo.TipoPersonaFiscalId,
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					CreadoPor = usr.Id
				};

				rsp = Comprador.Insert(m);
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
		public RespuestaFormato Put(int id, [FromBody] Comprador modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = Comprador.GetById(id);

				m.Nombre = modelo.Nombre;
				m.EmpresaId = modelo.EmpresaId;
				m.TipoPersonaFiscalId = modelo.TipoPersonaFiscalId;
				m.Descripcion = modelo.Descripcion;
				m.Activo = m.Activo;

				rsp = Comprador.Update(m);
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
				var m = Comprador.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = Comprador.Update(m);
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