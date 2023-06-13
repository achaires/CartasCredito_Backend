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
	public class EmpresasController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Empresa> Get()
		{
			return Empresa.Get(0);
		}

		// GET api/<controller>/5
		public Empresa Get(int id)
		{
			return Empresa.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Empresa modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var div = new Empresa()
				{
					DivisionId = modelo.DivisionId,
					Nombre = modelo.Nombre,
					RFC = modelo.RFC,
					Descripcion = modelo.Descripcion,
					CreadoPor = usr.Id
				};

				rsp = Empresa.Insert(div);
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
		public RespuestaFormato Put(int id, [FromBody] Empresa modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var div = Empresa.GetById(id);

				div.DivisionId = modelo.DivisionId;
				div.Nombre = modelo.Nombre;
				div.RFC = modelo.RFC;
				div.Descripcion = modelo.Descripcion;
				div.Activo = div.Activo;

				rsp = Empresa.Update(div);
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
				var modelo = Empresa.GetById(id);
				modelo.Activo = modelo.Activo ? false : true;

				rsp = Empresa.Update(modelo);
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