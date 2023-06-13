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
	public class BancosController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Banco> Get()
		{
			return Banco.Get(0);
		}

		// GET api/<controller>/5
		public Banco Get(int id)
		{
			return Banco.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Banco modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var div = new Banco()
				{
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					TotalLinea = modelo.TotalLinea,
					CreadoPor = usr.Id
				};

				rsp = Banco.Insert(div);
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
		public RespuestaFormato Put(int id, [FromBody] Banco modeloInput)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var modelo = Banco.GetById(id);

				modelo.Nombre = modeloInput.Nombre;
				modelo.Descripcion = modeloInput.Descripcion;
				modelo.TotalLinea = modeloInput.TotalLinea;
				modelo.Activo = modelo.Activo;

				rsp = Banco.Update(modelo);
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
				var modelo = Banco.GetById(id);
				modelo.Activo = modelo.Activo ? false : true;

				rsp = Banco.Update(modelo);
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