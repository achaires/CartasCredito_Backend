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
	public class ProyectosController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Proyecto> Get()
		{
			return Proyecto.Get(0);
		}

		// GET api/<controller>/5
		public Proyecto Get(int id)
		{
			return Proyecto.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Proyecto modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = new Proyecto()
				{
					EmpresaId = modelo.EmpresaId,
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					FechaApertura = modelo.FechaApertura,
					FechaCierre = modelo.FechaCierre,
					CreadoPor = usr.Id
				};

				rsp = Proyecto.Insert(m);
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
		public RespuestaFormato Put(int id, [FromBody] Proyecto modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = Proyecto.GetById(id);

				m.EmpresaId = modelo.EmpresaId;
				m.Nombre = modelo.Nombre;
				m.Descripcion = modelo.Descripcion;
				m.FechaApertura = modelo.FechaApertura;
				m.FechaCierre = modelo.FechaCierre;
				m.Activo = m.Activo;

				rsp = Proyecto.Update(m);
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
				var m = Proyecto.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = Proyecto.Update(m);
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