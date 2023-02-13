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
	public class ContactosController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Contacto> Get()
		{
			return Contacto.Get(0);
		}

		// GET api/<controller>/5
		public Contacto Get(int id)
		{
			return Contacto.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Contacto modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var m = new Contacto()
				{
					ModelId = modelo.ModelId,
					ModelNombre = modelo.ModelNombre,
					Nombre = modelo.Nombre,
					ApellidoPaterno = modelo.ApellidoPaterno,
					ApellidoMaterno = modelo.ApellidoMaterno,
					Telefono = modelo.Telefono,
					Email = modelo.Email,
					Celular = modelo.Celular,
					Fax = modelo.Fax,
					Descripcion = modelo.Descripcion,
				};

				rsp = Contacto.Insert(m);
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
		public RespuestaFormato Put(int id, [FromBody] Contacto modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var m = Contacto.GetById(id);

				m.Nombre = modelo.Nombre;
				m.ApellidoPaterno = modelo.ApellidoPaterno;
				m.ApellidoMaterno = modelo.ApellidoMaterno;
				m.Email = modelo.Email;
				m.Telefono = modelo.Telefono;
				m.Celular = modelo.Celular;
				m.Fax = modelo.Fax;
				m.Descripcion = modelo.Descripcion;

				rsp = Contacto.Update(m);
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
				var m = Contacto.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = Contacto.Update(m);
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