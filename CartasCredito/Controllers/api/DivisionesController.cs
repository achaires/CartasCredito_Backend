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
	public class DivisionesController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Division> Get()
		{
			return Division.Get(0);
		}

		// GET api/<controller>/5
		public Division Get(int id)
		{
			return Division.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Division modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var div = new Division()
				{
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					CreadoPor = usr
				};

				rsp = Division.Insert(div);
			} catch(Exception ex)
			{
				rsp.Flag = false;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}

		// PUT api/<controller>/5
		public RespuestaFormato Put(int id, [FromBody] Division modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var div = Division.GetById(id);

				div.Nombre = modelo.Nombre;
				div.Descripcion = modelo.Descripcion;
				div.Activo = div.Activo;
				
				rsp = Division.Update(div);
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
				var div = Division.GetById(id);
				div.Activo = div.Activo ? false : true;

				rsp = Division.Update(div);
			} catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.DataString = ex.Message;
			}

			return rsp;
		}
	}
}