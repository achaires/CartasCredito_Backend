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
	public class AgentesAduanalesController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<AgenteAduanal> Get()
		{
			return AgenteAduanal.Get(0);
		}

		// GET api/<controller>/5
		public AgenteAduanal Get(int id)
		{
			return AgenteAduanal.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] AgenteAduanal modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var div = new AgenteAduanal()
				{
					EmpresaId = modelo.EmpresaId,
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					CreadoPor = usr
				};

				rsp = AgenteAduanal.Insert(div);
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
		public RespuestaFormato Put(int id, [FromBody] AgenteAduanal modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var div = AgenteAduanal.GetById(id);

				div.EmpresaId = modelo.EmpresaId;
				div.Nombre = modelo.Nombre;
				div.Descripcion = modelo.Descripcion;
				div.Activo = div.Activo;

				rsp = AgenteAduanal.Update(div);
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
				var modelo = AgenteAduanal.GetById(id);
				modelo.Activo = modelo.Activo ? false : true;

				rsp = AgenteAduanal.Update(modelo);
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