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
	public class LineasDeCreditoController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<LineaDeCredito> Get()
		{
			return LineaDeCredito.Get(0);
		}

		// GET api/<controller>/5
		public LineaDeCredito Get(int id)
		{
			return LineaDeCredito.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] LineaDeCredito modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = new LineaDeCredito()
				{
					EmpresaId = modelo.EmpresaId,
					BancoId = modelo.BancoId,
					Monto = modelo.Monto,
					Cuenta = modelo.Cuenta,
					CreadoPor = usr.Id
				};

				rsp = LineaDeCredito.Insert(m);
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
		public RespuestaFormato Put(int id, [FromBody] LineaDeCredito modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = LineaDeCredito.GetById(id);

				m.EmpresaId = modelo.EmpresaId;
				m.BancoId = modelo.BancoId;
				m.Monto = modelo.Monto;
				m.Cuenta = modelo.Cuenta;
				m.Activo = m.Activo;

				rsp = LineaDeCredito.Update(m);
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
				var m = LineaDeCredito.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = LineaDeCredito.Update(m);
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