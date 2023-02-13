﻿using CartasCredito.Models;
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
	public class ComisionesController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Comision> Get()
		{
			return Comision.Get(0);
		}

		// GET api/<controller>/5
		public Comision Get(int id)
		{
			return Comision.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Comision modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var m = new Comision()
				{
					BancoId = modelo.BancoId,
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					Costo = modelo.Costo,
					SwiftApertura = modelo.SwiftApertura,
					SwiftOtro = modelo.SwiftOtro,
					PorcentajeIVA = modelo.PorcentajeIVA,
					CreadoPor = usr
				};

				rsp = Comision.Insert(m);
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
		public RespuestaFormato Put(int id, [FromBody] Comision modelo)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var m = Comision.GetById(id);

				m.BancoId = modelo.BancoId;
				m.Nombre = modelo.Nombre;
				m.Descripcion = modelo.Descripcion;
				m.Costo = modelo.Costo;
				m.SwiftApertura = modelo.SwiftApertura;
				m.SwiftOtro = modelo.SwiftOtro;
				m.PorcentajeIVA = modelo.PorcentajeIVA;
				m.Activo = m.Activo;

				rsp = Comision.Update(m);
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
				var m = Comision.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = Comision.Update(m);
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