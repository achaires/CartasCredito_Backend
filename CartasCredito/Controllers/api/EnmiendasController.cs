using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Web.Configuration;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class EnmiendasController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<string> Get()
		{
			return new string[] { "value1", "value2" };
		}

		// GET api/<controller>/5
		public string Get(int id)
		{
			return "value";
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] EnmiendaInsertDTO dtoEnmienda)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var modelo = new Enmienda()
				{
					CartaCreditoId = dtoEnmienda.CartaCreditoId,
					ConsideracionesAdicionales = dtoEnmienda.ConsideracionesAdicionales,
					DescripcionMercancia = dtoEnmienda.DescripcionMercancia,
					FechaLimiteEmbarque = dtoEnmienda.FechaLimiteEmbarque,
					FechaVencimiento = dtoEnmienda.FechaVencimiento,
					ImporteLC = dtoEnmienda.ImporteLC,
					InstruccionesEspeciales = dtoEnmienda.InstruccionesEspeciales,
					CreadoPor = usr
				};

				rsp = Enmienda.Insert(modelo);
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
		public RespuestaFormato Put(int id, [FromBody] EnmiendaUpdateDTO dtoEnmienda)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				//var modelo = Enmienda.GetById(id);
				var cc = CartaCredito.GetById(dtoEnmienda.CartaCreditoId);
				var modelo = cc.Enmiendas.Find(enm => enm.Id == dtoEnmienda.Id);

				modelo.Estatus = dtoEnmienda.Estatus;
				modelo.Prev_ImporteLC = cc.MontoOriginalLC;
				modelo.Prev_FechaVencimiento = cc.FechaVencimiento;
				modelo.Prev_FechaVencimiento = cc.FechaLimiteEmbarque;
				modelo.Prev_DescripcionMercancia = cc.DescripcionMercancia;
				modelo.Prev_ConsideracionesAdicionales = cc.ConsideracionesAdicionales;
				modelo.Prev_InstruccionesEspeciales = cc.InstruccionesEspeciales;

				cc.Estatus = dtoEnmienda.Estatus;
				cc.MontoOriginalLC = dtoEnmienda.ImporteLC;
				cc.FechaVencimiento = dtoEnmienda.FechaVencimiento;
				cc.FechaLimiteEmbarque = dtoEnmienda.FechaLimiteEmbarque;
				cc.DescripcionMercancia = dtoEnmienda.DescripcionMercancia;
				cc.ConsideracionesAdicionales = dtoEnmienda.ConsideracionesAdicionales;
				cc.InstruccionesEspeciales = dtoEnmienda.InstruccionesEspeciales;

				Enmienda.Update(modelo);

				rsp = CartaCredito.Update(cc);
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
		public void Delete(int id)
		{
		}
	}
}