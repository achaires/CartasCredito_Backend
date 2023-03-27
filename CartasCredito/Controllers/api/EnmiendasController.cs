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
		[HttpPost]
		[Route("api/enmiendas")]
		public RespuestaFormato Save([FromBody] EnmiendaInsertDTO dtoEnmienda)
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

		[HttpPost]
		[Route("api/enmiendas/aprobar/{id}")]
		public RespuestaFormato Aprobar(int id)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{
				var modeloEnmienda = Enmienda.GetById(id);
				var cc = CartaCredito.GetById(modeloEnmienda.CartaCreditoId);

				modeloEnmienda.Estatus = 2;
				modeloEnmienda.Prev_ImporteLC = cc.MontoOriginalLC;
				modeloEnmienda.Prev_FechaVencimiento = cc.FechaVencimiento;
				modeloEnmienda.Prev_FechaVencimiento = cc.FechaLimiteEmbarque;
				modeloEnmienda.Prev_DescripcionMercancia = cc.DescripcionMercancia;
				modeloEnmienda.Prev_ConsideracionesAdicionales = cc.ConsideracionesAdicionales;
				modeloEnmienda.Prev_InstruccionesEspeciales = cc.InstruccionesEspeciales;

				cc.Estatus = 2;

				if ( modeloEnmienda.ImporteLC != null && modeloEnmienda.ImporteLC > 0 )
				{
					cc.MontoOriginalLC = (decimal)modeloEnmienda.ImporteLC;
				}

				if ( modeloEnmienda.FechaVencimiento != null  )
				{
					cc.FechaVencimiento = (DateTime)modeloEnmienda.FechaVencimiento;
				}

				if ( modeloEnmienda.FechaLimiteEmbarque != null )
				{
					cc.FechaLimiteEmbarque = (DateTime)modeloEnmienda.FechaLimiteEmbarque;
				}
				
				if ( modeloEnmienda.DescripcionMercancia.Trim().Length > 5 )
				{
					cc.DescripcionMercancia = modeloEnmienda.DescripcionMercancia;
				}

				if (modeloEnmienda.ConsideracionesAdicionales.Trim().Length > 5)
				{
					cc.ConsideracionesAdicionales = modeloEnmienda.ConsideracionesAdicionales;
				}

				if ( modeloEnmienda.InstruccionesEspeciales.Trim().Length > 5 )
				{
					cc.InstruccionesEspeciales = modeloEnmienda.InstruccionesEspeciales;
				}

				Enmienda.Update(modeloEnmienda);

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

		[HttpPut]
		[Route("api/enmiendas/{id}")]
		public RespuestaFormato Update(int id, [FromBody] EnmiendaUpdateDTO dtoEnmienda)
		{
			var rsp = new RespuestaFormato();
			var usr = "12cb7342-837e-45d9-892c-6818a38a3816";

			/*
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
			*/

			return rsp;
		}

		// DELETE api/<controller>/5
		[HttpDelete]
		[Route("api/enmiendas/{id}")]
		public void Delete(int id)
		{
		}
	}
}