using CartasCredito.Models.DTOs;
using CartasCredito.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Helpers;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class ConversionMonedaController : ApiController
    {
		[HttpPost]
		public RespuestaFormato Convertir(ConversionMonedaDTO operacion)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var rspVal = 0M;

				switch ( operacion.MonedaInput )
				{
					case "USD":
						rspVal = 19.8632M;
						break;
					case "EUR":
						rspVal = 22.1557M;
						break;
					case "JPY":
						rspVal = 0.161M;
						break;
					case "RMB":
						rspVal = 0.161M;
						break;
				}

				rsp.Flag = true;
				rsp.DataString = rspVal.ToString();
				rsp.DataDecimal = rspVal;

				/*
				var clnt = new ConversionMonedaService.BPELToolsClient();
				var req = new ConversionMonedaService.processRequest();
				var res = new ConversionMonedaService.processResponse();

				req.process = new ConversionMonedaService.process();
				req.process.P_USER_CONVERSION_TYPE = "Financiero Venta";
				req.process.P_CONVERSION_DATESpecified = true;
				//req.process.P_CONVERSION_DATE = DateTime.Parse(operacion.Fecha);
				req.process.P_CONVERSION_DATE = DateTime.Parse(operacion.Fecha);
				req.process.P_FROM_CURRENCY = operacion.MonedaInput;
				req.process.P_TO_CURRENCY = operacion.MonedaOutput;

				res = clnt.process(req.process);

				if ( res.X_CONVERSION_RATE != null && res.X_MNS_ERROR == null )
				{
					rsp.Flag = true;
					rsp.DataString = res.X_CONVERSION_RATE.ToString();
					rsp.DataDecimal = res.X_CONVERSION_RATE.Value;
				} else
				{
					rsp.Flag = false;
					rsp.Errors.Add(res.X_MNS_ERROR);
					rsp.DataString = res.X_MNS_ERROR;
				}
				*/
			} catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.Errors.Add(ex.Message);
				rsp.DataString = ex.Message;
			}

			return rsp;
		}
	}
}
