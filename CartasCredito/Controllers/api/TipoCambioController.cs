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
	public class TipoCambioController : ApiController
	{
		// GET api/<controller>
		public List<TipoDeCambio> Get()
		{
			List<TipoDeCambio> res = new List<TipoDeCambio>();
			List<Moneda> monedas = Moneda.Get();
			//DateTime fecha = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-ddH") + "T00:00:00");
			DateTime fecha = DateTime.Parse("2023-12-01T00:00:00");
			foreach(Moneda moneda in monedas)
            {

				TipoDeCambio tc = new TipoDeCambio();
				tc.Fecha = fecha;
				tc.MonedaOriginal = moneda.Abbr.Trim().ToUpper();
				tc.MonedaNueva = "MXP";
				tc.Get();

				//
				if (tc.ConversionStr == "-1")
				{
					tc.Conversion = Utility.GetTipoDeCambio(tc.MonedaOriginal, tc.MonedaNueva, tc.Fecha);
					if (tc.Conversion > -1)
					{
						TipoDeCambio.Insert(tc);
					}
				}
				else
				{
					tc.Conversion = Decimal.Parse(tc.ConversionStr);
				}
				//
				res.Add(tc);

			}

			return res;
		}

		// GET api/<controller>/5
		public Comision Get(int id)
		{
			return Comision.GetById(id);
		}
	}
}