using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using OfficeOpenXml;
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
	public class PFEProgramasController : ApiController
	{
		[Route("api/pfeprogramas/buscar")]
		[HttpPost]
		public PFEPrograma Buscar(PFEProgramaBuscarDTO pfeProgramaBuscarDTO)
		{
			var rsp = new PFEPrograma();

			try
			{
				var buscarModelo = new PFEPrograma()
				{
					Anio = pfeProgramaBuscarDTO.Anio,
					Periodo = pfeProgramaBuscarDTO.Periodo,
					EmpresaId = pfeProgramaBuscarDTO.EmpresaId
				};

				rsp = PFEPrograma.Buscar(buscarModelo);

				if ( rsp.Id < 1 )
				{
					var dateNow = DateTime.Now;
					var periodoUltDia = DateTime.DaysInMonth(pfeProgramaBuscarDTO.Anio, pfeProgramaBuscarDTO.Periodo);

					var fechaPeriodoIni = new DateTime(pfeProgramaBuscarDTO.Anio, pfeProgramaBuscarDTO.Periodo, 1,0,0,0);
					var fechaPeriodoFin = new DateTime(pfeProgramaBuscarDTO.Anio, pfeProgramaBuscarDTO.Periodo, periodoUltDia, 23, 59, 59);

					rsp.Anio = pfeProgramaBuscarDTO.Anio;
					rsp.Periodo = pfeProgramaBuscarDTO.Periodo;
					rsp.EmpresaId = pfeProgramaBuscarDTO.EmpresaId;
					rsp.Pagos = Pago.GetProgramados()
							.Where(p => p.EmpresaId == pfeProgramaBuscarDTO.EmpresaId)
							.Where(p => p.FechaVencimiento >= fechaPeriodoIni)
							.Where(p => p.FechaVencimiento <= fechaPeriodoFin)
							.ToList();
				}
			} catch (Exception ex)
			{
				Utility.Logger.Error(ex);
			}

			return rsp;
		}

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
		public RespuestaFormato Post([FromBody] PFEProgramaInsertarDTO pfePrograma)
		{
			var rsp = new RespuestaFormato();
			try
			{
				var newProg = new PFEPrograma()
				{
					EmpresaId = pfePrograma.EmpresaId,
					Anio = pfePrograma.Anio,
					Periodo = pfePrograma.Periodo,
				};

				rsp = PFEPrograma.Insert(newProg);

				if ( rsp.Flag )
				{
					foreach ( var progPago in pfePrograma.Pagos)
					{
						PFEPrograma.InsertPFEProgramaPago(rsp.DataInt, progPago.Id);
					}

					foreach (var procTipoCambio in pfePrograma.TiposCambio)
					{
						var newTc = new PFETipoCambio()
						{
							ProgramaId = rsp.DataInt,
							MonedaId = procTipoCambio.MonedaId,
							PA = procTipoCambio.PA,
							PA1 = procTipoCambio.PA1,
							PA2 = procTipoCambio.PA2
						};

						PFEPrograma.InsertTipoCambio(newTc);
					}
				}

				rsp.DataString = "Programa creado con éxito";
				rsp.Flag = true;
			} catch (Exception ex)
			{
				rsp.DataString = ex.Message;
				rsp.Flag = false;
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

		// PUT api/<controller>/5
		public void Put(int id, [FromBody] string value)
		{
		}

		// DELETE api/<controller>/5
		public void Delete(int id)
		{
		}
	}
}