using CartasCredito.Models;
using CartasCredito.Models.DTOs;
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
		public PFEPrograma buscar([FromBody] PFEProgramaBuscarDTO programaBuscarDto)
		{
			var rsp = new PFEPrograma();
			try
			{
				var model = new PFEPrograma();
				model.Anio = programaBuscarDto.Anio;
				model.Periodo = programaBuscarDto.Periodo;
				model.EmpresaId = programaBuscarDto.EmpresaId;

				rsp = PFEPrograma.Buscar(model);
				
				if ( rsp.Id == 0 )
				{
					// buscar pagos del periodo seleccionado
					var daysInPeriodo = DateTime.DaysInMonth(programaBuscarDto.Anio,programaBuscarDto.Periodo);
					var fechaInicio = new DateTime(programaBuscarDto.Anio, programaBuscarDto.Periodo, 1,0,0,0);
					var fechaFin = new DateTime(programaBuscarDto.Anio, programaBuscarDto.Periodo, daysInPeriodo, 23, 59, 59);
					var pagos = Pago.Get(fechaInicio,fechaFin);

					rsp.Pagos = pagos;
				}

			} catch (Exception ex) {
				Utility.Logger.Error(ex);
			}

			return rsp;
		}

		[Route("api/pfeprogramas/agregar")]
		[HttpPost]
		public RespuestaFormato Agregar([FromBody] PFEProgramaInsertDTO programaDto)
		{
			var rsp = new RespuestaFormato();
			var userId = "12cb7342-837e-45d9-892c-6818a38a3816";
			try
			{
				var model = new PFEPrograma();
				model.Anio = programaDto.Anio;
				model.Periodo = programaDto.Periodo;
				model.EmpresaId = programaDto.EmpresaId;
				model.CreadoPor = userId;

				rsp = PFEPrograma.Insert(model);

				if ( rsp.DataInt > 0 )
				{
					programaDto.PagosId.ForEach(p => {
						var newPfePago = new PFEPago();
						newPfePago.PagoId = p;
						newPfePago.PFEProgramaId = rsp.DataInt;
						PFEPago.Insert(newPfePago);
					});
				}
			}
			catch (Exception ex)
			{
				rsp.DataInt = 0;
				rsp.Errors.Add(ex.ToString());
				rsp.DataString1 = ex.ToString();
				Utility.Logger.Error(ex);
			}

			return rsp;
		}
	}
}