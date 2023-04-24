using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class OperacionesController : ApiController
	{
		[HttpPost]
		public IEnumerable<CartaCredito> Filtrar([FromBody] CartasCreditoFiltrarDTO filtros)
		{
			var fechaIni = new DateTime(filtros.FechaInicio.Year, filtros.FechaInicio.Month, filtros.FechaInicio.Day, 0, 0, 0);
			var fechaFin = new DateTime(filtros.FechaFin.Year, filtros.FechaFin.Month, filtros.FechaFin.Day, 23, 59, 59);

			filtros.FechaInicio = fechaIni;
			filtros.FechaFin = fechaFin;
			
			return CartaCredito.Filtrar(filtros);
		}

		[Route("api/operaciones/cambiarestatus/{id}")]
		[HttpPost]
		public RespuestaFormato CambiarEstatus(string id, [FromBody] CartaCredito modelo)
		{
			var cc = CartaCredito.GetById(id);
			var rsp = CartaCredito.UpdateStatus(id, modelo.Estatus);

			return rsp;
		}

		[Route("api/operaciones/adjuntarswift/{id}")]
		[HttpPost]
		public async Task<RespuestaFormato> AdjuntarSwift([FromUri] string id)
		{
			var cc = CartaCredito.GetById(id);
			var rf = new RespuestaFormato();


			if (!Request.Content.IsMimeMultipartContent())
			{
				throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
			}

			string root = HttpContext.Current.Server.MapPath("~/Uploads");
			var provider = new MultipartFormDataStreamProvider(root);

			try
			{
				// Read the form data.
				await Request.Content.ReadAsMultipartAsync(provider);
				var swiftFilename = "";

				// This illustrates how to get the file names.
				foreach (MultipartFileData file in provider.FileData)
				{
					swiftFilename = file.Headers.ContentDisposition.FileName.Trim('\"');
					swiftFilename = DateTime.Now.ToString("yyyyMMddHmmss") + "-" + swiftFilename.Trim();
					string swiftFinalName = Path.Combine(root, swiftFilename);
					File.Move(file.LocalFileName,swiftFinalName);

					//Trace.WriteLine(file.Headers.ContentDisposition.FileName);
					//Trace.WriteLine("Server file path: " + file.LocalFileName);
				}

				rf = CartaCredito.UpdateSwiftFile(id, HttpContext.Current.Request.Form["NumCarta"], swiftFilename);
			}
			catch (System.Exception ex)
			{
				rf.DataInt = 0;
				rf.DataString = "";
				rf.Flag = false;
				rf.Errors.Add(ex.Message);
			}


			return rf;
		}

		[Route("api/operaciones/adjuntarswift-enmienda/{id}")]
		[HttpPost]
		public async Task<RespuestaFormato> AdjuntarSwiftEnmienda([FromUri] int id)
		{
			var enmienda = Enmienda.GetById(id);
			var cc = CartaCredito.GetById(enmienda.CartaCreditoId);
			var rf = new RespuestaFormato();


			if (!Request.Content.IsMimeMultipartContent())
			{
				throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
			}

			string root = HttpContext.Current.Server.MapPath("~/Uploads");
			var provider = new MultipartFormDataStreamProvider(root);

			try
			{
				// Read the form data.
				await Request.Content.ReadAsMultipartAsync(provider);
				var swiftFilename = "";

				// This illustrates how to get the file names.
				foreach (MultipartFileData file in provider.FileData)
				{
					swiftFilename = file.Headers.ContentDisposition.FileName.Trim('\"');
					swiftFilename = DateTime.Now.ToString("yyyyMMddHmmss") + "-" + swiftFilename.Trim();
					string swiftFinalName = Path.Combine(root, swiftFilename);
					File.Move(file.LocalFileName, swiftFinalName);

					//Trace.WriteLine(file.Headers.ContentDisposition.FileName);
					//Trace.WriteLine("Server file path: " + file.LocalFileName);
				}

				rf = Enmienda.UpdateSwiftFile(id, cc.Id, swiftFilename);
			}
			catch (System.Exception ex)
			{
				rf.DataInt = 0;
				rf.DataString = "";
				rf.Flag = false;
				rf.Errors.Add(ex.Message);
			}


			return rf;
		}

		[Route("api/operaciones/clonar")]
		[HttpPost]
		public async Task<RespuestaFormato> Clonar([FromBody] string cartaId)
		{
			var cc = CartaCredito.GetById(cartaId);
			var rf = new RespuestaFormato();


			try
			{
				var ccClon = new CartaCredito();

				ccClon.TipoCarta = cc.TipoCarta;
				ccClon.TipoCartaId = cc.TipoCartaId;
				ccClon.TipoActivoId = cc.TipoActivoId;
				ccClon.ProyectoId = cc.ProyectoId;
				ccClon.BancoId = cc.BancoId;
				ccClon.ProveedorId = cc.ProveedorId;
				ccClon.EmpresaId = cc.EmpresaId;
				ccClon.AgenteAduanalId = cc.AgenteAduanalId;
				ccClon.MonedaId = cc.MonedaId;
				ccClon.TipoPago = cc.TipoPago;
				ccClon.Responsable = cc.Responsable;
				ccClon.CompradorId = cc.CompradorId;
				ccClon.PorcentajeTolerancia = cc.PorcentajeTolerancia;
				ccClon.CostoApertura = cc.CostoApertura;
				ccClon.MontoOrdenCompra = cc.MontoOrdenCompra;
				ccClon.MontoOriginalLC = cc.MontoOriginalLC;
				ccClon.FechaApertura = cc.FechaApertura;
				ccClon.Incoterm = cc.Incoterm;
				ccClon.FechaLimiteEmbarque = cc.FechaLimiteEmbarque ;
				ccClon.FechaVencimiento = cc.FechaVencimiento;
				ccClon.EmbarquesParciales = cc.EmbarquesParciales;
				ccClon.Transbordos = cc.Transbordos;
				ccClon.PuntoEmbarque = cc.PuntoEmbarque;
				ccClon.PuntoDesembarque = cc.PuntoDesembarque;
				ccClon.DescripcionMercancia = cc.DescripcionMercancia;
				ccClon.DescripcionCartaCredito = cc.DescripcionCartaCredito;
				ccClon.InstruccionesEspeciales = cc.InstruccionesEspeciales;
				ccClon.PagoCartaAceptacion = cc.PagoCartaAceptacion;
				ccClon.ConsignacionMercancia = cc.ConsignacionMercancia;
				ccClon.ConsideracionesAdicionales = cc.ConsideracionesAdicionales;
				ccClon.ConsideracionesReclamacion = cc.ConsideracionesReclamacion;
				ccClon.DiasParaPresentarDocumentos = cc.DiasParaPresentarDocumentos;
				ccClon.DiasPlazoProveedor = cc.DiasPlazoProveedor;
				ccClon.CondicionesPago = cc.CondicionesPago;
				ccClon.NumeroPeriodos = cc.NumeroPeriodos;
				ccClon.BancoCorresponsalId = cc.BancoCorresponsalId;
				ccClon.SeguroPorCuenta = cc.SeguroPorCuenta;
				ccClon.GastosComisionesCorresponsal = cc.GastosComisionesCorresponsal;
				ccClon.ConfirmacionBancoNotificador = cc.ConfirmacionBancoNotificador;
				ccClon.TipoEmision = cc.TipoEmision;
				ccClon.CreadoPor = cc.CreadoPor;

				rf = CartaCredito.Insert(ccClon);
			} 
			catch (Exception ex) {
				rf.DataInt = 0;
				rf.DataString = "";
				rf.Flag = false;
				rf.Errors.Add(ex.Message);
			}

			return rf;
		}

		/*
		[Route("api/operaciones/adjuntarswift/{id}")]
		[HttpPost]
		public RespuestaFormato AdjuntarSwift(string id, [FromBody] CartaCreditoAdjuntarSwiftDTO adjuntarSwiftDTO)
		{
			var cc = CartaCredito.GetById(id);
			var rf = new RespuestaFormato();
			

			try
			{
				rf = CartaCredito.UpdateSwiftFile(id, adjuntarSwiftDTO.NumCarta, "prueba file");
			} catch ( Exception ex )
			{
				rf.DataInt = 0;
				rf.DataString = "";
				rf.Flag = false;
				rf.Errors.Add(ex.Message);
			}

			return rf;
		}
		*/
	}
}