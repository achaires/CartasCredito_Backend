using CartasCredito.Models;
using CartasCredito.Models.DTOs;
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
	public class CartasCreditoController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<CartaCredito> Get(DateTime fechaInicio, DateTime fechaFin)
		{
			return CartaCredito.Get(fechaInicio, fechaFin, 1);
		}

		// GET api/<controller>/5
		public CartaCredito Get(string id)
		{
			return CartaCredito.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] CartaCredito modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				modelo.CreadoPor = usr.Id;


				if ( modelo.TipoCartaId == 17 )
				{
					rsp = CartaCredito.Insert(modelo);
				} else
				{
					rsp = CartaCredito.InsertStandBy(modelo);
				}

				

				if ( rsp.Flag )
				{
					var bm = new BitacoraMovimiento();
					bm.Titulo = "Nueva carta de crédito";
					bm.CreadoPorId = usr.Id;
					bm.Descripcion = "Se ha creado una nueva carta de crédito";
					bm.CartaCreditoId = rsp.DataString;
					//bm.ModeloNombre = "CartaCredito";
					//bm.ModeloId = rsp.DataString;

					BitacoraMovimiento.Insert(bm);
				}
			}
			catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.DataString = ex.Message;
				rsp.Errors.Add("Ocurrió un error al intentar guardar la Carta de Crédito. Verifique los datos ingresados");
				//rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}

		// PUT api/<controller>/5
		public RespuestaFormato Put(string id, [FromBody] CartaCredito modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				var m = CartaCredito.GetById(id);
				/*
				m.EmpresaId = modelo.EmpresaId;
				m.Nombre = modelo.Nombre;
				m.Descripcion = modelo.Descripcion;
				m.Activo = m.Activo;
				*/

				//rsp = CartaCredito.Update(modelo);
				m.DocumentoANegociar = modelo.DocumentoANegociar;


				if (m.TipoCartaId == 17)
				{
					m.TipoActivoId = modelo.TipoActivoId;
					m.ProyectoId = modelo.ProyectoId;
					m.BancoId = modelo.BancoId;
					m.ProveedorId = modelo.ProveedorId;
					m.EmpresaId = modelo.EmpresaId;
					m.AgenteAduanalId = modelo.AgenteAduanalId;
					m.MonedaId = modelo.MonedaId;
					m.TipoPago = modelo.TipoPago;
					m.Responsable = modelo.Responsable;
					m.CompradorId = modelo.CompradorId;
					m.PorcentajeTolerancia = modelo.PorcentajeTolerancia;
					m.NumOrdenCompra = modelo.NumOrdenCompra;
					m.CostoApertura = modelo.CostoApertura;
					m.MontoOrdenCompra = modelo.MontoOrdenCompra;
					m.MontoOriginalLC = modelo.MontoOriginalLC;
					m.FechaApertura = modelo.FechaApertura;
					m.FechaLimiteEmbarque = modelo.FechaLimiteEmbarque;
					m.FechaVencimiento = modelo.FechaVencimiento;
					m.Incoterm = modelo.Incoterm;
					m.EmbarquesParciales = modelo.EmbarquesParciales;
					m.Transbordos = modelo.Transbordos;
					m.PuntoEmbarque = modelo.PuntoEmbarque;
					m.PuntoDesembarque = modelo.PuntoDesembarque;
					m.DescripcionMercancia = modelo.DescripcionMercancia;
					m.DescripcionCartaCredito = modelo.DescripcionCartaCredito;
					m.PagoCartaAceptacion = modelo.PagoCartaAceptacion;
					m.ConsignacionMercancia = modelo.ConsignacionMercancia;
					m.ConsideracionesAdicionales = modelo.ConsideracionesAdicionales;
					m.DiasParaPresentarDocumentos = modelo.DiasParaPresentarDocumentos;
					m.DiasPlazoProveedor = modelo.DiasPlazoProveedor;
					m.CondicionesPago = modelo.CondicionesPago;
					m.NumeroPeriodos = modelo.NumeroPeriodos;
					m.BancoCorresponsalId = modelo.BancoCorresponsalId;
					m.SeguroPorCuenta = modelo.SeguroPorCuenta;
					m.GastosComisionesCorresponsal = modelo.GastosComisionesCorresponsal;
					m.ConfirmacionBancoNotificador = modelo.ConfirmacionBancoNotificador;
					m.TipoEmision = modelo.TipoEmision;
					rsp = CartaCredito.Update(m);
				}
				else
				{
					m.TipoStandBy = modelo.TipoStandBy;
					m.BancoId = modelo.BancoId;
					m.EmpresaId = modelo.EmpresaId;
					m.MonedaId = modelo.MonedaId;
					m.CompradorId = modelo.CompradorId;
					m.MontoOriginalLC = modelo.MontoOriginalLC;
					m.FechaApertura = modelo.FechaApertura;
					m.FechaLimiteEmbarque = modelo.FechaLimiteEmbarque;
					m.FechaVencimiento = modelo.FechaVencimiento;
					m.ConsideracionesReclamacion = modelo.ConsideracionesReclamacion;
					m.ConsideracionesAdicionales = modelo.ConsideracionesAdicionales;
					m.BancoCorresponsalId = modelo.BancoCorresponsalId;
					m.TipoCoberturaId = modelo.TipoCoberturaId;
					rsp = CartaCredito.UpdateStandBy(m);
				}
				
				if(rsp.Flag == true)
                {
					RespuestaFormato delDocs = CartaCreditoDocumentoANegociar.Delete(m.Id);

					foreach (CartaCreditoDocumentoANegociar item in m.DocumentoANegociar)
					{
						item.IdCartaCredito = m.Id;
						item.IdDocumento = item.DocId;
						item.Insert();
					}
				}
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
		public RespuestaFormato Delete(string id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var modelo = CartaCredito.GetById(id);
				modelo.Activo = modelo.Activo ? false : true;

				rsp = CartaCredito.Update(modelo);
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