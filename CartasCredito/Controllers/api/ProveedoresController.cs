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
	public class ProveedoresController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<Proveedor> Get()
		{
			return Proveedor.Get(1);
		}

		// GET api/<controller>/5
		public Proveedor Get(int id)
		{
			return Proveedor.GetById(id);
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] Proveedor modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);
			var paisId = 0;

			try
			{
				var paises = Pais.Get();
				var findPais = paises.Find(p => p.Nombre.Trim().ToLower() == modelo.Pais.Trim().ToLower());
				
				if ( findPais != null)
				{
					paisId = findPais.Id;
				} else
				{
					var nuevoPais = new Pais();
					nuevoPais.Nombre = modelo.Pais;
					var nuevoPaisRsp = Pais.Insert(nuevoPais);
					paisId = nuevoPaisRsp.DataInt;
				}

			} catch (Exception ex)
			{
				paisId = 1;
			}

			try
			{
				var m = new Proveedor()
				{
					EmpresaId = modelo.EmpresaId,
					PaisId = paisId,
					Nombre = modelo.Nombre,
					Descripcion = modelo.Descripcion,
					CreadoPor = usr.Id
				};

				rsp = Proveedor.Insert(m);

				var mContacto = modelo.Contacto;
				mContacto.ModelNombre = "Proveedor";
				mContacto.ModelId = rsp.DataInt;

				Contacto.Insert(mContacto);
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
		public RespuestaFormato Put(int id, [FromBody] Proveedor modelo)
		{
			var rsp = new RespuestaFormato();
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);
			var paisId = 0;

			try
			{
				var paises = Pais.Get();
				var findPais = paises.Find(p => p.Nombre.Trim().ToLower() == modelo.Pais.Trim().ToLower());

				if (findPais != null)
				{
					paisId = findPais.Id;
				}
				else
				{
					var nuevoPais = new Pais();
					nuevoPais.Nombre = modelo.Pais;
					var nuevoPaisRsp = Pais.Insert(nuevoPais);
					paisId = nuevoPaisRsp.DataInt;
				}

			}
			catch (Exception ex)
			{
				paisId = 1;
			}

			try
			{
				var m = Proveedor.GetById(id);

				m.Nombre = modelo.Nombre;
				m.EmpresaId = modelo.EmpresaId;
				m.PaisId = paisId;
				m.Descripcion = modelo.Descripcion;
				m.Activo = m.Activo;

				rsp = Proveedor.Update(m);

				var mContacto = Contacto.GetByModelNombreAndId(id,"Proveedor");
				mContacto.ModelNombre = "Proveedor";
				mContacto.ModelId = rsp.DataInt;
				mContacto.Nombre = modelo.Contacto.Nombre;
				mContacto.ApellidoPaterno = modelo.Contacto.ApellidoPaterno;
				mContacto.ApellidoMaterno = modelo.Contacto.ApellidoMaterno;
				mContacto.Telefono = modelo.Contacto.Telefono;
				mContacto.Fax = modelo.Contacto.Fax;
				mContacto.Email = modelo.Contacto.Email;

				if ( mContacto.Id > 0 )
				{
					Contacto.Update(mContacto);
				} else
				{
					Contacto.Insert(mContacto);
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
		public RespuestaFormato Delete(int id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var m = Proveedor.GetById(id);
				m.Activo = m.Activo ? false : true;

				rsp = Proveedor.Update(m);
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