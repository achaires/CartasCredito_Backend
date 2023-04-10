using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using Microsoft.Ajax.Utilities;
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
	public class RolesController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<AspNetRole> Get()
		{
			var rsp = new List<AspNetRole>();
			try
			{
				rsp = AspNetRole.Get(0);
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex, ex.Message);
			}

			return rsp;
		}

		// GET api/<controller>/5
		public AspNetRole Get(string id)
		{
			var rsp = new AspNetRole();

			try
			{
				rsp = AspNetRole.GetById(id);
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex, ex.Message);
			}

			return rsp;
		}

		// POST api/<controller>
		public RespuestaFormato Post([FromBody] RoleInsertDTO roleDto)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var r = new AspNetRole();
				r.Name = roleDto.Name;

				rsp = AspNetRole.Insert(r);

				if (rsp.DataInt > 0)
				{
					foreach (var permissionId in roleDto.Permissions)
					{
						AspNetRole.InsertPermission(rsp.DataString, permissionId);
					}
				}
			}
			catch (Exception ex)
			{
				rsp.DataInt = 0;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

		// PUT api/<controller>/5
		public RespuestaFormato Put(string id, [FromBody] RoleUpdateDTO roleDto)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var r = AspNetRole.GetById(id);
				r.Name = roleDto.Name;

				rsp = AspNetRole.Update(r);

				if (rsp.DataInt > 0)
				{
					r.Permissions.ForEach(permission => {
						AspNetRole.DeletePermission(r.Id, permission.Id);
					});

					foreach (var permissionId in roleDto.Permissions)
					{
						AspNetRole.InsertPermission(rsp.DataString, permissionId);
					}
				}
			}
			catch (Exception ex)
			{
				rsp.DataInt = 0;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

		// DELETE api/<controller>/5
		public RespuestaFormato Delete(string id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var r = AspNetRole.GetById(id);
				r.Activo = r.Activo ? false : true;

				rsp = AspNetRole.Update(r);
			}
			catch (Exception ex)
			{
				rsp.DataInt = 0;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}
	}
}