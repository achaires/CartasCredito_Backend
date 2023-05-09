using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Helpers;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class InvitacionesController : ApiController
	{

		[HttpPost]
		[Route("api/validainvitacion/{tokenString}")]
		public IHttpActionResult Validainvitacion(string tokenString)
		{
			var invitacion = Invitacion.GetByToken(tokenString);

			if (invitacion == null)
			{
				return Unauthorized();
			}
			else
			{
				return Ok(invitacion);
			}
		}

		[HttpPost]
		[Route("api/invitaciones/finalizaregistro")]
		public RespuestaFormato Finalizaregistro(InvitacionRegistroDTO invitacionRegistroDTO)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var usr = AspNetUser.GetByUserName(invitacionRegistroDTO.UserName);

				if ( usr != null && usr.Id.Length > 0 )
				{
					var tryLogin_Gis = Utility.Login_GIS(usr.Email, invitacionRegistroDTO.Password);
					//var tryLogin_Gis = Utility.Login_GIS("jesus.sanchez.ext@gis.com.mx", "O}8CuTL&[TAdZ8Z");
					if (tryLogin_Gis.Flag != false)
					{
						usr.PasswordHash = Crypto.HashPassword(invitacionRegistroDTO.Password);
						rsp = AspNetUser.UpdatePassword(usr);
					}
					else
					{
						throw new Exception(tryLogin_Gis.Description);
					}
				} else
				{
					throw new Exception("Usuario no encontrado");
				}
			} catch (Exception ex)
			{
				rsp.Flag = false;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}
	}
}