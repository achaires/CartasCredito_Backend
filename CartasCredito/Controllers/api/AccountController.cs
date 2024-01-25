using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Threading;
using System.Web.Helpers;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class AccountController : ApiController
	{
		[HttpPost]
		public IHttpActionResult Login(LoginDTO loginDTO)
		{
			try
			{
				if (!ModelState.IsValid)
				{
					return BadRequest(ModelState);
				}

				var usuario = AspNetUser.GetByUserName(loginDTO.UserName);

				// Verifica que sea un usuario activo
				if (usuario.Activo == false)
				{
					return BadRequest();
				}

				// Verifica que el password sea correcto
				bool isLocalCredentialValid = Crypto.VerifyHashedPassword(usuario.PasswordHash, loginDTO.Password);

				// Validar cuenta en GIS
				var tryLogin_Gis = Utility.Login_GIS(loginDTO.UserName, loginDTO.Password);

				// Si no pasa la validacion en GIS, regresar mensaje - bypass cuenta softdepot
				if (tryLogin_Gis.Flag == false && loginDTO.UserName.ToLower() != "softdepot")
				{
					return Content(HttpStatusCode.Unauthorized, "LOGIN_GIS");
				}

				if (isLocalCredentialValid)
				{
					var token = TokenGenerator.GenerateTokenJwt(loginDTO.UserName);

					return Ok(token);
				}
				else
				{
					return Content(HttpStatusCode.Unauthorized, "LOGIN_APP");
				}
			}
			catch (Exception ex)
			{
				return Content(HttpStatusCode.Unauthorized, "LOGIN_APP");
			}
		}

		[HttpGet]
		public AspNetUser User()
		{
			var user = new AspNetUser();
			try
			{
				var identity = Thread.CurrentPrincipal.Identity;
				user = AspNetUser.GetByUserName(identity.Name);
			} catch (Exception ex)
			{
				Utility.Logger.Error(ex);
			}

			return user;
		}


	}
}