using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
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
			if (!ModelState.IsValid)
			{
				return BadRequest(ModelState);
			}

			var hashPass = Crypto.HashPassword(loginDTO.Password);

			var usuario = AspNetUser.GetByUserName(loginDTO.UserName);

			if (usuario.Activo == false)
			{
				return BadRequest();
			}

			bool isCredentialValid = Crypto.VerifyHashedPassword(usuario.PasswordHash, loginDTO.Password);

			if (isCredentialValid)
			{
				var token = TokenGenerator.GenerateTokenJwt(loginDTO.UserName);

				return Ok(token);
			}
			else
			{
				return Unauthorized();
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