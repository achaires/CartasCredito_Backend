using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Web;

namespace CartasCredito.Controllers
{
	public class TokenGenerator
	{
		public static string GenerateTokenJwt(string username)
		{
			// appsetting for Token JWT
			var secretKey = ConfigurationManager.AppSettings["JwtKey"];
			var audienceToken = ConfigurationManager.AppSettings["JwtAudienceToken"];
			var issuerToken = ConfigurationManager.AppSettings["JwtIssuer"];
			var expireTime = ConfigurationManager.AppSettings["JwtExpireMinutes"];

			var securityKey = new SymmetricSecurityKey(System.Text.Encoding.Default.GetBytes(secretKey));
			var signingCredentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256Signature);

			// create a claimsIdentity
			ClaimsIdentity claimsIdentity = new ClaimsIdentity(new[] { new Claim(ClaimTypes.Name, username) });

			// create token to the user
			var tokenHandler = new System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler();
			var jwtSecurityToken = tokenHandler.CreateJwtSecurityToken(
				audience: audienceToken,
				issuer: issuerToken,
				subject: claimsIdentity,
				notBefore: DateTime.UtcNow,
				expires: DateTime.UtcNow.AddMinutes(Convert.ToInt32(expireTime)),
				signingCredentials: signingCredentials);

			var jwtTokenString = tokenHandler.WriteToken(jwtSecurityToken);
			return jwtTokenString;
		}
	}
}