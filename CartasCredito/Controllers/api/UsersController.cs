using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using Microsoft.Ajax.Utilities;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Threading;
using System.Web;
using System.Web.Helpers;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	[Authorize]
	public class UsersController : ApiController
	{
		// GET api/users
		public IEnumerable<AspNetUser> Get()
		{
			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			var rsp = new List<AspNetUser>();
			try
			{
				rsp = AspNetUser.Get(0);
			} catch (Exception ex)
			{
				Utility.Logger.Error(ex,ex.Message);
			}

			return rsp;
		}

		// GET api/users/5
		public AspNetUser Get(string id)
		{
			var rsp = new AspNetUser();

			try
			{
				rsp = AspNetUser.GetById(id);
			} catch (Exception ex)
			{
				Utility.Logger.Error(ex, ex.Message);
			}

			return rsp;
		}

		// POST api/users
		public RespuestaFormato Post([FromBody] UserInsertDTO userDto)
		{
			var rsp = new RespuestaFormato();

			var identity = Thread.CurrentPrincipal.Identity;
			var usr = AspNetUser.GetByUserName(identity.Name);

			try
			{
				// Crear invitación
				var invitacion = new Invitacion()
				{
					Email = userDto.Email,
					UserName = userDto.Email,
					CreadoPorId = usr.Id
				};

				var rspInv = Invitacion.Insert(invitacion);

				if (rspInv.DataInt > 0)
				{
					var invDb = Invitacion.GetById(rspInv.DataInt);
					var btnLink = Utility.HostUrl + "/#/registro/" + invDb.Token;
					EnviaEmail(userDto.Email,btnLink);
				}

				var curDate = new DateTime();
				var u = new AspNetUser();
				u.UserName = userDto.Email;
				u.Email = userDto.Email;
				u.PasswordHash = Crypto.HashPassword("RandomPassword" + curDate.ToString());
				u.PhoneNumber = userDto.PhoneNumber;
				u.Activo = false;
				
				rsp = AspNetUser.Insert(u);

				if ( rsp.DataInt > 0 )
				{
					AspNetUser.InsertRole(rsp.DataString,userDto.RolId);

					var uProfile = new AspNetUserProfile();
					uProfile.UserId = rsp.DataString;
					uProfile.Name = userDto.Name;
					uProfile.LastName = userDto.LastName;
					uProfile.DisplayName = userDto.Name + " " + userDto.LastName;
					uProfile.Notes = userDto.Notes;
					AspNetUserProfile.Insert(uProfile);

                    foreach (var empresaId in userDto.Empresas)
                    {
                        AspNetUser.InsertEmpresa(rsp.DataString, empresaId);
                    }
                }
			} catch (Exception ex)
			{
				rsp.DataInt = 0;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

		protected static RespuestaFormato EnviaEmail(string sendEmailTo, string btnLink)
		{
			RespuestaFormato res = new RespuestaFormato();
			try
			{
				dll_Gis.Funciones fn = new dll_Gis.Funciones();
				int port = 0;
				string host = "";
				string username = "";
				string mailFrom = "cartas.credito@gis.com.mx";
				string password = "";
				bool ssl = false;
				bool defaultCredentials = false;
				string deliveryMethod = "";

				Int32.TryParse(ConfigurationManager.AppSettings["SMTPPort"].ToString(), out port);
				host = fn.Desencriptar(ConfigurationManager.AppSettings["SMTPHost"].ToString());
				username = fn.Desencriptar(ConfigurationManager.AppSettings["SMTPUsername"].ToString());

				Boolean.TryParse(ConfigurationManager.AppSettings["SMTPSSL"].ToString(), out ssl);
				Boolean.TryParse(ConfigurationManager.AppSettings["SMTPDefaultCredentials"].ToString(), out defaultCredentials);
				deliveryMethod = ConfigurationManager.AppSettings["SMTPDeliveryMethod"].ToString();

				MailMessage mail = new MailMessage();
				mail.IsBodyHtml = true;
				SmtpClient SmtpServer = new SmtpClient(host);

				mail.From = new MailAddress(mailFrom, "Portal Cartas de Crédito");

				var mailTo = new MailAddress(sendEmailTo);
				mail.To.Add(mailTo);
				mail.Subject = "Registro en Portal Cartas de Crédito";

				string mailBody = System.IO.File.ReadAllText(HttpContext.Current.Server.MapPath("~/Views/Shared/NotificacionRegistro.html"));
				mailBody = mailBody.Replace("#AppUrl#", btnLink);

				mail.Body = mailBody;


				mail.ReplyToList.Add(new MailAddress(mailFrom, "reply-to"));

				SmtpServer.Port = port;
				SmtpServer.UseDefaultCredentials = defaultCredentials;
				if (deliveryMethod == "Network")
				{
					SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
				}
				else
				{
					password = fn.Desencriptar(ConfigurationManager.AppSettings["SMTPPassword"].ToString());
					SmtpServer.Credentials = new System.Net.NetworkCredential(username, password);
				}
				SmtpServer.EnableSsl = ssl;
				//ACTIVAR CORREOS
				SmtpServer.Send(mail);
				res.Flag = true;
			}
			catch (Exception ex)
			{
				res.Flag = false;
				res.Errors.Add(ex.Message);
			}

			return res;
		}

		// PUT api/users/5
		public RespuestaFormato Put(string id, [FromBody] UserUpdateDTO userDto)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var u = AspNetUser.GetById(id);
				u.UserName = userDto.Email;
				u.Email = userDto.Email;
				u.PhoneNumber = userDto.PhoneNumber;

				rsp = AspNetUser.Update(u);

				if (rsp.DataInt > 0)
				{
					AspNetUser.UpdateRole(id, userDto.RolId);

					var uProfile = AspNetUserProfile.GetByUserId(id);
					uProfile.Name = userDto.Name;
					uProfile.LastName = userDto.LastName;
					uProfile.DisplayName = userDto.Name + " " + userDto.LastName;
					uProfile.Notes = userDto.Notes;
					AspNetUserProfile.Update(uProfile);

					u.Empresas.ForEach(emp => {
						AspNetUser.DeleteEmpresa(u.Id, emp.Id);
					});

					foreach (var empresaId in userDto.Empresas)
					{
						AspNetUser.InsertEmpresa(u.Id, empresaId);
					}
				}
			} catch (Exception ex) {
				rsp.DataInt = 0;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);

				Utility.Logger.Error(ex.Message);
			}

			return rsp;
		}

		// DELETE api/users/5
		public RespuestaFormato Delete(string id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var u = AspNetUser.GetById(id);
				u.Activo = u.Activo ? false : true ;

				rsp = AspNetUser.Update(u);
			} catch ( Exception ex )
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