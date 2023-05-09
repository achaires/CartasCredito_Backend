﻿using CartasCredito.Models;
using CartasCredito.Models.DTOs;
using Microsoft.Ajax.Utilities;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Web.Helpers;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class UsersController : ApiController
	{
		// GET api/users
		public IEnumerable<AspNetUser> Get()
		{
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
			var usrId = "12cb7342-837e-45d9-892c-6818a38a3816";

			try
			{

				var curDate = new DateTime();
				var u = new AspNetUser();
				u.UserName = userDto.UserName;
				u.Email = userDto.Email;
				u.PasswordHash = Crypto.HashPassword("RandomPassword" + curDate.ToString());
				u.PhoneNumber = userDto.PhoneNumber;
				u.Activo = false;
				
				rsp = AspNetUser.Insert(u);

				// Crear invitación
				var invitacion = new Invitacion()
				{
					Email = userDto.Email,
					UserName = userDto.UserName,
					CreadoPorId = usrId
				};

				var rspInv = Invitacion.Insert(invitacion);

				if ( rsp.DataInt > 0 )
				{
					var invDb = Invitacion.GetById(rsp.DataInt);

					// enviar invitacion
					var smtpClient = new SmtpClient(Utility.SmtpHost)
					{
						Port = int.Parse(Utility.SmtpPort),
						Credentials = new NetworkCredential(Utility.SmtpUser, Utility.SmtpPass),
						EnableSsl = false,
					};

					smtpClient.Send("cartasdecredito@gis.com.mx", "j.sanchez@softdepot.mx", "Invitación a Sistema Cartas de Crédito", "Ha sido invitado al sistema de Cartas de Crédito: " + invDb.Token);
				}

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

		// PUT api/users/5
		public RespuestaFormato Put(string id, [FromBody] UserUpdateDTO userDto)
		{
			var rsp = new RespuestaFormato();

			try
			{
				var u = AspNetUser.GetById(id);
				u.UserName = userDto.UserName;
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