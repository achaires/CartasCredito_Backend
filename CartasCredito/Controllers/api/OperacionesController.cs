﻿using CartasCredito.Models;
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