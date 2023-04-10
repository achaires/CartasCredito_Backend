using CartasCredito.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CartasCredito.Controllers.api
{
	[AllowAnonymous]
	[EnableCors(origins: "*", headers: "*", methods: "*")]
	public class PermissionsController : ApiController
	{
		// GET api/<controller>
		public IEnumerable<AspNetPermission> Get()
		{
			var rsp = new List<AspNetPermission>();
			try
			{
				rsp = AspNetPermission.Get(0);
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex, ex.Message);
			}

			return rsp;
		}

		// GET api/<controller>/5
		public string Get(int id)
		{
			return "value";
		}

		// POST api/<controller>
		public void Post([FromBody] string value)
		{
		}

		// PUT api/<controller>/5
		public void Put(int id, [FromBody] string value)
		{
		}

		// DELETE api/<controller>/5
		public void Delete(int id)
		{
		}
	}
}