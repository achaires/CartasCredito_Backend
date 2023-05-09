using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class InvitacionRegistroDTO
	{
		public string UserName { get; set; }
		public string Password { get; set; }
		public string Token { get; set; }
	}
}