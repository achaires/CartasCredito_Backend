using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class UserUpdatePasswordDTO
	{
		public string Id { get; set; }
		public string Password { get; set; }
	}
}