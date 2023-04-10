using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class UserUpdateDTO
	{
		public string Id { get; set; }
		public string UserName { get; set; }
		public string Name { get; set; }
		public string LastName { get; set; }
		public string Email { get; set; }
		public string PhoneNumber { get; set; }
		public string Notes { get; set; }
		public List<int> Empresas { get; set; }
		public string RolId { get; set; }
	}
}