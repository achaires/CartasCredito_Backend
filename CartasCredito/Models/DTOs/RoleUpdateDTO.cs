using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class RoleUpdateDTO
	{
		public string Id { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public bool Activo { get; set; }
		public List<int> Permissions { get; set; }
	}
}