using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class RoleInsertDTO
	{
		public string Name { get; set; }
		public string Description { get; set; }
		public List<int> Permissions { get; set; }

		public RoleInsertDTO() {
			Permissions = new List<int>();
		}
	}
}