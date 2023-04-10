using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class PFEProgramaInsertDTO
	{
		public int Anio { get; set; }
		public int Periodo { get; set; }
		public int EmpresaId { get; set; }

		public List<int> PagosId { get; set; }
	}
}