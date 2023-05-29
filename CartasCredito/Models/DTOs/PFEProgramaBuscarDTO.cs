using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class PFEProgramaBuscarDTO
	{
		public int Anio { get; set; }
		public int Periodo { get; set; }
		public int EmpresaId { get; set; }
	}
}