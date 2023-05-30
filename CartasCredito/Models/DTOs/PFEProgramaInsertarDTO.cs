using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class PFEProgramaInsertarDTO
	{
		public int Anio { get; set; }
		public int Periodo { get; set; }
		public int EmpresaId { get; set; }
		public List<Pago> Pagos { get; set; }
		public List<PFETipoCambioDTO> TiposCambio { get; set; }
	}

	public class PFETipoCambioDTO
	{
		public int MonedaId { get; set; }
		public decimal PA { get; set; }
		public decimal PA1 { get; set; }
		public decimal PA2 { get; set; }
	}
}