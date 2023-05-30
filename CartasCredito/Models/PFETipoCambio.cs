using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class PFETipoCambio
	{
		public int Id { get; set; }
		public int ProgramaId { get; set; }
		public int MonedaId { get; set; }
		public decimal PA { get; set; }
		public decimal PA1 { get; set; }
		public decimal PA2 { get; set; }
	}
}