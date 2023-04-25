using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class ComisionesPorTipoComision
	{
		public DateTime FechaInicio { get; set; }
		public DateTime FechaFin { get; set; }
		public decimal GranTotal { get; set; }

		public List<ReporteEmpresa> Empresas { get; set; }

		public class ReporteEmpresa
		{
			public string Empresa { get; set; }
			public decimal Total { get; set; }
			public List<EmpresaComision> EmpresaComisiones { get; set; }

			public class EmpresaComision 
			{
				public decimal Total { get; set; }
				public List<Comision> Comisiones { get; set; }
			}
		}
	}
}