using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class SolicitudReporteDTO
	{
		public int TipoReporteId { get; set; }
		public DateTime FechaInicio { get; set; }
		public DateTime FechaFin { get; set; }
		public int EmpresaId { get; set; }
		public SolicitudReporteDTO ()
		{
			var fechaActual = DateTime.Now;
			TipoReporteId = 1;
			FechaInicio = new DateTime(fechaActual.Year,fechaActual.Month,1,0,0,0);
			FechaFin = new DateTime(fechaActual.Year, fechaActual.Month, fechaActual.Day, 23, 59, 59);
		}
	}
}