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
		public DateTime FechaDivisas { get; set; }
		public DateTime FechaVencimientoInicio { get; set; }
		public DateTime FechaVencimientoFin { get; set; }
		public SolicitudReporteDTO ()
		{
			var fechaActual = DateTime.Now;
			TipoReporteId = 1;
			//FechaInicio = new DateTime(fechaActual.Year, fechaActual.Month, 1, 0, 0, 0);
			//FechaFin = new DateTime(fechaActual.Year, fechaActual.Month, fechaActual.Day, 23, 59, 59);
			FechaVencimientoInicio = new DateTime(1969, 1, 1, 0, 0, 0);
			FechaVencimientoFin = new DateTime(1969, 1, 1, 23, 59, 59);
			FechaInicio = new DateTime(1969, 1, 1, 0, 0, 0);
			FechaFin = new DateTime(1969, 1, 1, 23, 59, 59);
		}
	}
}