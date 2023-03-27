using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class EnmiendaInsertDTO
	{
		public string ConsideracionesAdicionales {get;set;}
		public string DescripcionMercancia {get;set;}
		public DateTime? FechaLimiteEmbarque {get;set;}
		public DateTime? FechaVencimiento {get;set;}
		public decimal? ImporteLC {get;set;}
		public string InstruccionesEspeciales {get;set;}
		public string CartaCreditoId { get; set; }

		public EnmiendaInsertDTO()
		{
			ConsideracionesAdicionales = string.Empty;
			DescripcionMercancia = string.Empty;
			FechaLimiteEmbarque = null;
			FechaVencimiento = null;
			ImporteLC = null;
			InstruccionesEspeciales = string.Empty;
		}
	}
}