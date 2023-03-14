using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class EnmiendaUpdateDTO
	{
		public int Id { get; set; }
		public int Estatus { get; set; }
		public string ConsideracionesAdicionales { get; set; }
		public string DescripcionMercancia { get; set; }
		public DateTime FechaLimiteEmbarque { get; set; }
		public DateTime FechaVencimiento { get; set; }
		public decimal ImporteLC { get; set; }
		public string InstruccionesEspeciales { get; set; }
		public string CartaCreditoId { get; set; }
	}
}