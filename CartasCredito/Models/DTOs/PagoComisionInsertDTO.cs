using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class PagoComisionInsertDTO
	{
		public string CartaCreditoId { get; set; }
		public DateTime FechaPago {get; set;}
		public decimal MontoPago {get; set;}
		public int ComisionId {get; set;}
		public int MonedaId {get; set;}
		public decimal TipoCambio {get; set;}
		public string Comentarios { get; set;}
	}
}