using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class CartaCreditoComisionInsertDTO
	{
		public string CartaCreditoId { get; set; }
		public int ComisionId { get; set; }
		public DateTime FechaCargo {get; set;}
		public int MonedaId {get; set;}
		public decimal Monto {get; set;}
		public int NumReferencia { get; set;}
	}
}