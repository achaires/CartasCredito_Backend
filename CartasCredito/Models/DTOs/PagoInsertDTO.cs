using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class PagoInsertDTO
	{
		public string CartaCreditoId { get; set; }
		public string NumeroFactura { get; set; }
		public DateTime FechaVencimiento { get; set; }
		public decimal MontoPago { get; set; }

	}
}