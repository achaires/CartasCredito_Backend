using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class PagoUpdateDTO
	{
		public int Id { get; set; }
		public DateTime FechaPago { get; set; }
		public decimal Monto { get; set; }
	}
}