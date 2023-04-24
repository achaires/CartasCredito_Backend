using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class ConversionMonedaDTO
	{
		public string MonedaInput { get; set; }
		public string MonedaOutput { get; set; }
		public string Fecha { get; set; }
	}
}