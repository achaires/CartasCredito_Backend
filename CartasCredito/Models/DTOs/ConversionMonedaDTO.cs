using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class ConversionMonedaDTO
	{
		public int MonedaInput { get; set; }
		public int MonedaOutput { get; set; }
		public string Fecha { get; set; }
	}
}