using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class BitacoraMovimientosFiltrarDTO
	{
		public double DateStart { get; set; }
		public double DateEnd { get; set; }
		public string CartaCreditoId { get; set; }
		public string UsuarioId { get; set; }
	}
}