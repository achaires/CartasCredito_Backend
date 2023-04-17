using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class BitacoraMovimientosFiltrarDTO
	{
		public DateTime DateStart { get; set; }
		public DateTime DateEnd { get; set; }
		public string CartaCreditoId { get; set; }
		public string UsuarioId { get; set; }

		public BitacoraMovimientosFiltrarDTO() {
			DateTime dateNow = DateTime.Now;
			DateStart = new DateTime(dateNow.Year, dateNow.Month, 1);
			DateEnd = dateNow;
		}
	}
}