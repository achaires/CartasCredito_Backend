using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class CartaCreditoAdjuntarSwiftDTO
	{
		public string NumCarta { get; set; }
		public HttpPostedFileBase SwiftFile { get; set; }

		public CartaCreditoAdjuntarSwiftDTO()
		{
			NumCarta = "";
			SwiftFile = null;
		}
	}
}