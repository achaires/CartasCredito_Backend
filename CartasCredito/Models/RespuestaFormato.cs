using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class RespuestaFormato
	{
		public bool Flag { get; set; }
		public string Description { get; set; }
		public List<string> Errors { get; set; }
		public ArrayList Content { get; set; }
		public int DataInt { get; set; }
		public string DataString { get; set; }
		public string Info { get; set; }
		public decimal DataDecimal { get; set; }
		public string DataString1 { get; set; }
		public string DataString2 { get; set; }
		public string DataString3 { get; set; }

		public RespuestaFormato()
		{
			Flag = false;
			Description = "No se ha ejecutado";
			Errors = new List<string>();
			Content = new ArrayList();
			DataInt = 0;
			Info = "";
			DataDecimal = 0;
			DataString = "";
			DataString1 = "";
			DataString2 = "";
			DataString3 = "";
		}
	}
}