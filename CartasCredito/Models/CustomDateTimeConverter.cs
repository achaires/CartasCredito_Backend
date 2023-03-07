using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class CustomDateTimeConverter : IsoDateTimeConverter
	{
		public CustomDateTimeConverter()
		{
			base.DateTimeFormat = "yyyy-MM-dd";
		}

	}
}