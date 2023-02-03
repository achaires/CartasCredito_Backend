using dll_Gis;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class DataAccess
	{
		readonly BaseDatos bd = new BaseDatos();
		readonly String conexion = String.Empty;
		readonly Funciones fn = new Funciones();

		public DataAccess()
		{

			try
			{
				string connstringEnc = ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
				string connstring = fn.Desencriptar(connstringEnc);
				conexion = connstring;

			}
			catch (Exception)
			{
				throw;
			}

		}

		public DataAccess(string con)
		{

			try
			{
				//string connstringEnc = ConfigurationManager.ConnectionStrings["DefaultConnection"].ToString();
				string connstring = fn.Desencriptar(con);
				conexion = connstring;
			}
			catch (Exception)
			{
				throw;
			}

		}
	}
}