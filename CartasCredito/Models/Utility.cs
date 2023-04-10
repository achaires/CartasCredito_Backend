using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Utility
	{
		public static string HostUrl = ConfigurationManager.AppSettings.Get("HostUrl");
		public static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

		public static RespuestaFormato Login_GIS(string username, string password)
		{
			RespuestaFormato res = new RespuestaFormato();

			try
			{
				dll_Gis.Funciones fn = new dll_Gis.Funciones();
				string outMsg = "";
				fn.fs_LogIn(username, password, out outMsg);
				if (outMsg != "")
				{
					res.Flag = false;
					res.Description = outMsg;
				}
				else
				{
					res.Flag = true;
				}
			}
			catch (Exception ex)
			{
				res.Flag = false;
				res.Errors.Add(ex.Message);
				res.Description = "Ocurrió un error.";
			}
			finally
			{
			}

			return res;
		}
	}
}