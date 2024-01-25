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

		public static decimal GetRateEx(int monedaIdIn, int monedaIdOut, DateTime fecha)
		{
			var rateEx = 1M;
			try
			{
				var monedaInDb = Moneda.GetById(monedaIdIn);
				var monedaOutDb = Moneda.GetById(monedaIdOut);

				var clnt = new ConversionMonedaService.BPELToolsClient();
				var req = new ConversionMonedaService.processRequest();
				var res = new ConversionMonedaService.processResponse();

				/*var timeoutSpan = new TimeSpan(0, 0, 1);
				clnt.Endpoint.Binding.CloseTimeout = timeoutSpan;
				clnt.Endpoint.Binding.OpenTimeout = timeoutSpan;
				clnt.Endpoint.Binding.ReceiveTimeout = timeoutSpan;
				clnt.Endpoint.Binding.SendTimeout = timeoutSpan;*/

				req.process = new ConversionMonedaService.process();
				req.process.P_USER_CONVERSION_TYPE = "Financiero Venta";
				req.process.P_CONVERSION_DATESpecified = true;
				req.process.P_CONVERSION_DATE = fecha;
				req.process.P_FROM_CURRENCY = monedaInDb.Abbr.Trim();
				req.process.P_TO_CURRENCY = monedaOutDb.Abbr.Trim();

				res = clnt.process(req.process);

				if (res.X_CONVERSION_RATE != null && res.X_MNS_ERROR == null)
				{
					rateEx = res.X_CONVERSION_RATE.Value;
				}

				Utility.Logger.Info("Conversion Rate " + res.X_CONVERSION_RATE.Value.ToString());
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
				rateEx = 1M;
			}

			return rateEx;
		}

		public static decimal GetTipoDeCambio(string monedaIdIn, string monedaIdOut, DateTime fecha)
		{
			var rateEx = 1M;
			try
			{
				var clnt = new ConversionMonedaService.BPELToolsClient();
				var req = new ConversionMonedaService.processRequest();
				var res = new ConversionMonedaService.processResponse();

				/*var timeoutSpan = new TimeSpan(0, 0, 1);
				clnt.Endpoint.Binding.CloseTimeout = timeoutSpan;
				clnt.Endpoint.Binding.OpenTimeout = timeoutSpan;
				clnt.Endpoint.Binding.ReceiveTimeout = timeoutSpan;
				clnt.Endpoint.Binding.SendTimeout = timeoutSpan;*/

				req.process = new ConversionMonedaService.process();
				req.process.P_USER_CONVERSION_TYPE = "Financiero Venta";
				req.process.P_CONVERSION_DATESpecified = true;
				req.process.P_CONVERSION_DATE = fecha;
				req.process.P_FROM_CURRENCY = monedaIdIn.Trim();
				req.process.P_TO_CURRENCY = monedaIdOut.Trim();

				res = clnt.process(req.process);

				if (res.X_CONVERSION_RATE != null && res.X_MNS_ERROR == null)
				{
					rateEx = res.X_CONVERSION_RATE.Value;
				}

				Utility.Logger.Info("Conversion Rate " + res.X_CONVERSION_RATE.Value.ToString());
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
				rateEx = -1;
			}

			return rateEx;
		}

		public static string ExcelColumnFromNumber(int column)
		{
			string columnString = "";
			decimal columnNumber = column;
			while (columnNumber > 0)
			{
				decimal currentLetterNumber = (columnNumber - 1) % 26;
				char currentLetter = (char)(currentLetterNumber + 65);
				columnString = currentLetter + columnString;
				columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
			}
			return columnString;
		}
	}
}