using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class Reporte
	{
		public int TipoReporteId { get; set; }

		public static List<CartaCredito> GetReporteAnalisisCartasCredito(DateTime fechaInicio, DateTime fechaFin, int empresaId = 0)
		{
			List<CartaCredito> res = new List<CartaCredito>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteAnalisisCartasCredito(out dt, out errores, fechaInicio, fechaFin, empresaId))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCredito();

							item.Consecutive = int.Parse(row[idx].ToString()); idx++;
							item.Id = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.TipoCartaId = int.TryParse(row[idx].ToString(), out int tipoCartaVal) ? tipoCartaVal : 0; idx++;
							item.TipoCarta = item.TipoCartaId == 17 ? "Comercial" : "StandBy";
							item.EmpresaId = int.TryParse(row[idx].ToString(), out int empidval) ? empidval : 0; idx++;
							item.BancoId = int.TryParse(row[idx].ToString(), out int bncidval) ? bncidval : 0; idx++;
							item.ProveedorId = int.TryParse(row[idx].ToString(), out int prvidval) ? prvidval : 0; idx++;
							item.MonedaId = int.TryParse(row[idx].ToString(), out int mndidval) ? mndidval : 0; idx++;
							item.DescripcionMercancia = row[idx].ToString(); idx++;
							item.PuntoEmbarque = row[idx].ToString(); idx++;
							item.DescripcionCartaCredito = row[idx].ToString(); idx++;
							item.MontoOriginalLC = decimal.TryParse(row[idx].ToString(), out decimal montooriginalval) ? montooriginalval : 0; idx++;
							item.DiasPlazoProveedor = int.TryParse(row[idx].ToString(), out int diasPlazoVal) ? diasPlazoVal : 0; idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.Banco = row[idx].ToString(); idx++;
							item.Proveedor = row[idx].ToString(); idx++;
							item.Moneda = row[idx].ToString(); idx++;

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				res = new List<CartaCredito>();

				// Get stack trace for the exception with source file information
				var st = new StackTrace(ex, true);
				// Get the top stack frame
				var frame = st.GetFrame(0);
				// Get the line number from the stack frame
				var line = frame.GetFileLineNumber();

				var errorMsg = ex.ToString();
			}

			return res;
		}

		public static List<CartaCreditoComision> GetReporteComisionesPorTipoComision(DateTime fechaInicio, DateTime fechaFin, int empresaId = 0)
		{
			List<CartaCreditoComision> res = new List<CartaCreditoComision>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ReporteComisionesPorTipoComision(out dt, out errores, fechaInicio, fechaFin, empresaId))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoComision();

							item.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;
							item.Empresa = row[idx].ToString(); idx++;
							item.Comision = row[idx].ToString(); idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.MonedaOriginal = row[idx].ToString(); idx++;
							item.Monto = decimal.TryParse(row[idx].ToString(), out decimal montoval) ? montoval : 0; idx++;
							item.MontoPagado = decimal.TryParse(row[idx].ToString(), out decimal montopagadoval) ? montopagadoval : 0; idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int estidval) ? estidval : 0; idx++;
							item.ComisionId = int.TryParse(row[idx].ToString(),out int comidval) ? comidval : 0; idx++;
							item.MonedaId = int.TryParse(row[idx].ToString(), out int midval) ? midval : 0; idx++;
							item.EstatusText = CartaCredito.GetStatusText(item.Estatus);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				res = new List<CartaCreditoComision>();

				// Get stack trace for the exception with source file information
				var st = new StackTrace(ex, true);
				// Get the top stack frame
				var frame = st.GetFrame(0);
				// Get the line number from the stack frame
				var line = frame.GetFileLineNumber();

				var errorMsg = ex.ToString();
			}

			return res;
		}
	}
}