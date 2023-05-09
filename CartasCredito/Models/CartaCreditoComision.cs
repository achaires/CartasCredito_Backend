using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class CartaCreditoComision
	{
		public int NumeroComision { get; set; }
		public int Id { get; set; }
		public string CartaCreditoId { get; set; }
		public int ComisionId { get; set; }
		// [JsonConverter(typeof(CustomDateTimeConverter))]
		public DateTime? FechaCargo { get; set; }
		//[JsonConverter(typeof(CustomDateTimeConverter))]
		public DateTime? FechaPago { get; set; }
		public int MonedaId { get; set; }
		public decimal TipoCambio { get; set; }
		public decimal Monto { get; set; }
		public decimal MontoPagado { get; set; }
		public bool Activo { get; set; }
		public string CreadoPor { get; set; }
		public int Estatus { get; set; }
		public string EstatusText { get; set; }
		public string EstatusClass { get; set; }
		public string Moneda { get; set; }
		public string Comision { get; set; }
		public int PagoId { get; set; }
		public int PagoMonedaId { get; set; }
		public string Comentarios { get; set; }

		// Para Reportes
		// public string Empresa { get; set; }
		public string NumCartaCredito { get; set; }
		// public string MonedaOriginal { get; set; }
		//public Pago Pago { get; set; }
		public int EstatusCartaId { get; set; }
		public string EstatusCartaText { get; set; }

		public static List<CartaCreditoComision> GetByCartaCreditoId(string cartaCreditoId)
		{
			List<CartaCreditoComision> res = new List<CartaCreditoComision>();

			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.Cons_ComisionesByCartaCreditoId(cartaCreditoId, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new CartaCreditoComision();

							item.NumeroComision = int.TryParse(row[idx].ToString(), out int numval) ? numval : 0; idx++;
							item.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;
							item.CartaCreditoId = row[idx].ToString(); idx++;
							item.ComisionId = int.TryParse(row[idx].ToString(), out int comisionIdVal) ? comisionIdVal : 0; idx++;
							item.FechaCargo = DateTime.TryParse(row[idx].ToString(), out DateTime fechaCargoVal) ? fechaCargoVal : DateTime.Parse("2000-01-01"); idx++;
							item.MonedaId = int.TryParse(row[idx].ToString(), out int monedaIdVal) ? monedaIdVal : 0; idx++;
							item.Monto = decimal.TryParse(row[idx].ToString(), out decimal montoVal) ? montoVal : 0; idx++;
							item.Activo = bool.TryParse(row[idx].ToString(), out bool activoVal) ? activoVal : false; idx++;
							item.CreadoPor = row[idx].ToString(); idx++;
							item.Estatus = int.TryParse(row[idx].ToString(), out int estatusVal) ? estatusVal : -1; idx++;
							item.Moneda = row[idx].ToString(); idx++;
							item.Comision = row[idx].ToString(); idx++;
							item.PagoId = int.TryParse(row[idx].ToString(), out int pagoidval) ? pagoidval : 0; idx++;
							item.PagoMonedaId = int.TryParse(row[idx].ToString(), out int pagoMonId) ? pagoMonId : 0; idx++;
							item.Comentarios = row[idx].ToString(); idx++;

							if (row[idx].ToString().Length > 0 )
							{
								item.FechaPago = DateTime.Parse(row[idx].ToString());
							} else
							{
								item.FechaPago = null;
							}
							idx++;


							item.TipoCambio = decimal.TryParse(row[idx].ToString(), out decimal tipoCambioVal) ? tipoCambioVal : 0; idx++;
							item.MontoPagado = decimal.TryParse(row[idx].ToString(), out decimal montoPVal) ? montoPVal : 0; idx++;
							item.NumCartaCredito = row[idx].ToString(); idx++;
							item.EstatusCartaId = int.TryParse(row[idx].ToString(), out int estcartaid) ? estcartaid : 0; idx++;
							item.EstatusCartaText = CartaCredito.GetStatusText(item.EstatusCartaId);
							item.EstatusText = GetStatusText(item.Estatus);
							item.EstatusClass = GetStatusClass(item.Estatus);

							res.Add(item);
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.Write(ex);
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

		public static RespuestaFormato Insert(CartaCreditoComision modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Ins_CartaCreditoComision(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						int id = 0;
						Int32.TryParse(row[3].ToString(), out id);

						if (id > 0)
						{
							rsp.Flag = true;
							rsp.DataInt = id;
						}
					}
				}
				else
				{
					rsp.Description = "Ocurrió un error";
					rsp.Errors.Add(errores);
				}
			}
			catch (Exception ex)
			{
				rsp.Errors.Add(ex.Message);
				rsp.Description = "Ocurrió un error";
			}

			return rsp;
		}

		public static CartaCreditoComision GetById(string id)
		{
			CartaCreditoComision rsp = new CartaCreditoComision();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_CartaComisionById(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var idx = 0;
						var row = dt.Rows[0];

						rsp.Id = int.TryParse(row[idx].ToString(), out int idval) ? idval : 0; idx++;
						rsp.CartaCreditoId = row[idx].ToString(); idx++;
						rsp.ComisionId = int.TryParse(row[idx].ToString(), out int comisionIdVal) ? comisionIdVal : 0; idx++;
						rsp.FechaCargo = DateTime.TryParse(row[idx].ToString(), out DateTime fechaCargoVal) ? fechaCargoVal : DateTime.Parse("2000-01-01"); idx++;
						rsp.MonedaId = int.TryParse(row[idx].ToString(), out int monedaIdVal) ? monedaIdVal : 0; idx++;
						rsp.Monto = decimal.TryParse(row[idx].ToString(), out decimal montoVal) ? montoVal : 0; idx++;
						rsp.Activo = bool.TryParse(row[idx].ToString(), out bool activoVal) ? activoVal : false; idx++;
						rsp.CreadoPor = row[idx].ToString(); idx++;
						rsp.Estatus = int.TryParse(row[idx].ToString(), out int estatusVal) ? estatusVal : -1; idx++;
						rsp.Moneda = row[idx].ToString(); idx++;
						rsp.Comision = row[idx].ToString(); idx++;

					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new CartaCreditoComision();
			}

			return rsp;
		}

		public static string GetStatusText(int estatusInt)
		{
			string rsp = "";

			switch (estatusInt)
			{
				case 4:
					rsp = "Refinanciado";
					break;

				case 1:
					rsp = "Pendiente";
					break;

				case 3:
					rsp = "Pagado";
					break;
			}

			return rsp;
		}

		public static string GetStatusClass(int estatusInt)
		{
			string rsp = "";

			switch (estatusInt)
			{
				case 4:
					rsp = "";
					break;

				case 1:
					rsp = "bg-yellow-100";
					break;

				case 3:
					rsp = "bg-green-300";
					break;
			}

			return rsp;
		}

	}
}