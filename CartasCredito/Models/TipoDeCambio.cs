using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace CartasCredito.Models
{
	public class TipoDeCambio
	{
		public int Id { get; set; }
		public string MonedaOriginal { get; set; }
		public string MonedaNueva { get; set; }
		public decimal Conversion { get; set; }
		public string ConversionStr { get; set; }
		public DateTime Fecha { get; set; }
		public string Fecha_str { get; set; }

		public TipoDeCambio()
		{
			Id = 0;
			MonedaOriginal = "";
			MonedaNueva = "";
			Conversion = -1;
			ConversionStr = "-1";
			Fecha_str = "";
		}

		public RespuestaFormato Get()
		{
			RespuestaFormato res = new RespuestaFormato();
			try
			{
				DataAccess da = new DataAccess();

				var dt = new System.Data.DataTable();
				var errores = "";
				if (da.cons_TiposDeCambio_AlDia(this, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							var idx = 0;
							var row = dt.Rows[0];
							this.ConversionStr = row[idx].ToString(); idx++;
						}
						res.Flag = true;
					}
                    else
                    {
						this.ConversionStr = "-1";
						this.Conversion = -1;
						res.Flag = false;
						res.Description = "No hay información";
					}
				}
				else
				{
					this.ConversionStr = "-1";
					this.Conversion = -1;
					res.Flag = false;
					res.Description = errores;
				}

			}
			catch (Exception ex)
			{
				this.ConversionStr = "-1";
				this.Conversion = -1;
				res.Flag = false;
				res.Description = ex.Message;
			}

			return res;
		}

		public static RespuestaFormato Insert(TipoDeCambio modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.ins_cat_TipoDeCambio(modelo, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						var row = dt.Rows[0];
						int id = 0;
						Int32.TryParse(row[3].ToString(), out id);

						if (id > 0)
						{
							rsp.Description = "Modelo insertado";
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

		public static RespuestaFormato Update(Banco modelo)
		{
			RespuestaFormato rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Upd_Banco(modelo, out dt, out errores))
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

		public static List<TipoDeCambio> TiposDeCambioAlDia(DateTime fechaDivisa)
		{
			RespuestaFormato rsp = new RespuestaFormato();
			List<TipoDeCambio> tiposDeCambio = new List<TipoDeCambio>();
			try
			{
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "JPY",
					MonedaNueva = "USD"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "MXP",
					MonedaNueva = "USD"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "EUR",
					MonedaNueva = "USD"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "USD",
					MonedaNueva = "USD"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "RMB",
					MonedaNueva = "USD"
				});

				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "JPY",
					MonedaNueva = "MXP"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "MXP",
					MonedaNueva = "MXP"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "EUR",
					MonedaNueva = "MXP"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "USD",
					MonedaNueva = "MXP"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "RMB",
					MonedaNueva = "MXP"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "JPY",
					MonedaNueva = "JPY"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "EUR",
					MonedaNueva = "EUR"
				});
				tiposDeCambio.Add(new TipoDeCambio()
				{
					Fecha = fechaDivisa,
					MonedaOriginal = "RMB",
					MonedaNueva = "RMB"
				});

				for (int tc_idx = 0; tc_idx < tiposDeCambio.Count; tc_idx++)
                {
					tiposDeCambio[tc_idx].Get();

					if (tiposDeCambio[tc_idx].ConversionStr == "-1")
					{
						tiposDeCambio[tc_idx].Conversion = Utility.GetTipoDeCambio(tiposDeCambio[tc_idx].MonedaOriginal, tiposDeCambio[tc_idx].MonedaNueva, tiposDeCambio[tc_idx].Fecha);
						if (tiposDeCambio[tc_idx].Conversion > -1)
						{
							TipoDeCambio.Insert(tiposDeCambio[tc_idx]);
						}
					}
					else
					{
						tiposDeCambio[tc_idx].Conversion = Decimal.Parse(tiposDeCambio[tc_idx].ConversionStr);
					}
				}

				if(tiposDeCambio.Where(i=>i.Conversion == -1).Count() > 0)
                {
					tiposDeCambio = new List<TipoDeCambio>();
                }
			}
			catch (Exception ex)
			{
				rsp.Errors.Add(ex.Message);
				rsp.Description = "Ocurrió un error";
				tiposDeCambio = new List<TipoDeCambio>();
			}

			return tiposDeCambio;
		}

	}
}