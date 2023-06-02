using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models
{
	public class PFETipoCambio
	{
		public int Id { get; set; }
		public int ProgramaId { get; set; }
		public int MonedaId { get; set; }
		public decimal PA { get; set; }
		public decimal PA1 { get; set; }
		public decimal PA2 { get; set; }
		public Moneda Moneda { get; set; }

		public static List<PFETipoCambio> GetByProgramaId(int id)
		{
			List<PFETipoCambio> rsp = new List<PFETipoCambio>();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Cons_PFETipoCambioByProgramaId(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						for (int i = 0; i < dt.Rows.Count; i++)
						{
							int idx = 0;
							var row = dt.Rows[i];
							var item = new PFETipoCambio();

							item.Id = int.Parse(row[idx].ToString()); idx++;
							item.ProgramaId = int.TryParse(row[idx].ToString(), out int progidval) ? progidval : 0; idx++;
							item.MonedaId = int.TryParse(row[idx].ToString(), out int mndidval) ? mndidval : 0; idx++;
							item.PA = decimal.TryParse(row[idx].ToString(), out decimal paval) ? paval : 0M; idx++;
							item.PA1 = decimal.TryParse(row[idx].ToString(), out decimal pa1val) ? pa1val : 0M; idx++;
							item.PA2 = decimal.TryParse(row[idx].ToString(), out decimal pa2val) ? pa2val : 0M; idx++;
							item.Moneda = Moneda.GetById(item.MonedaId);

							rsp.Add(item);
						}
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				rsp = new List<PFETipoCambio>();
			}

			return rsp;
		}

		public static RespuestaFormato DelByProgramaId(int id)
		{
			var rsp = new RespuestaFormato();

			try
			{
				DataAccess da = new DataAccess();
				var dt = new System.Data.DataTable();
				var errores = "";

				if (da.Del_PFETipoCambioByProgramaId(id, out dt, out errores))
				{
					if (dt.Rows.Count > 0)
					{
						rsp.Flag = true;
						rsp.DataString = "Registro eliminado";
					}
				}
			}
			catch (Exception ex)
			{
				Utility.Logger.Error(ex.Message);
				rsp.Flag = false;
				rsp.DataString = ex.Message;
				rsp.Errors.Add(ex.Message);
			}

			return rsp;
		}
	}
}