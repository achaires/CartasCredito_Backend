using dll_Gis;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using CartasCredito.Models.DTOs;

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

		#region Divisiones
		public Boolean Cons_Divisiones(out DataTable dt, out String msgError, int activo = -1)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Divisiones", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Division(Division modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];

				@params[0] = new SqlParameter("@Nombre", modelo.Nombre);
				@params[1] = new SqlParameter("@Descripcion", modelo.Descripcion);
				@params[2] = new SqlParameter("@CreadoPor", modelo.CreadoPor);

				if (!bd.ExecuteProcedure(conexion, "ins_Division", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_DivisionById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_DivisionById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Division(Division modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];

				@params[0] = new SqlParameter("@Id", modelo.Id);
				@params[1] = new SqlParameter("@Nombre", modelo.Nombre);
				@params[2] = new SqlParameter("@Descripcion", modelo.Descripcion);
				@params[3] = new SqlParameter("@Activo", modelo.Activo);

				if (!bd.ExecuteProcedure(conexion, "upd_Division", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Empresas
		public Boolean Cons_Empresas(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Empresas", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Empresa(Empresa modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@DivisionId", modelo.DivisionId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@RFC", modelo.RFC); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Empresa", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_EmpresaById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_EmpresaById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}


		public Boolean Cons_EmpresasByUserId(string id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@UserId", id);

				if (!bd.ExecuteProcedure(conexion, "cons_EmpresasByUserId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Empresa(Empresa modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[6];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@DivisionId", modelo.DivisionId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@RFC", modelo.RFC); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Empresa", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Bancos
		public Boolean Cons_Bancos(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Bancos", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Banco(Banco modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@TotalLinea", modelo.TotalLinea); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Banco", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_BancoById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_BancoById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Banco(Banco modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@TotalLinea", modelo.TotalLinea); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Banco", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region TiposCobertura
		public Boolean Cons_TiposCobertura(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_TiposCobertura", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_TipoCobertura(TipoCobertura modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_TipoCobertura", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_TipoCoberturaById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_TipoCoberturaById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_TipoCobertura(TipoCobertura modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_TipoCobertura", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region TiposComision
		public Boolean Cons_TiposComision(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_TiposComision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_TipoComision(TipoComision modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_TipoComision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_TipoComisionById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_TipoComisionById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_TipoComision(TipoComision modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_TipoComision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region TiposPersonaFiscal
		public Boolean Cons_TiposPersonaFiscal(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_TiposPersonaFiscal", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_TipoPersonaFiscal(TipoPersonaFiscal modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_TipoPersonaFiscal", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_TipoPersonaFiscalById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_TipoPersonaFiscalById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_TipoPersonaFiscal(TipoPersonaFiscal modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_TipoPersonaFiscal", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region AgentesAduanales
		public Boolean Cons_AgentesAduanales(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_AgentesAduanales", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_AgenteAduanal(AgenteAduanal modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_AgenteAduanal", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_AgenteAduanalById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_AgenteAduanalById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_AgenteAduanal(AgenteAduanal modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_AgenteAduanal", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Comisiones
		public Boolean Cons_Comisiones(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Comisiones", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Comision(Comision modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[9];
				var pix = 0;

				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@TipoComisionId", modelo.TipoComisionId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;
				@params[pix] = new SqlParameter("@Costo", modelo.Costo); pix++;
				@params[pix] = new SqlParameter("@SwiftApertura", modelo.SwiftApertura); pix++;
				@params[pix] = new SqlParameter("@SwiftOtro", modelo.SwiftOtro); pix++;
				@params[pix] = new SqlParameter("@PorcentajeIVA", modelo.PorcentajeIVA); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Comision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_ComisionById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_ComisionById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Comision(Comision modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[10];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@TipoComisionId", modelo.TipoComisionId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Costo", modelo.Costo); pix++;
				@params[pix] = new SqlParameter("@SwiftApertura", modelo.SwiftApertura); pix++;
				@params[pix] = new SqlParameter("@SwiftOtro", modelo.SwiftOtro); pix++;
				@params[pix] = new SqlParameter("@PorcentajeIVA", modelo.PorcentajeIVA); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Comision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Compradores
		public Boolean Cons_Compradores(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Compradores", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Comprador(Comprador modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@TipoPersonaFiscalId", modelo.TipoPersonaFiscalId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Comprador", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_CompradorById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_CompradorById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Comprador(Comprador modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[6];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@TipoPersonaFiscalId", modelo.TipoPersonaFiscalId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Comprador", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Documentos
		public Boolean Cons_Documentos(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Documentos", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Documento(Documento modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Documento", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_DocumentoById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_DocumentoById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Documento(Documento modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Documento", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Monedas
		public Boolean Cons_Monedas(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Monedas", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Moneda(Moneda modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Abbr", modelo.Abbr); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Moneda", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_MonedaById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_MonedaById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Moneda(Moneda modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Abbr", modelo.Abbr); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Moneda", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Proveedores
		public Boolean Cons_Proveedores(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Proveedores", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Proveedor(Proveedor modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@PaisId", modelo.PaisId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Proveedor", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_ProveedorById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_ProveedorById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Proveedor(Proveedor modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[6];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@PaisId", modelo.PaisId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Proveedor", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Paises
		public Boolean Cons_Paises(out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[0];
				//@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Paises", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Pais(Pais modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Alpha2Code", modelo.Alpha2Code); pix++;
				@params[pix] = new SqlParameter("@Alpha3Code", modelo.Alpha3Code); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Pais", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		/*
		public Boolean Cons_PaisById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_PaisById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Pais(Pais modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Alpha2Code", modelo.Alpha2Code); pix++;
				@params[pix] = new SqlParameter("@Alpha3Code", modelo.Alpha3Code); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Pais", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		*/
		#endregion

		#region Proyectos
		public Boolean Cons_Proyectos(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Proyectos", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Proyecto(Proyecto modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[6];
				var pix = 0;

				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@FechaApertura", modelo.FechaApertura); pix++;
				@params[pix] = new SqlParameter("@FechaCierre", modelo.FechaCierre); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Proyecto", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_ProyectoById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_ProyectoById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Proyecto(Proyecto modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[7];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@FechaApertura", modelo.FechaApertura); pix++;
				@params[pix] = new SqlParameter("@FechaCierre", modelo.FechaCierre); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Proyecto", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region TiposActivo
		public Boolean Cons_TiposActivo(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_TiposActivo", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_TipoActivo(TipoActivo modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Responsable", modelo.Responsable); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_TipoActivo", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_TipoActivoById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_TipoActivoById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_TipoActivo(TipoActivo modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Responsable", modelo.Responsable); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_TipoActivo", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region LineasDeCredito
		public Boolean Cons_LineasDeCredito(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_LineasDeCredito", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_LineaDeCredito(LineaDeCredito modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@Cuenta", modelo.Cuenta); pix++;
				@params[pix] = new SqlParameter("@Monto", modelo.Monto); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_LineaDeCredito", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_LineaDeCreditoById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_LineaDeCreditoById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_LineaDeCredito(LineaDeCredito modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[6];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@Monto", modelo.Monto); pix++;
				@params[pix] = new SqlParameter("@Cuenta", modelo.Cuenta); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_LineaDeCredito", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region Contactos
		public Boolean Cons_Contactos(out DataTable dt, out String msgError, int activo = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_Contactos", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_Contacto(Contacto modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[10];
				var pix = 0;

				@params[pix] = new SqlParameter("@ModelNombre", modelo.ModelNombre); pix++;
				@params[pix] = new SqlParameter("@ModelId", modelo.ModelId); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@ApellidoPaterno", modelo.ApellidoPaterno); pix++;
				@params[pix] = new SqlParameter("@ApellidoMaterno", modelo.ApellidoMaterno); pix++;
				@params[pix] = new SqlParameter("@Telefono", modelo.Telefono); pix++;
				@params[pix] = new SqlParameter("@Fax", modelo.Fax); pix++;
				@params[pix] = new SqlParameter("@Email", modelo.Email); pix++;
				@params[pix] = new SqlParameter("@Celular", modelo.Celular); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Contacto", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_ContactoById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_ContactoById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_ContactoByModelNombreAndId(int modelId, string modelNombre, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];
				@params[0] = new SqlParameter("@ModelId", modelId);
				@params[1] = new SqlParameter("@ModelNombre", modelNombre);

				if (!bd.ExecuteProcedure(conexion, "cons_ContactoByModelNombreAndId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_Contacto(Contacto modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[10];
				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Nombre", modelo.Nombre); pix++;
				@params[pix] = new SqlParameter("@ApellidoPaterno", modelo.ApellidoPaterno); pix++;
				@params[pix] = new SqlParameter("@ApellidoMaterno", modelo.ApellidoMaterno); pix++;
				@params[pix] = new SqlParameter("@Telefono", modelo.Telefono); pix++;
				@params[pix] = new SqlParameter("@Fax", modelo.Fax); pix++;
				@params[pix] = new SqlParameter("@Email", modelo.Email); pix++;
				@params[pix] = new SqlParameter("@Celular", modelo.Celular); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Contacto", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region CartasdeCredito
		public Boolean Ins_CartaCredito(CartaCredito modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[39];
				int pix = 0;

				@params[pix] = new SqlParameter("@TipoCarta", modelo.TipoCarta); pix++;
				@params[pix] = new SqlParameter("@TipoActivoId", modelo.TipoActivoId); pix++;
				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@ProyectoId", modelo.ProyectoId); pix++;
				@params[pix] = new SqlParameter("@ProveedorId", modelo.ProveedorId); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@AgenteAduanalId", modelo.AgenteAduanalId); pix++;
				@params[pix] = new SqlParameter("@MonedaId", modelo.MonedaId); pix++;
				@params[pix] = new SqlParameter("@TipoPago", modelo.TipoPago); pix++;
				@params[pix] = new SqlParameter("@Responsable", modelo.Responsable); pix++;
				@params[pix] = new SqlParameter("@CompradorId", modelo.CompradorId); pix++;
				@params[pix] = new SqlParameter("@PorcTolerancia", modelo.PorcentajeTolerancia); pix++;
				@params[pix] = new SqlParameter("@NumeroOrdenCompra", modelo.NumOrdenCompra); pix++;
				@params[pix] = new SqlParameter("@CostoApertura", modelo.CostoApertura); pix++;
				@params[pix] = new SqlParameter("@MontoOrdenCompra", modelo.MontoOrdenCompra); pix++;
				@params[pix] = new SqlParameter("@MontoOriginal", modelo.MontoOriginalLC); pix++;
				@params[pix] = new SqlParameter("@FechaApertura", modelo.FechaApertura.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@Incoterm", modelo.Incoterm); pix++;
				@params[pix] = new SqlParameter("@FechaLimiteEmbarque", modelo.FechaLimiteEmbarque.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@EmbarquesParciales", modelo.EmbarquesParciales); pix++;
				@params[pix] = new SqlParameter("@Transbordos", modelo.Transbordos); pix++;
				@params[pix] = new SqlParameter("@PuntoEmbarque", modelo.PuntoEmbarque); pix++;
				@params[pix] = new SqlParameter("@PuntoDesembarque", modelo.PuntoDesembarque); pix++;
				@params[pix] = new SqlParameter("@DescripcionMercancia", modelo.DescripcionMercancia); pix++;
				@params[pix] = new SqlParameter("@DescripcionCartaCredito", modelo.DescripcionCartaCredito); pix++;
				@params[pix] = new SqlParameter("@PagoCartaAceptacion", modelo.PagoCartaAceptacion); pix++;
				@params[pix] = new SqlParameter("@ConsignacionMercancia", modelo.ConsignacionMercancia); pix++;
				@params[pix] = new SqlParameter("@ConsideracionesAdicionales", modelo.ConsideracionesAdicionales); pix++;
				@params[pix] = new SqlParameter("@DiasPresentarDocumentos", modelo.DiasParaPresentarDocumentos); pix++;
				@params[pix] = new SqlParameter("@DiasPlazoProveedor", modelo.DiasPlazoProveedor); pix++;
				@params[pix] = new SqlParameter("@CondicionesPago", modelo.CondicionesPago); pix++;
				@params[pix] = new SqlParameter("@NumeroPeriodos", modelo.NumeroPeriodos); pix++;
				@params[pix] = new SqlParameter("@BancoCorresponsalId", modelo.BancoCorresponsalId); pix++;
				@params[pix] = new SqlParameter("@SeguroPorCuenta", modelo.SeguroPorCuenta); pix++;
				@params[pix] = new SqlParameter("@GastosComisionesCorresponsal", modelo.GastosComisionesCorresponsal); pix++;
				@params[pix] = new SqlParameter("@ConfirmacionBancoNotificador", modelo.ConfirmacionBancoNotificador); pix++;
				@params[pix] = new SqlParameter("@TipoEmision", modelo.TipoEmision); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;


				if (!bd.ExecuteProcedure(conexion, "ins_CartaCredito", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{


					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
					else
					{
						string ccId = dt.Rows[0][3].ToString();

						//Ins_CartaCreditoDocumentos(ccId, modelo.DocumentosRequeridosIds, out dt, out msgError);
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Ins_CartaCreditoStandBy(CartaCredito modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[13];
				int pix = 0;

				@params[pix] = new SqlParameter("@TipoCarta", modelo.TipoCartaId); pix++;
				@params[pix] = new SqlParameter("@TipoStandBy", modelo.TipoStandBy); pix++;
				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@MonedaId", modelo.MonedaId); pix++;
				@params[pix] = new SqlParameter("@CompradorId", modelo.CompradorId); pix++;
				@params[pix] = new SqlParameter("@MontoOriginal", modelo.MontoOriginalLC); pix++;
				@params[pix] = new SqlParameter("@FechaApertura", modelo.FechaApertura.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaLimiteEmbarque", modelo.FechaLimiteEmbarque.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@ConsideracionesReclamacion", modelo.ConsideracionesReclamacion); pix++;
				@params[pix] = new SqlParameter("@ConsideracionesAdicionales", modelo.ConsideracionesAdicionales); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_CartaStandBy", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{


					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
					else
					{
						string ccId = dt.Rows[0][3].ToString();

						//Ins_CartaCreditoDocumentos(ccId, modelo.DocumentosRequeridosIds, out dt, out msgError);
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_CartaCredito(CartaCredito modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				SqlParameter[] @params = new SqlParameter[42];

				@params[pix] = new SqlParameter("@CartaCreditoId", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Estatus", modelo.Estatus); pix++;
				@params[pix] = new SqlParameter("@TipoCarta", modelo.TipoCartaId); pix++;
				@params[pix] = new SqlParameter("@TipoActivoId", modelo.TipoActivoId); pix++;
				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@ProyectoId", modelo.ProyectoId); pix++;
				@params[pix] = new SqlParameter("@ProveedorId", modelo.ProveedorId); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@AgenteAduanalId", modelo.AgenteAduanalId); pix++;
				@params[pix] = new SqlParameter("@MonedaId", modelo.MonedaId); pix++;
				@params[pix] = new SqlParameter("@TipoPago", modelo.TipoPago); pix++;
				@params[pix] = new SqlParameter("@Responsable", modelo.Responsable); pix++;
				@params[pix] = new SqlParameter("@CompradorId", modelo.CompradorId); pix++;
				@params[pix] = new SqlParameter("@PorcTolerancia", modelo.PorcentajeTolerancia); pix++;
				@params[pix] = new SqlParameter("@NumeroOrdenCompra", modelo.NumOrdenCompra); pix++;
				@params[pix] = new SqlParameter("@CostoApertura", modelo.CostoApertura); pix++;
				@params[pix] = new SqlParameter("@MontoOrdenCompra", modelo.MontoOrdenCompra); pix++;
				@params[pix] = new SqlParameter("@MontoOriginal", modelo.MontoOriginalLC); pix++;
				@params[pix] = new SqlParameter("@FechaApertura", modelo.FechaApertura.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@Incoterm", modelo.Incoterm); pix++;
				@params[pix] = new SqlParameter("@FechaLimiteEmbarque", modelo.FechaLimiteEmbarque.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@EmbarquesParciales", modelo.EmbarquesParciales); pix++;
				@params[pix] = new SqlParameter("@Transbordos", modelo.Transbordos); pix++;
				@params[pix] = new SqlParameter("@PuntoEmbarque", modelo.PuntoEmbarque); pix++;
				@params[pix] = new SqlParameter("@PuntoDesembarque", modelo.PuntoDesembarque); pix++;
				@params[pix] = new SqlParameter("@DescripcionMercancia", modelo.DescripcionMercancia); pix++;
				@params[pix] = new SqlParameter("@DescripcionCartaCredito", modelo.DescripcionCartaCredito); pix++;
				@params[pix] = new SqlParameter("@PagoCartaAceptacion", modelo.PagoCartaAceptacion); pix++;
				@params[pix] = new SqlParameter("@ConsignacionMercancia", modelo.ConsignacionMercancia); pix++;
				@params[pix] = new SqlParameter("@ConsideracionesAdicionales", modelo.ConsideracionesAdicionales); pix++;
				@params[pix] = new SqlParameter("@DiasPresentarDocumentos", modelo.DiasParaPresentarDocumentos); pix++;
				@params[pix] = new SqlParameter("@DiasPlazoProveedor", modelo.DiasPlazoProveedor); pix++;
				@params[pix] = new SqlParameter("@CondicionesPago", modelo.CondicionesPago); pix++;
				@params[pix] = new SqlParameter("@BancoCorresponsalId", modelo.BancoCorresponsalId); pix++;
				@params[pix] = new SqlParameter("@SeguroPorCuenta", modelo.SeguroPorCuenta); pix++;
				@params[pix] = new SqlParameter("@GastosComisionesCorresponsal", modelo.GastosComisionesCorresponsal); pix++;
				@params[pix] = new SqlParameter("@ConfirmacionBancoNotificador", modelo.ConfirmacionBancoNotificador); pix++;
				@params[pix] = new SqlParameter("@TipoEmision", modelo.TipoEmision); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;
				@params[pix] = new SqlParameter("@InstruccionesEspeciales", modelo.InstruccionesEspeciales); pix++;


				if (!bd.ExecuteProcedure(conexion, "upd_CartaCredito", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{


					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
					else
					{
						string ccId = dt.Rows[0][3].ToString();

						//Ins_CartaCreditoDocumentos(ccId, modelo.DocumentosRequeridosIds, out dt, out msgError);
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_StandBy(CartaCredito modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				SqlParameter[] @params = new SqlParameter[15];

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@TipoCarta", modelo.TipoCartaId); pix++;
				@params[pix] = new SqlParameter("@TipoStandBy", modelo.TipoStandBy); pix++;
				@params[pix] = new SqlParameter("@BancoId", modelo.BancoId); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", modelo.EmpresaId); pix++;
				@params[pix] = new SqlParameter("@MonedaId", modelo.MonedaId); pix++;
				@params[pix] = new SqlParameter("@CompradorId", modelo.CompradorId); pix++;
				@params[pix] = new SqlParameter("@MontoOriginal", modelo.MontoOriginalLC); pix++;
				@params[pix] = new SqlParameter("@FechaApertura", modelo.FechaApertura.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaLimiteEmbarque", modelo.FechaLimiteEmbarque.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@ConsideracionesReclamacion", modelo.ConsideracionesReclamacion); pix++;
				@params[pix] = new SqlParameter("@ConsideracionesAdicionales", modelo.ConsideracionesAdicionales); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;
				@params[pix] = new SqlParameter("@Estatus", modelo.Estatus); pix++;


				if (!bd.ExecuteProcedure(conexion, "upd_CartaStandBy", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{


					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
					else
					{
						string ccId = dt.Rows[0][3].ToString();

						//Ins_CartaCreditoDocumentos(ccId, modelo.DocumentosRequeridosIds, out dt, out msgError);
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_CartaCreditoSwift(string ccid, string numeroCartaComercial, string swiftFilename, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				SqlParameter[] @params = new SqlParameter[3];
				@params[pix] = new SqlParameter("@NumeroCarta", numeroCartaComercial); pix++;
				@params[pix] = new SqlParameter("@SwiftFilename", swiftFilename); pix++;
				@params[pix] = new SqlParameter("@CartaComercialId", ccid); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_CartaCreditoSwift", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
					else
					{
						string ccId = dt.Rows[0][3].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_CartaCreditoStatus(string ccid, int newStatus, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				SqlParameter[] @params = new SqlParameter[2];
				@params[pix] = new SqlParameter("@Estatus", newStatus); pix++;
				@params[pix] = new SqlParameter("@CartaComercialId", ccid); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_CartaCreditoStatus", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
					else
					{
						string ccId = dt.Rows[0][3].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_CartasCredito(out DataTable dt, out String msgError, DateTime fechaInicio, DateTime fechaFin, int activo = -1)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				
				SqlParameter[] @params = new SqlParameter[3];
				@params[pix] = new SqlParameter("@Activo", activo); pix++;
				@params[pix] = new SqlParameter("@FechaInicio", fechaInicio.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaFin", fechaFin.ToString("yyyy-MM-dd HH:mm:ss")); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_CartasCredito", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_CartasCreditoFiltrar(CartasCreditoFiltrarDTO model, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				SqlParameter[] @params = new SqlParameter[10];
				@params[pix] = new SqlParameter("@FechaInicio", model.FechaInicio.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaFin", model.FechaFin.ToString("yyyy-MM-dd HH:mm:ss")); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_CartasCreditoFiltrar", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_ReporteCartas(int empresaId, DateTime fechaInicio, DateTime fechaFin, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				SqlParameter[] @params = new SqlParameter[3];
				@params[pix] = new SqlParameter("@EmpresaId", empresaId); pix++;
				@params[pix] = new SqlParameter("@FechaInicio", fechaInicio.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaFin", fechaFin.ToString("yyyy-MM-dd HH:mm:ss")); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_ReporteCartas", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_CartaCreditoById(string id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_CartaById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_SwiftByCartaCreditoId(string id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@CartaCreditoId", id);

				if (!bd.ExecuteProcedure(conexion, "cons_SwiftByCartaCreditoId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}
		#endregion

		#region Pagos
		public Boolean Ins_Pago(Pago modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				int pix = 0;

				@params[pix] = new SqlParameter("@NumeroFactura", modelo.NumeroFactura); pix++;
				@params[pix] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@MontoPago", modelo.MontoPago); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;
				@params[pix] = new SqlParameter("@CartaCreditoId", modelo.CartaCreditoId); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Pago", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_Pago(Pago modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[8];
				int pix = 0;
				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@NumeroFactura", modelo.NumeroFactura); pix++;
				@params[pix] = new SqlParameter("@FechaPago", modelo.FechaPago?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@MontoPago", modelo.MontoPago); pix++;
				@params[pix] = new SqlParameter("@MontoPagado", modelo.MontoPagado); pix++;
				@params[pix] = new SqlParameter("@Estatus", modelo.Estatus); pix++;
				@params[pix] = new SqlParameter("@CartaCreditoId", modelo.CartaCreditoId); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Pago", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_PagosByCartaCreditoId(string cartaCreditoId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@CartaCreditoId", cartaCreditoId);

				if (!bd.ExecuteProcedure(conexion, "cons_PagosByCartaCreditoId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_PagosProgramados(out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[0];

				if (!bd.ExecuteProcedure(conexion, "cons_PagosProgramados", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_PagosVencidos(out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[0];

				if (!bd.ExecuteProcedure(conexion, "cons_PagosVencidos", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_Pagos(out DataTable dt, out String msgError, DateTime dateStart, DateTime dateEnd)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				var pix = 0;
				
				SqlParameter[] @params = new SqlParameter[2];
				@params[pix] = new SqlParameter("FechaInicio", dateStart.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("FechaFin", dateEnd.ToString("yyyy-MM-dd HH:mm:ss")); pix++;


				if (!bd.ExecuteProcedure(conexion, "cons_Pagos", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_PagoById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_PagoById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}
		#endregion

		#region BitacoraMovimientos
		public Boolean Ins_BitacoraMovimiento(BitacoraMovimiento modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[7];
				var pix = 0;

				@params[pix] = new SqlParameter("@Titulo", modelo.Titulo); pix++;
				@params[pix] = new SqlParameter("@Descripcion", modelo.Descripcion); pix++;
				@params[pix] = new SqlParameter("@CartaCreditoId", modelo.CartaCreditoId); pix++;
				@params[pix] = new SqlParameter("@ModeloNombre", modelo.ModeloNombre); pix++;
				@params[pix] = new SqlParameter("@ModeloId", modelo.ModeloId); pix++;
				@params[pix] = new SqlParameter("@CreadoPorId", modelo.CreadoPorId); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_BitacoraMovimiento", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_BitacoraMovimientos(DateTime dateStart, DateTime dateEnd, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];
				var pix = 0;

				@params[pix] = new SqlParameter("@FechaInicio", dateStart.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaFin", dateEnd.ToString("yyyy-MM-dd HH:mm:ss")); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_BitacoraMovimientos", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region CartaComisiones
		public Boolean Cons_ComisionesByCartaCreditoId(string cartaCreditoId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@CartaCreditoId", cartaCreditoId);

				if (!bd.ExecuteProcedure(conexion, "cons_ComisionesByCartaCreditoId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_CartaCreditoComision(CartaCreditoComision modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[7];

				var pix = 0;

				@params[pix] = new SqlParameter("@CartaCreditoId", modelo.CartaCreditoId); pix++;
				@params[pix] = new SqlParameter("@ComisionId", modelo.ComisionId); pix++;
				@params[pix] = new SqlParameter("@FechaCargo", modelo.FechaCargo?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@MonedaId", modelo.MonedaId); pix++;
				@params[pix] = new SqlParameter("@Monto", modelo.Monto); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;
				@params[pix] = new SqlParameter("@NumReferencia", modelo.PagoId); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_CartaCreditoComision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_CartaComisionById(string id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];

				int i = 0;
				@params[0] = new SqlParameter("@Id", id);

				i++;
				if (!bd.ExecuteProcedure(conexion, "cons_CartaComisionById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos a mostrar";
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Ins_PagoComision(Pago modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[7];
				int pix = 0;

				@params[pix] = new SqlParameter("@FechaPago", modelo.FechaPago?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@MontoPago", modelo.MontoPago); pix++;
				@params[pix] = new SqlParameter("@ComisionId", modelo.ComisionId); pix++;
				@params[pix] = new SqlParameter("@MonedaId", modelo.MonedaId); pix++;
				@params[pix] = new SqlParameter("@TipoCambio", modelo.TipoCambio); pix++;
				@params[pix] = new SqlParameter("@Estatus", modelo.Estatus); pix++;
				@params[pix] = new SqlParameter("@Comentarios", modelo.Comentarios); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_PagoComision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_PagoByComisionId(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_PagoByComisionId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}
		#endregion

		#region Enmiendas
		public Boolean Upd_Enmienda(Enmienda modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;

				SqlParameter[] @params = new SqlParameter[17];

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				//@params[pix] = new SqlParameter("@DocumentoSwift", modelo.DocumentoSwift); pix++;
				@params[pix] = new SqlParameter("@CartaCreditoId", modelo.CartaCreditoId); pix++;
				@params[pix] = new SqlParameter("@ImporteLC", modelo.ImporteLC); pix++;
				@params[pix] = new SqlParameter("@FechaLimiteEmbarque", modelo.FechaLimiteEmbarque?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@DescripcionMercancia", modelo.DescripcionMercancia); pix++;
				@params[pix] = new SqlParameter("@ConsideracionesAdicionales", modelo.ConsideracionesAdicionales); pix++;
				@params[pix] = new SqlParameter("@InstruccionesEspeciales", modelo.InstruccionesEspeciales); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;
				@params[pix] = new SqlParameter("@Estatus", modelo.Estatus); pix++;
				@params[pix] = new SqlParameter("@Prev_ImporteLC", modelo.Prev_ImporteLC); pix++;
				@params[pix] = new SqlParameter("@Prev_FechaLimiteEmbarque", modelo.Prev_FechaLimiteEmbarque?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@Prev_FechaVencimiento", modelo.Prev_FechaVencimiento?.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@Prev_DescripcionMercancia", modelo.Prev_DescripcionMercancia); pix++;
				@params[pix] = new SqlParameter("@Prev_ConsideracionesAdicionales", modelo.Prev_ConsideracionesAdicionales); pix++;
				@params[pix] = new SqlParameter("@Prev_InstruccionesEspeciales", modelo.Prev_InstruccionesEspeciales); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_Enmienda", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_EnmiendaSwift(int enmiendaId, string ccid, string swiftFilename, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;
				SqlParameter[] @params = new SqlParameter[3];
				@params[pix] = new SqlParameter("@EnmiendaId", enmiendaId); pix++;
				@params[pix] = new SqlParameter("@SwiftFilename", swiftFilename); pix++;
				@params[pix] = new SqlParameter("@CartaComercialId", ccid); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_EnmiendaSwift", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
					else
					{
						string ccId = dt.Rows[0][3].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Ins_Enmienda(Enmienda modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[8];

				@params[0] = new SqlParameter("@CartaCreditoId", modelo.CartaCreditoId);
				@params[1] = new SqlParameter("@ImporteLC", modelo.ImporteLC);
				@params[2] = new SqlParameter("@FechaLimiteEmbarque", modelo.FechaLimiteEmbarque?.ToString("yyyy-MM-dd HH:mm:ss"));
				@params[3] = new SqlParameter("@FechaVencimiento", modelo.FechaVencimiento?.ToString("yyyy-MM-dd HH:mm:ss"));
				@params[4] = new SqlParameter("@DescripcionMercancia", modelo.DescripcionMercancia);
				@params[5] = new SqlParameter("@ConsideracionesAdicionales", modelo.ConsideracionesAdicionales);
				@params[6] = new SqlParameter("@InstruccionesEspeciales", modelo.InstruccionesEspeciales);
				@params[7] = new SqlParameter("@CreadoPor", modelo.CreadoPor);

				if (!bd.ExecuteProcedure(conexion, "ins_Enmienda", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_EnmiendasByCartaCreditoId(out DataTable dt, out String msgError, string cartaCreditoId)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@CartaCreditoId", cartaCreditoId);

				if (!bd.ExecuteProcedure(conexion, "cons_EnmiendasByCartaCreditoId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_EnmiendaById(out DataTable dt, out String msgError, int id)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_EnmiendaById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}
		#endregion

		#region Reportes
		public Boolean Ins_Reporte(Reporte modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[4];
				var pix = 0;

				@params[pix] = new SqlParameter("@TipoReporte", modelo.TipoReporte); pix++;
				@params[pix] = new SqlParameter("@CreadoPor", modelo.CreadoPor); pix++;
				@params[pix] = new SqlParameter("@CreadoPorId", modelo.CreadoPorId); pix++;
				@params[pix] = new SqlParameter("@Filename", modelo.Filename); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_Reporte", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_Reportes(out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[0];

				if (!bd.ExecuteProcedure(conexion, "cons_Reportes", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		/*
		public Boolean Cons_ReporteAnalisisCartasCredito(out DataTable dt, out String msgError, DateTime fechaInicio, DateTime fechaFin, int empresaId = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;

				SqlParameter[] @params = new SqlParameter[3];
				@params[pix] = new SqlParameter("@EmpresaId", empresaId); pix++;
				@params[pix] = new SqlParameter("@FechaInicio", fechaInicio.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaFin", fechaFin.ToString("yyyy-MM-dd HH:mm:ss")); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_ReporteAnalisisCartasCredito", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_ReporteComisionesPorTipoComision(out DataTable dt, out String msgError, DateTime fechaInicio, DateTime fechaFin, int empresaId = 0)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				int pix = 0;

				SqlParameter[] @params = new SqlParameter[3];
				@params[pix] = new SqlParameter("@EmpresaId", empresaId); pix++;
				@params[pix] = new SqlParameter("@FechaInicio", fechaInicio.ToString("yyyy-MM-dd HH:mm:ss")); pix++;
				@params[pix] = new SqlParameter("@FechaFin", fechaFin.ToString("yyyy-MM-dd HH:mm:ss")); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_ReporteComisionesPorTipoComision", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}
		*/
		#endregion

		#region AspNetUsers
		public Boolean Cons_AspNetUsers(out DataTable dt, out String msgError, int activo = -1)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetUsers", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
				Utility.Logger.Error("DataAccess - " + msgError);
			}

			return boolProcess;
		}

		public Boolean Ins_AspNetUser(AspNetUser modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@Email", modelo.Email); pix++;
				@params[pix] = new SqlParameter("@PasswordHash", modelo.PasswordHash); pix++;
				@params[pix] = new SqlParameter("@PhoneNumber", modelo.PhoneNumber); pix++;
				@params[pix] = new SqlParameter("@UserName", modelo.UserName); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_AspNetUser", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
				Utility.Logger.Error("DataAccess - " + msgError);
			}
			return boolProcess;
		}

		public Boolean Cons_AspNetUserById(string id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetUserById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_AspNetUserByUserName(out DataTable dt, out String msgError, string userName)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];

				@params[0] = new SqlParameter("@Usuario", userName);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetUserByUserName", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos a mostrar";
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		public Boolean Upd_AspNetUser(AspNetUser modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];

				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Email", modelo.Email); pix++;
				@params[pix] = new SqlParameter("@PhoneNumber", modelo.PhoneNumber); pix++;
				@params[pix] = new SqlParameter("@UserName", modelo.UserName); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_AspNetUser", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_AspNetUserPassword(AspNetUser modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];

				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Password", modelo.PasswordHash); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_AspNetUserPassword", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Ins_AspNetUserRole(string userId, string roleId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];

				var pix = 0;

				@params[pix] = new SqlParameter("@UserId", userId); pix++;
				@params[pix] = new SqlParameter("@RoleId", roleId); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_AspNetUserRole", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Upd_AspNetUserRole(string userId, string roleId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];

				var pix = 0;

				@params[pix] = new SqlParameter("@UserId", userId); pix++;
				@params[pix] = new SqlParameter("@RoleId", roleId); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_AspNetUserRole", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Ins_AspNetUserEmpresa(string userId, int empresaId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];

				var pix = 0;

				@params[pix] = new SqlParameter("@UserId", userId); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", empresaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_AspNetUserEmpresa", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Del_AspNetUserEmpresa(string userId, int empresaId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];

				var pix = 0;

				@params[pix] = new SqlParameter("@UserId", userId); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", empresaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "del_AspNetUserEmpresa", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region AspNetUserProfile
		public Boolean Ins_AspNetUserProfile(AspNetUserProfile modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;

				@params[pix] = new SqlParameter("@Name", modelo.Name); pix++;
				@params[pix] = new SqlParameter("@LastName", modelo.LastName); pix++;
				@params[pix] = new SqlParameter("@DisplayName", modelo.DisplayName); pix++;
				@params[pix] = new SqlParameter("@UserId", modelo.UserId); pix++;
				@params[pix] = new SqlParameter("@Notes", modelo.Notes); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_AspNetUserProfile", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
				Utility.Logger.Error("DataAccess - " + msgError);
			}
			return boolProcess;
		}

		public Boolean Cons_AspNetUserProfileByUserId(string userId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@UserId", userId);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetUserProfileByUserId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_AspNetUserProfile(AspNetUserProfile modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[6];

				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Name", modelo.Name); pix++;
				@params[pix] = new SqlParameter("@LastName", modelo.LastName); pix++;
				@params[pix] = new SqlParameter("@DisplayName", modelo.DisplayName); pix++;
				@params[pix] = new SqlParameter("@UserId", modelo.UserId); pix++;
				@params[pix] = new SqlParameter("@Notes", modelo.Notes); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_AspNetUserProfile", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region AspNetRoles
		public Boolean Cons_AspNetRoles(out DataTable dt, out String msgError, int activo = -1)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetRoles", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
				Utility.Logger.Error("DataAccess - " + msgError);
			}

			return boolProcess;
		}

		public Boolean Ins_AspNetRole(AspNetRole modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				var pix = 0;

				@params[pix] = new SqlParameter("@Name", modelo.Name); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_AspNetRole", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
				Utility.Logger.Error("DataAccess - " + msgError);
			}
			return boolProcess;
		}

		public Boolean Cons_AspNetRoleById(string id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetRoleById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Upd_AspNetRole(AspNetRole modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];

				var pix = 0;

				@params[pix] = new SqlParameter("@Id", modelo.Id); pix++;
				@params[pix] = new SqlParameter("@Name", modelo.Name); pix++;
				@params[pix] = new SqlParameter("@Activo", modelo.Activo); pix++;

				if (!bd.ExecuteProcedure(conexion, "upd_AspNetRole", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Ins_AspNetRolePermission(string roleId, int permissionId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];

				var pix = 0;

				@params[pix] = new SqlParameter("@RoleId", roleId); pix++;
				@params[pix] = new SqlParameter("@PermissionId", permissionId); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_AspNetRolePermission", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Del_AspNetRolePermission(string roleId, int permissionId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];

				var pix = 0;

				@params[pix] = new SqlParameter("@RoleId", roleId); pix++;
				@params[pix] = new SqlParameter("@PermissionId", permissionId); pix++;

				if (!bd.ExecuteProcedure(conexion, "del_AspNetRolePermission", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}
		#endregion

		#region AspNetPermissions
		public Boolean Cons_AspNetPermissions(out DataTable dt, out String msgError, int activo = 1)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Activo", activo);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetPermissions", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_AspNetPermissionsByRoleId(string roleId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@RoleId", roleId);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetPermissionsByRoleId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_AspNetPermissionById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_AspNetPermissionsById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No hay datos para mostrar";
					}
				}
			}
			catch (Exception e)
			{
				boolProcess = false;
				msgError = e.ToString();
			}

			return boolProcess;
		}
		#endregion

		#region Invitaciones
		public Boolean Ins_Invitacion(Invitacion modelo, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];

				@params[0] = new SqlParameter("@Email", modelo.Email);
				@params[1] = new SqlParameter("@UserName", modelo.UserName);
				@params[2] = new SqlParameter("@CreadoPorId", modelo.CreadoPorId);

				if (!bd.ExecuteProcedure(conexion, "ins_Invitacion", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (!dt.Rows[0][0].ToString().Equals("0"))
					{
						boolProcess = false;
						msgError = dt.Rows[0][1].ToString();
					}
				}

			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}
			return boolProcess;
		}

		public Boolean Cons_InvitacionById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Id", id);

				if (!bd.ExecuteProcedure(conexion, "cons_InvitacionById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_InvitacionByToken(string token, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				@params[0] = new SqlParameter("@Token", token);

				if (!bd.ExecuteProcedure(conexion, "cons_InvitacionByToken", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}
		#endregion

		#region PFEPrograma
		public Boolean Ins_PFEPrograma(PFEPrograma model, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;
				@params[pix] = new SqlParameter("@Anio", model.Anio); pix++;
				@params[pix] = new SqlParameter("@Periodo", model.Periodo); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", model.EmpresaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_PFEPrograma", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_PFEProgramaBuscar(PFEPrograma model, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[3];
				var pix = 0;
				@params[pix] = new SqlParameter("@Anio", model.Anio); pix++;
				@params[pix] = new SqlParameter("@Periodo", model.Periodo); pix++;
				@params[pix] = new SqlParameter("@EmpresaId", model.EmpresaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_PFEProgramaBuscar", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_PFEProgramaById(int id, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				var pix = 0;
				@params[pix] = new SqlParameter("@Id", id); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_PFEProgramaById", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_PFEPagos(int pfeProgramaId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				var pix = 0;
				@params[pix] = new SqlParameter("@ProgramaId", pfeProgramaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_PFEPagosByProgramaId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_PFEPago(int programaId, int pagoId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[2];
				var pix = 0;
				@params[pix] = new SqlParameter("@ProgramaId", programaId); pix++;
				@params[pix] = new SqlParameter("@PagoId", pagoId); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_PFEProgramaPago", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Ins_PFETipoCambio(PFETipoCambio model, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[5];
				var pix = 0;
				@params[pix] = new SqlParameter("@ProgramaId", model.ProgramaId); pix++;
				@params[pix] = new SqlParameter("@MonedaId", model.MonedaId); pix++;
				@params[pix] = new SqlParameter("@PA", model.PA); pix++;
				@params[pix] = new SqlParameter("@PA1", model.PA1); pix++;
				@params[pix] = new SqlParameter("@PA2", model.PA2); pix++;

				if (!bd.ExecuteProcedure(conexion, "ins_PFETipoCambio", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Cons_PFETipoCambioByProgramaId(int programaId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				var pix = 0;
				@params[pix] = new SqlParameter("@ProgramaId", programaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "cons_PFETipoCambioByProgramaId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Del_PFETipoCambioByProgramaId(int programaId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				var pix = 0;
				@params[pix] = new SqlParameter("@ProgramaId", programaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "del_PFETipoCambioByProgramaId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}

		public Boolean Del_PFEPagosByProgramaId(int programaId, out DataTable dt, out String msgError)
		{
			bool boolProcess = true;
			dt = new DataTable();
			msgError = string.Empty;

			try
			{
				SqlParameter[] @params = new SqlParameter[1];
				var pix = 0;
				@params[pix] = new SqlParameter("@ProgramaId", programaId); pix++;

				if (!bd.ExecuteProcedure(conexion, "del_PFEPagosByProgramaId", @params, out dt, 1000))
				{
					boolProcess = false;
					msgError = bd._error.ToString();
				}
				else
				{
					if (dt.Rows.Count < 1)
					{
						boolProcess = false;
						msgError = "No se pudo encontrar el registro";
					}
				}
			}
			catch (Exception ex)
			{
				boolProcess = false;
				msgError = ex.ToString();
			}

			return boolProcess;
		}
		#endregion
	}
}