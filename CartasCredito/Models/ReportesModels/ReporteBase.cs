using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public abstract class ReporteBase
	{
		protected DateTime FechaInicio { get; set; }
		protected DateTime FechaFin { get; set; }
		protected int EmpresaId { get; set; }
		protected DateTime fechaDivisa { get; set; }
		protected string ColIni { get; set; }
		protected string ColFin { get; set; }
		protected ExcelPackage EPackage { get; set; }
		protected ExcelWorksheet ESheet { get; set; }
		protected Reporte ReporteResultado { get; set; }
		protected string ReporteNombre { get; set; }
		protected string ReporteFilename { get; set; }
		protected ReporteBase(DateTime fi, DateTime ff, int eid, DateTime fd, string colIni, string colFin, string reporteNombre)//, DateTime fvi, DateTime fvf)
		{
			var curDate = DateTime.Now;

			fechaDivisa = fd;
			EPackage = new ExcelPackage();
			ESheet = EPackage.Workbook.Worksheets.Add("Reporte");
			FechaInicio = fi;
			FechaFin = ff;
			EmpresaId = eid;
			ColIni = colIni;
			ColFin = colFin;
			ReporteNombre = reporteNombre;
			ReporteFilename = ReporteNombre.Trim().Replace(' ', '-').ToLower() + "-" + curDate.Year.ToString() + curDate.Month.ToString() + curDate.Day.ToString() + curDate.Hour.ToString() + curDate.Minute.ToString() + curDate.Second.ToString() + ".xlsx";

			ESheet.Cells.Style.Font.Size = 10;
			ESheet.Cells["B4:"+colFin+"4"].Style.Font.Bold = true;

			ESheet.Cells["B1:"+ colFin +"1"].Style.Font.Size = 18;
			ESheet.Cells["B1:"+colFin+"1"].Style.Font.Bold = true;
			ESheet.Cells["B1:"+colFin+"1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
			ESheet.Cells["B1"].Value = "Grupo Industrial Saltillo, S.A.B. de C.V.";

			ESheet.Cells["B2:"+colFin+"2"].Style.Font.Size = 13;
			ESheet.Cells["B2:"+colFin+"2"].Style.Font.Bold = true;
			ESheet.Cells["B2:"+colFin+"2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
			ESheet.Cells["B2"].Value = ReporteNombre;

			ESheet.Cells["B3:"+colFin+"3"].Style.Font.Size = 13;
			ESheet.Cells["B3:"+colFin+"3"].Style.Font.Bold = true;
			ESheet.Cells["B3:"+colFin+"3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
			if(EmpresaId==0){
				ESheet.Cells["B3"].Value = "Empresa: Todas";
			}else{
				var empresa = Empresa.GetById(EmpresaId);
				ESheet.Cells["B3"].Value = "Empresa: " + empresa.Nombre;
			}

			ESheet.Cells["B4:"+colFin+"4"].Style.Font.Size = 13;
			ESheet.Cells["B4:"+colFin+"4"].Style.Font.Bold = false;
			ESheet.Cells["B4:"+colFin+"4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

			//----
			string tipoCambio = "";
			if(fechaDivisa != null)
            {
				tipoCambio = "| Tipo de Cambio al: " + fechaDivisa.ToString("dd-MM-yyyy");
				/*ESheet.Cells["B5"].Value = tipoCambio;
				ESheet.Cells["B5:" + colFin + "4"].Style.Font.Size = 14;
				ESheet.Cells["B5:" + colFin + "4"].Style.Font.Bold = false;
				ESheet.Cells["B5:" + colFin + "4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;*/
			}
			//----

			string periodoText = "";

			if (reporteNombre == "Reporte de Vencimientos de Cartas de Crédito" ||
					reporteNombre == "Reporte de Cartas de Crédito Stand By")
			{
				periodoText += "";

				if (fechaDivisa != null)
				{
					tipoCambio = "Tipo de Cambio al: " + fechaDivisa.ToString("dd-MM-yyyy");
					/*ESheet.Cells["B5"].Value = tipoCambio;
					ESheet.Cells["B5:" + colFin + "4"].Style.Font.Size = 14;
					ESheet.Cells["B5:" + colFin + "4"].Style.Font.Bold = false;
					ESheet.Cells["B5:" + colFin + "4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;*/
				}
			}
			else
			{
				if (FechaInicio.Year != 1969)
				{
					periodoText += "Periodo: Del " + FechaInicio.ToString("dd-MM-yyyy") + " Al " + FechaFin.ToString("dd-MM-yyyy");
				}
				else
				{
					periodoText += "Periodo: Histórico";
				}
			}

			

			ESheet.Cells["B4"].Value = periodoText + tipoCambio;

			ESheet.Cells["B1:E1"].Merge = true;
			ESheet.Cells["B2:E2"].Merge = true;
			ESheet.Cells["B3:E3"].Merge = true;
			ESheet.Cells["B4:E4"].Merge = true;

			/*
			ESheet.Cells["B1:" + colFin + "1"].Merge = true;
			ESheet.Cells["B2:" + colFin + "2"].Merge = true;
			ESheet.Cells["B4:" + colFin + "4"].Merge = true;
			*/

		}


		public abstract Reporte Generar();
	}
}