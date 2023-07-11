﻿using CartasCredito.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.ReportesModels
{
	public class ComisionesPorTipoComision : IGeneradorReporte
	{
		public bool Verificar(int reporteId)
		{
			return reporteId == 2;
		}

		public void Generar(DateTime fechaIni, DateTime fechaFin)
		{
			throw new NotImplementedException();
		}
	}
}