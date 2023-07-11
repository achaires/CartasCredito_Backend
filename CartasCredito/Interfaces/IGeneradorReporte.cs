using CartasCredito.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CartasCredito.Interfaces
{
	public interface IGeneradorReporte
	{
		bool Verificar(int reporteId);
		Reporte Generar(int empresaId, DateTime fechaIni, DateTime fechaFin);
	}
}
