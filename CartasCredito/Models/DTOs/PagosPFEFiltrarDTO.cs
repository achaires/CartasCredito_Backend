using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class PagosPFEFiltrarDTO
	{
		public List<int> Filtro_Anios { get; set; }
		public List<PeriodoContable> Filtro_Periodos { get; set; }
		public List<Empresa> Filtro_Empresas { get; set; }

		public List<Pago> PagosProgramados { get; set; }

		// Opciones seleccionadas
		public int Anio { get; set; }
		public int PeriodoId { get; set; }
		public int EmpresaId { get; set; } // Tambien se envia a sp

		// Opciones para SP
		public DateTime FechaInicio { get; set; }
		public DateTime FechaFin { get; set; }
		public int Activo { get; set; }

		public PagosPFEFiltrarDTO()
		{
			DateTime curDate = DateTime.Now;

			Filtro_Anios = new List<int>();

			for (int i = curDate.Year + 5; i >= 2008; i--)
			{
				Filtro_Anios.Add(i);
			}

			Filtro_Periodos = new List<PeriodoContable>()
			{
				new PeriodoContable()
				{
					Id = 1,
					Name = "Enero",
				},
				new PeriodoContable()
				{
					Id = 2,
					Name = "Febrero",
				},
				new PeriodoContable()
				{
					Id = 3,
					Name = "Marzo",
				},
				new PeriodoContable()
				{
					Id = 4,
					Name = "Abril",
				},
				new PeriodoContable()
				{
					Id = 5,
					Name = "Mayo",
				},
				new PeriodoContable()
				{
					Id = 6,
					Name = "Junio",
				},
				new PeriodoContable()
				{
					Id = 7,
					Name = "Julio",
				},
				new PeriodoContable()
				{
					Id = 8,
					Name = "Agosto",
				},
				new PeriodoContable()
				{
					Id = 9,
					Name = "Septiembre",
				},
				new PeriodoContable()
				{
					Id = 10,
					Name = "Octubre",
				},
				new PeriodoContable()
				{
					Id = 11,
					Name = "Noviembre",
				},
				new PeriodoContable()
				{
					Id = 12,
					Name = "Diciembre",
				},
			};

			Filtro_Empresas = Empresa.Get(1);
		}
	}
}