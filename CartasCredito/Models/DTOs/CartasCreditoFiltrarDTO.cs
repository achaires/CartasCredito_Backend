using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CartasCredito.Models.DTOs
{
	public class CartasCreditoFiltrarDTO
	{
		public List<TipoActivo> TiposActivo { get; set; }
		public List<Moneda> Monedas { get; set; }
		public List<Proveedor> Proveedores { get; set; }
		public List<Empresa> Empresas { get; set; }
		public List<Banco> Bancos { get; set; }

		public string NumCarta { get; set; }
		public string TipoCarta { get; set; }
		public int TipoCartaId { get; set; }
		public int TipoActivoId { get; set; }
		public int MonedaId { get; set; }
		public int ProveedorId { get; set; }
		public int EmpresaId { get; set; }
		public int BancoId { get; set; }
		public int Estatus { get; set; }

		public DateTime FechaInicio { get; set; }
		public DateTime FechaFin { get; set; }

		public CartasCreditoFiltrarDTO()
		{
			NumCarta = "";
			TipoCarta = "";

			DateTime dateNow = DateTime.Now;
			FechaInicio = new DateTime(dateNow.Year, dateNow.Month, 1);
			FechaFin = dateNow;

			TiposActivo = TipoActivo.Get(1);
			TiposActivo.Insert(0, new TipoActivo
			{
				Nombre = "--",
				Id = 0
			});

			Monedas = Moneda.Get(1);
			Monedas.Insert(0, new Moneda
			{
				Nombre = "--",
				Id = 0
			});

			Empresas = Empresa.Get(1);
			Empresas.Insert(0, new Empresa
			{
				Nombre = "--",
				Id = 0
			});

			Proveedores = Proveedor.Get(1);
			Proveedores.Insert(0, new Proveedor
			{
				Nombre = "--",
				Id = 0
			});

			Bancos = Banco.Get(1);
			Bancos.Insert(0, new Banco
			{
				Nombre = "--",
				Id = 0
			});
		}
	}
}