using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CartasCredito.Controllers
{
	public class HomeController : Controller
	{
		public ActionResult Index()
		{
			dll_Gis.Funciones fn = new dll_Gis.Funciones();
			string a = fn.Desencriptar("XBo46ROG61Y5w6rhWPFTnRXX8hGWRd1yxq1ndmoai7V19C9CZLpXM1RRRS/P90Folv96pxOqsdVor6C4qbX6uS8xwooE2Xb0iVinU7f9H2AnBA60zCN7M7x45/Bm89xsoLHpEPsfFiU47jRfv/GdVon6qPB7btgOeH/p10GDJEp00J3WtF4O0Pe9QukbzNDts1ECWx9U37lw0M4SJ9OaHjGjHCv6jIRbqt9tb2gsA101MgX4+GyL6ZufSn6KiYmw");
			string a2 = fn.Desencriptar("XBo46ROG61Y5w6rhWPFTnRXX8hGWRd1yxq1ndmoai7V19C9CZLpXM1RRRS/P90Folv96pxOqsdVor6C4qbX6uS8xwooE2Xb0iVinU7f9H2AnBA60zCN7M7x45/Bm89xsoLHpEPsfFiU47jRfv/GdVon6qPB7btgOeH/p10GDJEp00J3WtF4O0Pe9QukbzNDts1ECWx9U37lw0M4SJ9OaHjGjHCv6jIRbqt9tb2gsA101MgX4+GyL6ZufSn6KiYmw");
			//Server=SRGISMTY2-1713; Database=GIS_CartasCredito_desa; Trusted_Connection=False; MultipleActiveResultSets=true; User Id=sql.dev.softdepot; pwd=@G1sD3v#8426$2022; Trusted_Connection=False
			return View();
		}
	}
}