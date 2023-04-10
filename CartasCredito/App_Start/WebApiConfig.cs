using CartasCredito.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace CartasCredito
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
			// Web API configuration and services
			config.EnableCors();
			// Web API routes
			config.MapHttpAttributeRoutes();

			config.MessageHandlers.Add(new TokenValidationHandler());
			config.Formatters.XmlFormatter.SupportedMediaTypes.Add(new System.Net.Http.Headers.MediaTypeHeaderValue("multipart/form-data"));


			config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
