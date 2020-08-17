using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace API_Parser.Configuration
{
    public class API_Config
    {        
        public static void Register(HttpConfiguration config)
        {            
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "parserapi/{controller}/{action}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}