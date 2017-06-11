using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace Vidly
{
	public class RouteConfig
	{
		public static void RegisterRoutes(RouteCollection routes)
		{
			routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

			routes.MapRoute(
				"MoviesByReleaseDate",
				"Movies/Released/{year}/{month}",
				new { controller = "Movies", action = "ByReleaseDate" },
				new { year = @"\d{4}", month = @"\d{2}" }
				);

			routes.MapRoute(
				name: "Customers",
				url: "Customers",
				defaults: new { controller = "Customers", action = "Index" }
			);

			routes.MapRoute(
				name: "CustomersDetails",
				url: "Customers/{id}",
				defaults: new { controller = "Customers", action = "CustomersDetails", Id = UrlParameter.Optional }
			);

			routes.MapRoute(
				name: "Movies",
				url: "Movies",
				defaults: new { controller = "Movies", action = "Index" }
			);

            routes.MapRoute(
                name: "Random",
                url: "Movies/Random",
                defaults: new { controller = "Movies", action = "Random" }
            );

            routes.MapRoute(
				name: "MoviesDetails",
				url: "Movies/{id}",
				defaults: new { controller = "Movies", action = "MoviesDetails", Id = UrlParameter.Optional }
			);

			routes.MapRoute(
				name: "Default",
				url: "{controller}/{action}/{id}",
				defaults: new { controller = "Home", action = "Index", Id = UrlParameter.Optional }
			);
		}
	}
}
