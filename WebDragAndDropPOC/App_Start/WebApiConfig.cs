using Autofac;
using Autofac.Integration.WebApi;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Web.Http;
using WebDragAndDropPOC.Controllers;
using WebDragAndDropPOC.Middleware;

namespace WebDragAndDropPOC
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services
            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );

            var builder = new ContainerBuilder();

            // Register your Web API controllers.
            builder.RegisterApiControllers(Assembly.GetExecutingAssembly());

            // OPTIONAL: Register the Autofac filter provider.
            builder.RegisterWebApiFilterProvider(config);

            // OPTIONAL: Register the Autofac model binder provider.
            builder.RegisterWebApiModelBinderProvider();

            var filterConfig = ConfigurationManager.AppSettings["AttachmentBlackList"].Split(',');
            var filterSet = new HashSet<string>(filterConfig);
            builder.RegisterInstance<IMsgParser>(new MsgParser(filterSet));
            var container = builder.Build();

            // Set the dependency resolver to be Autofac.
            config.DependencyResolver = new AutofacWebApiDependencyResolver(container);
        }
    }
}
