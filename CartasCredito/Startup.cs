using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(CartasCredito.Startup))]
namespace CartasCredito
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
