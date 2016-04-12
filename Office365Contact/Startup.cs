using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Office365Contact.Startup))]
namespace Office365Contact
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
