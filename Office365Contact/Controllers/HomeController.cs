using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using Office365Contact.Utils;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace Office365Contact.Controllers
{
    public class HomeController : Controller
    {
        public async Task<ActionResult> Index()
        {
            //AuthenticationContext authContext = new AuthenticationContext(string.Format("{0}/{1}", SettingsHelper.AuthorizationUri, "dataonline.co.nz")); //566cc084-238e-4821-a4e0-9ae4e1c6140d

            //var exClient = new OutlookServicesClient(new Uri("https://outlook.office365.com/api/v1.0"),
            //    async () =>
            //    {
            //        var authResult = authContext.AcquireToken("https://outlook.office365.com/", new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey));

            //        return authResult.AccessToken;
            //    });

            //var contactsResult = await exClient.Me.Contacts.ExecuteAsync();

            AuthenticationContext authContext = new AuthenticationContext("https://login.microsoftonline.com/" + SettingsHelper.TenantId + "/oauth2/token");

            var authResult = authContext.AcquireToken("https://graph.windows.net/", new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey));

            var client = new HttpClient();
            var queryString = HttpUtility.ParseQueryString(string.Empty);
            //Bearer
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);

            queryString["api-version"] = "1.6";
            var uri = "https://graph.windows.net/" + SettingsHelper.TenantId + "/contacts?" + queryString;


            var response = await client.GetAsync(uri);

            if (response.Content != null)
            {
                var responseString = await response.Content.ReadAsStringAsync();
                return View(responseString);
            }
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}