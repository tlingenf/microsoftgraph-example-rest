using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace Drive.Graph
{
    class Program
    {
        public static IConfigurationRoot _configuration;

        static void Main(string[] args)
        {
            ConfigureAppAsync(args).Wait();
            MainAsync(args).Wait();
        }

        public static async Task MainAsync(string[] args)
        {
            //var accessToken = await InteractiveAuthAsync(new string[] { "https://graph.microsoft.com/.default" });        // delegated auth
            var accessToken = await ClientCredentialsAuthAsync(new string[] { "https://graph.microsoft.com/.default" });    // app auth
            var graph = new GraphRestService(_configuration, accessToken);

            var itemInfo = await graph.GetDriveItemAsync(_configuration["driveItemUri"]);
            var path = await graph.DownloadDriveItemAsync(itemInfo);

            Console.WriteLine($"File Downloaded to: {path}");
        }

        public static async Task<string> InteractiveAuthAsync(IEnumerable<string> scopes)
        {
            IPublicClientApplication app = PublicClientApplicationBuilder
                .Create(_configuration["appId"])
                .WithTenantId(_configuration["tenantId"])
                .WithRedirectUri("http://localhost")
                .Build();

            AuthenticationResult auth = await app.AcquireTokenInteractive(scopes)
                .ExecuteAsync();

            return auth.AccessToken;
        }

        public static async Task<string> ClientCredentialsAuthAsync(IEnumerable<string> scopes)
        {
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                .Create(_configuration["appId"])
                .WithClientSecret(_configuration["appSecret"])
                .WithTenantId(_configuration["tenantId"])
                .Build();

            AuthenticationResult auth = await app.AcquireTokenForClient(scopes)
                .ExecuteAsync();

            return auth.AccessToken;
        }

        static async Task ConfigureAppAsync(string[] args)
        {
            _configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetParent(AppContext.BaseDirectory).FullName)
                .AddJsonFile("appsettings.json", false)
                .Build();
        }
    }
}
