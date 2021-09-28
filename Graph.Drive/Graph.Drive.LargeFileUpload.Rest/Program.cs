using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Graph.Drive.LargeFileUpload.Rest
{
    class Program
    {
        public static IConfigurationRoot _configuration;

        static void Main(string[] args)
        {
            ConfigureAppAsync(args).Wait();
            MainAsync(args).Wait();
        }

        private static async Task MainAsync(string[] args)
        {
            var accessToken = await ClientCredentialsAuthAsync(new string[] { "https://graph.microsoft.com/.default" });


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
                .AddJsonFile("appsettings.json.user", true)
                .Build();
        }
    }
}
