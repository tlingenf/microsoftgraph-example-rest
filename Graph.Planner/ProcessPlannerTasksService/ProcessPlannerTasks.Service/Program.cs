using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ProcessPlannerTasks.Service
{
    class Program
    {
        private const string clientId = "bc6a651d-94d6-4624-8e92-20653be4ad51";
        private const string tenantId = "organizations";
        static readonly string[] scopes = { "Tasks.ReadWrite.Shared", "Tasks.ReadWrite", "Group.Read.All" };

        static async Task Main(string[] args)
        {
            await InteractiveAuth();
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        static async Task InteractiveAuth()
        {
            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                .Create(clientId)
                .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                .Build();

            InteractiveAuthenticationProvider authProvider = new InteractiveAuthenticationProvider(publicClientApplication, scopes);

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            await DoWorkAsync(graphClient);
        }

        private static async Task DoWorkAsync(GraphServiceClient graphClient)
        {
            var plans = await graphClient.Me.Planner.Plans
                .Request()
                .GetAsync();

            var ready2Process = new List<PlannerTask>();
            foreach (var plan in plans)
            {
                var foundTasks = from task in (await graphClient.Planner.Plans[plan.Id].Tasks.Request().GetAsync())
                                 where (task.AppliedCategories.Category2 == true)
                                 select task;

                ready2Process.AddRange(foundTasks);
            }

            foreach (var currentTask in ready2Process)
            {
                await CompleteTask(currentTask, graphClient);
            }
        }

        private static async Task CompleteTask(PlannerTask task, GraphServiceClient graphClient)
        {
            Console.WriteLine("Updating Task {0}", task.Title);
            var updateTask = new PlannerTask();
            updateTask.AppliedCategories = new PlannerAppliedCategories();
            updateTask.AppliedCategories.Category2 = false;
            updateTask.AppliedCategories.Category4 = true;
            await graphClient.Planner.Tasks[task.Id]
                .Request()
                .Header("If-Match", task.GetEtag())
                .UpdateAsync(updateTask);
        }

        private static SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }
    }
}
