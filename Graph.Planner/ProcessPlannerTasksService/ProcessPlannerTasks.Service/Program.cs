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
        private const string clientId = "4b7bce2b-0b0c-4b3c-bf1d-ee2607d95b08";
        private const string tenantId = "34f6951c-e759-4967-9c6a-f05de4922e5c";
        private const string username = "svc.planner@M365x570719.onmicrosoft.com";
        private const string password = "Buc87698";
        static readonly string[] scopes = { "Tasks.Read.Shared", "Group.ReadWrite.All", "offline_access" };

        static async Task Main(string[] args)
        {
            //await ResourceOwnerAuth(args);
            await DeviceCodeAuth(args);
        }

        static async Task ResourceOwnerAuth(string[] args)
        {
            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                        .Create(clientId)
                        .WithTenantId(tenantId)
                        .Build();

            UsernamePasswordProvider authProvider = new UsernamePasswordProvider(publicClientApplication, scopes);

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var me = await graphClient.Me
                .Request()
                .WithUsernamePassword(username, ConvertToSecureString(password))
                .GetAsync();

            await DoWorkAsync(graphClient);
        }

        static async Task DeviceCodeAuth(string[] args)
        {

            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
            .Create(clientId)
            .WithTenantId(tenantId)
            .Build();

            Func<DeviceCodeResult, Task> deviceCodeReadyCallback = async dcr => await System.Console.Out.WriteLineAsync(dcr.Message);
            DeviceCodeProvider authProvider = new DeviceCodeProvider(publicClientApplication, scopes, deviceCodeReadyCallback);
            //DeviceCodeProvider authProvider = new DeviceCodeProvider(publicClientApplication, scopes);

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
                await ScheduleTask(currentTask, graphClient);
            }
        }

        private static async Task ScheduleTask(PlannerTask task, GraphServiceClient graphClient)
        {
            Console.WriteLine("Updating Task {0}", task.Title);
            var updateTask = new PlannerTask();
            updateTask.AppliedCategories = new PlannerAppliedCategories();
            updateTask.AppliedCategories.Category2 = false;
            updateTask.AppliedCategories.Category1 = true;
            await graphClient.Planner.Tasks[task.Id]
                .Request()
                .Header("If-Match", task.GetEtag())
                .UpdateAsync(updateTask);

            //task.AppliedCategories.Category2 = null;
            //task.AppliedCategories.Category1 = true;
            //await graphClient.Planner.Tasks[task.Id]
            //    .Request()
            //    .UpdateAsync(task);

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
