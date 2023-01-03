using System;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

namespace EWS_OAuth2
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            // Using Microsoft.Identity.Client
            var cca = ConfidentialClientApplicationBuilder
                .Create("513f26ae-dee9-48a1-ba0c-bc541d15cefd")      //appId
                .WithClientSecret("b3F8Q~Ijl9f6stBHxRH2e5NA.agbdeqXgn4qaaYR") // client secrete: b3F8Q~Ijl9f6stBHxRH2e5NA.agbdeqXgn4qaaYR, client id: 41ff9c4c-f5b4-4495-98b8-ba690ce64755
                .WithTenantId("f6ed3d45-259e-4bcb-b07c-976fc1e5ee08") // tenantId
                .Build();

            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

            try
            {
                // Get token
                var authResult = await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService(ExchangeVersion.Exchange2016);
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);


                // To use application permissions, you will also need to explicitly impersonate a mailbox that you would like to access.
                ewsClient.ImpersonatedUserId =
                    new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "yuanliang@yulian.onmicrosoft.com");

                //Include x-anchormailbox header
                ewsClient.HttpHeaders.Add("X-AnchorMailbox", "yuanliang@yulian.onmicrosoft.com");

                // Make an EWS call to list folders on exhange online
                var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(50));
                foreach (var folder in folders.Result)
                {
                    Console.WriteLine($"Folder: {folder.DisplayName}");
                }

                // Make an EWS call to read 50 emails (last 5 days) from Inbox folder
                TimeSpan ts = new TimeSpan(-360, 0, 0, 0);
                DateTime date = DateTime.Now.Add(ts);
                SearchFilter.IsGreaterThanOrEqualTo filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);
                var findResults = ewsClient.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(50));
                foreach (var mailItem in findResults.Result)
                {
                    Console.WriteLine($"Subject: {mailItem.Subject}");
                }
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }

            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("Hit any key to exit...");
                Console.ReadKey();
            }
        }
    }
}
