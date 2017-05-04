using System;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using BotAuth.Models;
using System.Configuration;
using BotAuth.Dialogs;
using BotAuth.AADv2;
using System.Threading;
using System.Net.Http;
using BotAuth;
using System.Collections.Generic;

namespace MsftGraphBotQuickStart.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        public Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);

            return Task.CompletedTask;
        }

        // Initialize selection and query options
        private QueryOption selection;
        private List<QueryOption> options = new List<QueryOption>() {
            new QueryOption() { Text = "photo", Endpoint = "photo/$value" },
            new QueryOption() { Text = "files", Endpoint = "drive/root/children" },
            new QueryOption() { Text = "mail", Endpoint = "mailfolders/inbox/messages" },
            new QueryOption() { Text = "events", Endpoint = $"calendarview?startdatetime={DateTime.Now.ToString("yyyy-MM-ddTHH:mmzzz")}&enddatetime={DateTime.Now.AddDays(7).ToString("yyyy-MM-ddTHH:mmzzz")}" }
        };

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        {
            var activity = await result as Activity;
            context.ConversationData.SetValue<Activity>("OriginalMessage", activity);

            // Prompt the user to select a Microsoft Graph query
            PromptDialog.Choice(context, this.choiceCallback, options, "What do you want me to query with the Microsoft Graph?:");
        }

        private async Task choiceCallback(IDialogContext context, IAwaitable<QueryOption> result)
        {
            selection = await result;

            // Initialize AuthenticationOptions with details from AAD v2 app registration (https://apps.dev.microsoft.com)
            AuthenticationOptions authConfig = new AuthenticationOptions()
            {
                Authority = ConfigurationManager.AppSettings["aad:Authority"],
                ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
                ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
                Scopes = new string[] { "User.Read", "Calendars.Read", "Mail.Read", "Files.Read" },
                RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
            };

            // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
            await context.Forward(new AuthDialog(new MSALAuthProvider(), authConfig), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
            {
                var tokenInfo = await authResult;

                // Check which query to perform
                var endpoint = $"https://graph.microsoft.com/v1.0/me/{selection.Endpoint}";
                if (selection.Text == "photo")
                {
                    // Photo is special because we need to convert stream to base64 string
                    var bytes = await new HttpClient().GetStreamWithAuthAsync(tokenInfo.AccessToken, endpoint);
                    var pic = "data:image/png;base64," + Convert.ToBase64String(bytes);

                    // Output the picture as a base64 encoded string
                    var msg = authContext.MakeMessage();
                    msg.Text = "Check it out...I got your picture using the Microsoft Graph!";
                    msg.Attachments.Add(new Attachment("image/png", pic));
                    await authContext.PostAsync(msg);
                }
                else
                {
                    // Perform a get and output the results based on option
                    var json = await new HttpClient().GetWithAuthAsync(tokenInfo.AccessToken, endpoint);
                    if (selection.Text == "files")
                        await authContext.PostAsync($"I located {((Newtonsoft.Json.Linq.JArray)json.SelectToken("value")).Count} folders/files in the root of your OneDrive.");
                    else if (selection.Text == "mail")
                        await authContext.PostAsync($"The last email you recieved had the following subject: \"{((Newtonsoft.Json.Linq.JArray)json.SelectToken("value")).First.Value<string>("subject")}\"");
                    else if (selection.Text == "events")
                        await authContext.PostAsync($"Your next meeting is: \"{((Newtonsoft.Json.Linq.JArray)json.SelectToken("value")).First.Value<string>("subject")}\"");
                }
                
                // Prompt the user to select a Microsoft Graph query (recursive)
                PromptDialog.Choice(authContext, this.choiceCallback, options, "What do you want me to query with the Microsoft Graph now?:");
            }, context.ConversationData.GetValue<Activity>("OriginalMessage"), CancellationToken.None);
        }
    }

    [Serializable]
    public class QueryOption
    {
        public string Text { get; set; }
        public string Endpoint { get; set; }
        public override string ToString()
        {
            return this.Text;
        }
    }
}