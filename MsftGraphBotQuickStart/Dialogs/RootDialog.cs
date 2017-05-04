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

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        {
            var activity = await result as Activity;

            // Initialize AuthenticationOptions with details from AAD v2 app registration (https://apps.dev.microsoft.com)
            AuthenticationOptions options = new AuthenticationOptions()
            {
                Authority = ConfigurationManager.AppSettings["aad:Authority"],
                ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
                ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
                Scopes = new string[] { "User.Read" },
                RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
            };

            // Forward the dialog to the AuthDialog to sign the user in and get an access token for calling the Microsoft Graph
            await context.Forward(new AuthDialog(new MSALAuthProvider(), options), async (IDialogContext authContext, IAwaitable<AuthResult> authResult) =>
            {
                var tokenInfo = await authResult;

                // Get the users profile photo from the Microsoft Graph
                var bytes = await new HttpClient().GetStreamWithAuthAsync(tokenInfo.AccessToken, "https://graph.microsoft.com/v1.0/me/photo/$value");
                var pic = "data:image/png;base64," + Convert.ToBase64String(bytes);

                // Output the picture as a base64 encoded string
                var msg = authContext.MakeMessage();
                msg.Text = "Check it out...I got your picture using the Microsoft Graph!";
                msg.Attachments.Add(new Attachment("image/png", pic));
                await authContext.PostAsync(msg);
            }, activity, CancellationToken.None);
        }
    }
}