using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace GoContactSyncMod
{
    //own implementation to append login_hint parameter to uri

    public class GCSMOAuth2WebAuthorizationBroker : GoogleWebAuthorizationBroker
    {
        public new static async Task<UserCredential> AuthorizeAsync(ClientSecrets clientSecrets,
            IEnumerable<string> scopes, string user, CancellationToken taskCancellationToken,
            IDataStore dataStore = null)
        {
            var initializer = new GCSMAuthorizationCodeFlow.Initializer
            {
                ClientSecrets = clientSecrets,
                Scopes = scopes,
                DataStore = dataStore ?? new FileDataStore(Folder)
            };

            var flow = new GCSMAuthorizationCodeFlow(initializer, user);

            // Create an authorization code installed app instance and authorize the user.
            return await new AuthorizationCodeInstalledApp(flow, new LocalServerCodeReceiver()).AuthorizeAsync
                (user, taskCancellationToken).ConfigureAwait(false);
        }
    }

    public class GCSMAuthorizationCodeFlow : GoogleAuthorizationCodeFlow
    {
        private readonly string userId;

        /// <summary>Constructs a new Google authorization code flow.</summary>
        public GCSMAuthorizationCodeFlow(Initializer initializer, string userId)
            : base(initializer)
        {
            this.userId = userId;
        }

        public override AuthorizationCodeRequestUrl CreateAuthorizationCodeRequest(string redirectUri)
        {
            return new GoogleAuthorizationCodeRequestUrl(new Uri(AuthorizationServerUrl))
            {
                ClientId = ClientSecrets.ClientId,
                Scope = string.Join(" ", Scopes),
                //appen duser to url
                LoginHint = userId,
                RedirectUri = redirectUri
            };
        }
    }
}
