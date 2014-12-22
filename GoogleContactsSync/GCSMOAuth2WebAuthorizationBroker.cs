using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Json;
using Google.Apis.Requests;
using Google.Apis.Requests.Parameters;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace GoContactSyncMod
{
    //own implementation to append login_hint parameter to uri

    public class GCSMOAuth2WebAuthorizationBroker
    {
        /// <summary>The folder which is used by the <see cref="Google.Apis.Util.Store.FileDataStore"/>.</summary>
        /// <remarks>
        /// The reason that this is not 'private const' is that a user can change it and store the credentials in a
        /// different location.
        /// </remarks>
        public static string Folder = "Google.Apis.Auth";

        /// <summary>Asynchronously authorizes the specified user.</summary>
        /// <remarks>
        /// In case no data store is specified, <see cref="Google.Apis.Util.Store.FileDataStore"/> will be used by 
        /// default.
        /// </remarks>
        /// <param name="clientSecrets">The client secrets.</param>
        /// <param name="scopes">
        /// The scopes which indicate the Google API access your application is requesting.
        /// </param>
        /// <param name="user">The user to authorize.</param>
        /// <param name="taskCancellationToken">Cancellation token to cancel an operation.</param>
        /// <param name="dataStore">The data store, if not specified a file data store will be used.</param>
        /// <returns>User credential.</returns>
        public static async Task<UserCredential> AuthorizeAsync(ClientSecrets clientSecrets,
            IEnumerable<string> scopes, string user, CancellationToken taskCancellationToken,
            IDataStore dataStore = null)
        {
            var initializer = new GCSMAuthorizationCodeFlow.Initializer
            {
                ClientSecrets = clientSecrets,
            };
            return await AuthorizeAsyncCore(initializer, scopes, user, taskCancellationToken, dataStore)
                .ConfigureAwait(false);
        }

        /// <summary>Asynchronously authorizes the specified user.</summary>
        /// <remarks>
        /// In case no data store is specified, <see cref="Google.Apis.Util.Store.FileDataStore"/> will be used by 
        /// default.
        /// </remarks>
        /// <param name="clientSecretsStream">
        /// The client secrets stream. The authorization code flow constructor is responsible for disposing the stream.
        /// </param>
        /// <param name="scopes">
        /// The scopes which indicate the Google API access your application is requesting.
        /// </param>
        /// <param name="user">The user to authorize.</param>
        /// <param name="taskCancellationToken">Cancellation token to cancel an operation.</param>
        /// <param name="dataStore">The data store, if not specified a file data store will be used.</param>
        /// <returns>User credential.</returns>
        public static async Task<UserCredential> AuthorizeAsync(Stream clientSecretsStream,
            IEnumerable<string> scopes, string user, CancellationToken taskCancellationToken,
            IDataStore dataStore = null)
        {
            var initializer = new GCSMAuthorizationCodeFlow.Initializer
            {
                ClientSecretsStream = clientSecretsStream,
            };
            return await AuthorizeAsyncCore(initializer, scopes, user, taskCancellationToken, dataStore)
                .ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously reauthorizes the user. This method should be called if the users want to authorize after 
        /// they revoked the token.
        /// </summary>
        /// <param name="userCredential">The current user credential. Its <see cref="UserCredential.Token"/> will be
        /// updated. </param>
        /// <param name="taskCancellationToken">Cancellation token to cancel an operation.</param>
        public static async Task ReauthorizeAsync(UserCredential userCredential,
            CancellationToken taskCancellationToken)
        {
            // Create an authorization code installed app instance and authorize the user.
            UserCredential newUserCredential = await new AuthorizationCodeInstalledApp(userCredential.Flow,
                new LocalServerCodeReceiver()).AuthorizeAsync
                (userCredential.UderId, taskCancellationToken).ConfigureAwait(false);
            userCredential.Token = newUserCredential.Token;
        }

        /// <summary>The core logic for asynchronously authorizing the specified user.</summary>
        /// <param name="initializer">The authorization code initializer.</param>
        /// <param name="scopes">
        /// The scopes which indicate the Google API access your application is requesting.
        /// </param>
        /// <param name="user">The user to authorize.</param>
        /// <param name="taskCancellationToken">Cancellation token to cancel an operation.</param>
        /// <param name="dataStore">The data store, if not specified a file data store will be used.</param>
        /// <returns>User credential.</returns>
        private static async Task<UserCredential> AuthorizeAsyncCore(
            GCSMAuthorizationCodeFlow.Initializer initializer, IEnumerable<string> scopes, string user,
            CancellationToken taskCancellationToken, IDataStore dataStore = null)
        {
            initializer.Scopes = scopes;
            initializer.DataStore = dataStore ?? new FileDataStore(Folder);
            var flow = new GCSMAuthorizationCodeFlow(initializer, user);

            // Create an authorization code installed app instance and authorize the user.
            return await new AuthorizationCodeInstalledApp(flow, new LocalServerCodeReceiver()).AuthorizeAsync
                (user, taskCancellationToken).ConfigureAwait(false);
        }
    }

    public class GCSMAuthorizationCodeFlow : AuthorizationCodeFlow
    {
        private readonly string revokeTokenUrl;
        private readonly string userId;

        /// <summary>Gets the token revocation URL.</summary>
        public string RevokeTokenUrl { get { return revokeTokenUrl; } }

        /// <summary>Constructs a new Google authorization code flow.</summary>
        public GCSMAuthorizationCodeFlow(Initializer initializer, string userId)
            : base(initializer)
        {
            revokeTokenUrl = initializer.RevokeTokenUrl;
            this.userId = userId;
        }

        public override AuthorizationCodeRequestUrl CreateAuthorizationCodeRequest(string redirectUri)
        {
            return new GoogleAuthorizationCodeRequestUrl(new Uri(AuthorizationServerUrl))
            {
                ClientId = ClientSecrets.ClientId,
                Scope = string.Join(" ", Scopes),
                LoginHint = userId,
                RedirectUri = redirectUri
            };
        }

        public override async Task RevokeTokenAsync(string userId, string token,
            CancellationToken taskCancellationToken)
        {
            GoogleRevokeTokenRequest request = new GoogleRevokeTokenRequest(new Uri(RevokeTokenUrl))
            {
                Token = token
            };
            var httpRequest = new HttpRequestMessage(HttpMethod.Get, request.Build());

            var response = await HttpClient.SendAsync(httpRequest, taskCancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                var error = NewtonsoftJsonSerializer.Instance.Deserialize<TokenErrorResponse>(content);
                throw new TokenResponseException(error);
            }

            await DeleteTokenAsync(userId, taskCancellationToken);
        }

        /// <summary>An initializer class for Google authorization code flow. </summary>
        public new class Initializer : AuthorizationCodeFlow.Initializer
        {
            /// <summary>Gets or sets the token revocation URL.</summary>
            public string RevokeTokenUrl { get; set; }

            /// <summary>
            /// Constructs a new initializer. Sets Authorization server URL to 
            /// <see cref="Google.Apis.Auth.OAuth2.GoogleAuthConsts.AuthorizationUrl"/>, and Token server URL to 
            /// <see cref="Google.Apis.Auth.OAuth2.GoogleAuthConsts.TokenUrl"/>.
            /// </summary>
            public Initializer()
                : base(GoogleAuthConsts.AuthorizationUrl, GoogleAuthConsts.TokenUrl)
            {
                RevokeTokenUrl = GoogleAuthConsts.RevokeTokenUrl;
            }
        }
    }

    class GoogleRevokeTokenRequest
    {
        private readonly Uri revokeTokenUrl;
        /// <summary>Gets the URI for token revocation.</summary>
        public Uri RevokeTokenUrl
        {
            get { return revokeTokenUrl; }
        }

        /// <summary>Gets or sets the token to revoke.</summary>
        [Google.Apis.Util.RequestParameterAttribute("token")]
        public string Token { get; set; }

        public GoogleRevokeTokenRequest(Uri revokeTokenUrl)
        {
            this.revokeTokenUrl = revokeTokenUrl;
        }

        /// <summary>Creates a <see cref="System.Uri"/> which is used to request the authorization code.</summary>
        public Uri Build()
        {
            var builder = new RequestBuilder()
            {
                BaseUri = revokeTokenUrl
            };
            ParameterUtils.InitParameters(builder, this);
            return builder.BuildUri();
        }
    }

}
