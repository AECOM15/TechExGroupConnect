// <copyright file="SingleTenantBotFrameworkHttpAdapter.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

namespace Microsoft.Teams.Apps.DIConnect.Common.Adapter
{
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.DIConnect.Common.Services.CommonBot;

    /// <summary>
    /// A BotFrameworkHttpAdapter that supports single-tenant bot registrations
    /// by including the tenant ID when building credentials.
    /// </summary>
    public class SingleTenantBotFrameworkHttpAdapter : BotFrameworkHttpAdapter
    {
        private readonly ICredentialProvider credentialProvider;
        private readonly string tenantId;

        /// <summary>
        /// Initializes a new instance of the <see cref="SingleTenantBotFrameworkHttpAdapter"/> class.
        /// </summary>
        /// <param name="credentialProvider">The credential provider.</param>
        /// <param name="botOptions">The bot options containing the tenant ID.</param>
        public SingleTenantBotFrameworkHttpAdapter(
            ICredentialProvider credentialProvider,
            IOptions<BotOptions> botOptions)
            : base(credentialProvider)
        {
            this.credentialProvider = credentialProvider;
            this.tenantId = botOptions?.Value?.TenantId;
        }

        /// <inheritdoc/>
        protected override async Task<AppCredentials> BuildCredentialsAsync(string appId, string oAuthScope = null)
        {
            var appPassword = await this.credentialProvider.GetAppPasswordAsync(appId);

            if (!string.IsNullOrEmpty(this.tenantId))
            {
                return new MicrosoftAppCredentials(appId, appPassword, this.tenantId);
            }

            return new MicrosoftAppCredentials(appId, appPassword);
        }
    }
}
