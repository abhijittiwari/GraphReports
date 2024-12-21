using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using Azure.Core;
using System.Text.Json;
using Microsoft.Graph.Models;
using System.Diagnostics.Metrics;
public class TokenProvider
{
    private readonly string[] _scopes;
    private readonly string _tenantId;
    private readonly string _clientId;
    private readonly InteractiveBrowserCredentialOptions _options;

    public TokenProvider(string tenantId, string clientId, string[] scopes)
    {
        _tenantId = tenantId;
        _clientId = clientId;
        _scopes = scopes;
        _options = new InteractiveBrowserCredentialOptions
        {
            TenantId = _tenantId,
            ClientId = _clientId,
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            RedirectUri = new Uri("http://localhost"),
        };
    }

    public async Task<AccessToken> GetTokenAsync()
    {
        var interactiveCredential = new InteractiveBrowserCredential(_options);
        var requestContext = new TokenRequestContext(_scopes);
        return await interactiveCredential.GetTokenAsync(requestContext);
    }
}
