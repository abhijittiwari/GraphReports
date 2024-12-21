
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using Azure.Core;
using Azure.Identity;
using System.Text.Json;
using Microsoft.Graph.Models;
using System.Diagnostics.Metrics;
using Microsoft.Graph.Beta;
using Microsoft.Extensions.Azure;
using Microsoft.Kiota.Serialization;



namespace GraphReports
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }



        private async void buttonGetAllUsers_Click(object sender, EventArgs e)
        {
            try
            {
                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Text = "Authenticating"; // Set text in progress bar
                var scopes = new[] { "User.Read.All", "Directory.Read.All", "AuditLog.Read.All" };

                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);

                var graphClient = new GraphServiceClient(interactiveCredential, scopes);
                progressBar1.Text = "Fetching Users";
                var usersResponse = await graphClient.Users.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "UserType","displayName", "Mail", "jobTitle", "UserPrincipalName", "Id", "OnPremisesSyncEnabled", "CreatedDateTime", "ProxyAddresses", "AssignedLicenses", "AssignedPlans", "ServiceProvisioningErrors", "SignInSessionsValidFromDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName", "OnPremisesLastSyncDateTime","AccountEnabled","Manager","SignInActivity"
                    };
                });

                if (usersResponse?.Value != null)
                {
                    var users = usersResponse.Value.Select(user => new
                    {
                        UserType = user.UserType ?? "Not Available",
                        DisplayName = user.DisplayName ?? "Not Available",
                        Email = user.Mail ?? "Not Available",
                        JobTitle = user.JobTitle ?? "Not Available",
                        UserPrincipalName = user.UserPrincipalName ?? "Not Available",
                        LastSignInActivity = user.SignInActivity?.LastSignInDateTime?.UtcDateTime.ToString() ?? "No Sign In Activity",
                        ID = user.Id ?? "Not Available",
                        AccountEnabled = user.AccountEnabled?.ToString() ?? "Not Available",
                        Manager = user.Manager?.Id != null ? GetUserPrincipalName(graphClient, user.Manager.Id).Result : "Not Available",
                        City = user.City ?? "Not Available",
                        Country = user.Country ?? "Not Available",
                        Department = user.Department ?? "Not Available",
                        OfficeLocation = user.OfficeLocation ?? "Not Available",
                        LastPasswordChangeDateTime = user.LastPasswordChangeDateTime?.UtcDateTime.ToString() ?? "Not Available",
                        SignInSessionsValidFromDateTime = user.SignInSessionsValidFromDateTime?.UtcDateTime.ToString() ?? "No Sign In Sessions",
                        OnPremisesImmutableId = user.OnPremisesImmutableId ?? "No Immutable Id",
                        OnPremisesDistinguishedName = user.OnPremisesDistinguishedName ?? "No Distinguished Name",
                        OnPremisesLastSyncDateTime = user.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "No Last Sync DateTime",
                        SyncEnabled = user.OnPremisesSyncEnabled?.ToString() ?? "Not Synced",
                        WhenCreated = user.CreatedDateTime?.UtcDateTime.ToString() ?? "Not Available",
                        ProxyAddresses = user.ProxyAddresses != null && user.ProxyAddresses.Any() ? string.Join(", ", user.ProxyAddresses) : "No Proxy Address",
                        AssignedLicenses = user.AssignedLicenses != null && user.AssignedLicenses.Any() ? string.Join(", ", user.AssignedLicenses.Where(license => license.SkuId.HasValue).Select(license => Mapping.GetProductNameBySkuId(license.SkuId.Value.ToString()))) : "No Assigned Licenses",
                        DisabledPlans = user.AssignedLicenses != null && user.AssignedLicenses.Any() ? string.Join(", ", user.AssignedLicenses.SelectMany(license => license.DisabledPlans?.Where(planId => planId.HasValue).Select(planId => ServicePlanMapping.GetServicePlanById(planId.Value.ToString())) ?? Enumerable.Empty<string>())) : "No Disabled Plans"
                    }).ToList();

                    dataGridView1.DataSource = users;
                }
                else
                {
                    MessageBox.Show("No users found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async Task<string> GetUserPrincipalName(GraphServiceClient graphClient, string managerId)
        {
            var manager = await graphClient.Users[managerId].GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new[] { "UserPrincipalName" };
            });

            return manager?.UserPrincipalName ?? "Not Available";
        }




        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private async void buttonGetSynced_Click(object sender, EventArgs e)
        {
            try
            {
                progressBar1.Text = "Authenticating";
                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                var scopes = new[] { "User.Read.All", "Directory.Read.All" };

                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);

                var graphClient = new GraphServiceClient(interactiveCredential, scopes);
                progressBar1.Text = "Getting Synced Users";
                var usersResponse = await graphClient.Users.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "onPremisesSyncEnabled eq true";
                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "UserType","displayName", "Mail", "jobTitle", "UserPrincipalName", "Id", "OnPremisesSyncEnabled", "CreatedDateTime", "ProxyAddresses", "AssignedLicenses", "AssignedPlans", "ServiceProvisioningErrors", "SignInSessionsValidFromDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName", "OnPremisesLastSyncDateTime","AccountEnabled","Manager"
                    };
                });

                if (usersResponse?.Value != null)
                {
                    var users = usersResponse.Value.Select(user => new
                    {
                        UserType = user.UserType ?? "Not Available",
                        DisplayName = user.DisplayName ?? "Not Available",
                        Email = user.Mail ?? "Not Available",
                        JobTitle = user.JobTitle ?? "Not Available",
                        UserPrincipalName = user.UserPrincipalName ?? "Not Available",
                        Manager = user.Manager?.Id != null ? GetUserPrincipalName(graphClient, user.Manager.Id).Result : "Not Available",
                        ID = user.Id ?? "Not Available",
                        AccountEnabled = user.AccountEnabled?.ToString() ?? "Not Available",
                        City = user.City ?? "Not Available",
                        Country = user.Country ?? "Not Available",
                        Department = user.Department ?? "Not Available",
                        OfficeLocation = user.OfficeLocation ?? "Not Available",
                        LastPasswordChangeDateTime = user.LastPasswordChangeDateTime?.UtcDateTime.ToString() ?? "Not Available",
                        SignInSessionsValidFromDateTime = user.SignInSessionsValidFromDateTime?.UtcDateTime.ToString() ?? "No Sign In Sessions",
                        OnPremisesImmutableId = user.OnPremisesImmutableId ?? "No Immutable Id",
                        OnPremisesDistinguishedName = user.OnPremisesDistinguishedName ?? "No Distinguished Name",
                        OnPremisesLastSyncDateTime = user.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "No Last Sync DateTime",
                        SyncEnabled = user.OnPremisesSyncEnabled?.ToString() ?? "Not Synced",
                        WhenCreated = user.CreatedDateTime?.ToString() ?? "Not Available",
                        ProxyAddresses = user.ProxyAddresses != null && user.ProxyAddresses.Any() ? string.Join(", ", user.ProxyAddresses) : "No Proxy Address",
                        AssignedLicenses = user.AssignedLicenses != null && user.AssignedLicenses.Any() ? string.Join(", ", user.AssignedLicenses.Where(license => license.SkuId.HasValue).Select(license => Mapping.GetProductNameBySkuId(license.SkuId.Value.ToString()))) : "No Assigned Licenses",
                        DisabledPlans = user.AssignedLicenses != null && user.AssignedLicenses.Any() ? string.Join(", ", user.AssignedLicenses.SelectMany(license => license.DisabledPlans?.Where(planId => planId.HasValue).Select(planId => ServicePlanMapping.GetServicePlanById(planId.Value.ToString())) ?? Enumerable.Empty<string>())) : "No Disabled Plans"

                    }).ToList();
                    if (users.Count > 0)
                    {

                        dataGridView1.DataSource = users;
                    }
                    else
                    {
                        MessageBox.Show("No users found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async void buttonGetGuests_Click(object sender, EventArgs e)
        {
            progressBar1.Text = "Authenticating";
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "User.Read.All", "Directory.Read.All" };

                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);

                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // Initialize an empty list to store all users
                var allUsers = new List<object>();

                // Fetch the first page of guest users
                progressBar1.Text = "Getting Guest Users";
                var usersPage = await graphClient.Users.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "UserType eq 'Guest'";
                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "UserType","displayName", "Mail", "jobTitle", "UserPrincipalName", "Id", "OnPremisesSyncEnabled", "CreatedDateTime", "ProxyAddresses", "AssignedLicenses", "AssignedPlans", "ServiceProvisioningErrors", "SignInSessionsValidFromDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName", "OnPremisesLastSyncDateTime"
                    };
                });

                // Process each page of results
                while (usersPage != null)
                {
                    if (usersPage.Value != null)
                    {
                        var users = usersPage.Value.Select(user => new
                        {
                            UserType = user.UserType ?? "Not Available",
                            DisplayName = user.DisplayName ?? "Not Available",
                            Email = user.Mail ?? "Not Available",
                            JobTitle = user.JobTitle ?? "Not Available",
                            UserPrincipalName = user.UserPrincipalName ?? "Not Available",
                            ID = user.Id ?? "Not Available",
                            AccountEnabled = user.AccountEnabled?.ToString() ?? "Not Available",
                            City = user.City ?? "Not Available",
                            Country = user.Country ?? "Not Available",
                            Department = user.Department ?? "Not Available",
                            OfficeLocation = user.OfficeLocation ?? "Not Available",
                            LastPasswordChangeDateTime = user.LastPasswordChangeDateTime?.UtcDateTime.ToString() ?? "Not Available",
                            SignInSessionsValidFromDateTime = user.SignInSessionsValidFromDateTime?.UtcDateTime.ToString() ?? "No Sign In Sessions",
                            SyncEnabled = user.OnPremisesSyncEnabled?.ToString() ?? "Not Synced",
                            WhenCreated = user.CreatedDateTime?.ToString() ?? "Not Available",
                            ProxyAddresses = user.ProxyAddresses != null && user.ProxyAddresses.Any() ? string.Join(", ", user.ProxyAddresses) : "No Proxy Address",
                        });

                        // Add the current page of users to the full list
                        allUsers.AddRange(users);
                    }

                    if (usersPage.OdataNextLink != null)
                    {
                        usersPage = await graphClient.Users.GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Filter = "UserType eq 'Guest'";
                            requestConfiguration.QueryParameters.Select = new[]
                            {
                                "UserType","displayName", "Mail", "jobTitle", "UserPrincipalName", "Id", "OnPremisesSyncEnabled", "CreatedDateTime", "ProxyAddresses", "AssignedLicenses", "AssignedPlans", "ServiceProvisioningErrors", "SignInSessionsValidFromDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName", "OnPremisesLastSyncDateTime"
                            };
                        });
                    }
                    else
                    {
                        usersPage = null;
                    }
                }

                // Bind the aggregated user data to the DataGridView
                if (allUsers.Any())
                {
                    dataGridView1.DataSource = allUsers;
                }
                else
                {
                    MessageBox.Show("No users found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }


        private async void buttonGetUnlicensed_Click(object sender, EventArgs e)
        {
            progressBar1.Text = "Authenticating";
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "User.Read.All", "Directory.Read.All" };

                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);

                var graphClient = new GraphServiceClient(interactiveCredential, scopes);
                progressBar1.Text = "Getting Unlicensed Users";
                var usersResponse = await graphClient.Users.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[]
                    {
        "UserType", "displayName", "Mail", "jobTitle", "UserPrincipalName", "Id", "OnPremisesSyncEnabled",
        "CreatedDateTime", "ProxyAddresses", "AssignedLicenses", "AssignedPlans", "ServiceProvisioningErrors",
        "SignInSessionsValidFromDateTime", "OnPremisesImmutableId", "OnPremisesDistinguishedName",
        "OnPremisesLastSyncDateTime","AccountEnabled","Manager"
    };
                });

                if (usersResponse?.Value != null)
                {
                    // Filter for users with no assigned licenses
                    var unlicensedUsers = usersResponse.Value
    .Where(user => user.AssignedLicenses == null || !user.AssignedLicenses.Any())
    .Select(user => new
    {
        UserType = user.UserType ?? "Not Available",
        DisplayName = user.DisplayName ?? "Not Available",
        Email = user.Mail ?? "Not Available",
        JobTitle = user.JobTitle ?? "Not Available",
        UserPrincipalName = user.UserPrincipalName ?? "Not Available",
        Manager = user.Manager?.Id != null ? GetUserPrincipalName(graphClient, user.Manager.Id).Result : "Not Available",
        ID = user.Id ?? "Not Available",
        AccountEnabled = user.AccountEnabled?.ToString() ?? "Not Available",
        City = user.City ?? "Not Available",
        Country = user.Country ?? "Not Available",
        Department = user.Department ?? "Not Available",
        OfficeLocation = user.OfficeLocation ?? "Not Available",
        LastPasswordChangeDateTime = user.LastPasswordChangeDateTime?.UtcDateTime.ToString() ?? "Not Available",
        SignInSessionsValidFromDateTime = user.SignInSessionsValidFromDateTime?.UtcDateTime.ToString() ?? "No Sign In Sessions",
        OnPremisesImmutableId = user.OnPremisesImmutableId ?? "No Immutable Id",
        OnPremisesDistinguishedName = user.OnPremisesDistinguishedName ?? "No Distinguished Name",
        OnPremisesLastSyncDateTime = user.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "No Last Sync DateTime",
        SyncEnabled = user.OnPremisesSyncEnabled?.ToString() ?? "Not Synced",
        WhenCreated = user.CreatedDateTime?.UtcDateTime.ToString() ?? "Not Available",
        ProxyAddresses = user.ProxyAddresses != null && user.ProxyAddresses.Any()
            ? string.Join(", ", user.ProxyAddresses)
            : "No Proxy Address",
        AssignedLicenses = "No Assigned Licenses",
        DisabledPlans = "No Disabled Plans"
    }).ToList();

                    if (unlicensedUsers.Any())
                    {
                        dataGridView1.DataSource = unlicensedUsers;
                    }
                    else
                    {
                        MessageBox.Show("No unlicensed users found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("No users found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async void buttonGetAllGroups_Click(object sender, EventArgs e)
        {
            try
            {
                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // List to hold all groups (handles paginated responses)
                var allGroups = new List<dynamic>();

                // Initialize progress bar
                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Text = "Getting Groups";

                // Paginated request
                var groupsPage = await graphClient.Groups.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                        "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                        "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                    };
                });

                // Loop through all pages
                while (groupsPage != null)
                {
                    if (groupsPage.Value != null)
                    {
                        foreach (var group in groupsPage.Value)
                        {
                            var memberCount = await graphClient.Groups[group.Id].Members.Count.GetAsync((requestConfiguration) =>
                            {
                                requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                            });
                            allGroups.Add(new
                            {
                                Id = group.Id ?? "Not Available",
                                DisplayName = group.DisplayName ?? "Not Available",
                                Description = group.Description ?? "Not Available",
                                Mail = group.Mail ?? "Not Available",
                                MailEnabled = group.MailEnabled?.ToString() ?? "Not Available",
                                SecurityEnabled = group.SecurityEnabled?.ToString() ?? "Not Available",
                                Visibility = group.Visibility ?? "Not Available",
                                GroupTypes = group.GroupTypes != null && group.GroupTypes.Any()
                                    ? string.Join(", ", group.GroupTypes)
                                    : "Not Available",
                                LicenseProcessingState = group.LicenseProcessingState?.State ?? "Not Available",
                                Team = group.Team != null ? "Team Enabled" : "Not a Team",
                                OnPremisesSyncEnabled = group.OnPremisesSyncEnabled?.ToString() ?? "Not Available",
                                OnPremisesLastSyncDateTime = group.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "Not Available",
                                OnPremisesSecurityIdentifier = group.OnPremisesSecurityIdentifier ?? "Not Available",
                                OnPremisesDomainName = group.OnPremisesDomainName ?? "Not Available",
                                MemberCount = memberCount
                            });
                        }
                    }

                    // Get the next page if available
                    groupsPage = groupsPage.OdataNextLink != null
                        ? await graphClient.Groups.GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Select = new[]
                            {
                                "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                                "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                                "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                            };
                            requestConfiguration.QueryParameters.Skip = allGroups.Count;
                        })
                        : null;
                }

                // Hide progress bar
                progressBar1.Visible = false;

                // Display the data in a DataGridView
                if (allGroups.Any())
                {
                    dataGridView1.DataSource = allGroups;
                }
                else
                {
                    MessageBox.Show("No groups found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void buttonUnifiedGroups_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // List to hold all groups (handles paginated responses)
                var allGroups = new List<dynamic>();
                progressBar1.Text = "Getting Unified Groups";
                // Paginated request
                var groupsPage = await graphClient.Groups.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "groupTypes/any(g:g eq 'Unified')";
                    requestConfiguration.QueryParameters.Select = new[]
                    {
                "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                "Visibility", "GroupTypes", "LicenseProcessingState", "Team","CreatedDateTime"
            };
                });

                // Loop through all pages
                while (groupsPage != null)
                {
                    if (groupsPage.Value != null)
                    {
                        allGroups.AddRange(groupsPage.Value.Select(group => new
                        {
                            Id = group.Id ?? "Not Available",
                            DisplayName = group.DisplayName ?? "Not Available",
                            Description = group.Description ?? "Not Available",
                            Mail = group.Mail ?? "Not Available",
                            MailEnabled = group.MailEnabled?.ToString() ?? "Not Available",
                            SecurityEnabled = group.SecurityEnabled?.ToString() ?? "Not Available",
                            Visibility = group.Visibility ?? "Not Available",
                            GroupTypes = group.GroupTypes != null && group.GroupTypes.Any()
                                ? string.Join(", ", group.GroupTypes)
                                : "Not Available",
                            LicenseProcessingState = group.LicenseProcessingState?.State ?? "Not Available",
                            Team = group.Team != null ? "Team Enabled" : "Not a Team",
                            CreatedDateTime = group.CreatedDateTime?.UtcDateTime.ToString() ?? "Not Available"
                        }));
                    }

                    // Get the next page if available
                    groupsPage = groupsPage.OdataNextLink != null
                        ? await graphClient.Groups.GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Filter = "groupTypes/any(g:g eq 'Unified')";

                            requestConfiguration.QueryParameters.Select = new[]
                            {
                                "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                                "Visibility", "GroupTypes", "LicenseProcessingState", "Team","CreatedDateTime"
                            };
                            requestConfiguration.QueryParameters.Skip = allGroups.Count;
                        })
                        : null;
                }

                // Display the data in a DataGridView
                if (allGroups.Any())
                {
                    dataGridView1.DataSource = allGroups;
                }
                else
                {
                    MessageBox.Show("No groups found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {
                    using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "CSV|*.csv", FileName = "DataExport.csv" })
                    {
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            using (StreamWriter sw = new StreamWriter(sfd.FileName))
                            {
                                // Write the header
                                var headers = dataGridView1.Columns.Cast<DataGridViewColumn>();
                                sw.WriteLine(string.Join(",", headers.Select(column => "\"" + column.HeaderText + "\"")));

                                // Write the data
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    var cells = row.Cells.Cast<DataGridViewCell>();
                                    sw.WriteLine(string.Join(",", cells.Select(cell => "\"" + cell.Value?.ToString() + "\"")));
                                }
                            }
                            MessageBox.Show("Data successfully exported.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No data available to export.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void SaveReportToFile(string reportContent)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
                saveFileDialog.Title = "Save Office 365 Active User Report";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog.FileName, reportContent);
                }
            }
        }
        private async void buttonActiveUserCount_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All", "Reports.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // Set the Accept header to application/json
                var activeUsersReport = await graphClient.Reports.GetOffice365ActiveUserCountsWithPeriod("D180").GetAsGetOffice365ActiveUserCountsWithPeriodGetResponseAsync();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async void buttonDistributionGroups_Click(object sender, EventArgs e)
        {
            try
            {
                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // List to hold all groups (handles paginated responses)
                var allGroups = new List<dynamic>();

                // Initialize progress bar
                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Text = "Getting Groups";

                // Paginated request
                var groupsPage = await graphClient.Groups.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "mailEnabled eq true and securityEnabled eq false";

                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                        "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                        "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                    };
                });

                // Loop through all pages
                while (groupsPage != null)
                {
                    if (groupsPage.Value != null)
                    {
                        allGroups.AddRange(groupsPage.Value.Select(group => new
                        {
                            Id = group.Id ?? "Not Available",
                            DisplayName = group.DisplayName ?? "Not Available",
                            Description = group.Description ?? "Not Available",
                            Mail = group.Mail ?? "Not Available",
                            MailEnabled = group.MailEnabled?.ToString() ?? "Not Available",
                            SecurityEnabled = group.SecurityEnabled?.ToString() ?? "Not Available",
                            Visibility = group.Visibility ?? "Not Available",
                            GroupTypes = group.GroupTypes != null && group.GroupTypes.Any()
                                ? string.Join(", ", group.GroupTypes)
                                : "Not Available",
                            LicenseProcessingState = group.LicenseProcessingState?.State ?? "Not Available",
                            Team = group.Team != null ? "Team Enabled" : "Not a Team",
                            OnPremisesSyncEnabled = group.OnPremisesSyncEnabled?.ToString() ?? "Not Available",
                            OnPremisesLastSyncDateTime = group.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "Not Available",
                            OnPremisesSecurityIdentifier = group.OnPremisesSecurityIdentifier ?? "Not Available",
                            OnPremisesDomainName = group.OnPremisesDomainName ?? "Not Available",
                        }));
                    }

                    // Get the next page if available
                    groupsPage = groupsPage.OdataNextLink != null
                        ? await graphClient.Groups.GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Select = new[]
                            {
                                "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                                "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                                "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                            };
                            requestConfiguration.QueryParameters.Skip = allGroups.Count;
                        })
                        : null;
                }

                // Hide progress bar
                progressBar1.Visible = false;

                // Display the data in a DataGridView
                if (allGroups.Any())
                {
                    var filteredGroups = allGroups.Where(group =>
                                         !group.GroupTypes.Contains("Unified")).ToList();
                    dataGridView1.DataSource = filteredGroups;

                }
                else
                {
                    MessageBox.Show("No groups found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private async void button1_Click(object sender, EventArgs e)

        {
            try
            {

                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // List to hold all groups (handles paginated responses)
                var allGroups = new List<dynamic>();

                // Initialize progress bar
                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Text = "Getting Groups";

                // Paginated request
                var groupsPage = await graphClient.Groups.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "mailEnabled eq false and securityEnabled eq true";

                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                        "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                        "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                    };
                });

                // Loop through all pages
                while (groupsPage != null)
                {
                    if (groupsPage.Value != null)
                    {
                        allGroups.AddRange(groupsPage.Value.Select(group => new
                        {
                            Id = group.Id ?? "Not Available",
                            DisplayName = group.DisplayName ?? "Not Available",
                            Description = group.Description ?? "Not Available",
                            Mail = group.Mail ?? "Not Available",
                            MailEnabled = group.MailEnabled?.ToString() ?? "Not Available",
                            SecurityEnabled = group.SecurityEnabled?.ToString() ?? "Not Available",
                            Visibility = group.Visibility ?? "Not Available",
                            GroupTypes = group.GroupTypes != null && group.GroupTypes.Any()
                                ? string.Join(", ", group.GroupTypes)
                                : "Not Available",
                            LicenseProcessingState = group.LicenseProcessingState?.State ?? "Not Available",
                            Team = group.Team != null ? "Team Enabled" : "Not a Team",
                            OnPremisesSyncEnabled = group.OnPremisesSyncEnabled?.ToString() ?? "Not Available",
                            OnPremisesLastSyncDateTime = group.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "Not Available",
                            OnPremisesSecurityIdentifier = group.OnPremisesSecurityIdentifier ?? "Not Available",
                            OnPremisesDomainName = group.OnPremisesDomainName ?? "Not Available",
                        }));
                    }

                    // Get the next page if available
                    groupsPage = groupsPage.OdataNextLink != null
                        ? await graphClient.Groups.GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Select = new[]
                            {
                                "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                                "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                                "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                            };
                            requestConfiguration.QueryParameters.Skip = allGroups.Count;
                        })
                        : null;
                }

                // Hide progress bar
                progressBar1.Visible = false;

                // Display the data in a DataGridView
                if (allGroups.Any())
                {
                    var filteredGroups = allGroups.Where(group =>
                                         !group.GroupTypes.Contains("Unified")).ToList();
                    dataGridView1.DataSource = filteredGroups;

                }
                else
                {
                    MessageBox.Show("No groups found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private async void buttonMailEnabledSec_Click(object sender, EventArgs e)
        {
            try
            {

                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // List to hold all groups (handles paginated responses)
                var allGroups = new List<dynamic>();

                // Initialize progress bar
                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Text = "Getting Groups";

                // Paginated request
                var groupsPage = await graphClient.Groups.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "mailEnabled eq true and securityEnabled eq true";

                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                        "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                        "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                    };
                });

                // Loop through all pages
                while (groupsPage != null)
                {
                    if (groupsPage.Value != null)
                    {
                        allGroups.AddRange(groupsPage.Value.Select(group => new
                        {
                            Id = group.Id ?? "Not Available",
                            DisplayName = group.DisplayName ?? "Not Available",
                            Description = group.Description ?? "Not Available",
                            Mail = group.Mail ?? "Not Available",
                            MailEnabled = group.MailEnabled?.ToString() ?? "Not Available",
                            SecurityEnabled = group.SecurityEnabled?.ToString() ?? "Not Available",
                            Visibility = group.Visibility ?? "Not Available",
                            GroupTypes = group.GroupTypes != null && group.GroupTypes.Any()
                                ? string.Join(", ", group.GroupTypes)
                                : "Not Available",
                            LicenseProcessingState = group.LicenseProcessingState?.State ?? "Not Available",
                            Team = group.Team != null ? "Team Enabled" : "Not a Team",
                            OnPremisesSyncEnabled = group.OnPremisesSyncEnabled?.ToString() ?? "Not Available",
                            OnPremisesLastSyncDateTime = group.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "Not Available",
                            OnPremisesSecurityIdentifier = group.OnPremisesSecurityIdentifier ?? "Not Available",
                            OnPremisesDomainName = group.OnPremisesDomainName ?? "Not Available",
                        }));
                    }

                    // Get the next page if available
                    groupsPage = groupsPage.OdataNextLink != null
                        ? await graphClient.Groups.GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Select = new[]
                            {
                                "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                                "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                                "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                            };
                            requestConfiguration.QueryParameters.Skip = allGroups.Count;
                        })
                        : null;
                }

                // Hide progress bar
                progressBar1.Visible = false;

                // Display the data in a DataGridView
                if (allGroups.Any())
                {
                    var filteredGroups = allGroups.Where(group =>
                                         !group.GroupTypes.Contains("Unified")).ToList();
                    dataGridView1.DataSource = filteredGroups;

                }
                else
                {
                    MessageBox.Show("No groups found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private async void buttonGetSubs_Click(object sender, EventArgs e)
        {
            try
            {
                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All", "Organization.ReadWrite.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // Initialize progress bar
                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Text = "Getting Subscriptions";

                // Fetch subscriptions
                var subscriptions = await graphClient.SubscribedSkus.GetAsync();

                if (subscriptions?.Value != null)
                {
                    var subscriptionList = subscriptions.Value.Select(sub => new
                    {
                        Id = sub.Id ?? "Not Available",
                        SKUID = sub.SkuId?.ToString() ?? "Not Available",
                        Product = Mapping.GetProductNameBySkuId(sub.SkuId?.ToString()) ?? "Not Available",
                        SkuPartNumber = sub.SkuPartNumber ?? "Not Available",
                        ConsumedUnits = sub.ConsumedUnits?.ToString() ?? "Not Available",
                        PrepaidUnitsEnabled = sub.PrepaidUnits?.Enabled?.ToString() ?? "Not Available",
                        PrepaidUnitsSuspended = sub.PrepaidUnits?.Suspended?.ToString() ?? "Not Available",
                        PrepaidUnitsWarning = sub.PrepaidUnits?.Warning?.ToString() ?? "Not Available"
                    }).ToList();

                    dataGridView1.DataSource = subscriptionList;
                }
                else
                {
                    MessageBox.Show("No subscriptions found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async void buttonGetAdmins_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "RoleManagement.Read.Directory", "User.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // Initialize progress bar
                progressBar1.Text = "Getting RBAC Role Members";

                // Fetch roles
                var roles = await graphClient.DirectoryRoles.GetAsync();

                if (roles?.Value != null)
                {
                    var roleMembers = new List<dynamic>();

                    foreach (var role in roles.Value)
                    {
                        var members = await graphClient.DirectoryRoles[role.Id].Members.GetAsync();

                        if (members?.Value != null)
                        {
                            foreach (var member in members.Value)
                            {
                                string displayName = "Not Available";
                                if (member.OdataType == "#microsoft.graph.user")
                                {
                                    var user = await graphClient.Users[member.Id].GetAsync();
                                    displayName = user?.DisplayName ?? "Not Available";
                                }
                                else if (member.OdataType == "#microsoft.graph.group")
                                {
                                    var group = await graphClient.Groups[member.Id].GetAsync();
                                    displayName = group?.DisplayName ?? "Not Available";
                                }

                                else if (member.OdataType == "#microsoft.graph.servicePrincipal")
                                {
                                    var servicename = await graphClient.ServicePrincipals[member.Id].GetAsync();
                                    displayName = servicename?.AppDisplayName ?? "Not Available";

                                }
                                roleMembers.Add(new
                                {
                                    RoleName = role.DisplayName ?? "Not Available",
                                    MemberType = member.OdataType ?? "Not Available",
                                    DisplayName = displayName,
                                    ObjectId = member.Id ?? "Not Available"
                                });
                            }
                        }
                    }

                    if (roleMembers.Any())
                    {
                        dataGridView1.DataSource = roleMembers;
                    }
                    else
                    {
                        MessageBox.Show("No role members found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("No roles found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async void buttonGetDomains_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "Domain.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // Initialize progress bar
                progressBar1.Text = "Getting Domains";

                // Fetch domains
                var domains = await graphClient.Domains.GetAsync();

                if (domains?.Value != null)
                {
                    var domainList = domains.Value.Select(domain => new
                    {
                        DomainName = domain.Id ?? "Not Available",
                        IsVerified = domain.IsVerified?.ToString() ?? "Not Available",
                        IsDefault = domain.IsDefault?.ToString() ?? "Not Available",
                        AuthenticationType = domain.AuthenticationType ?? "Not Available",
                        SupportedServices = domain.SupportedServices != null && domain.SupportedServices.Any()
                            ? string.Join(", ", domain.SupportedServices)
                            : "Not Available"

                    }).ToList();

                    dataGridView1.DataSource = domainList;
                }
                else
                {
                    MessageBox.Show("No domains found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async void buttonGetDomainDependency_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "Domain.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // Initialize progress bar
                progressBar1.Text = "Getting Domain Dependencies";

                // Fetch domains
                var domains = await graphClient.Domains[textBoxDomainName.Text].GetAsync();

                if (domains?.Id != null)
                {
                    var domainDependencies = new List<dynamic>();


                    var domainName = textBoxDomainName.Text;
                    var domainObjects = await graphClient.Domains[domainName].DomainNameReferences.GetAsync();

                    if (domainObjects?.Value != null)
                    {
                        domainDependencies.AddRange(domainObjects.Value.Select(obj => new
                        {
                            DomainName = domainName,
                            ObjectId = obj.Id ?? "Not Available",
                            ObjectType = obj.OdataType ?? "Not Available"
                        }));
                    }


                    if (domainDependencies.Any())
                    {
                        dataGridView1.DataSource = domainDependencies;
                    }
                    else
                    {
                        MessageBox.Show("No domain dependencies found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("No domains found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }

        private async void buttonGetLicensedGroups_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // List to hold all groups (handles paginated responses)
                var allGroups = new List<dynamic>();

                // Initialize progress bar
                progressBar1.Text = "Getting Licensed Groups";

                // Paginated request
                var groupsPage = await graphClient.Groups.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = "assignedLicenses/any()";
                    requestConfiguration.QueryParameters.Select = new[]
                    {
                        "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                        "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                        "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName","assignedLicenses"
                    };
                });

                // Loop through all pages
                while (groupsPage != null)
                {
                    if (groupsPage.Value != null)
                    {
                        allGroups.AddRange(groupsPage.Value.Select(group => new
                        {
                            Id = group.Id ?? "Not Available",
                            DisplayName = group.DisplayName ?? "Not Available",
                            Description = group.Description ?? "Not Available",
                            AssignedLicenses = group.AssignedLicenses != null && group.AssignedLicenses.Any()
                                ? string.Join(", ", group.AssignedLicenses.Select(license => Mapping.GetProductNameBySkuId(license.SkuId.ToString())))
                                : "No Assigned Licenses",
                            Mail = group.Mail ?? "Not Available",
                            MailEnabled = group.MailEnabled?.ToString() ?? "Not Available",
                            SecurityEnabled = group.SecurityEnabled?.ToString() ?? "Not Available",
                            Visibility = group.Visibility ?? "Not Available",
                            GroupTypes = group.GroupTypes != null && group.GroupTypes.Any()
                                ? string.Join(", ", group.GroupTypes)
                                : "Not Available",
                            LicenseProcessingState = group.LicenseProcessingState?.State ?? "Not Available",
                            Team = group.Team != null ? "Team Enabled" : "Not a Team",
                            OnPremisesSyncEnabled = group.OnPremisesSyncEnabled?.ToString() ?? "Not Available",
                            OnPremisesLastSyncDateTime = group.OnPremisesLastSyncDateTime?.UtcDateTime.ToString() ?? "Not Available",
                            OnPremisesSecurityIdentifier = group.OnPremisesSecurityIdentifier ?? "Not Available",
                            OnPremisesDomainName = group.OnPremisesDomainName ?? "Not Available",
                        }));
                    }

                    // Get the next page if available
                    groupsPage = groupsPage.OdataNextLink != null
                        ? await graphClient.Groups.GetAsync(requestConfiguration =>
                        {
                            requestConfiguration.QueryParameters.Filter = "assignedLicenses/any()";
                            requestConfiguration.QueryParameters.Select = new[]
                            {
                                "Id", "DisplayName", "Description", "Mail", "MailEnabled", "SecurityEnabled",
                                "Visibility", "GroupTypes", "LicenseProcessingState", "Team", "OnPremisesSyncEnabled",
                                "OnPremisesLastSyncDateTime", "OnPremisesSecurityIdentifier", "OnPremisesDomainName"
                            };
                            requestConfiguration.QueryParameters.Skip = allGroups.Count;
                        })
                        : null;
                }

                // Hide progress bar
                progressBar1.Visible = false;

                // Display the data in a DataGridView
                if (allGroups.Any())
                {
                    dataGridView1.DataSource = allGroups;
                }
                else
                {
                    MessageBox.Show("No licensed groups found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                progressBar1.Visible = false;
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void buttonGetGroupMembers_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            try
            {
                var scopes = new[] { "Group.Read.All", "GroupMember.Read.All", "Directory.Read.All" };

                // Tenant ID and Client ID from textboxes
                var tenantId = textBoxTenant.Text;
                var clientId = textBoxClientID.Text;

                // Interactive browser credential options
                var options = new InteractiveBrowserCredentialOptions
                {
                    TenantId = tenantId,
                    ClientId = clientId,
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri("http://localhost"),
                };

                var interactiveCredential = new InteractiveBrowserCredential(options);
                var graphClient = new GraphServiceClient(interactiveCredential, scopes);

                // Initialize progress bar
                progressBar1.Text = "Getting Group Members";

                // Fetch group by display name
                var groupName = textBoxGroupName.Text;
                var groups = await graphClient.Groups.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = $"startswith(displayName,'{groupName}')";

                    requestConfiguration.QueryParameters.Select = new[] { "Id", "DisplayName" };
                });

                if (groups?.Value != null && groups.Value.Any())
                {
                    var groupId = groups.Value.First().Id;

                    // Fetch group members
                    var members = await graphClient.Groups[groupId].Members.GetAsync(requestConfiguration=>
                    {
                        requestConfiguration.QueryParameters.Select = new[] { "Id", "displayName", "mail" };
                        requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");

                    });

                    if (members?.Value != null)
                    {
                        var memberList = members.Value.Select(member => new
                        {
                            DisplayName = member.AdditionalData?.ContainsKey("displayName") == true
                                ? member.AdditionalData["displayName"]?.ToString()
                                : "Not Available",
                            mail = member.AdditionalData?.ContainsKey("mail") == true
                                ? member.AdditionalData["userPrincipalName"]?.ToString()
                                : "Not Available",
                            Id = member.Id ?? "Not Available"
                        }).ToList();

                        dataGridView1.DataSource = memberList;
                    }
                    else
                    {
                        MessageBox.Show("No members found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Group not found", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.Visible = false;
        }
    }
}
