using System.Collections.Concurrent;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

Console.WriteLine("Starting to count the AAD objects...");
Console.WriteLine();

TokenCredential authProvider = null; // TODO
GraphServiceClient graphClient = new GraphServiceClient(authProvider);

var deletedUsers = await GetDeletedUsersCount_viaHttpCLient();
var deletedGroups = await GetDeletedGroupsCount_viaHttpCLient();
var deletedApplications = await GetDeletedApplicationsCount_viaHttpCLient();
var deletedDevices = await GetDeletedDevicesCount_viaHttpCLient();
var devices = await GetDevices();
var users = await GetUsers();
var orgContacts = await GetContacts();
var directoryRoles = await GetDirectoryRoles();
var applications = await GetApplications();
var administrativeUnits = await GetAdministrativeUnits();
var groups = await GetGroups();
var servicePrincipals = await GetServicePrincipals();
var oAuth2PermissionGrants = await GetOAuth2PermissionGrants();
var appRoleAssignments = await GetAppRoleAssignments(users);

Console.WriteLine($"Number of Applications              : {applications.Count}");
Console.WriteLine($"Number of AdministrativeUnits       : {administrativeUnits.Count}");
Console.WriteLine($"Number of App Role Assignments      : {appRoleAssignments.Count}");
Console.WriteLine($"Number of Directory Roles           : {directoryRoles.Count}");
Console.WriteLine($"Number of Devices                   : {devices.Count}");
Console.WriteLine($"Number of Groups                    : {groups.Count}");
Console.WriteLine($"Number of Org Contacts              : {orgContacts.Count}");
Console.WriteLine($"Number of GetOAuth2PermissionGrants : {oAuth2PermissionGrants.Count}");
Console.WriteLine($"Number of Service Principals        : {servicePrincipals.Count}");
Console.WriteLine($"Number of Users                     : {users.Count}");
Console.WriteLine($"Number of Deleted Users             : {deletedUsers} ");
Console.WriteLine($"Number of Deleted Groups            : {deletedGroups} ");
Console.WriteLine($"Number of Deleted Applications      : {deletedApplications} ");
Console.WriteLine($"Number of Deleted Devices           : {deletedDevices} ");

Console.WriteLine();
Console.WriteLine("SUM: " + (applications.Count +
                             administrativeUnits.Count +
                             appRoleAssignments.Count +
                             directoryRoles.Count +
                             devices.Count +
                             groups.Count +
                             orgContacts.Count +
                             oAuth2PermissionGrants.Count +
                             servicePrincipals.Count +
                             users.Count +
                             deletedUsers +
                             deletedGroups +
                             deletedApplications +
                             deletedDevices));
Console.WriteLine();

Console.WriteLine("Finished.");


async Task<List<Application>> GetApplications()
{
    var all = new List<Application>();

    var currentSet = await graphClient.Applications
        .Request()
        .Top(900)
        .Select("id")
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<List<ServicePrincipal>> GetServicePrincipals()
{
    var all = new List<ServicePrincipal>();

    var currentSet = await graphClient.ServicePrincipals
        .Request()
        .Top(900)
        .Select("id")
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<List<Device>> GetDevices()
{
    var all = new List<Device>();

    var currentSet = await graphClient.Devices
        .Request()
        .Top(900)
        // .Select("id")
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<List<OrgContact>> GetContacts()
{
    var all = new List<OrgContact>();

    var currentSet = await graphClient.Contacts
        .Request()
        .Top(900)
        // .Select("id")
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<List<User>> GetUsers()
{
    var all = new List<User>();

    var currentSet = await graphClient.Users
        .Request()
        .Top(900)
        // .Select("id")
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<int> GetDeletedUsersCount_viaHttpCLient()
{
    var responseString = await ExecuteMsGraphHttpCallForDeletedObjects("microsoft.graph.user");

    var jsonObject = JsonNode.Parse(responseString)!;

    return jsonObject["@odata.count"]!.GetValue<int>();
}

async Task<int> GetDeletedGroupsCount_viaHttpCLient()
{
    var responseString = await ExecuteMsGraphHttpCallForDeletedObjects("microsoft.graph.group");

    var jsonObject = JsonNode.Parse(responseString)!;

    return jsonObject["@odata.count"]!.GetValue<int>();
}

async Task<int> GetDeletedApplicationsCount_viaHttpCLient()
{
    var responseString = await ExecuteMsGraphHttpCallForDeletedObjects("microsoft.graph.application");

    var jsonObject = JsonNode.Parse(responseString)!;

    return jsonObject["@odata.count"]!.GetValue<int>();
}

async Task<int> GetDeletedDevicesCount_viaHttpCLient()
{
    var responseString = await ExecuteMsGraphHttpCallForDeletedObjects("microsoft.graph.device");

    var jsonObject = JsonNode.Parse(responseString)!;

    return jsonObject["@odata.count"]!.GetValue<int>();
}


async Task<List<AppRoleAssignment>> GetAppRoleAssignments(IEnumerable<User> users)
{
    int failed = 0;
    var all = new ConcurrentBag<AppRoleAssignment>();

    await Parallel.ForEachAsync(users, new ParallelOptions() { MaxDegreeOfParallelism = 50, }, async (user, token) =>
    {
        try
        {
            var result = await graphClient.Users[user.Id].AppRoleAssignments
                .Request()
                .GetAsync();

            foreach (var appRoleAssignment in result.CurrentPage)
            {
                all.Add(appRoleAssignment);
            }

            if (result.NextPageRequest != null)
            {
                throw new NotImplementedException();
            }
        }
        catch (Exception e)
        {
            Interlocked.Increment(ref failed);
            Console.WriteLine(e);
            throw;
        }
    });

    return all.ToList();
}



async Task<List<Group>> GetGroups()
{
    var all = new List<Group>();

    var currentSet = await graphClient.Groups
        .Request()
        .Top(900)
        .Select("id")
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<List<DirectoryRole>> GetDirectoryRoles()
{
    var all = new List<DirectoryRole>();

    var currentSet = await graphClient.DirectoryRoles
        .Request()
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<List<OAuth2PermissionGrant>> GetOAuth2PermissionGrants()
{
    var all = new List<OAuth2PermissionGrant>();

    var currentSet = await graphClient.Oauth2PermissionGrants
        .Request()
        .Top(900)
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}


async Task<List<AdministrativeUnit>> GetAdministrativeUnits()
{
    var all = new List<AdministrativeUnit>();

    var currentSet = await graphClient.Directory.AdministrativeUnits
        .Request()
        .Top(900)
        .Select("id")
        .GetAsync();

    while (currentSet.Count > 0)
    {
        all.AddRange(currentSet.CurrentPage);

        if (currentSet.NextPageRequest != null)
        {
            currentSet = await currentSet.NextPageRequest.GetAsync();
        }
        else
        {
            break;
        }
    }

    return all;
}

async Task<string> ExecuteMsGraphHttpCallForDeletedObjects(string directoryObjectType)
{
    var accessToken = await authProvider.GetTokenAsync(new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }, null, null), CancellationToken.None);

    using HttpClient client = new HttpClient();
    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.Token);
    client.DefaultRequestHeaders.Add("ConsistencyLevel", "eventual");

    string requestUrl = $"https://graph.microsoft.com/v1.0/directory/deletedItems/{directoryObjectType}?$count=true&$select=id,DisplayName";
    using var request = new HttpRequestMessage(new HttpMethod("GET"), requestUrl);
    using var response = await client.SendAsync(request);
    return await response.Content.ReadAsStringAsync();
}
