# Teams meeting transcript by URL

This command-line sample resolves a Microsoft Teams meeting join URL with Microsoft Graph and prints the latest available transcript in WebVTT format.

## Prerequisites

- Node.js 20 or later and npm
- A Microsoft 365 tenant in which you can create an app registration
- An administrator who can grant Microsoft Graph application permissions and configure a Teams application access policy
- A Teams meeting organized by a user in that tenant, with transcription enabled and a transcript available

## 1. Register an application

1. In the [Microsoft Entra admin center](https://entra.microsoft.com/), open **Identity > Applications > App registrations**.
2. Select **New registration**, enter a name, keep **Accounts in this organizational directory only**, and select **Register**. No redirect URI is required.
3. On the app's **Overview** page, save these values:
   - **Directory (tenant) ID**
   - **Application (client) ID**
4. Open **Certificates & secrets > Client secrets**, create a client secret, and immediately save its **Value**. The value is displayed only once.

## 2. Add Microsoft Graph permissions

In the app registration:

1. Open **API permissions > Add a permission > Microsoft Graph > Application permissions**.
2. Add:
   - `OnlineMeetings.Read.All`
   - `OnlineMeetingTranscript.Read.All`
3. Select **Grant admin consent** for the tenant.

These must be **Application** permissions, not Delegated permissions.

## 3. Configure an application access policy

Microsoft Graph requires a Teams application access policy in addition to the permissions above. A Teams administrator should run the following in PowerShell, replacing the placeholders with the client ID from step 1 and the object ID or user principal name of the meeting organizer. 
Run these commands from **Windows** machine from a PowerShell console run **as administrator**:

```powershell
Install-Module MicrosoftTeams -Scope CurrentUser

# 853dd5d3-bae1-4355-ae45-d6b0add10afa is tenantId of asanadeveloper.onmicrosoft.com
Connect-MicrosoftTeams -UseDeviceAuthentication -TenantId 853dd5d3-bae1-4355-ae45-d6b0add10afa

# optionally check if the command above worked correctly:
Get-CsTenant

New-CsApplicationAccessPolicy `
  -Identity <select any name for the policy> `
  -AppIds "<CLIENT_ID>" `
  -Description "Allow the transcript sample to access organizers' meetings"

Grant-CsApplicationAccessPolicy `
  -PolicyName <policy name defined above> `
  -Global
```

Policy changes can take up to 30 minutes to affect Microsoft Graph calls. See [Configure application access to online meetings](https://learn.microsoft.com/graph/cloud-communication-online-meeting-application-access-policy).

## 4. Configure the project

From the repository root:

```bash
cd meeting-transcript-by-url
npm install
cp .env.example .env
```

Fill in `.env`:

```dotenv
TENANT_ID=853dd5d3-bae1-4355-ae45-d6b0add10afa
CLIENT_ID=<Application (client) ID from point 1.3>
CLIENT_SECRET=<secret value from point 1.4>
TARGET_USER_ID=<meeting-organizer-object-id>
```

To find the organizer object id, go to https://entra.microsoft.com/ -> Users -> find your user -> copy Object ID
The user id must be the user that you will use to create meetings.

## 5. Run the sample

Create a meeting in Outlook, make it a Teams meeting. Copy the join link from meeting description.

```bash
npm install
npm start -- "<copied url>"
```

```bash
e.g. ```
npm start -- https://teams.microsoft.com/l/meetup-join/19%3ameeting_NmZhOGUwZDctZGMzNC00OTg1LTg5NjctMTZkM2Y1YmYxZGM0%40thread.v2/0\?context\=%7b%22Tid%22%3a%22853dd5d3-bae1-4355-ae45-d6b0add10afa%22%2c%22Oid%22%3a%22881d0d67-82a0-4927-b2a3-1730357326b5%22%7d
```


If the command prints `(no transcript available)`, confirm that the meeting has ended or its transcript has been generated, and that transcription was enabled. For `Forbidden` errors, verify admin consent and the organizer's application access policy; allow up to 30 minutes after changing the policy.




