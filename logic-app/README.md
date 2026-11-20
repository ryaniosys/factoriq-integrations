# factoriq.ai Email-to-Extraction - Azure Logic App

Automatically process PDF email attachments through factoriq.ai using an Azure Logic App.

## How It Works

1. The Logic App monitors your Outlook inbox for new emails with attachments (polling every ~3 minutes)
2. For each PDF attachment, it pushes the file to factoriq.ai's mail ingest endpoint (`POST /api/ingest/mail`), together with the mail's sender, subject and received time
3. factoriq.ai stores the document and runs AI-powered data extraction on it. The call returns once the extraction has finished
4. Results are reviewed in the factoriq.ai UI. factoriq.ai keeps the sender, subject and received time with each document

PDFs are matched by content type or by the `.pdf` file extension, since some mail clients send PDFs as `application/octet-stream`.

## Prerequisites

- An Azure subscription
- A Microsoft 365 account with Outlook
- A factoriq.ai account with an API key ([generate one here](https://inbox.factoriq.ai))

## Quick Deploy

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fryaniosys%2Ffactoriq-integrations%2Fmain%2Flogic-app%2Fmain.json)

> **Note:** After clicking the button, you will need to fill in:
> - **Logic App Name** - A name for the Logic App resource (e.g., `factoriq-email-processor`)
> - **Mailbox Address** - The email address to monitor (e.g., `orders@yourcompany.com`)
> - **Inbox Api Key** - Your factoriq.ai API key (starts with `idpe_`)
> - **Document Type** - The type of documents you receive (default: `order`)

After deployment completes, you **must authorize the Outlook connection** (see [Post-Deployment Setup](#post-deployment-setup)).

## Manual Setup (Azure Portal)

If you prefer to create the Logic App manually instead of using the Deploy button:

### Step 1: Create a Logic App

1. Go to the [Azure Portal](https://portal.azure.com)
2. Click **Create a resource** > search for **Logic App**
3. Select **Logic App (Consumption)** and click **Create**
4. Fill in:
   - **Subscription** and **Resource Group** (create new if needed)
   - **Logic App name**: e.g., `factoriq-email-processor`
   - **Region**: Choose your preferred region
5. Click **Review + create** > **Create**

### Step 2: Design the Workflow

1. Open the Logic App and click **Logic App Designer**
2. Search for **Office 365 Outlook** and select **When a new email arrives (V2)**
3. Sign in with your Microsoft 365 account when prompted
4. Configure the trigger:
   - **Folder**: Inbox
   - **Only with Attachments**: Yes
   - **Include Attachments**: Yes
5. Click **New step** > **Control** > **For each**
   - Select **Attachments** from the dynamic content for the input
   - Click the **...** menu on the For each and select **Settings**
   - Set **Concurrency Control** to **On**, **Degree of Parallelism** to **1**
6. Inside the For each, add a **Condition** that matches PDFs. Switch to the code view and use:
   ```
   @or(equals(items('For_each')?['ContentType'], 'application/pdf'), endsWith(toLower(items('For_each')?['Name']), '.pdf'))
   ```
7. In the **True** branch, add an **HTTP** action:
   - **Method**: POST
   - **URI**: `https://inbox.factoriq.ai/api/ingest/mail`
   - **Headers**:
     - `Authorization`: `Bearer <your-api-key>`
     - `Content-Type`: `application/json`
   - **Body**:
     ```json
     {
       "message_id": "@{triggerBody()?['Id']}/@{items('For_each')?['Id']}",
       "sender_address": "@{triggerBody()?['From']}",
       "subject": "@{triggerBody()?['Subject']}",
       "received_at": "@{triggerBody()?['DateTimeReceived']}",
       "body_content": "@{triggerBody()?['Body']}",
       "body_content_type": "@{if(equals(triggerBody()?['IsHtml'], true), 'html', 'text')}",
       "document_type": "order",
       "attachments": [
         {
           "name": "@{items('For_each')?['Name']}",
           "content_type": "application/pdf",
           "content_base64": "@{items('For_each')?['ContentBytes']}",
           "is_inline": false
         }
       ]
     }
     ```
   - `message_id` combines the mail and the attachment, so each PDF becomes its own document and a retry of the same PDF is ignored
   - `content_type` is always `application/pdf`, because the condition above already matched the file as a PDF, even when the mail client sent `application/octet-stream`
8. **Settings** > **Retry Policy**: **Exponential**, 2 retries. Retrying is safe, see [Known Limitations](#known-limitations)
9. **Save** the Logic App

### Step 3: Secure the API Key

1. Click the HTTP action > **Settings**
2. Enable **Secure Inputs** to hide the API key from run history
3. Save the Logic App

## Post-Deployment Setup

After deploying via the button or creating manually, you must authorize the Outlook connection:

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to your **Resource Group**
3. Find the resource of type **API Connection** (named `<logic-app-name>-office365`)
4. Click **Edit API connection**
5. Click **Authorize** and sign in with the Microsoft 365 account that has access to the mailbox
6. Click **Save**

The Logic App will begin monitoring the mailbox on its next polling cycle (~3 minutes).

## Verify Your Integration

1. Send a test email **with a PDF attachment** to the monitored mailbox
2. Wait 3-5 minutes for the Logic App to trigger
3. Check the Logic App **Run History** in the Azure Portal:
   - Go to your Logic App > **Overview** > **Runs history**
   - You should see a successful run with green checkmarks
4. Log in to [factoriq.ai](https://inbox.factoriq.ai) and verify the document appears in your document list

If the run shows errors, see [Troubleshooting](#troubleshooting).

## Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `logicAppName` | Yes | - | Name for the Logic App resource |
| `mailboxAddress` | Yes | - | Outlook email address to monitor |
| `inboxApiUrl` | No | `https://inbox.factoriq.ai` | factoriq.ai API base URL |
| `inboxApiKey` | Yes | - | API key (starts with `idpe_`) |
| `documentType` | No | `order` | Document type: `order`, `invoice`, `data_sheet`, `delivery_note`, `quote_request` |

All resources are deployed to the **resource group's region** (selected via the Region dropdown during deployment).

## Known Limitations

1. **Quota usage**: Each PDF consumes **1 quota unit** of your monthly quota, charged when its extraction starts. If the extraction fails, the unit is refunded.

2. **Extraction is synchronous**: The ingest call returns only when the extraction has finished. A retry is safe: factoriq.ai recognises the same `message_id` and returns the existing document with `"status": "duplicate"`, without storing or charging it again. If the extraction fails, the document stays in factoriq.ai marked as failed and can be reprocessed from the UI.

3. **Polling interval**: The Logic App checks for new emails every ~3 minutes. It is not real-time.

4. **No automatic alerting**: If the Logic App fails (e.g., expired API key), you must check the run history manually. Consider setting up [Azure Monitor alerts](https://learn.microsoft.com/en-us/azure/logic-apps/monitor-logic-apps) for Logic App failure rates.

5. **User mailboxes only**: The template watches the mailbox of the account that authorizes the Outlook connection. To watch a shared mailbox, open the Logic App Designer and replace the trigger with **When a new email arrives in a shared mailbox (V2)**, setting **Original Mailbox Address** to the shared mailbox. The authorizing account needs Full Access to it. The [Power Automate guide](../power-automate/#run-the-flow-under-a-service-account) explains why that account should be a dedicated service account.

6. **API key in deployment history**: When using the Deploy to Azure button, the API key is stored in the Azure deployment history. For sensitive environments, consider using [Azure Key Vault](https://learn.microsoft.com/en-us/azure/key-vault/general/overview) to store the key.

## Troubleshooting

### Logic App runs show "Failed" status

- **Check the error details** in the run history. Click the failed run to see which action failed.
- **401 Unauthorized**: Your API key is invalid or expired. Generate a new key in factoriq.ai and update the Logic App parameters.
- **400 Bad Request**: The PDF file may be corrupted, too large, or not a valid PDF.
- **`"status": "duplicate"`** in the response: this PDF from this mail was already received. This is informational, not an error.
- **429 Too Many Requests**: The monthly quota is used up. Nothing is stored. The call is retried automatically (up to 2 times with exponential backoff).

### Logic App doesn't trigger

- **Check the Outlook connection**: Go to the API Connection resource and verify it shows "Connected". If not, re-authorize it.
- **Check the mailbox**: Ensure emails are arriving in the Inbox (not a subfolder or spam).
- **Check the Logic App is enabled**: Go to Logic App > Overview and ensure the status is "Enabled".

### Documents appear in factoriq.ai but extraction fails

- The document stays in factoriq.ai with the status `extraction_failed`. Open it in the UI and reprocess it there.
- Resending the same email does not help: factoriq.ai ignores a mail and attachment it has already received.

### OAuth connection expired

- Microsoft 365 OAuth tokens refresh automatically, but in rare cases the refresh token may expire (e.g., after password change or admin revocation).
- Re-authorize the connection: API Connection resource > Edit > Authorize > Save.
