# factoriq.ai Email-to-Extraction - Power Automate

Automatically process PDF email attachments through factoriq.ai using Microsoft Power Automate.

## How It Works

1. A cloud flow watches a user or shared mailbox for new emails with attachments
2. For each PDF attachment, it uploads the file to factoriq.ai via a Custom Connector
3. It then runs AI-powered data extraction on the uploaded document. The call waits until the extraction is finished and returns the result
4. Results are reviewed in the factoriq.ai UI

The flow runs in your Microsoft 365 tenant. factoriq.ai never receives mailbox credentials, only the attachments the flow sends.

## Prerequisites

- A Microsoft 365 account with Power Automate access (see [Run the flow under a service account](#run-the-flow-under-a-service-account))
- A factoriq.ai account with an API key ([generate one here](https://inbox.factoriq.ai))
- If you watch a shared mailbox: **Full Access** permission on that mailbox for the account that owns the flow

## Setup

There is no ready-made Solution package (`.zip`) yet, so the flow is built by hand.

### Step 1: Create the Custom Connector

1. Go to [Power Automate](https://make.powerautomate.com)
2. Click **More** > **Discover all** > **Custom connectors** (under Data)
3. Click **New custom connector** > **Import an OpenAPI file**
4. Name: `factoriq.ai Document Processing`
5. Import the file: `custom-connector/apiDefinition.swagger.json`
6. On the **Security** tab:
   - Authentication type: **API Key**
   - Parameter label: `API Key`
   - Parameter name: `Authorization`
   - Parameter location: **Header**
7. Click **Create connector**
8. On the **Test** tab, create a new connection:
   - Enter your API key in the format: `Bearer idpe_<your-api-key>`
   - Click **Create connection**

If you imported an earlier version of the connector, open it, choose **Update from OpenAPI file**, and import the file again. This adds the `quote_request` document type and the full extraction response.

### Step 2: Create the Cloud Flow

1. Go to **My flows** > **New flow** > **Automated cloud flow**
2. Name: `factoriq.ai Email Processor`
3. Choose the trigger (Office 365 Outlook):
   - **User mailbox:** **When a new email arrives (V2)**
   - **Shared mailbox:** **When a new email arrives in a shared mailbox (V2)**, and set **Original Mailbox Address** to the shared mailbox, e.g. `orders@yourcompany.com`
4. Configure the trigger:
   - **Folder**: Inbox
   - **Only with Attachments**: Yes
   - **Include Attachments**: Yes
5. Click **New step** > **Apply to each**
   - Input: Select **Attachments** from the trigger dynamic content
   - Click **...** > **Settings** > set **Concurrency Control** to **On**, degree **1**
6. Inside Apply to each, add a **Condition** that matches PDFs. Some mail clients send PDFs as `application/octet-stream`, so check the file name as well. Switch the condition to **Edit in advanced mode** and enter:
   ```
   @or(equals(items('Apply_to_each')?['contentType'], 'application/pdf'), endsWith(toLower(items('Apply_to_each')?['name']), '.pdf'))
   ```
7. In the **Yes** branch, add the **factoriq.ai - Upload a document** action:
   - **file**: Select **Attachment Content** from dynamic content
   - **document_type**: `order` (or `invoice`, `data_sheet`, `delivery_note`, `quote_request`)
8. Add **factoriq.ai - Run extraction for an uploaded document**:
   - **document_id**: Select **id** from the Upload response
   - Open **...** > **Settings** and set **Retry Policy** to **None** (see [Known Limitations](#known-limitations))
9. **Save** the flow

### Step 3: Test

1. Click **Test** > **Manually** > **Test**
2. Send an email with a PDF attachment to the monitored mailbox
3. Wait for the flow to trigger and verify all steps succeed

## Run the flow under a service account

A flow runs with the identity of whoever created its Outlook connection. If that is a personal account, the flow stops when the person changes their password, gets a new MFA method, or leaves the company.

For production use:

1. Create a dedicated account, e.g. `svc-factoriq@yourcompany.com`, with a Power Automate license
2. Grant it **Full Access** to the mailbox it should watch, and nothing more
3. Sign in as that account to create the Outlook connection and the factoriq.ai connection
4. Add a second person as **co-owner** of the flow so it can be maintained without the service account password

The Outlook connection has the same mailbox rights as the account behind it. Keep that account's permissions limited to the order mailbox.

## For maintainers: creating a Solution package

Once the flow is built and tested, it can be exported as a Solution package so the next customer can import it instead of building it by hand:

1. Go to [Power Automate](https://make.powerautomate.com) > **Solutions**
2. Click **New solution**:
   - Display name: `factoriq.ai Email-to-Extraction`
   - Publisher: Select or create your publisher
3. Open the solution and click **Add existing** > **Cloud flow** > select your flow
4. Click **Add existing** > **Custom connector** > select `factoriq.ai Document Processing`
5. Click **Export** > **As unmanaged** > download the `.zip`
6. Commit the `.zip` to `power-automate/solution/` and add an import section to this guide

## Verify Your Integration

1. Send a test email **with a PDF attachment** to the monitored mailbox
2. Wait 3-5 minutes for the flow to trigger
3. Check the flow **Run history** in Power Automate:
   - Go to **My flows** > click the flow > **Run history**
   - You should see a successful run
4. Log in to [factoriq.ai](https://inbox.factoriq.ai) and verify the document appears

## Known Limitations

1. **Quota usage**: Each document consumes **2 quota units** (1 for upload + 1 for extraction) of your monthly quota.

2. **Extraction is synchronous**: The extraction action returns only when the extraction has finished. If it fails, factoriq.ai deletes the document so the same file can be uploaded again. An automatic retry would either fail with 404 or run (and bill) the extraction a second time, so keep the retry policy on that action set to **None**.

3. **Polling interval**: Power Automate checks for new emails periodically (~1-5 minutes depending on your license tier).

4. **No email tracing**: The factoriq.ai UI does not show which email a document came from. Check the flow run history for traceability.

5. **PDF attachments only**: The flow sends PDF attachments. The API also accepts PNG, JPEG, TXT and HTML, but sending images from emails would also send logos and signatures.

6. **API key storage**: The API key is stored in the Custom Connector connection, encrypted at rest. It is not visible as a plaintext environment variable.

7. **License requirements**: Running cloud flows requires a Power Automate license. Check your organization's Microsoft 365 plan for included entitlements.

## Troubleshooting

### Flow runs show "Failed" status

- **Check the error details** in the run history. Click the failed run to see which action failed.
- **401 Unauthorized**: Your API key is invalid or expired. Update the Custom Connector connection with a new key.
- **400 Bad Request** on upload: The file may be corrupted, too large, or of an unsupported type.
- **400 / 500** on extraction: The document could not be processed and has been deleted. Resend the email to retry.
- **409 Conflict**: Duplicate document (already uploaded). This is informational.
- **429 Too Many Requests**: Rate limit or monthly quota reached.

### Flow doesn't trigger

- Ensure the Outlook connection is active (not expired)
- Verify emails are arriving in the Inbox folder
- For a shared mailbox: check that the connection account has Full Access and that **Original Mailbox Address** is set
- Check that the flow is turned on

### Custom Connector test fails

- Verify the API key format: `Bearer idpe_<your-key>` (include the `Bearer ` prefix)
- Ensure `inbox.factoriq.ai` is reachable from your network
- Check that your factoriq.ai account is active and approved

### Outlook connection expired

- Go to **Data** > **Connections** in Power Automate
- Find the Office 365 Outlook connection
- Click **...** > **Fix connection** and re-authenticate
