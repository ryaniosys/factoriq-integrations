<p align="center">
  <h1 align="center">factoriq.ai Integrations</h1>
  <p align="center">
    Deploy email-to-extraction workflows in your Azure or Microsoft 365 tenant.
    <br />
    <a href="https://inbox.factoriq.ai/docs"><strong>API Docs</strong></a>
    &middot;
    <a href="https://inbox.factoriq.ai"><strong>factoriq.ai</strong></a>
  </p>
</p>

---

## Azure Logic App

Automatically process incoming email through factoriq.ai. Monitors your Outlook inbox and pushes every new email, with its sender, subject and received time, to factoriq.ai, which runs the AI-powered data extraction. Each PDF attachment becomes a document; an email without a PDF is sent as its body.

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fryaniosys%2Ffactoriq-integrations%2Fmain%2Flogic-app%2Fmain.json)

### What you'll need

- An Azure subscription
- A Microsoft 365 account with Outlook
- A factoriq.ai API key &mdash; [get one here](https://inbox.factoriq.ai)

### How it works

```
Email arrives in Outlook
  └─ Logic App triggers (polls every ~3 min)
       ├─ Has PDF attachments → one push per PDF
       └─ No PDF               → one push with the email body
            └─ factoriq.ai stores the document and runs the extraction
```

### Deploy

1. Click **Deploy to Azure** above
2. Fill in your mailbox address, API key, and document type
3. After deployment, **authorize the Outlook connection**:
   - Go to your Resource Group &rarr; find the API Connection resource
   - Click **Edit API connection** &rarr; **Authorize** &rarr; sign in &rarr; **Save**
4. Send a test email with a PDF attachment to verify

**[Full setup guide &rarr;](logic-app/)**

---

## Power Automate

The same email-to-extraction flow as a Power Automate cloud flow with a Custom Connector. Works with user and shared mailboxes. The flow runs in your tenant, so factoriq.ai never receives mailbox credentials.

### Deploy

1. Import the Custom Connector from [`power-automate/custom-connector/apiDefinition.swagger.json`](power-automate/custom-connector/)
2. Create a connection with your API key (format: `Bearer idpe_<your-key>`)
3. Build the flow following the [step-by-step guide](power-automate/)
4. For production, run the flow under a [dedicated service account](power-automate/#run-the-flow-under-a-service-account)

**[Full setup guide &rarr;](power-automate/)**

---

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| **logicAppName** | &mdash; | Name for the Logic App resource |
| **mailboxAddress** | &mdash; | Outlook email to monitor |
| **inboxApiUrl** | `https://inbox.factoriq.ai` | API base URL |
| **inboxApiKey** | &mdash; | Your API key (`idpe_` prefix) |
| **documentType** | `order` | `order` / `invoice` / `data_sheet` / `delivery_note` / `quote_request` |

## Known Limitations

| Limitation | Details |
|------------|---------|
| Quota usage | Each PDF uses 1 unit of the monthly quota, charged when its extraction starts |
| Synchronous extraction | The ingest call returns when the extraction has finished. A retry is safe: the same mail and attachment is never stored or charged twice |
| Polling interval | ~3 minutes (Logic App Consumption tier) |
| Shared mailboxes | Supported in Power Automate; the Logic App template needs its trigger swapped ([details](logic-app/#known-limitations)) |
| Attachments | PDF only (matched by content type or `.pdf` extension). An email without a PDF is sent as its body; other attachments such as images are not sent |

## License

MIT
