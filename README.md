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

Automatically process PDF email attachments through factoriq.ai. Monitors your Outlook inbox, extracts PDF attachments, uploads them, and triggers AI-powered data extraction.

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fryaniosys%2Ffactoriq-integrations%2Fmain%2Flogic-app%2Fmain.json)

### What you'll need

- An Azure subscription
- A Microsoft 365 account with Outlook
- A factoriq.ai API key &mdash; [get one here](https://inbox.factoriq.ai)

### How it works

```
Email arrives in Outlook
  └─ Logic App triggers (polls every ~3 min)
       └─ For each PDF attachment:
            ├─ Upload to factoriq.ai
            └─ Trigger extraction
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

The same email-to-extraction flow as a Power Automate cloud flow with a Custom Connector.

### Deploy

1. Import the Custom Connector from [`power-automate/custom-connector/apiDefinition.swagger.json`](power-automate/custom-connector/)
2. Create a connection with your API key (format: `Bearer idpe_<your-key>`)
3. Build the flow following the [step-by-step guide](power-automate/)

**[Full setup guide &rarr;](power-automate/)**

---

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| **logicAppName** | &mdash; | Name for the Logic App resource |
| **mailboxAddress** | &mdash; | Outlook email to monitor |
| **inboxApiUrl** | `https://inbox.factoriq.ai` | API base URL |
| **inboxApiKey** | &mdash; | Your API key (`idpe_` prefix) |
| **documentType** | `order` | `order` / `invoice` / `data_sheet` / `delivery_note` |

## Known Limitations

| Limitation | Details |
|------------|---------|
| Quota usage | Each document uses 2 quota units (upload + extract) |
| Polling interval | ~3 minutes (Logic App Consumption tier) |
| Email tracing | No link between source email and processed document in factoriq.ai |
| Shared mailboxes | Not supported in current templates |

## License

MIT
