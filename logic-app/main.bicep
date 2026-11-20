// Logic App Template - Email-to-Extraction Flow
// Monitors an Outlook mailbox and pushes every new email to factoriq.ai's mail ingest
// endpoint, with its sender, subject and received time: one push per PDF attachment,
// or one push with the body when the mail has no PDF. factoriq.ai starts the extraction.
//
// Post-deployment: The O365 Outlook connection must be authorized manually
// in the Azure Portal by a user with mailbox access.

// === Parameters ===

@description('Name for the Logic App resource')
param logicAppName string

@description('Outlook mailbox email address to monitor')
param mailboxAddress string

@description('factoriq.ai API base URL (must be HTTPS)')
param inboxApiUrl string = 'https://inbox.factoriq.ai'

@secure()
@description('factoriq.ai API key (idpe_ prefix)')
param inboxApiKey string

@description('Document type for extraction')
@allowed([
  'order'
  'invoice'
  'data_sheet'
  'delivery_note'
  'quote_request'
])
param documentType string = 'order'

// === Variables ===

var office365ConnectionName = '${logicAppName}-office365'
var office365ConnectionDisplayName = 'Outlook - ${mailboxAddress}'

// === Resources ===

// Office 365 Outlook API connection
// After deployment, a user must authorize this connection in the Azure Portal
resource office365Connection 'Microsoft.Web/connections@2016-06-01' = {
  name: office365ConnectionName
  location: resourceGroup().location
  properties: {
    api: {
      id: subscriptionResourceId('Microsoft.Web/locations/managedApis', resourceGroup().location, 'office365')
    }
    displayName: office365ConnectionDisplayName
  }
}

// Logic App (Consumption tier)
resource logicApp 'Microsoft.Logic/workflows@2019-05-01' = {
  name: logicAppName
  location: resourceGroup().location
  tags: {
    purpose: 'email-to-extraction'
    'managed-by': 'factoriq-integrations'
  }
  properties: {
    state: 'Enabled'
    definition: {
      '$schema': 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#'
      contentVersion: '1.0.0.0'
      parameters: {
        '$connections': {
          defaultValue: {}
          type: 'Object'
        }
        inboxApiUrl: {
          type: 'String'
          defaultValue: ''
        }
        inboxApiKey: {
          type: 'SecureString'
          defaultValue: ''
        }
        documentType: {
          type: 'String'
          defaultValue: 'order'
        }
      }
      triggers: {
        When_a_new_email_arrives: {
          type: 'ApiConnection'
          inputs: {
            host: {
              connection: {
                name: '@parameters(\'$connections\')[\'office365\'][\'connectionId\']'
              }
            }
            method: 'get'
            path: '/v2/Mail/OnNewEmail'
            queries: {
              folderId: 'Inbox'
              importance: 'Any'
              fetchOnlyWithAttachment: false
              includeAttachments: true
            }
          }
          recurrence: {
            frequency: 'Minute'
            interval: 3
          }
          splitOn: '@triggerBody()?[\'value\']'
        }
      }
      actions: {
        // PDFs are matched by content type or extension, because some mail clients
        // send them as application/octet-stream.
        Filter_PDF_attachments: {
          type: 'Query'
          inputs: {
            from: '@coalesce(triggerBody()?[\'Attachments\'], json(\'[]\'))'
            where: '@or(equals(item()?[\'ContentType\'], \'application/pdf\'), endsWith(toLower(item()?[\'Name\']), \'.pdf\'))'
          }
          runAfter: {}
        }
        Check_has_PDF: {
          type: 'If'
          expression: {
            and: [
              {
                greater: [
                  '@length(body(\'Filter_PDF_attachments\'))'
                  0
                ]
              }
            ]
          }
          actions: {
            // One push per PDF, so every PDF in the mail becomes its own document.
            // The id combines the mail and the attachment, which makes a retry a no-op.
            For_each_PDF: {
              type: 'Foreach'
              foreach: '@body(\'Filter_PDF_attachments\')'
              actions: {
                Try_ingest_PDF: {
                  type: 'Scope'
                  actions: {
                    Ingest_mail_PDF: {
                      type: 'Http'
                      inputs: {
                        method: 'POST'
                        uri: '@{parameters(\'inboxApiUrl\')}/api/ingest/mail'
                        headers: {
                          Authorization: 'Bearer @{parameters(\'inboxApiKey\')}'
                          'Content-Type': 'application/json'
                        }
                        body: {
                          message_id: '@{triggerBody()?[\'Id\']}/@{items(\'For_each_PDF\')?[\'Id\']}'
                          sender_address: '@{triggerBody()?[\'From\']}'
                          subject: '@{triggerBody()?[\'Subject\']}'
                          received_at: '@{triggerBody()?[\'DateTimeReceived\']}'
                          body_content: '@{triggerBody()?[\'Body\']}'
                          body_content_type: '@{if(equals(triggerBody()?[\'IsHtml\'], true), \'html\', \'text\')}'
                          document_type: '@{parameters(\'documentType\')}'
                          attachments: [
                            {
                              name: '@{items(\'For_each_PDF\')?[\'Name\']}'
                              // Declared as PDF even when the mail client sent application/octet-stream:
                              // the filter already matched it as a PDF, and factoriq.ai picks the
                              // document by content type.
                              content_type: 'application/pdf'
                              content_base64: '@{items(\'For_each_PDF\')?[\'ContentBytes\']}'
                              is_inline: false
                            }
                          ]
                        }
                        // Safe to retry: a repeated message_id returns the existing message
                        // and is not stored or charged again.
                        retryPolicy: {
                          type: 'exponential'
                          count: 2
                          interval: 'PT15S'
                          minimumInterval: 'PT10S'
                          maximumInterval: 'PT1M'
                        }
                      }
                      runAfter: {}
                      runtimeConfiguration: {
                        secureData: {
                          properties: [
                            'inputs'
                          ]
                        }
                      }
                    }
                  }
                  runAfter: {}
                }
                Catch_ingest_PDF: {
                  type: 'Scope'
                  actions: {
                    Log_ingest_PDF: {
                      type: 'Compose'
                      inputs: {
                        message: 'Failed to send PDF attachment'
                            attachment: '@items(\'For_each_PDF\')?[\'Name\']'
                        error: '@result(\'Try_ingest_PDF\')'
                        timestamp: '@utcNow()'
                      }
                      runAfter: {}
                    }
                  }
                  runAfter: {
                    Try_ingest_PDF: [
                      'Failed'
                      'TimedOut'
                    ]
                  }
                }
              }
              runAfter: {}
              runtimeConfiguration: {
                concurrency: {
                  repetitions: 1
                }
              }
            }
          }
          'else': {
            actions: {
              // A mail without a PDF is sent with no attachments; factoriq.ai stores its body
              // as the document.
              Try_ingest_body: {
                type: 'Scope'
                actions: {
                  Ingest_mail_body: {
                    type: 'Http'
                    inputs: {
                      method: 'POST'
                      uri: '@{parameters(\'inboxApiUrl\')}/api/ingest/mail'
                      headers: {
                        Authorization: 'Bearer @{parameters(\'inboxApiKey\')}'
                        'Content-Type': 'application/json'
                      }
                      body: {
                        message_id: '@{triggerBody()?[\'Id\']}'
                        sender_address: '@{triggerBody()?[\'From\']}'
                        subject: '@{triggerBody()?[\'Subject\']}'
                        received_at: '@{triggerBody()?[\'DateTimeReceived\']}'
                        body_content: '@{triggerBody()?[\'Body\']}'
                        body_content_type: '@{if(equals(triggerBody()?[\'IsHtml\'], true), \'html\', \'text\')}'
                        document_type: '@{parameters(\'documentType\')}'
                        // No attachments: factoriq.ai stores the mail body as the document.
                        attachments: []
                      }
                      // Safe to retry: a repeated message_id returns the existing message
                      // and is not stored or charged again.
                      retryPolicy: {
                        type: 'exponential'
                        count: 2
                        interval: 'PT15S'
                        minimumInterval: 'PT10S'
                        maximumInterval: 'PT1M'
                      }
                    }
                    runAfter: {}
                    runtimeConfiguration: {
                      secureData: {
                        properties: [
                          'inputs'
                        ]
                      }
                    }
                  }
                }
                runAfter: {}
              }
              Catch_ingest_body: {
                type: 'Scope'
                actions: {
                  Log_ingest_body: {
                    type: 'Compose'
                    inputs: {
                      message: 'Failed to send mail body'
                      error: '@result(\'Try_ingest_body\')'
                      timestamp: '@utcNow()'
                    }
                    runAfter: {}
                  }
                }
                runAfter: {
                  Try_ingest_body: [
                    'Failed'
                    'TimedOut'
                  ]
                }
              }
            }
          }
          runAfter: {
            Filter_PDF_attachments: [
              'Succeeded'
            ]
          }
        }
      }
      outputs: {}
    }
    parameters: {
      '$connections': {
        value: {
          office365: {
            id: subscriptionResourceId('Microsoft.Web/locations/managedApis', resourceGroup().location, 'office365')
            connectionId: office365Connection.id
            connectionName: office365Connection.name
          }
        }
      }
      inboxApiUrl: {
        value: inboxApiUrl
      }
      inboxApiKey: {
        value: inboxApiKey
      }
      documentType: {
        value: documentType
      }
    }
  }
}

// === Outputs ===

output logicAppId string = logicApp.id
output logicAppName string = logicApp.name
output office365ConnectionName string = office365Connection.name
