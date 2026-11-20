// Logic App Template - Email-to-Extraction Flow
// Monitors an Outlook mailbox for incoming emails and pushes each PDF attachment,
// with the mail's sender, subject and received time, to factoriq.ai's mail ingest
// endpoint. factoriq.ai starts the extraction itself.
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
              fetchOnlyWithAttachment: true
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
        Check_has_attachments: {
          type: 'If'
          expression: {
            and: [
              {
                greater: [
                  '@length(triggerBody()?[\'Attachments\'])'
                  0
                ]
              }
            ]
          }
          actions: {
            For_each_attachment: {
              type: 'Foreach'
              foreach: '@triggerBody()?[\'Attachments\']'
              actions: {
                Check_is_PDF: {
                  type: 'If'
                  // Some mail clients send PDFs as application/octet-stream, so match the extension too
                  expression: {
                    or: [
                      {
                        equals: [
                          '@items(\'For_each_attachment\')?[\'ContentType\']'
                          'application/pdf'
                        ]
                      }
                      {
                        endsWith: [
                          '@toLower(items(\'For_each_attachment\')?[\'Name\'])'
                          '.pdf'
                        ]
                      }
                    ]
                  }
                  actions: {
                    Try_process_attachment: {
                      type: 'Scope'
                      actions: {
                        Ingest_mail: {
                          type: 'Http'
                          inputs: {
                            method: 'POST'
                            uri: '@{parameters(\'inboxApiUrl\')}/api/ingest/mail'
                            headers: {
                              Authorization: 'Bearer @{parameters(\'inboxApiKey\')}'
                              'Content-Type': 'application/json'
                            }
                            // One push per PDF, so every PDF in the mail becomes its own document.
                            // The id combines the mail and the attachment, which makes a retry a no-op.
                            body: {
                              message_id: '@{triggerBody()?[\'Id\']}/@{items(\'For_each_attachment\')?[\'Id\']}'
                              sender_address: '@{triggerBody()?[\'From\']}'
                              subject: '@{triggerBody()?[\'Subject\']}'
                              received_at: '@{triggerBody()?[\'DateTimeReceived\']}'
                              body_content: '@{triggerBody()?[\'Body\']}'
                              body_content_type: '@{if(equals(triggerBody()?[\'IsHtml\'], true), \'html\', \'text\')}'
                              document_type: '@{parameters(\'documentType\')}'
                              attachments: [
                                {
                                  name: '@{items(\'For_each_attachment\')?[\'Name\']}'
                                  // Declared as PDF even when the mail client sent application/octet-stream:
                                  // the condition above already matched it as a PDF, and factoriq.ai picks
                                  // the document by content type.
                                  content_type: 'application/pdf'
                                  content_base64: '@{items(\'For_each_attachment\')?[\'ContentBytes\']}'
                                  is_inline: false
                                }
                              ]
                            }
                            // Safe to retry: a repeated message_id returns the existing document
                            // and is not charged again.
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
                        Parse_ingest_response: {
                          type: 'ParseJson'
                          inputs: {
                            content: '@body(\'Ingest_mail\')'
                            schema: {
                              type: 'object'
                              properties: {
                                message_id: {
                                  type: 'integer'
                                }
                                document_id: {
                                  type: [
                                    'integer'
                                    'null'
                                  ]
                                }
                                status: {
                                  type: 'string'
                                }
                              }
                            }
                          }
                          runAfter: {
                            Ingest_mail: [
                              'Succeeded'
                            ]
                          }
                        }
                      }
                      runAfter: {}
                    }
                    Catch_processing_error: {
                      type: 'Scope'
                      actions: {
                        Log_error: {
                          type: 'Compose'
                          inputs: {
                            message: 'Failed to process PDF attachment'
                            attachment: '@items(\'For_each_attachment\')?[\'Name\']'
                            error: '@result(\'Try_process_attachment\')'
                            timestamp: '@utcNow()'
                          }
                          runAfter: {}
                        }
                      }
                      runAfter: {
                        Try_process_attachment: [
                          'Failed'
                          'TimedOut'
                        ]
                      }
                    }
                  }
                  'else': {
                    actions: {}
                  }
                  runAfter: {}
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
            actions: {}
          }
          runAfter: {}
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
