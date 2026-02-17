// Logic App Template - Email-to-Extraction Flow
// Monitors an Outlook mailbox for incoming emails, extracts PDF attachments,
// uploads them to factoriq.ai, and triggers extraction.
//
// Post-deployment: The O365 Outlook connection must be authorized manually
// in the Azure Portal by a user with mailbox access.

// === Parameters ===

@description('Azure region for all resources. Leave blank to use the resource group location.')
@metadata({
  strongType: 'location'
})
param location string

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
  location: location
  properties: {
    api: {
      id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'office365')
    }
    displayName: office365ConnectionDisplayName
  }
}

// Logic App (Consumption tier)
resource logicApp 'Microsoft.Logic/workflows@2019-05-01' = {
  name: logicAppName
  location: location
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
                  expression: {
                    and: [
                      {
                        equals: [
                          '@items(\'For_each_attachment\')?[\'ContentType\']'
                          'application/pdf'
                        ]
                      }
                    ]
                  }
                  actions: {
                    Try_process_attachment: {
                      type: 'Scope'
                      actions: {
                        Upload_document: {
                          type: 'Http'
                          inputs: {
                            method: 'POST'
                            uri: '@{parameters(\'inboxApiUrl\')}/api/documents/upload'
                            headers: {
                              Authorization: 'Bearer @{parameters(\'inboxApiKey\')}'
                            }
                            body: {
                              '$content-type': 'multipart/form-data'
                              '$multipart': [
                                {
                                  headers: {
                                    'Content-Disposition': 'form-data; name="file"; filename="@{items(\'For_each_attachment\')?[\'Name\']}"'
                                  }
                                  body: '@base64ToBinary(items(\'For_each_attachment\')?[\'ContentBytes\'])'
                                }
                                {
                                  headers: {
                                    'Content-Disposition': 'form-data; name="document_type"'
                                  }
                                  body: '@{parameters(\'documentType\')}'
                                }
                              ]
                            }
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
                        Parse_upload_response: {
                          type: 'ParseJson'
                          inputs: {
                            content: '@body(\'Upload_document\')'
                            schema: {
                              type: 'object'
                              properties: {
                                id: {
                                  type: 'integer'
                                }
                                filename: {
                                  type: 'string'
                                }
                                status: {
                                  type: 'string'
                                }
                              }
                            }
                          }
                          runAfter: {
                            Upload_document: [
                              'Succeeded'
                            ]
                          }
                        }
                        Trigger_extraction: {
                          type: 'Http'
                          inputs: {
                            method: 'POST'
                            uri: '@{parameters(\'inboxApiUrl\')}/api/documents/@{body(\'Parse_upload_response\')?[\'id\']}/extract'
                            headers: {
                              Authorization: 'Bearer @{parameters(\'inboxApiKey\')}'
                            }
                            retryPolicy: {
                              type: 'exponential'
                              count: 2
                              interval: 'PT15S'
                              minimumInterval: 'PT10S'
                              maximumInterval: 'PT1M'
                            }
                          }
                          runAfter: {
                            Parse_upload_response: [
                              'Succeeded'
                            ]
                          }
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
            id: subscriptionResourceId('Microsoft.Web/locations/managedApis', location, 'office365')
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
