import { Client } from '@microsoft/microsoft-graph-client'

import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials'
import { InteractiveBrowserCredential } from '@azure/identity'

const tenantId = 'common'
const clientId = '263164c8-7259-4d24-8fbd-37b803eac572'
const workbookId = 'D056DC955413101C!14506'

const currentTableName = 'CurrentExpenses'
const currentPivotName = 'CurrentPivot'
const currentSheetName = 'Budget 2022-23'

async function login(): Promise<Client> {
  const scopes = ['User.Read', 'Files.ReadWrite.All']

  const client = Client.initWithMiddleware({
    authProvider: new TokenCredentialAuthenticationProvider(
      new InteractiveBrowserCredential({
        tenantId,
        clientId,
        redirectUri: window.location.href
      }),
      { scopes }
    )
  })

  return client
}

let cachedClient: Client | null = null

export async function addRow(
  date: string,
  description: string,
  amount: string
) {
  const client = cachedClient ?? (cachedClient = await login())

  const sessionId = (
    await client
      .api(`/me/drive/items/${workbookId}/workbook/createSession`)
      .post({
        persistChanges: true
      })
  ).id

  try {
    await client
      .api(
        `/me/drive/items/${workbookId}/workbook/tables/${currentTableName}/rows`
      )
      .headers({ ['workbook-session-id']: sessionId })
      .post({
        values: [[description, date, amount, null, null, null]],
        index: null
      })

    const worksheets = await client
      .api(`/me/drive/items/${workbookId}/workbook/worksheets`)
      .headers({ ['workbook-session-id']: sessionId })
      .get()

    const worksheetId = (worksheets.value as any[]).find(
      (i) => i.name == currentSheetName
    ).id

    await client
      .api(
        `/me/drive/items/${workbookId}/workbook/worksheets/${worksheetId}/pivotTables/${currentPivotName}/refresh`
      )
      .headers({ ['workbook-session-id']: sessionId })
      .post({})
  } finally {
    await client
      .api(`/me/drive/items/${workbookId}/workbook/closeSession`)
      .headers({ ['workbook-session-id']: sessionId })
      .post({})
  }
}
