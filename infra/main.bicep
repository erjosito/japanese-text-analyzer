@description('Azure region for the dedicated Japanese Text Analyzer backend.')
param location string = resourceGroup().location

@description('Globally unique Foundry resource and endpoint name.')
param accountName string = 'foundry-japanese-text-analyzer-a8fbd8'

@description('Object ID of the Entra user who can run inference from the app.')
param inferenceUserObjectId string

var tags = {
  application: 'japanese-text-analyzer'
  purpose: 'Live Japanese reading app backend - not a temporary demo'
  repository: 'https://github.com/erjosito/japanese-text-analyzer'
}

resource foundry 'Microsoft.CognitiveServices/accounts@2025-06-01' = {
  name: accountName
  location: location
  kind: 'AIServices'
  sku: {
    name: 'S0'
  }
  identity: {
    type: 'SystemAssigned'
  }
  tags: tags
  properties: {
    allowProjectManagement: true
    customSubDomainName: accountName
    disableLocalAuth: true
    // The static PWA calls inference directly using the signed-in user's token.
    publicNetworkAccess: 'Enabled'
  }
}

resource project 'Microsoft.CognitiveServices/accounts/projects@2025-06-01' = {
  parent: foundry
  name: 'japanese-text-analyzer'
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  tags: tags
  properties: {
    displayName: 'Japanese Text Analyzer'
    description: 'Dedicated project for the live Japanese reading and textbook exercise app.'
  }
  // Foundry rejects simultaneous project and deployment updates on one account.
  dependsOn: [
    model
  ]
}

resource model 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: foundry
  name: 'japanese-text-analyzer'
  sku: {
    name: 'GlobalStandard'
    capacity: 30
  }
  properties: {
    model: {
      format: 'OpenAI'
      name: 'gpt-5.4-mini'
      version: '2026-03-17'
    }
    versionUpgradeOption: 'OnceCurrentVersionExpired'
  }
}

var inferenceRoleId = subscriptionResourceId('Microsoft.Authorization/roleDefinitions', '5e0bd9bd-7b93-4f28-af87-19fc36ad61bd')

resource inferenceAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(foundry.id, inferenceUserObjectId, inferenceRoleId)
  scope: foundry
  properties: {
    roleDefinitionId: inferenceRoleId
    principalId: inferenceUserObjectId
    principalType: 'User'
  }
}

output endpoint string = 'https://${accountName}.openai.azure.com'
output deploymentName string = model.name
output projectId string = project.id
