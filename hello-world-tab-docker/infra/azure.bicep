param resourceBaseName string

param aadAppClientId string
param aadAppTenantId string
param aadAppOauthAuthorityHost string
@secure()
param aadAppClientSecret string

param location string = resourceGroup().location
var oauthAuthority = uri(aadAppOauthAuthorityHost, aadAppTenantId)

@description('Specifies the docker container image to deploy.')
param containerImage string = 'mcr.microsoft.com/azuredocs/containerapps-helloworld:latest'

@description('Minimum number of replicas that will be deployed')
@minValue(0)
@maxValue(25)
param minReplica int = 1

@description('Maximum number of replicas that will be deployed')
@minValue(0)
@maxValue(25)
param maxReplica int = 3

var teamsMobileOrDesktopAppClientId = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
var teamsWebAppClientId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'
var officeWebAppClientId1 = '4345a7b9-9a63-4910-a426-35363201d503'
var officeWebAppClientId2 = '4765445b-32c6-49b0-83e6-1d93765276ca'
var outlookDesktopAppClientId = 'd3590ed6-52b3-4102-aeff-aad2292ab01c'
var outlookWebAppClientId = '00000002-0000-0ff1-ce00-000000000000'
var allowedClientApplications = [aadAppClientId, teamsMobileOrDesktopAppClientId, teamsWebAppClientId, officeWebAppClientId1, officeWebAppClientId2, outlookDesktopAppClientId, outlookWebAppClientId]

module acr 'containerRegistry.bicep' = {
  name: 'acr'
  params: {
    containerRegistryName: resourceBaseName
  }
}

resource containerAppEnv 'Microsoft.App/managedEnvironments@2023-05-01' = {
  name: resourceBaseName
  location: location
  properties: {}
}

resource frontendApp 'Microsoft.App/containerApps@2023-05-01' = {
  name: '${resourceBaseName}-tab'
  location: location
  properties: {
    managedEnvironmentId: containerAppEnv.id
    configuration: {
      ingress: {
        external: true
        targetPort: 80
        allowInsecure: false
        traffic: [
          {
            latestRevision: true
            weight: 100
          }
        ]
      }
    }
    template: {
      containers: [
        {
          name: '${resourceBaseName}-tab'
          image: containerImage
          resources: {
            cpu: json('.25')
            memory: '.5Gi'
          }
        }
      ]
      scale: {
        minReplicas: minReplica
        maxReplicas: maxReplica
        rules: [
          {
            name: 'http-requests'
            http: {
              metadata: {
                concurrentRequests: '10'
              }
            }
          }
        ]
      }
    }
  }
}

var siteDomain = frontendApp.properties.configuration.ingress.fqdn
var aadApplicationIdUri = 'api://${siteDomain}/${aadAppClientId}'

resource backendApp 'Microsoft.App/containerApps@2023-05-01' = {
  name: '${resourceBaseName}-api'
  location: location
  properties: {
    managedEnvironmentId: containerAppEnv.id
    configuration: {
      secrets: [
        {
          name: 'm365-client-id'
          value: aadAppClientId
        }
        {
          name: 'm365-client-secret'
          value: aadAppClientSecret
        }
        {
          name: 'm365-tenant-id'
          value: aadAppTenantId
        }
        {
          name: 'm365-oauth-authority-host'
          value: aadAppOauthAuthorityHost
        }
      ]
      ingress: {
        external: true
        targetPort: 80
        allowInsecure: false
        traffic: [
          {
            latestRevision: true
            weight: 100
          }
        ]
        corsPolicy: {
          allowCredentials: true
          allowedOrigins: [
            'https://${siteDomain}'
          ]
        }
      }
    }
    template: {
      containers: [
        {
          name: '${resourceBaseName}-api'
          image: containerImage
          resources: {
            cpu: json('.25')
            memory: '.5Gi'
          }
        }
      ]
      scale: {
        minReplicas: minReplica
        maxReplicas: maxReplica
        rules: [
          {
            name: 'http-requests'
            http: {
              metadata: {
                concurrentRequests: '10'
              }
            }
          }
        ]
      }
    }
  }
}

resource authSettings 'Microsoft.App/containerApps/authConfigs@2023-05-01' = {
  name: 'current'
  parent: backendApp
  properties: {
    platform: {
      enabled: true
    }
    identityProviders: {
      azureActiveDirectory: {
        enabled: true
        registration: {
          clientId: aadAppClientId
          openIdIssuer: '${oauthAuthority}/v2.0'
        }
        validation: {
          allowedAudiences: [
            aadAppClientId
            aadApplicationIdUri
          ]
          defaultAuthorizationPolicy: {
            allowedApplications: allowedClientApplications
          }
        }
      }
    }
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output AZURE_CONTAINER_REGISTRY_SERVER string = acr.outputs.loginServer
output AZURE_CONTAINER_REGISTRY_NAME string = acr.outputs.name
output API_FUNCTION_ENDPOINT string = 'https://${backendApp.properties.configuration.ingress.fqdn}'
output BACKEND_APP_NAME string = backendApp.name
output FRONTEND_APP_NAME string = frontendApp.name
output TAB_DOMAIN string = siteDomain
output TAB_ENDPOINT string = 'https://${siteDomain}'
output AZURE_CONTAINER_APP_RESOURCE_ID string = backendApp.id
