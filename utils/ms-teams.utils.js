const config = require('config')
const { HTTP_METHODS } = require('../constants/constants')
const requestUtils = require('./request.utils')

const microsoftAuthUrl = config.externalApiUrls.microsoftAuthUrl
const graphApiUrl = config.externalApiUrls.graphApiUrl

const tenancyId = config.msTeamsSecrets.tenancyId
const clientId = config.msTeamsSecrets.clientId
const clientSecret = config.msTeamsSecrets.clientSecret

class MSTeamsUtils {
  /**   
     * @function
     * @name getGraphToken
     * @param {String} idToken
     * @param {String} clientId
     * @param {String} clientSecret
     * @returns
     */
  async getGraphToken (authToken) {
    const options = {
      method: HTTP_METHODS.POST,
      url: `${microsoftAuthUrl}${tenancyId}/oauth2/v2.0/token`,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      form: {
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default',
        roles:["Calls.JoinGroupCall.All",
        "Calls.InitiateGroupCall.All",
        "Calls.JoinGroupCallAsGuest.All",
        "User.Read.All",
        "Calls.AccessMedia.All",
        "Calls.Initiate.All"]
      },
      json: true
    }
    const result = await requestUtils.request(options)
    console.log('MSTeamUtils:getGraphToken: Result ' + JSON.stringify(result))
    return { refreshToken: result.refresh_token, accessToken: result.access_token }
  }

  /**
     * @function
     * @name getTeamsProfile
     * @param {String} token
     * @param {String} userId
     * @returns
     */
  async getTeamsProfile (token, userId) {
    console.log('MSTeamUtils:getTeamsProfile: Getting teams profile ', userId)
    const options = {
      method: HTTP_METHODS.GET,
      url: `${graphApiUrl}users/${userId}`,
      headers: {
        Authorization: `Bearer ${token}`
      },
      json: true
    }
    const response = await requestUtils.request(options)
    console.log('MSTeamUtils:getTeamsProfile: Team profile details ', response)
    return response
  }

  /**
     * @function
     * @name createGroup
     * @param {String} token
     * @param {String} projectName
     * @param {Array} ownerIds
     * @returns
     */
  async createGroup (token, projectName, ownerIds) {
    console.log('MSTeamUtils:createGroup: ', ownerIds)
    const now = new Date().getTime()
    const ownersString = ownerIds.map(owner => `${graphApiUrl}users/${owner}`)
    const options = {
      method: HTTP_METHODS.POST,
      url: `${graphApiUrl}groups`,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: {
        description: projectName,
        displayName: projectName,
        groupTypes: ['Unified'],
        mailEnabled: false,
        mailNickname: `Dummy_${now}`,
        securityEnabled: true,
        'owners@odata.bind': ownersString
      },
      json: true
    }
    const response = await requestUtils.request(options)
    return response
  }

  /**
     * @function
     * @name createTeamsTenancyId
     * @param {String} token
     * @param {String} groupId
     */
  async createTeam (token, groupId) {
    console.log('MSTeamUtils:createTeam: ', groupId)
    const options = {
      method: HTTP_METHODS.PUT,
      url: `${graphApiUrl}groups/${groupId}/team`,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: {
        memberSettings: { allowCreateUpdateChannels: true },
        messagingSettings: { allowUserEditMessages: true, allowUserDeleteMessages: true },
        funSettings: { allowGiphy: true, giphyContentRating: 'strict' }
      },
      json: true
    }
    console.log('MSTeamUtils:createTeam: Teams request ' + JSON.stringify(options))
    const response = await requestUtils.request(options, true)
    return response
  }

  /**
     * @function
     * @name addMembers
     * @param {String} token
     * @param {String} groupId
     * @param {String} memberId
     * @returns
     */
  async addMembers (token, groupId, memberId) {
    const options = {
      method: HTTP_METHODS.POST,
      url: `${graphApiUrl}groups/${groupId}/members/$ref`,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: { '@odata.id': `${graphApiUrl}directoryObjects/${memberId}` },
      json: true
    }
    const response = await requestUtils.request(options)
    return response
  }



/**
     * @function
     * @name call
     * @param {String} token
     * @returns
     */
 async call (token) {
    const options = {
      method: HTTP_METHODS.POST,
      url: `${graphApiUrl}communications/calls`,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body:    
      
      {
        "@odata.type": "#microsoft.graph.call",
        "callbackUri": "https://ambitious-river-009967503.1.azurestaticapps.net/callback",
        "requestedModalities": [
          "audio"
        ],
        
        "targets":[{
          "@odata.type": "#microsoft.graph.invitationParticipantInfo",
          "identity": {
            "@odata.type": "#microsoft.graph.identitySet",
            "user": {
                  "id": "f55773ac-9336-427d-aeae-0d188296976c",
                  "@odata.type": "microsoft.graph.identity",
                  "displayName": "VK",
                  "tenantId": "3bed7566-57ac-4ec6-b36c-82ce68dd4afd"
              }
          },
          "allowConversationWithoutHost": true
        }],
        "tenantId":"3bed7566-57ac-4ec6-b36c-82ce68dd4afd",
        "participantId": "1d297e2a-1110-4544-9b01-1704add834e8"
      },
      json: true
    }
    const response = await requestUtils.request(options)
    return response
  }
}


module.exports = new MSTeamsUtils()
