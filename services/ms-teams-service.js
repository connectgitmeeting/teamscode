const generalUtils = require('../utils/general.utils')
const msTeamUtils = require('../utils/ms-teams.utils')
const { STATUS_CODES } = require('../constants/constants')

class MSTeamsService {
  /**
   * @function
   * @name createTeam
   * @param {*} payload
   * @returns
   */
  async createTeam (payload) {
    try {
      const { authorization, teamName, teamMembers } = payload
      // Get access token from AD to start team creation process
      const { accessToken } = await msTeamUtils.getGraphToken(authorization)
      if (accessToken) {
      // This user
        const adminCreationPromises = teamMembers.filter(p => p.role === 'admin')
          .map(p => msTeamUtils.getTeamsProfile(accessToken, p.userId)
            .then(response => response.id))
        //   Admin team profile ids
        const admins = await Promise.all(adminCreationPromises)
        const group = await msTeamUtils.createGroup(accessToken, teamName, admins)
        if (!group) {
          logger.error('MSTeamsService:createTeam: Error, Failed to create group', { group })
          throw new CustomError('Failed to create group', STATUS_CODES.INTERNAL_SERVER_ERROR)
        }

        // After creation of group, it needs some time to sync data
        // If tried before timeout, will return 404 error
        await generalUtils.wait(2 * 60 * 1000)
        const teamReqResponse = await msTeamUtils.createTeam(accessToken, group.id)
        if (!teamReqResponse) {
          logger.error('MSTeamsService:createTeam: Error, No response received after teams creation')
          throw CustomError('No response received after teams creation', STATUS_CODES.INTERNAL_SERVER_ERROR)
        }
        // If status code is not 200, 201 or 202, then team creation will be failed
        if (teamReqResponse.statusCode !== STATUS_CODES.OK &&
            teamReqResponse.statusCode !== STATUS_CODES.CREATED &&
            teamReqResponse.statusCode !== STATUS_CODES.ACCEPTED) {
          logger.error('MSTeamsService:createTeam: Error, Error while creating team ', { error: teamReqResponse.body })
          throw new CustomError('Error while creating team', teamReqResponse.statusCode)
        }

        try {
          // Get teams profile for members to add
          const teamProfiles = await Promise
            .all(teamMembers
              .filter(p => p.role === 'member')
              .map(async p => {
                return await msTeamUtils
                  .getTeamsProfile(accessToken, p.userId).then(response => response.id)
              }))

          const promises = teamProfiles.map(async id => await msTeamUtils.addMembers(accessToken, group.id, id))
          const addedMembers = await Promise.all(promises)
          console.log('MSTeamsService:createTeam: Members added to team ', addedMembers)
          return { message: 'Team created successfully' }
        } catch (error) {
          console.log('MSTeamsService:createTeam: Error while adding members to team ', error.message)
          throw new CustomError('Error while adding members to team', STATUS_CODES.INTERNAL_SERVER_ERROR)
        }
      } else {
        logger.error('No response received/Error while fetching access token')
        throw new CustomError('No response received/Error while fetching access token', STATUS_CODES.INTERNAL_SERVER_ERROR)
      }
    } catch (error) {
      throw new CustomError(error.message, error.statusCode || STATUS_CODES.INTERNAL_SERVER_ERROR)
    }
  }

  async teamsCall(){
    try {
        let  authorization =''
        // Get access token from AD to start team creation process
        const { accessToken } = await msTeamUtils.getGraphToken(authorization)
        if (accessToken) {
        // This user
        msTeamUtils.call(accessToken).then(res=>res.id);
          }}
  catch (error) {
    console.error(error.message, error.statusCode || STATUS_CODES.INTERNAL_SERVER_ERROR)
}
  }}
module.exports = new MSTeamsService()
