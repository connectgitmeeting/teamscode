const msTeamsService = require('./services/ms-teams-service')
class Index {
  async invoke () {
    try {
      const { message } = await msTeamsService.teamsCall()
      console.info('Index:invoke: Team created ', { message })
    } catch (error) {
      console.error('Index:invoke: Error during team creation ', { error })
    }
  }
}

new Index()
  .invoke()
