const request = require('requestretry')

/**
 * @class
 * @description Contains all functions related to request
 */
class RequestUtils {
  /**
     * @function
     * @name request
     * @param {request.RequestRetryOptions} options
     * @param {Boolean} sendWholeResponse
     * @returns
     */
  request (options, sendWholeResponse = false) {
    return new Promise((resolve, reject) => {
      request(options, (err, response, body) => {
        if (err) {
          console.error('RequestUtils:request: Error in response ', { err })
          return reject(err)
        }
        console.info('RequestUtils:request: Response received ', { url: options.url, statusCode: response.statusCode })
        return resolve(sendWholeResponse ? response : body)
      })
    })
  }
}

module.exports = new RequestUtils()
