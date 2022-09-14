/**
 * @class
 * @description Contains functions related to general utilities
 */
 class GeneralUtils {
    /**
       * @function
       * @name wait
       * @param {Number} timeout
       */
    wait (timeout) {
      return new Promise((resolve, reject) => {
        setTimeout(resolve, timeout)
      })
    }
  }
  
  module.exports = new GeneralUtils()
  