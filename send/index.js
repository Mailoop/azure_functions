// module.exports = async function (context, req) {
//     context.log('JavaScript HTTP trigger function processed a request.');

//     const name = (req.query.name || (req.body && req.body.name));
//     const responseMessage = name
//         ? "Hello, " + name + ". This HTTP triggered function executed successfully."
//         : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";

//     context.res = {
//         // status: 200, /* Defaults to 200 */
//         body: responseMessage
//     };
// }


var request = require('request');
const getToken = require('./getToken')

function sendMail(token, email_address, mailbox_message_id) {
  return new Promise((resolve, reject) => {
    const options = {
      method: 'POST',
      url: 'https://graph.microsoft.com/v1.0/users/' + email_address + '/' + mailbox_message_id + '/send',
      headers: {
        'Authorization': 'Bearer ' + token,
        'content-type': 'application/json'
      }
    };
    
    request(options, (error, response, body) => {
      const result = JSON.parse(body);
      if (!error && response.statusCode == 204) {
        resolve(result.value);
      } else {
        reject(result);
      }
    });
  });
}

// For dev purpose only

function listMail(token, email_address) {
    return new Promise((resolve, reject) => {
      const options = {
        method: 'GET',
        url: 'https://graph.microsoft.com/v1.0/users/' + email_address + '/messages',
        headers: {
          'Authorization': 'Bearer ' + token,
          'content-type': 'application/json'
        }
      };
      
      request(options, (error, response, body) => {
        const result = JSON.parse(body);
        if (!error ) {
          resolve(result.value, response.statusCode);
        } else {
          reject(result);
        }
      });
    });
  }

module.exports = function (context, req) {
  try {
    context.log('Starting function');
    if (req.query.email_address) {
      const email_address = req.query.email_address;
      const mailbox_message_id = req.query.mailbox_message_id;

      getToken().then(token => {
        listMail(token, email_address, mailbox_message_id)
          .then((result, statusCode) => {
            context.res = {
              status: statusCode,
              body: JSON.stringify(result),
              headers: {
                'Content-Type': 'application/json'
              }
            };
            context.done();
          }).catch(() => {
            context.log('An error occurred while asking MS Graph API');
            context.done();
          });
      }).catch(()=>{
        context.res = {
            status: 400,
            body: "Impossible to get Token"
          };
          context.done();  
      });
    } else {
      context.res = {
        status: 400,
        body: "Please pass an email_address on the query string"
      };
      context.done();        
    }
  } catch (e) {
    context.res = {
        status: 500,
        body: e
      };
      context.done();        
    }
};