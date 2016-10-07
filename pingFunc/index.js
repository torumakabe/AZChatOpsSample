const nconf = require('nconf');
var request = require('request');

//Main flow
module.exports = (context, myTimer) => {

  nconf.argv().env();

  const urlString = `https://${nconf.get('FUNCTION_APP_NAME')}.azurewebsites.net/api/postJob?code=${nconf.get('SLACK_SLASHCMD_TOKEN')}`;

  const options = {
      url: urlString,
      method: 'POST',
      form: {'token': nconf.get('SLACK_SLASHCMD_TOKEN'),
             'team_id': 'dummy',
             'team_domain': 'dummy',
             'channel_id': 'dummy',
             'channel_nam': 'dummy',
             'user_id': 'dummy',
             'command': 'dummy',
             'text': 'ping',
             'response_ur': 'dummy'
            }
  }

  request(options, function (error, response, body) {
      if (!error && response.statusCode == 200) {
          context.log('Pong');
      } else {
          context.log('Error occured: ' + body);
      }
  });

  context.done();

};