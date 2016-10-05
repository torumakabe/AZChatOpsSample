const nconf = require('nconf');
const uuid = require('node-uuid');
const ArmClient = require('armclient');
const qs = require('querystring');
const Promise = require('bluebird');

const Queue = require('../lib/queue');

// Initialize configuration.
nconf.argv()
  .env()
  .file({ file: '../config.json' });

// ARM client.
const armClient = ArmClient({ 
  subscriptionId: nconf.get('SUBSCRIPTION_ID'),
  auth: ArmClient.clientCredentials({
    tenantId: nconf.get('TENANT_ID'), 
    clientId: nconf.get('CLIENT_ID'),
    clientSecret: nconf.get('CLIENT_SECRET')
  })
});

// Runbook queue.
const queue = Queue(nconf.get('STORAGE_ACCOUNT'), nconf.get('STORAGE_ACCOUNT_KEY'), 'azure-runslash-jobs');
queue.create()
  .catch((err) => { 
    throw err;
  });

// Helper method to execute a runbook.
const executeRunbook = (channel, requestedBy, name, params) => {
  const jobId = uuid.v4();
  const request = {
    properties: {
      runbook: {
        name
      },
      parameters: {
        context: JSON.stringify(params),
        MicrosoftApplicationManagementStartedBy: "\"azure-runslash\"",
        MicrosoftApplicationManagementRequestedBy: `\"${requestedBy}\"`
      }
    },
    tags: {}
  };
  
  return new Promise((resolve, reject) => {
    return queue.send({ posted: new Date(), jobId: jobId, channel: channel, requestedBy: requestedBy, runbook: name })
      .then(() => {
        return armClient.provider(nconf.get('AUTOMATION_RESOURCE_GROUP'), 'Microsoft.Automation')
          .put(`/automationAccounts/${nconf.get('AUTOMATION_ACCOUNT')}/Jobs/${jobId}`, { 'api-version': '2015-10-31' }, request)})
      .then((data) => {
        resolve(data);
        })
      .catch((err) => {
        reject(err);
      });
  });
};

//Main flow
module.exports = function (context, data) {

  context.log('receive a request.');
 
  const body = qs.parse(data);

  // Runbook name is required.
  if (!body.text || body.text.length === 0) {
    context.res = {
      response_type: "in_channel",
      attachments: [{
        color: '#F35A00',
        fallback: `Unable to execute Azure Automation Runbook: The runbook name is required.`,
        text: `Unable to execute Azure Automation Runbook: The runbook name is required.`
      }]
    };
//    context.done();
  }
  
  // Collect information.
  const input = body.text.trim().split(' ');
  const runbook = input[0];
  const params = input.splice(1);
  
  // Execute the runbook.
  executeRunbook(`#${body.channel_name}`, body.user_name, runbook, params)
    .then((data) => {
      const subscriptionsUrl = 'https://portal.azure.com/#resource/subscriptions';
      const runbookUrl = `${subscriptionsUrl}/${nconf.get('SUBSCRIPTION_ID')}/resourceGroups/${nconf.get('AUTOMATION_RESOURCE_GROUP')}/providers/Microsoft.Automation/automationAccounts/${nconf.get('AUTOMATION_ACCOUNT')}/runbooks/${runbook}`;

      context.log('check return value after posting runbook: ' + JSON.stringify(data));
      
      context.res = {
        response_type: 'in_channel',
        attachments: [{
          color: '#00BCF2',
          mrkdwn_in: ['text'],
          fallback: `Azure Automation Runbook ${runbook} has been queued.`,
          text: `Azure Automation Runbook *${runbook}* has been queued (<${runbookUrl}|Open Runbook>).`,
          fields: [
            { 'title': 'Automation Account', 'value': nconf.get('AUTOMATION_ACCOUNT'), 'short': true },
            { 'title': 'Runbook', 'value': runbook, 'short': true },
            { 'title': 'Job ID', 'value': data.body.properties.jobId, 'short': true },
            { 'title': 'Parameters', 'value': `"${params.join('", "')}"`, 'short': true },
          ],
        }]
      };

      context.log('In executeRunbook: ' + JSON.stringify(context.res));

    })
    .catch((err) => {
      if (err) {
//        logger.error(err);  
          context.log(err);
      }
      
      context.res = {
        response_type: 'in_channel',
        attachments: [{
          color: '#F35A00',
          fallback: `Unable to execute Azure Automation Runbook: ${err.message || err.details && err.details.message || err.status}.`,
          text: `Unable to execute Azure Automation Runbook: ${err.message || err.details && err.details.message || err.status}.`
        }]
      };
    });

  context.log('Before context.done: ' + context.res);
  context.done();

};