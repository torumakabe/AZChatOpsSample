const ArmClient = require('armclient');
const async = require('async');
const azureStorage = require('azure-storage');
const nconf = require('nconf');
const uuid = require('node-uuid');
const qs = require('querystring');

//Main flow
module.exports = (context, data) => {

  nconf.argv().env();

  const armClient = ArmClient({ 
    subscriptionId: nconf.get('SUBSCRIPTION_ID'),
    auth: ArmClient.clientCredentials({
      tenantId: nconf.get('TENANT_ID'), 
      clientId: nconf.get('CLIENT_ID'),
      clientSecret: nconf.get('CLIENT_SECRET')
      })
  });
 
  const queue = Queue(nconf.get('STORAGE_ACCOUNT'), nconf.get('STORAGE_ACCOUNT_KEY'), 'azure-runslash-jobs');
  queue.create()
    .catch((err) => { 
      context.log('Error occured: ' + err);
    });
 
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

  } else if (!body.text == 'ping') {
      context.res = { status: 200, body: 'Pong' };
  } else {   

    // Collect information.
    const input = body.text.trim().split(' ');
    const name = input[0];
    const params = input.splice(1);
    
    const channel = `#${body.channel_name}`;
    const requestedBy = body.user_name;

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

    return queue.send({ posted: new Date(), jobId: jobId, channel: channel, requestedBy: requestedBy, runbook: name })
      .then(() => {
        return armClient.provider(nconf.get('AUTOMATION_RESOURCE_GROUP'), 'Microsoft.Automation')
          .put(`/automationAccounts/${nconf.get('AUTOMATION_ACCOUNT')}/Jobs/${jobId}`, { 'api-version': '2015-10-31' }, request)})
      .then((rbdata) => {
        const subscriptionsUrl = 'https://portal.azure.com/#resource/subscriptions';
        const runbookUrl = `${subscriptionsUrl}/${nconf.get('SUBSCRIPTION_ID')}/resourceGroups/${nconf.get('AUTOMATION_RESOURCE_GROUP')}/providers/Microsoft.Automation/automationAccounts/${nconf.get('AUTOMATION_ACCOUNT')}/runbooks/${name}`;
       
        context.res = {
          response_type: 'in_channel',
          attachments: [{
            color: '#00BCF2',
            mrkdwn_in: ['text'],
            fallback: `Azure Automation Runbook ${name} has been queued.`,
            text: `Azure Automation Runbook *${name}* has been queued (<${runbookUrl}|Open Runbook>).`,
            fields: [
              { 'title': 'Automation Account', 'value': nconf.get('AUTOMATION_ACCOUNT'), 'short': true },
              { 'title': 'Runbook', 'value': name, 'short': true },
              { 'title': 'Job ID', 'value': jobId, 'short': true },
              { 'title': 'Parameters', 'value': `"${params.join('", "')}"`, 'short': true },
            ],
          }]
        };
        })
      .catch((err) => {
        context.log('Error occured: ' + JSON.stringify(err));
        context.res = {
          response_type: 'in_channel',
          attachments: [{
            color: '#F35A00',
            fallback: `Unable to execute Azure Automation Runbook: ${err.message || err.details && err.details.message || err.status}.`,
            text: `Unable to execute Azure Automation Runbook: ${err.message || err.details && err.details.message || err.status}.`
          }]
        };
      });
  }

  context.done();

};


//Helper functions
const Queue = (accountName, accountKey, queueName) => {
  
  const queueService = azureStorage.createQueueService(accountName, accountKey);
  
  const client = {
    create: () => {
      return new Promise((resolve, reject) => {
        try {
          queueService.createQueueIfNotExists(queueName, (err, res, response) => {
            if (err) {
              return reject(err);
            }
            return resolve();
          });
        } catch (e) {
          return reject(e);
        }
      });
    },
    
    send: (msg) => {   
      return new Promise((resolve, reject) => {
        try {           
          queueService.createMessage(queueName, JSON.stringify(msg, null, 2), { messagettl: 240 * 60 }, (err) => {
            if (err) {
              return reject(err);
            }
            return resolve();
          });
        } catch (e) {
          return reject(e);
        }
      });
    },
    
    update: (msg, contents) => {
      return new Promise((resolve, reject) => {
        try {           
          queueService.updateMessage(queueName, msg.messageid, msg.popreceipt, 5, { messageText: JSON.stringify(contents, null, 2) }, (err, res) => {
            if (err) {
              return reject(err);
            }
            return resolve();
          });
        } catch (e) {
          return reject(e);
        }
      });
      
    },
    
    delete: (msg) => {   
      return new Promise((resolve, reject) => {
        try { 
          queueService.deleteMessage(queueName, msg.messageid, msg.popreceipt, (err, res) => {
            if (err) {
              return reject(err);
            }
            return resolve();
          });
        } catch (e) {
          return reject(e);
        }
      });
    },
    
    getMessages: (messageCallback) => {  
      return new Promise((resolve, reject) => {
        try { 
          queueService.getMessages(queueName, { numOfMessages: 15, visibilityTimeout: 5 }, (err, result, response) => {
            if (err) {
              return reject(err);
            }
            
            if (!result || !result.length) {
              return resolve();
            }
            
            async.each(result, messageCallback, (err) => {
              if (err) {
                return reject(err);
              }
              
              return resolve();
            });
          });
        } catch (e) {
          return reject(e);  
        }
      });
    },
    
    process: (messageCallback) => {
      return new Promise((resolve, reject) => {
        const processMessages = () => {
          client.getMessages(messageCallback)
            .catch(reject);
        };
        processMessages();
      });
    }
  };
  return client;
};