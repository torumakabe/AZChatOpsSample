const ArmClient = require('armclient');
const async = require('async');
const azureStorage = require('azure-storage');
const nconf = require('nconf');
const Slack = require('node-slackr');

//Main flow
module.exports = (context, myTimer) => {

  nconf.argv().env();

  const armClient = ArmClient({ 
    subscriptionId: nconf.get('SUBSCRIPTION_ID'),
    auth: ArmClient.clientCredentials({
      tenantId: nconf.get('TENANT_ID'), 
      clientId: nconf.get('CLIENT_ID'),
      clientSecret: nconf.get('CLIENT_SECRET')
    })
  });

  const slack = new Slack(nconf.get('SLACK_INCOMING_WEBHOOK_URL'), {
    channel: nconf.get('SLACK_CHANNEL')
  });
  
  const queue = Queue(nconf.get('STORAGE_ACCOUNT'), nconf.get('STORAGE_ACCOUNT_KEY'), 'azure-runslash-jobs');

  const worker = (message, cb) => {
    var job = JSON.parse(message.messagetext);
    
    armClient.provider(nconf.get('AUTOMATION_RESOURCE_GROUP'), 'Microsoft.Automation')
      .get(`/automationAccounts/${nconf.get('AUTOMATION_ACCOUNT')}/Jobs/${job.jobId}`, { 'api-version': '2015-10-31' })
      .then((res) => {
        return {
          status: res.body.properties.status,
          provisioningState: res.body.properties.provisioningState
        };
      })
      .then((currentJob) => {
        // The job hasn't changed.
        if (job.status === currentJob.status) {
          return cb();
        }
        
        // Update the current job.
        job.status = currentJob.status;
        job.provisioningState = currentJob.provisioningState;
        
        // Post an update to Slack.
        postToSlack(job);
        
        // We need to keep monitoring this job for updates.
        if (job.status !== 'Completed' && job.status !== 'Failed') {
          return queue.update(message, job)
            .then(cb)
            .catch(cb);
        }
        
        // Job will not receive any more updates, let's stop here.
        return queue.delete(message)
          .then(cb)
          .catch(cb);
      })
      .catch((err) => {
        context.log('Error occured: ' + JSON.stringify(err));
        cb();
      });
  }; 

  // Post the message to slack.
  const postToSlack = (job) => {
    var color;
    var message;
    
    switch (job.status) {
      case 'New':
        return;
      case 'Activating':
        message = `On your marks... Job ${job.jobId} is being activated!`;
        color = '#95A5A6';
        break;
      case 'Running':
        message = `Finally! Job ${job.jobId}  is running!`;
        color = '#95A5A6';
        break;
      case 'Completed':
        message = `Success! Job ${job.jobId}  completed!`;
        color = '#7CD197';
        break;
      case 'Failed':
        message = `Oops! Job ${job.jobId} failed!`;
        color = '#F35A00';
        break;
    }

    const subscriptionsUrl = 'https://portal.azure.com/#resource/subscriptions';
    const runbookUrl = `${subscriptionsUrl}/${nconf.get('SUBSCRIPTION_ID')}/resourceGroups/${nconf.get('AUTOMATION_RESOURCE_GROUP')}/providers/Microsoft.Automation/automationAccounts/${nconf.get('AUTOMATION_ACCOUNT')}/runbooks/${job.runbook}`;
    
    var msg = {
      attachments: [{
        color: color,
        fallback: message,
        mrkdwn_in: ['text'],
        text: `Status update for job '${job.jobId}' (<${runbookUrl}|Open Runbook>).`,
        fields: [
          { 'title': 'Automation Account', 'value': nconf.get('AUTOMATION_ACCOUNT'), 'short': true },
          { 'title': 'Runbook', 'value': job.runbook, 'short': true },
          { 'title': 'Job ID', 'value': job.jobId, 'short': true },
          { 'title': 'Status', 'value': job.status, 'short': true },
        ]
      }]
    };

    slack.notify(msg, (err, result) => {
      if (err) {
        context.log('Error occured: ' + JSON.stringify(err));
      }
    });
  };

  // Start listening.
  queue.create()
    .then(() => {
      context.log(`Listening for messages in ${nconf.get('STORAGE_ACCOUNT')}/azure-runslash-jobs.`);
      queue.process(worker);
      context.log('Check done.');
      context.done();
    })
    .catch((err) => { 
      context.log('Error occured: ' + JSON.stringify(err));
      context.done();
    });

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