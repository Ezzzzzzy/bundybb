//===
//SLACK dummy bot for botkit
//===

require('dotenv').config()//test environment

if (!process.env.SLACK_TOKEN) {
    console.log('Error: Specify SLACK_TOKEN in environment');
    process.exit(1);
 }

//get BotKit to spawn bot
var Botkit = require('botkit');
var express = require('express');
var fs = require('fs');
var os = require('os');
var sheetsLink = "https://docs.google.com/spreadsheets/d/1k6bNyz5a3r-zuG2Jkw-Yeg_FZ06543qOfpPiq7SBpsk/edit#gid=0g";
var app = express();
var GoogleSpreadsheet = require('google-spreadsheet');
var creds = require('./client_secret.json');
var moment = require('moment-timezone');
var spreadsheetId = '1k6bNyz5a3r-zuG2Jkw-Yeg_FZ06543qOfpPiq7SBpsk';
var doc = new GoogleSpreadsheet(spreadsheetId);
var worksheetNum = 1;

moment().format();
var controller = Botkit.slackbot({
 debug: false
});
var bot = controller.spawn({
  token: process.env.SLACK_TOKEN
});
bot.startRTM(function(err,bot,payload) {

});
controller.setupWebserver(process.env.PORT || 3001, function(err, webserver) {
    controller.createWebhookEndpoints(webserver, bot, function() {});
 });

/*************************
 *****bundy-functions*****
 *************************/
function timeIn(bot, message, tstamp, worksheetNum){
    var today = moment().format('DD/MM/YYYY');
    bot.api.users.info({user:message.user},function(err,response) {
        doc.useServiceAccountAuth(creds, function (err) {
            doc.addRow(worksheetNum, {
                Username: response.user.name,
                Name: response.user.profile.real_name,
                Date_In: today,
                Time_In: tstamp
            }, function(err, row){
                if (err) {
                    console.log(err);
                }
                bot.reply(message, response.user.profile.real_name + ', you have timed in.');
            });
        });
    });
 }
function timeOut(bot, message, tstamp, worksheetNum){
    var today = moment().format('DD/MM/YYYY');
    bot.api.users.info({user:message.user},function(err,response) {
        doc.useServiceAccountAuth(creds, function (err) {
            doc.addRow(worksheetNum, {
                Username: response.user.name,
                Name: response.user.profile.real_name,
                Date_Out: today,
                Time_Out: tstamp
            }, function(err, row){
                if (err) {
                    console.log(err);
                }
                bot.reply(message, response.user.profile.real_name + ', you have timed out.');
            });
        });
    });
 }

 

doc.useServiceAccountAuth(creds, function (err) {
    doc.getInfo(function(err, response) {
        worksheetNum = response.worksheets.length;
        console.log("Bundy initialized. Working with sheet number " + worksheetNum);
        //===
        //bot commands
        //===
        controller.hears(['^help$'], 'direct_message,direct_mention,mention', function(bot,message) {
            bot.reply(message, 'Command Format: \n' +
                '@bundy <command> \n' +
                'Timestamp Format: HH:MM:SS 24-hr format \n' +
                'Timezone: Asia/Manila \n\n' +
                'Commands:\n'+
                'in \n\t\t times user in\n' +
                'out \n\t\t times user out\n' +
                'in/out timestamp \n\t\t times in/out  user at specified timestamp\n' +
                'user in/out username\n\t\t times in/out or renews time in for specified user at current time\n'+ 
                'user in/out timestamp/username username/timestamp \n\t\t times in/out of specified user at timestamp \n' +
                'new sheet \n\t\t creates a new worksheet and sets that worksheet as the target for timing in and out'
            );
         })
        controller.hears(['^in$'], 'direct_message,direct_mention,mention', function(bot, message) {
            var timestamp =moment().format('HH:mm:ss');
            timeIn(bot, message, timestamp, worksheetNum);
         })
        controller.hears(['^out$'], 'direct_message,direct_mention,mention', function(bot,message) {
            var timestamp = moment().format('HH:mm:ss');
            timeOut(bot, message, timestamp, worksheetNum);  
         })
        controller.hears(['^new sheet$'], 'direct_message,direct_mention,mention', function(bot,message) {
            var worksheetTemp;
            doc.useServiceAccountAuth(creds, function (err) {
                if (err)
                    console.log(err);
                doc.addWorksheet({
                    'colCount':'6', 
                    'headers':['Username', 'Name', 'Date_In', 'Time_In', 'Date_Out' ,'Time_Out']
                }, 
                function(err){
                   doc.getInfo(function(err, response){
                        worksheetNum = response.worksheets.length;
                        bot.reply(message, 'You have created a new sheet.');
                   })
                });
            });
         });
        controller.hears(['^user (.*) (.*) (.*)$'], 'direct_message,direct_mention,mention', function(bot,message) {
            var command = message.match[1];
            var field1 = message.match[2];
            var field2 = message.match[3];
            var timeInput, userId;

            var today = moment().format('DD/MM/YYYY');
            var isTimeInputOk = false;

            var userFound = false;

            if(field2!=null &&field1!=null) {
                if(field1.substr(0,2)=='<@'){
                    userId = field1.substr(2, field1.length-3);
                    timeInput = field2;
                    isTimeInputOk = moment(timeInput,'HH:mm:ss').isValid();
                }
                else if(field2.substr(0,2)=='<@'){
                    userId = field2.substr(2, field2.length-3);
                    timeInput = field1;
                    isTimeInputOk = moment(timeInput,'HH:mm:ss').isValid();
                }
                if(isTimeInputOk){
                    if(command == 'in'){
                        bot.api.users.list({},function(err, response) {
                            for(var i = 0; i < response.members.length; i++){
                                if(response.members[i].id==userId){
                                    doc.useServiceAccountAuth(creds, function (err) {
                                        doc.addRow(worksheetNum, {
                                            Username: response.members[i].name,
                                            Name: response.members[i].real_name,
                                            Date_In: today,
                                            Time_In: timeInput
                                        }, function(err, row){
                                            if (err) {
                                                console.log(err);
                                            }
                                            bot.reply(message, response.members[i].real_name + ', you have timed in.');
                                            userFound = true;
                                        });
                                    });
                                    break;
                                }
                            }
                            if(!userFound){
                                bot.reply(message, 'Cannot locate user: ' + userId);
                            }
                        });
                    }
                    else if(command == 'out'){
                        bot.api.users.list({},function(err, response) {
                            for(var i = 0; i < response.members.length; i++){
                                if(response.members[i].id==userId){
                                    doc.useServiceAccountAuth(creds, function (err) {
                                        doc.addRow(worksheetNum, {
                                            Username: response.members[i].name,
                                            Name: response.members[i].real_name,
                                            Date_Out: today,
                                            Time_Out: timeInput
                                        }, function(err, row){
                                            if (err) {
                                                console.log(err);
                                            }
                                            bot.reply(message, response.members[i].real_name + ', you have timed out.');
                                            userFound = true;
                                        });
                                    });
                                    break;
                                }
                            }
                            if(!userFound){
                                bot.reply(message, 'Cannot locate user: ' + userId);
                            }
                        });
                    }
                    else{
                        bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands");
                    }
                }
                else {
                    bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands");
                }
            }
        })
        controller.hears(['^user (.*) (.*)$'], 'direct_message,direct_mention,mention', function(bot,message) {
            var userId;
            var command = message.match[1];
            var field1 = message.match[2];
            var timeInput = moment().format('HH:mm:ss'); 
            var today = moment().format('DD/MM/YYYY');
            var userFound = false;
            if(field1.substr(0,2)=='<@'){
                userId = field1.substr(2, field1.length-3);
            }
            
            if(command == 'in'){
                bot.api.users.list({},function(err, response) {
                    for(var i = 0; i < response.members.length; i++){
                        if(response.members[i].id==userId){
                            doc.useServiceAccountAuth(creds, function (err) {
                                doc.addRow(worksheetNum, {
                                    Username: response.members[i].name,
                                    Name: response.members[i].real_name,
                                    Date_In: today,
                                    Time_In: timeInput
                                }, function(err, row){
                                    if (err) {
                                        console.log(err);
                                    }
                                    bot.reply(message, response.members[i].real_name + ', you have timed in.');
                                    userFound = true;
                                });
                            });
                            break;
                        }
                    }
                    if(!userFound){
                        bot.reply(message, 'Cannot locate user: ' + userId);
                    }
                });
            }
            else if(command == 'out'){
                bot.api.users.list({},function(err, response) {
                    for(var i = 0; i < response.members.length; i++){
                        if(response.members[i].id==userId){
                            doc.useServiceAccountAuth(creds, function (err) {
                                doc.addRow(worksheetNum, {
                                    Username: response.members[i].name,
                                    Name: response.members[i].real_name,
                                    Date_Out: today,
                                    Time_Out: timeInput
                                }, function(err, row){
                                    if (err) {
                                        console.log(err);
                                    }
                                    bot.reply(message, response.members[i].real_name + ', you have timed out.');
                                    userFound = true;
                                });
                            });    
                            break;
                        }
                    }
                    if(!userFound){
                        bot.reply(message, 'Cannot locate user: ' + userId);
                    }
                });
            }
            else{
                bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands");
            }
        });
        controller.hears(['(.*) (.*)'], 'direct_message,direct_mention,mention', function(bot,message) {
            var command = message.match[1];
            var timestamp = message.match[2];
            if(moment(timestamp, 'HH:mm:ss')) {
                if(command=='in'){
                    timeIn(bot, message, timestamp, worksheetNum);
                }
                // else if(command=='renew'){
                //     renew(bot, message, timestamp, worksheetNum);
                // }
                else if(command=='out'){
                    timeOut(bot, message, timestamp, worksheetNum);
                }
                else{
                    bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands");
                }
            }
            else{
                bot.reply(message, "Please follow the time format (HH:MM:SS) 24-hours.");
            }
        });
    });
});