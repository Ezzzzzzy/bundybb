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
// var sheetsLink = "https://docs.google.com/spreadsheets/d/1k6bNyz5a3r-zuG2Jkw-Yeg_FZ06543qOfpPiq7SBpsk/edit#gid=0g"; //whitecloak
var sheetsLink = "https://docs.google.com/spreadsheets/d/1RNKEAZ2HRCWT--Vj8mV3yQOgPYa4qIQJdLEvGTHGc9A/edit#gid=823444506"; //botbrosAI
var app = express();
var GoogleSpreadsheet = require('google-spreadsheet');
var creds = require('./client_secret.json');
var moment = require('moment-timezone');
// var spreadsheetId = '1k6bNyz5a3r-zuG2Jkw-Yeg_FZ06543qOfpPiq7SBpsk'; //whitecloak
var spreadsheetId = '1RNKEAZ2HRCWT--Vj8mV3yQOgPYa4qIQJdLEvGTHGc9A'; //botbrosAI
var doc = new GoogleSpreadsheet(spreadsheetId);
var worksheetNum = 1;

moment().format();
moment.suppressDeprecationWarnings = true;
var controller = Botkit.slackbot({ debug: false });
var bot = controller.spawn({
    token: 'xoxb-265554471089-tjUOBTxm6tXkAE7L8zNEMMjc'
});

bot.startRTM(function(err,bot,payload) {});
controller.setupWebserver(process.env.PORT || 3001, function(err, webserver) {
    controller.createWebhookEndpoints(webserver, bot, function() {});
});

/*************************
 *****bundy-functions*****
 *************************/
function findById(id, rows){
    var i
    for(i in rows)
      if(rows[i].id == id)
        return i;
  }

function timeIn(bot, message, tstamp, worksheetNum, id){
    var today = moment().format('DD/MM/YYYY');
    bot.api.users.info({user:message.user},function(err,response) {
        var rowId = id +response.user.name
        var condi;
        console.log(rowId)
        doc.getRows(1,
            {
            offset: 1,
            },(err,row)=>{
             condi = findById(rowId,row)
             if(condi == undefined){
                doc.useServiceAccountAuth(creds, function (err) {
                    doc.addRow(1, {
                        Id: rowId,
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
            }
            else{
                bot.reply(message, response.user.profile.real_name + ', you cant time in twice a day');
            }
            
        })
    });
 
}

function timeOut(bot, message, tstamp, worksheetNum,id){
    var today = moment().format('DD/MM/YYYY');
    bot.api.users.info({user:message.user},function(err,response) {
        var rowId = id +response.user.name
        doc.useServiceAccountAuth(creds, function (err) {
            doc.getRows(1,
                {
                offset: 1,
                },(err,row)=>{
                if(findById(rowId,row) == undefined){
                    bot.reply(message, response.user.profile.real_name + ', I think you did not time in. Please use my time in command before you time out');
                }  
                else{
                    console.log(findById(rowId,row))
                    row[findById(rowId,row)].Date_Out = today
                    row[findById(rowId,row)].Time_Out = tstamp
                    row[findById(rowId,row)].save();
                    bot.reply(message, response.user.profile.real_name + ', you have timed out.');
                }
            })
            });
    });
}


function report(bot, message, username, fromDate, toDate){
    doc.useServiceAccountAuth(creds, function (err) {
        if (err) console.log(err);
        doc.getInfo((err, info)=>{
            if(err) console.log(err);
            info.worksheets[1].getRows({
                offset: 2,
                orderby: 'col3'
            }, (err, rows)=>{
                //Get Holiday 2018 //worksheetTitle = Holidays 2018
                var holidays = [];
                for(var x=0;x<rows.length;x++){
                    var holiday = {}
                    holiday['date'] = rows[x].date;
                    holiday['day'] = rows[x].day;
                    holiday['holiday'] = rows[x].holiday;
                    holiday['regular'] = rows[x].regular;
                    holidays.push(holiday);
                }
                info.worksheets[0].getRows({
                    offset: 2,
                    orderby: 'col2'
                }, (err, rows)=>{
                    if(err) console.log(err);
                    fromDate = moment(fromDate, 'YYYY-MM-DD').format('DD-MMM-YYYY');
                    toDate = moment(toDate, 'YYYY-MM-DD').format('DD-MMM-YYYY');
                    var arrayOfusername = [];
                    for(var x=0;x<rows.length;x++){
                        if(username === rows[x].username){
                            arrayOfusername.push(rows[x])
                        }
                    }

                    if(arrayOfusername.length != 0){
                        //create worksheet for report
                        doc.addWorksheet({
                            'title': username + " " + fromDate + " to " + toDate,
                            'colCount':'12', 
                            'headers':['Date', 'Day', 'Start_time', 'Finish_time',
                                        'Total_hours' ,'Notes','Reg_day', 'OTH',
                                        'HOL', 'SL', 'VL', 'Total_days']
                        },(err)=>{
                            doc.getInfo(function(err, response){
                                worksheetNum = response.worksheets.length;
                                var arrayOfRowsToBeAdded = [], rowsToBeAdded = [];
                                //get records from masterlist
                                for(var x=0;x<rows.length;x++){
                                    //check username
                                    if(rows[x].username === username){
                                        var dayin, row ={},
                                            todayWithoutTime = moment(rows[x].datein, 'DD-MM-YYYY').format('MMM-D-YY'), 
                                            todayAtTen = moment(todayWithoutTime + " 10:00:00").format('MMM-D-YY HH:mm:ss');
                                        //if date_in are in range, put it in variable "row"(JSON)
                                        if(rows[x].datein != "" && moment.utc(moment(rows[x].datein, 'DD/MM/YYYY').format('MMM-DD-YYYY')).isSameOrAfter(fromDate) && moment.utc(moment(rows[x].datein, 'DD/MM/YYYY').format('MMM-DD-YYYY')).isSameOrBefore(toDate)){
                                            dayin = moment(rows[x].datein + " " + rows[x].timein, 'DD-MM-YYYY HH:mm:ss').format('MMM-D-YY HH:mm:ss');
                                            row['Date'] = moment(rows[x].datein, 'DD/MM/YYYY').format('MMM-DD');
                                            row['Day'] = moment(rows[x].datein, 'DD/MM/YYYY').format('ddd');
                                            row['Start_time'] = rows[x].timein;
                                            row['Notes'] = moment(dayin).isAfter(moment(todayAtTen)) ? "late by: " + moment.utc(moment(dayin).diff(moment(todayAtTen))).format('H:mm:ss') : ""

                                            //check if date is holiday
                                            for(var k=0;k<holidays.length;k++){
                                                if(moment.utc(moment(holidays[k].date).format('MMM-D-YY')).isSame(moment.utc(moment(dayin).format('MMM-D-YY')))){
                                                    var typeOfHoliday = holidays[k].regular == 0 ? 'Special Non-Working Holiday' : 'Reg Holiday';
                                                    // row['HOL'] = holidays[k].holiday; //print in googlesheet the name of holiday
                                                    row['HOL'] = 1;
                                                    row['Reg_day'] = "";
                                                }else{
                                                    row['Reg_day'] = 1;
                                                }
                                            }
                                        }
                                        //if date_out are in range append in variable "row"(JSON)
                                        if(rows[x].dateout != "" && moment.utc(moment(rows[x].dateout, 'DD/MM/YYYY').format('MMM-DD-YYYY')).isSameOrAfter(fromDate) && moment.utc(moment(rows[x].dateout, 'DD/MM/YYYY').format('MMM-DD-YYYY')).isSameOrBefore(toDate)){
                                            var dayout = moment(rows[x].dateout + " " + rows[x].timeout, 'DD-MM-YYYY HH:mm:ss').format('MMM-D-YY HH:mm:ss');
                                            var totalHours = moment.utc(moment(dayout).diff(moment(dayin))).subtract({hours: 1}).format('HH:mm:ss');
                                            row['Finish_time'] = rows[x].timeout;
                                            row['Total_hours'] = totalHours;
                                        }

                                        //push to outer array
                                        if(Object.keys(row).length != 0){
                                            arrayOfRowsToBeAdded.push(row)
                                        }
                                    }
                                }

                                if(arrayOfRowsToBeAdded.length != 0){
                                    arrayOfRowsToBeAdded.map((values, i, arr)=>{
                                        if(i%2 === 0){
                                            var nullValues = { OTH: "", SL: "", VL: "", Total_days: 1,}
                                            rowsToBeAdded.push(Object.assign({},arr[i], nullValues ,arr[i+1]))
                                        }
                                    });

                                    //add the rows in the worksheet that was created
                                    for(var x=0; x<rowsToBeAdded.length;x++){
                                        doc.addRow(worksheetNum, rowsToBeAdded[x], (err,row)=>{
                                            if(err) console.log(err)
                                        })
                                    };
                                    bot.reply(message, "Your report has been made");
                                }else bot.reply(message, "Your report has been made");
                            })
                        })
                    }else bot.reply(message, "No user was found.");
                })
            });
        });
    })
}

doc.useServiceAccountAuth(creds, function (err) {
    doc.getInfo(function(err, response) {
        worksheetNum = response.worksheets.length;
        console.log("Bundy initialized. Working with sheet number " + worksheetNum);
        //===
        //bot commands
        //===

        controller.hears(['^help$'], 'direct_message,direct_mention,mention', function(bot,message){
            bot.reply(message, 'Command Format: \n' +
                '@time <command> \n' +
                'Timestamp Format: HH:MM:SS 24-hr format \n' +
                'Timezone: Asia/Manila \n\n' +
                'Commands:\n'+
                'in \n\t\t times user in\n' +
                'out \n\t\t times user out\n' +
                'in/out timestamp \n\t\t times in/out  user at specified timestamp\n' +
                'user in/out username\n\t\t times in/out or renews time in for specified user at current time\n'+ 
                'user in/out timestamp/username username/timestamp \n\t\t times in/out of specified user at timestamp \n' +
                'new sheet \n\t\t cr    eates a new worksheet and sets that worksheet as the target for timing in and out..'
            );
        });

        controller.hears(['^in$'], 'direct_message,direct_mention,mention', function(bot, message) {
            var timestamp=moment().tz('Asia/Manila').format('HH:mm:ss');
            var id = moment().tz('Asia/Manila').format('MDYY')
            timeIn(bot, message, timestamp, worksheetNum, id);
        });
        
        controller.hears(['^out$'], 'direct_message,direct_mention,mention', function(bot,message) {
            var timestamp = moment().tz('Asia/Manila').format('HH:mm:ss');
            var id = moment().tz('Asia/Manila').format('MDYY')
            timeOut(bot, message, timestamp, worksheetNum, id); 
        });

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

            if(field2!=null && field1!=null) {
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
                        bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands HEHEHEHE");
                    }
                }
                else {
                    bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands HEHEHEHE");
                }
            }
        });

        controller.hears(['^user (.*) (.*)$'], 'direct_message,direct_mention,mention', function(bot,message) {
            var userId;
            var command = message.match[1];
            var field1 = message.match[2];
            var timeInput = moment().tz('Asia/Manila').format('HH:mm:ss'); 
            var today = moment().tz('Asia/Manila').format('DD/MM/YYYY');
            var userFound = false;
            if(field1.substr(0,2)=='<@'){
                userId = field1.substr(2, field1.length-3);
            }
            
            if(command === 'in'){
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
            else if(command === 'out'){
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
            else if(command === 'report'){

            }
            else{
                bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands HEHEHEHE");
            }
        });
        
        controller.hears(['^report (.*) (.*) (.*)$'], 'direct_message,direct_mention,mention', function(bot,message){
            var username = message.match[1], fromDate = message.match[2], toDate = message.match[3];
            report(bot, message, username, fromDate, toDate);
        });

        controller.hears(['(.*) (.*)$'], 'direct_message,direct_mention,mention', function(bot,message) {
            var command = message.match[1];
            var timestamp = message.match[2];
            var id = moment().tz('Asia/Manila').format('MDYY')
            console.log(moment(timestamp, 'HH:mm:ss').isValid())
            if(moment(timestamp, 'HH:mm:ss')) {
                if(command=='in'){
                    timeIn(bot, message, timestamp, worksheetNum, id);
                }
                // else if(command=='renew'){
                //     renew(bot, message, timestamp, worksheetNum);
                // }
                else if(command=='out'){
                    timeOut(bot, message, timestamp, worksheetNum, id); 
                }
                else{
                    bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands.");
                }
            }
            else{
                bot.reply(message, "Please follow the time format (HH:MM:SS) 24-hours.");
            }
        });

    });
});