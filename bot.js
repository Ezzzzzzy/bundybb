//===
//SLACK dummy bot for botkit
//===

//test environment
if (!process.env.SLACK_TOKEN) {
    console.log('Error: Specify SLACK_TOKEN in environment');
    process.exit(1);
}

//get BotKit to spawn bot
var Botkit = require('botkit');
var fs = require('fs');
var os = require('os');
var sheetsLink = "https://docs.google.com/spreadsheets/d/1hmDypfJm73C6996CQngV1N5s-KPApycENq7Xzg33g0c/edit?usp=sharing";
var controller = Botkit.slackbot({
 debug: false
});
var maxRows = 100;
var bot = controller.spawn({
  token: process.env.SLACK_TOKEN
});

// start Slack RTM
bot.startRTM(function(err,bot,payload) {
  // handle errors...
});

//prepare the webhook
controller.setupWebserver(process.env.PORT || 3001, function(err, webserver) {
    controller.createWebhookEndpoints(webserver, bot, function() {
        // handle errors...
    });
});

var GoogleSpreadsheet = require('google-spreadsheet');
var creds = require('./client_secret.json');
var moment = require('moment');
moment().format();
var spreadsheetId = '1hmDypfJm73C6996CQngV1N5s-KPApycENq7Xzg33g0c';
var doc = new GoogleSpreadsheet(spreadsheetId);
//===
//bot commands
//===
/*************************
 *****misc--functions*****
 *************************/
function setCellValue(row_index, col_index, col_range, val_index, val){
    doc.useServiceAccountAuth(creds, function (err) {
        doc.getCells(1, {
            'min-row': row_index, 
            'max-row': row_index, 
            'min-col': col_index,
            'max-col': col_range,
            'return-empty': true
        }, 
        function (err, cells) {
            cells[val_index].setValue(val, function(err){
                if(err){
                    console.log(err);
                }
            });
        });
    });
 }
function dateDMY(){
    var today = new Date();
    var dd = today.getDate();
    var mm = today.getMonth() + 1; //January is 0!
    var yyyy = today.getFullYear();
    if(dd<10){
        dd='0'+dd;
    } 
    if(mm<10){
        mm='0'+mm;
    } 
    var today = dd+'/'+mm+'/'+yyyy;
    return today;
 }
function msToTime(duration) {
    var milliseconds = parseInt((duration%1000)/100)
        , seconds = parseInt((duration/1000)%60)
        , minutes = parseInt((duration/(1000*60))%60)
        , hours = parseInt((duration/(1000*60*60))%24);

    hours = (hours < 10) ? "0" + hours : hours;
    minutes = (minutes < 10) ? "0" + minutes : minutes;
    seconds = (seconds < 10) ? "0" + seconds : seconds;

    return hours + ":" + minutes + ":" + seconds + "." + milliseconds;
 }
function parseMonth(date){
    return parseInt(date.substr(4,6));
 }
function timestampToSeconds(timestamp){
    var tstampValues = timestamp.split(":");
    if(tstampValues[0]==''){
        return 0;
    }
    var sum = 0;

    tstampValues[0] = parseInt(tstampValues[0]*3600);
    tstampValues[1] = parseInt(tstampValues[1]*60);
    tstampValues[2] = parseInt(tstampValues[2]);

    for(var i = 0; i < 3; i++){
        sum += tstampValues[i];
    }
    return sum;
 }


/*************************
 *****bundy-functions*****
 *************************/
function timeIn(bot, message, renewMessage, tstamp){
    var skipTimeIn = false;

    //indexes that increment by three to point to the next name/date
    var nameSelectionIndex = 0;
    var dateSelectionIndex = 2;

    var rowEntryIndex = 2;
    var today = dateDMY();

    bot.api.users.info({user:message.user},function(err,response) {
        doc.useServiceAccountAuth(creds, function (err) {
            doc.getCells(1, {
                'min-row': 2, 
                'max-row': maxRows, 
                'min-col': 1,
                'max-col': 3,
                'return-empty': false
            }, 
            function (err, cells) {
                //iterates through all username values
                while(cells[nameSelectionIndex]!=null){
                    //checks if name is equal
                    if(cells[nameSelectionIndex].value == response.user.name){
                        //checks if there is a date
                        if(cells[dateSelectionIndex]!=null){
                            //checks if the date is the same date as today
                            if(cells[dateSelectionIndex].value == today){
                                skipTimeIn = true;
                                bot.reply(message, renewMessage);
                            }
                        }
                    }
                    //next name
                    nameSelectionIndex += 3;
                    //next date
                    dateSelectionIndex += 3;
                }
                //if user hasn't timed in yet
                rowEntryIndex += (nameSelectionIndex/3);
                if(!skipTimeIn) {
                    setCellValue(rowEntryIndex, 1, 1, 0, response.user.name);
                    setCellValue(rowEntryIndex, 2, 2, 0, response.user.profile.real_name);
                    setCellValue(rowEntryIndex, 3, 3, 0, today);
                    setCellValue(rowEntryIndex, 4, 4, 0, tstamp);
                    bot.reply(message, '<@'+message.user+'>, you have timed in.');
                }
            });
        });
    });
 }
function renew(bot, message, tstamp, date){
    var rowEntryIndex = 2;

    //indexes that increment by three to point to the next name/date
    var nameSelectionIndex = 0;
    var dateSelectionIndex = 2;

    var today = dateDMY();
    var hasTimedInAgain = false;

    bot.api.users.info({user:message.user},function(err,response) {
        doc.useServiceAccountAuth(creds, function (err) {
            doc.getCells(1, {
                'min-row': 2, 
                'max-row': maxRows, 
                'min-col': 1,
                'max-col': 3,
                'return-empty': false
            }, 
            function (err, cells) {
                //iterates through all username values
                while(cells[nameSelectionIndex]!=null){
                    //checks if name is equal
                    if(cells[nameSelectionIndex].value == response.user.name){
                        //checks if there is a date
                        if(cells[dateSelectionIndex]!=null){
                            //checks if the date is the same date as today
                            if(cells[dateSelectionIndex].value == today){
                                hasTimedInAgain = true;
                                setCellValue(rowEntryIndex, 4, 4, 0, tstamp);
                                setCellValue(rowEntryIndex, 5, 5, 0, "");
                                setCellValue(rowEntryIndex, 6, 6, 0, "");
                                bot.reply(message, '<@'+message.user+'>, you have timed in again.');
                            }
                        }
                    }
                    //next name and date
                    nameSelectionIndex += 3;
                    dateSelectionIndex += 3;
                    rowEntryIndex++;
                }
                //if user hasn't timed in yet
                if(!hasTimedInAgain) {
                    bot.reply(message, '<@'+message.user+'>, you have not timed in yet.');                    
                }
            });
        });
    });
 }
function timeOut(bot, message, tstamp){
    var today = dateDMY();
    bot.api.users.info({user:message.user},function(err,response) {
        doc.useServiceAccountAuth(creds, function (err) {
            doc.getCells(1, {
                'min-row': 2, 
                'max-row': maxRows, 
                'min-col': 1,
                'max-col': 3,
                'return-empty': false
            }, 
            function (err, cells) {
                //i and j are pertain to columns
                var i = 0; 
                var j = 2;
                var rowNum = 2;
                while(cells[i]!=null){
                    if(cells[i].value==response.user.name){
                        if(cells[j].value==today){
                            setCellValue(rowNum, 5, 5, 0, tstamp);
                            setCellValue(rowNum, 6, 6, 0, '=E'+rowNum+'-'+'D'+rowNum);
                            bot.reply(message, '<@'+message.user+'>, you have timed out.');
                            break;
                        }
                    }
                    else if (cells[i+1]==null){
                        bot.reply(message, '<@'+message.user+'>, you have not yet timed in. Please time in first');
                    }
                    i+=3;
                    j+=3;
                    rowNum++;
                }
            });
        });
    });
 }

controller.hears(['^help$'], 'direct_message,direct_mention,mention', function(bot,message) {
    bot.reply(message, 'Command Format: \n' +
        '@bundy <command> \n' +
        'Timestamp Format: HH:MM:SS 24-hr format \n\n' +
        'Limitations: Currently limited to 100 rows. After exceeding, please delete entries in the excel sheet\n'+
        'Commands:\n'+
        'in \n\t\t times user in\n' +
        'out \n\t\t times user out\n' +
        'renew \n\t\t renews user time in\n' +
        'in/out/renew timestamp \n\t\t times in/out or renews time in user at specified timestamp\n' +
        'user in/out/renew username\n\t\t times in/out or renews time in for specified user at current time\n'+ 
        'user in/out/renew timestamp/username username/timestamp \n\t\t times in/out or renews time in of specified user at timestamp' 
    );
 })
controller.hears(['^in$'], 'direct_message,direct_mention,mention', function(bot, message) {
    var timestamp =moment().format('HH:mm:ss');
    var renewMsg = 'Please type <renew> to time in again';
    timeIn(bot, message, renewMsg, timestamp);
 })
controller.hears(['^renew$'], 'direct_message,direct_mention,mention', function(bot,message){
    var timestamp = moment().format('HH:mm:ss');
    renew(bot, message, timestamp);
 })
controller.hears(['^out$'], 'direct_message,direct_mention,mention', function(bot,message) {
    var timestamp = moment().format('HH:mm:ss');
    console.log(timestamp);
    timeOut(bot, message, timestamp);  
 })
controller.hears(['^hours$'], 'direct_message,direct_mention,mention', function(bot,message) {
    var timeInput;

    var skipTimeIn = false;
    var skipTimeOutErrorMessage = false;

    //indexes that increment by three to point to the next name/date
    var nameSelectionIndex = 0;
    var timeOutSelectionIndex = 4;
    var hoursSelectionIndex = 5;
    var dateSelectionIndex = 2;


    var rowEntryIndex = 2;
    var today = dateDMY();
    var month = parseMonth(today);
    var totalSeconds = 0;

    bot.api.users.info({user:message.user}, function(err,response) {
        doc.useServiceAccountAuth(creds, function (err) {
            doc.getCells(1, {
                'min-row': 2, 
                'max-row': maxRows, 
                'min-col': 1,
                'max-col': 6,
                'return-empty': true
            }, 
            function (err, cells) {
                //iterates through all username values
                while(cells[nameSelectionIndex]!=null){
                    //checks if name is equal
                    if(cells[nameSelectionIndex].value == response.user.name){
                        //checks if there is a date
                        if(cells[dateSelectionIndex]!=null){
                            //checks if the month is the month today
                            if(parseMonth(cells[dateSelectionIndex].value) == month){
                                if(cells[timeOutSelectionIndex]!=null) {
                                    totalSeconds += (timestampToSeconds(cells[hoursSelectionIndex].value));
                                }
                            }
                        }
                    }
                    //next name
                    nameSelectionIndex += 6;
                    //next date
                    dateSelectionIndex += 6;
                    //next timeout
                    timeOutSelectionIndex +=6;
                    //next hours
                    hoursSelectionIndex +=6;
                }
                if(totalSeconds==0) {
                    bot.reply(message, "You haven\'t timed out yet during this month")
                }
                else {
                    bot.reply(message, "Total Hours for this month:\n" + 
                        parseInt(totalSeconds/3600) + " hours");
                }
            });
        });
    });
 })
controller.hears(['^user (.*) (.*) (.*)$'], 'direct_message,direct_mention,mention', function(bot,message) {
    var command = message.match[1];
    var field1 = message.match[2];
    var field2 = message.match[3];
    var timeInput, userId;

    var skipTimeIn = false;
    var skipTimeOutErrorMessage = false;

    //indexes that increment by three to point to the next name/date
    var nameSelectionIndex = 0;
    var dateSelectionIndex = 2;

    var rowEntryIndex = 2;
    var today = dateDMY();
    var isTimeInputOk = false;

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
                                doc.getCells(1, {
                                    'min-row': 2, 
                                    'max-row': maxRows, 
                                    'min-col': 1,
                                    'max-col': 3,
                                    'return-empty': false
                                }, 
                                function (err, cells) {
                                    //iterates through all username values
                                    while(cells[nameSelectionIndex]!=null){
                                        //checks if name is equal
                                        if(cells[nameSelectionIndex].value == response.members[i].name){
                                            //checks if there is a date
                                            if(cells[dateSelectionIndex]!=null){
                                                //checks if the date is the same date as today
                                                if(cells[dateSelectionIndex].value == today){
                                                    skipTimeIn = true;
                                                    bot.reply(message, 'Please type <user renew username/timestamp timestamp/username> to time in again');
                                                }
                                            }
                                        }
                                        //next name
                                        nameSelectionIndex += 3;
                                        //next date
                                        dateSelectionIndex += 3;
                                    }
                                    //if user hasn't timed in yet
                                    rowEntryIndex += (nameSelectionIndex/3);
                                    if(!skipTimeIn) {
                                        setCellValue(rowEntryIndex, 1, 1, 0, response.members[i].name);
                                        setCellValue(rowEntryIndex, 2, 2, 0, response.members[i].real_name);
                                        setCellValue(rowEntryIndex, 3, 3, 0, today);
                                        setCellValue(rowEntryIndex, 4, 4, 0, timeInput);
                                        bot.reply(message, '<@'+message.user+'>, you have timed in '+'<@'+userId+'>');
                                    }
                                });
                            });
                            break;
                        }
                    }
                });
            }
            else if(command == 'out'){
                bot.api.users.list({},function(err, response) {
                    for(var i = 0; i < response.members.length; i++){
                        if(response.members[i].id==userId){
                            doc.useServiceAccountAuth(creds, function (err) {
                                doc.getCells(1, {
                                    'min-row': 2, 
                                    'max-row': maxRows, 
                                    'min-col': 1,
                                    'max-col': 3,
                                    'return-empty': false
                                }, 
                                function (err, cells) {
                                    //iterates through all username values
                                    while(cells[nameSelectionIndex]!=null){
                                        //checks if name is equal
                                        if(cells[nameSelectionIndex].value == response.members[i].name){
                                            //checks if there is a date
                                            if(cells[dateSelectionIndex]!=null){
                                                //checks if the date is the same date as today
                                                if(cells[dateSelectionIndex].value == today){
                                                    skipTimeOutErrorMessage = true;
                                                    setCellValue(rowEntryIndex, 5, 5, 0, timeInput);
                                                    setCellValue(rowEntryIndex, 6, 6, 0, '=E'+rowEntryIndex+'-'+'D'+rowEntryIndex);
                                                    bot.reply(message, '<@'+message.user+'>, you have timed out '+'<@'+userId+'>');
                                                }
                                            }
                                        }
                                        //next name
                                        nameSelectionIndex += 3;
                                        //next date
                                        dateSelectionIndex += 3;
                                        rowEntryIndex ++;
                                    }
                                    //if user hasn't timed in yet
                                    if(!skipTimeOutErrorMessage) {
                                        bot.reply(message, '<@'+userId+'> has not yet timed in');
                                    }
                                });
                            });
                            break;
                        }
                    }
                });
            }
            else if(command=='renew'){
                bot.api.users.list({},function(err, response) {
                    for(var i = 0; i < response.members.length; i++){
                        if(response.members[i].id==userId){
                            doc.useServiceAccountAuth(creds, function (err) {
                                doc.getCells(1, {
                                    'min-row': 2, 
                                    'max-row': maxRows, 
                                    'min-col': 1,
                                    'max-col': 3,
                                    'return-empty': false
                                }, 
                                function (err, cells) {
                                    //iterates through all username values
                                    while(cells[nameSelectionIndex]!=null){
                                        //checks if name is equal
                                        if(cells[nameSelectionIndex].value == response.members[i].name){
                                            //checks if there is a date
                                            if(cells[dateSelectionIndex]!=null){
                                                //checks if the date is the same date as today
                                                if(cells[dateSelectionIndex].value == today){
                                                    skipTimeOutErrorMessage = true;
                                                    setCellValue(rowEntryIndex, 4, 4, 0, timeInput);
                                                    setCellValue(rowEntryIndex, 5, 5, 0, "");
                                                    setCellValue(rowEntryIndex, 6, 6, 0, "");
                                                    bot.reply(message, '<@'+message.user+'>, you have renewed '+'<@'+userId+'>\'s time in.');
                                                }
                                            }
                                        }
                                        //next name
                                        nameSelectionIndex += 3;
                                        //next date
                                        dateSelectionIndex += 3;
                                        rowEntryIndex ++;
                                    }
                                    //if user hasn't timed in yet
                                    if(!skipTimeOutErrorMessage) {
                                        bot.reply(message, '<@'+userId+'> has not yet timed in');
                                    }
                                });
                            });
                            break;
                        }
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
    var command = message.match[1];
    var field1 = message.match[2];
    var timeInput = moment().format('HH:mm:ss'); 
    var userId;

    var skipTimeIn = false;
    var skipTimeOutErrorMessage = false;

    //indexes that increment by three to point to the next name/date
    var nameSelectionIndex = 0;
    var dateSelectionIndex = 2;

    var rowEntryIndex = 2;
    var today = dateDMY();
    var isTimeInputOk = false;

    if(field1.substr(0,2)=='<@'){
        userId = field1.substr(2, field1.length-3);
    }
    
    if(command == 'in'){
        bot.api.users.list({},function(err, response) {
            for(var i = 0; i < response.members.length; i++){
                if(response.members[i].id==userId){
                    doc.useServiceAccountAuth(creds, function (err) {
                        doc.getCells(1, {
                            'min-row': 2, 
                            'max-row': maxRows, 
                            'min-col': 1,
                            'max-col': 3,
                            'return-empty': false
                        }, 
                        function (err, cells) {
                            //iterates through all username values
                            while(cells[nameSelectionIndex]!=null){
                                //checks if name is equal
                                if(cells[nameSelectionIndex].value == response.members[i].name){
                                    //checks if there is a date
                                    if(cells[dateSelectionIndex]!=null){
                                        //checks if the date is the same date as today
                                        if(cells[dateSelectionIndex].value == today){
                                            skipTimeIn = true;
                                            bot.reply(message, 'Please type <user renew username/timestamp timestamp/username> to time in again');
                                        }
                                    }
                                }
                                //next name
                                nameSelectionIndex += 3;
                                //next date
                                dateSelectionIndex += 3;
                            }
                            //if user hasn't timed in yet
                            rowEntryIndex += (nameSelectionIndex/3);
                            if(!skipTimeIn) {
                                setCellValue(rowEntryIndex, 1, 1, 0, response.members[i].name);
                                setCellValue(rowEntryIndex, 2, 2, 0, response.members[i].real_name);
                                setCellValue(rowEntryIndex, 3, 3, 0, today);
                                setCellValue(rowEntryIndex, 4, 4, 0, timeInput);
                                bot.reply(message, '<@'+message.user+'>, you have timed in '+'<@'+userId+'>');
                            }
                        });
                    });
                    break;
                }
            }
        });
    }
    else if(command == 'out'){
        bot.api.users.list({},function(err, response) {
            for(var i = 0; i < response.members.length; i++){
                if(response.members[i].id==userId){
                    doc.useServiceAccountAuth(creds, function (err) {
                        doc.getCells(1, {
                            'min-row': 2, 
                            'max-row': maxRows, 
                            'min-col': 1,
                            'max-col': 3,
                            'return-empty': false
                        }, 
                        function (err, cells) {
                            //iterates through all username values
                            while(cells[nameSelectionIndex]!=null){
                                //checks if name is equal
                                if(cells[nameSelectionIndex].value == response.members[i].name){
                                    //checks if there is a date
                                    if(cells[dateSelectionIndex]!=null){
                                        //checks if the date is the same date as today
                                        if(cells[dateSelectionIndex].value == today){
                                            skipTimeOutErrorMessage = true;
                                            setCellValue(rowEntryIndex, 5, 5, 0, timeInput);
                                            setCellValue(rowEntryIndex, 6, 6, 0, '=E'+rowEntryIndex+'-'+'D'+rowEntryIndex);
                                            bot.reply(message, '<@'+message.user+'>, you have timed out '+'<@'+userId+'>');
                                        }
                                    }
                                }
                                //next name
                                nameSelectionIndex += 3;
                                //next date
                                dateSelectionIndex += 3;
                                rowEntryIndex ++;
                            }
                            //if user hasn't timed in yet
                            if(!skipTimeOutErrorMessage) {
                                bot.reply(message, '<@'+userId+'> has not yet timed in');
                            }
                        });
                    });
                    break;
                }
            }
        });
    }
    else if(command=='renew'){
        bot.api.users.list({},function(err, response) {
            for(var i = 0; i < response.members.length; i++){
                if(response.members[i].id==userId){
                    doc.useServiceAccountAuth(creds, function (err) {
                        doc.getCells(1, {
                            'min-row': 2, 
                            'max-row': maxRows, 
                            'min-col': 1,
                            'max-col': 3,
                            'return-empty': false
                        }, 
                        function (err, cells) {
                            //iterates through all username values
                            while(cells[nameSelectionIndex]!=null){
                                //checks if name is equal
                                if(cells[nameSelectionIndex].value == response.members[i].name){
                                    //checks if there is a date
                                    if(cells[dateSelectionIndex]!=null){
                                        //checks if the date is the same date as today
                                        if(cells[dateSelectionIndex].value == today){
                                            skipTimeOutErrorMessage = true;
                                            setCellValue(rowEntryIndex, 4, 4, 0, timeInput);
                                            setCellValue(rowEntryIndex, 5, 5, 0, "");
                                            setCellValue(rowEntryIndex, 6, 6, 0, "");
                                            bot.reply(message, '<@'+message.user+'>, you have renewed '+'<@'+userId+'>\'s time in.');
                                        }
                                    }
                                }
                                //next name
                                nameSelectionIndex += 3;
                                //next date
                                dateSelectionIndex += 3;
                                rowEntryIndex ++;
                            }
                            //if user hasn't timed in yet
                            if(!skipTimeOutErrorMessage) {
                                bot.reply(message, '<@'+userId+'> has not yet timed in');
                            }
                        });
                    });
                    break;
                }
            }
        });
    }
    else{
        bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands");
    }
 })
controller.hears(['(.*) (.*)'], 'direct_message,direct_mention,mention', function(bot,message) {
    var command = message.match[1];
    var timestamp = message.match[2];
    var renewMsg = 'Please type <renew> <timestamp> to time in again at a particular time since you have already timed in';
    if(moment(timestamp, 'HH:mm:ss')) {
        if(command=='in'){
            timeIn(bot, message, renewMsg, timestamp);
        }
        else if(command=='renew'){
            renew(bot, message, timestamp);
        }
        else if(command=='out'){
            timeOut(bot, message, timestamp);
        }
        else{
            bot.reply(message, "I don\'t understand the command. Please type @bundy help for a list of all the commands");
        }
    }
    else{
        bot.reply(message, "Please follow the time format (HH:MM:SS) 24-hours.");
    }
 })
controller.hears(['hello','hi'],['direct_message','direct_mention','mention'],function(bot,message) {
    bot.reply(message,"Hello.");
});


controller.on('slash_command',function(bot,message) {
  // reply to slash command
  bot.replyPublic(message,'Everyone can see the results of this slash command');
});
