var clientID = '2a11ee6e18fe46ad89f9dcbd5507b76b';
var secret = 'cc88177b38774df2acfedd668be53824';
var code;
var selectedProjectID;

var markCompletedUUID;
var markCompletedItemID;

function search(nameKey, myArray){
    for (var i=0; i < myArray.length; i++) {
        if (myArray[i].id === nameKey) {
            return myArray[i].sortnum;
        }
    }
}

function hasduedate(id, myArray){
    for (var i=0; i < myArray.length; i++) {
        if (myArray[i].id === id) {
            if (myArray[i].due != null) 
              return true;
            else
              return false;
            
        }
    }
}

function createUUID() {
    // http://www.ietf.org/rfc/rfc4122.txt
    var s = [];
    var hexDigits = "0123456789abcdef";
    for (var i = 0; i < 36; i++) {
        s[i] = hexDigits.substr(Math.floor(Math.random() * 0x10), 1);
    }
    s[14] = "4";  // bits 12-15 of the time_hi_and_version field to 0010
    s[19] = hexDigits.substr((s[19] & 0x3) | 0x8, 1);  // bits 6-7 of the clock_seq_hi_and_reserved to 01
    s[8] = s[13] = s[18] = s[23] = "-";

    var uuid = s.join("");
    return uuid;
}

var xhrRequest = function (url, type, callback) {
  var xhr = new XMLHttpRequest();
  xhr.onload = function () {
      if(this.status == 200)
          callback(this.responseText);
      else
          sendErrorString("Error:Status " + this.status);
  };
  xhr.open(type, url);
  xhr.send();
};

function getWatchVersion()
{
    var platform;
    if(Pebble.getActiveWatchInfo) {
      var watchinfo= Pebble.getActiveWatchInfo();
      platform=watchinfo.platform;
      } else {
        platform="aplite";
      }
    return platform;
}

function isTimelineEnabled()
{
    if (localStorage.getItem("timelineEnabled") == "true")
        return true;
    else
        return false;
}

//if integer is 0 return a string 00. Used for times
function leadingZeroCheck(number)
{
    if (number < 10)
        return "0" + number;
    else
        return number;
}

function removeOutlookGarbage(str)
{
    if (startsWith(str,"[[outlook=id"))
    {
        var contentStart = str.indexOf(",") + 2;
        var contentEnd = str.indexOf("]]");
        return str.substring(contentStart, contentEnd);
    }
    else
    {
        return str;
    }
}

//defining my own startsWith as the built in one does not seem to work
//str = string to search
//strMatch = string to match
function startsWith(str, strMatch)
{
    for(var i = 0; i < str.length; i++)
    {
        if (str.substring(0, i) == strMatch)
            return true;
    }
    return false;
}

function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

function getItems(responseText)
{
        
        // responseText contains a JSON object with item info
        var json = JSON.parse(responseText);
        json = json.items;
    
        
       
        
    
        var isToday = 0;
        //if selected project ID is 0 then it means "today" was selected
        if (selectedProjectID === 0)
        {
            isToday = 1;
        }
    
        //sort the list based on the item order property, if today, sort by date
        //Changed to date added
        if (isToday)
        {
             json.sort(function(a, b) {
                 var d1 = new Date(a.date_added);
                 var d2 = new Date(b.date_added);
                 if (d1 > d2)
                     return 1;
                 else if (d1 == d2)
                     return 0;
                 else
                     return -1;
            });   
        }
        else
  
        for(var i=0;i<json.length;i++)
        {
	        if (json[i].parent_id == null)
	        {
		        json[i].sortnum = json[i].child_order * 1000;
	        }
	        else
	        continue;
        }
  
        for(var i=0;i<json.length;i++)
        {
        	if (json[i].parent_id != null)
        	{ 
        		json[i].sortnum = search(json[i].parent_id,json) + json[i].child_order;
        	}
        	else
        		continue;
        }
  
          
        {
            json.sort(function(a, b) {
                return a.sortnum  - b.sortnum;
            });
         
        }
            
        if (json[0])
        {
            if (!json[0].hasOwnProperty("id"))
            {
                sendErrorMessage(3);
                return;
            }
        }
    
        var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        
    
        // Conditions
        var itemNames = "";
        var itemIDs = "";
        var itemDates = "";
        var itemDueDates = "";
        var itemIndentation = "";
        
    
        //only put "Add New" if we are on basalt or chalk
        if (getWatchVersion() == "basalt" || getWatchVersion() == "chalk" || getWatchVersion() == "diorite" || getWatchVersion() == "emery")
        {
            //cannot add items to "Today" project so we only add the "Add New" button if it's not a today project
            if (!isToday)
            {
                itemNames = itemNames + "+ Add New |";
                itemIDs = itemIDs + "0|";
                itemDates = itemDates + "|";
                itemDueDates = itemDueDates + "|";
                itemIndentation = itemIndentation + "1|";   
            }
        }
    
        for(var i=0;i<json.length;i++)
        {
            
            if (isToday)
            {
                var today = new Date();
                today.setHours(0);
                today.setMinutes(0);
                today.setSeconds(0);
                //considered "Today" if due date is in the current day or less (overdue)
                today = addDays(today, 1);

                if ((json[i].due) == null)
                    continue;
                var d = new Date(json[i].due.date);
                if (d >= today)
                {
                    continue;
                }
            }
            else
            {
                //only proccess items that are in the selected project ID
                if (json[i].project_id != selectedProjectID)
                {
                    continue;
                }
            }
            //xhrRequest('https://hookb.in/Ew985m0g?stuff=' + JSON.stringify(json[i]), 'POST', deadFunction);
            
            
            //items added via outlook have an ID tag in their content and some really weird syntax. The below is to fix this and show it as a normal item
            json[i].content = removeOutlookGarbage(json[i].content);
            
            itemNames = itemNames + json[i].content.replace("|", "") + " |";
            itemIDs = itemIDs  + json[i].id + "|";
			if (json[i].due == null)
                itemDates = itemDates + "|";
            else
                itemDates = itemDates + json[i].due.string + "|";
      if (json[i].parent_id == null)
                itemIndentation = itemIndentation +  "0|";
            else
              itemIndentation = itemIndentation + "1|";
            if (json[i].due == null)
            {
                itemDueDates = itemDueDates + "|"; 
            }
            else
            {
                var d = new Date(json[i].due.date);
              //if the time is 23:59 this specifies "no time" //change to 00:00
                if ((d.getHours() == 00) && (d.getMinutes() == 00))
                    itemDueDates = itemDueDates + monthNames[d.getMonth()] + " " + d.getDate() + "|";
                else
                    itemDueDates = itemDueDates + monthNames[d.getMonth()] + " " + d.getDate() + " " + d.getHours() + ":" + leadingZeroCheck(d.getMinutes())  + "|";  
            }
        }
    

    
        var dictionary = 
        {
            "ITEM_NAMES": itemNames,
            "ITEM_IDS": itemIDs,
            "ITEM_DATES": itemDates,
            "ITEM_DUE_DATES": itemDueDates,
            "ITEM_INDENTATION": itemIndentation
        };
    
    
        // Send to Pebble
        Pebble.sendAppMessage(dictionary,
                              function(e) 
                              {
                                  
                              },
                              function(e) 
                              {
                                  sendErrorString(e.error.message);
                              });  
}

function getAllItemsForTimeline(responseText)
{
    try
    {
        // responseText contains a JSON object with item info
        var json = JSON.parse(responseText);
        json = json.items;
    
    //some testing code
    /*var json = [];
    for(var z=0;z<json.length;z++)
    {
            console.log("item " + z + "\n");
            insertUserPin(json[z], function(responseText) 
            { 
                    
            });
    }*/
    
        for(var i=0;i<json.length;i++)
        {
            if (json[i].due != null)
            {
                
                var date = new Date(json[i].due.date);
                var minDate = new Date();
                //min date is two days ago. Have to use this weird calculation because javascript is stupid
                minDate.setTime( minDate.getTime() - 2 * 86400000 );
                var maxDate = new Date();
                //max date is 1 year from now.
                maxDate.setTime( maxDate.getTime() + 364 * 86400000 );
                //don't want to enter in entries from a long time ago, or too far in the future. Timeline will reject anything outside the range of 2 days
                //in the past and 1 year in the future
                if (date < minDate)
                    continue;
                if (date > maxDate)
                    continue;
                
                //items added via outlook have an ID tag in their content and some really weird syntax. The below is to fix this and show it as a normal item
                json[i].content = removeOutlookGarbage(json[i].content);
                //truncate string if its longer than 256 characters as this is not allowed in timeline
                if(json[i].content.length > 256)
                {
                    json[i].content = json[i].content.substring(0, 255);
                }
                
                
                //Create timeline array of items that have due dates
                var pin = {
                "id": "TodoistMiniItem-" + json[i].id,
                "time": date.toISOString(),
                "layout": {
                  "type": "genericPin",
                  "title": json[i].content,
                  "tinyIcon": "system://images/SCHEDULED_EVENT"
                  }
                };
                
                //TEMP make sure to comment out in live versions
                //xhrRequestPost('http://www.bradpaugh.com/index.html', 'POST', JSON.stringify(pin));
                
                //test pin keep commented other than for testing
                //pin =  {"id":"TodoistMiniItem-21903246","time":"2017-03-01T22:59:59.000Z","layout":{"type":"genericPin","title":"Termin Confort d'eau - Wartung Boiler 20.3.","tinyIcon":"system://images/SCHEDULED_EVENT"}};
                //console.log("writing timeline" + JSON.stringify(pin));
                insertUserPin(pin, function(responseText) 
                { 
                    
                });
            }
        }
    
    //send message to watch to tell it that loading is complete
    var dictionary = 
        {
            "TIMELINE_COMPLETE": "0"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
    }
    catch (err)
    {
        sendErrorString(err.message);
    }
}

function getToken(responseText) 
{
    // responseText contains a JSON object with token info
    var json = JSON.parse(responseText);
    if (responseText == "\"LOGIN_ERROR\"")
    {
        sendErrorMessage(1);
        openConfig();
        return;
    }
    // Conditions
    var token = json.access_token;
    localStorage.setItem("todoistMiniTokenV8", token);
    getProjectsFromToken();
}

function getProjects(responseText)
{
    try
    {
        // responseText contains a JSON object with project data
        var json = JSON.parse(responseText);
        json = json.projects;
    
        if (json[0])
        {
            if (!json[0].hasOwnProperty("id"))
            {
                sendErrorMessage(2);
                return;
            }
        }
        // Conditions
        var projectNames = "";
        var projectIDs = "";
        var projectIndentation = "";
        
        //put today project in (custom)
        projectNames = "Today |";
        projectIDs = "0|";
        projectIndentation = "1|";
        
        //sort the list based on the item order property
        json.sort(function(a, b) {
            return parseInt(a.item_order) - parseInt(b.item_order);
        });
        
        for(var i=0;i<json.length;i++)
        {
            projectNames = projectNames + json[i].name.replace("|", "")  + " |";
            projectIDs = projectIDs  + json[i].id + "|";
            projectIndentation = projectIndentation + json[i].indent + "|";
        }
    
        var dictionary = 
        {
            "PROJECT_NAMES": projectNames,
            "PROJECT_IDs": projectIDs,
            "PROJECT_INDENTATION": projectIndentation
        };
    
        // Send to Pebble
        Pebble.sendAppMessage(dictionary,
                              function(e) 
                              {
                                  if ((Pebble.getActiveWatchInfo().firmware.major >= 3) && isTimelineEnabled())
                                      sendWaitingMessageAndPerformAction(5);
                              },
                              function(e) 
                              {
                                  sendErrorString(e.error.message);
                                  //xhrRequestPost('http://www.bradpaugh.com/index.html', 'POST', 'JAVASCIRPT ERROR MESSAGE: ' + e.error.message);
                              });   
    }
    catch (err)
    {
        sendErrorString(err.message);
    }
}

function markItem(responseText)
{
    if (responseText.search("ok") > 0)
    {
        var dictionary = 
        {
            "SELECTED_ITEM": "1"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
    }
    else
    {
        var dictionary = 
        {
            "SELECTED_ITEM": "0"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
    }
    
}

function markRecurringItem(responseText)
{
    if (responseText.search("ok") > 0)
    {
        var dictionary = 
        {
            "SELECTED_ITEM": "1"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
    }
    else
    {
        var dictionary = 
        {
            "SELECTED_ITEM": "0"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
    }
    
}

function uncompleteItem(responseText)
{
    if (responseText.search("ok") > 0)
    {
        var dictionary = 
        {
            "SELECTED_ITEM_UNCOMPLETE": "1"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
    }
    else
    {
        var dictionary = 
        {
            "SELECTED_ITEM_UNCOMPLETE": "0"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
    }
    
}

function addItem(responseText)
{
    //add response error handling here
        var dictionary = 
        {
            "ADD_NEW_ITEM": "1"
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });
}

//code 1 = waiting for config
//code 2 = waiting to load data
function sendWaitingMessageAndPerformAction(code)
{
    try
    {
        //when we send app message it just needs to be a 1 or 2 (config or loading) 3 = timeline loading
        var sendCode;
        if ((code == 3) || (code == 4))
            sendCode = 2;
        else if (code == 5)
            sendCode = 3;
        else
            sendCode = code;
        
     
        var dictionary = 
            {
                "WAITING": sendCode
            };
            Pebble.sendAppMessage(dictionary,
                              function(e) 
                              {
                                  if (code == 1)
                                  {
                                       openConfig();   
                                  }
                                  if (code == 2)
                                  {
                                       getProjectsFromToken(); 
                                  }
                                  if (code == 3)
                                  {
                                       processTodoistDataWithGoogle();
                                  }
                                  if (code == 4)
                                  {
                                       processTodoistData();  
                                  }
                                  if (code == 5)
                                  {
                                       pinTimelineItems();
                                  }
                              },
                              function(e) 
                              {
                                  sendErrorString(e.error.message);
                                  //xhrRequestPost('http://www.bradpaugh.com/index.html', 'POST', 'JAVASCRIPT ERROR MESSAGE sendWaitingMessage: ' + e.error.message);
                              });
    }
    catch (err)
    {
        sendErrorString(err.message);
    }
      
}

//code 1 = Login failed
function sendErrorMessage(code)
{
    var dictionary = 
        {
            "ERROR": code
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
}

function sendErrorString(errorMsg)
{
    var dictionary = 
        {
            "ERRORMSG": errorMsg
        };
        Pebble.sendAppMessage(dictionary,
                          function(e) 
                          {
                              
                          },
                          function(e) 
                          {
                              sendErrorString(e.error.message);
                          });   
}

function processTodoistData() 
{
    var url = "https://todoist.com/oauth/access_token?client_id=" + encodeURIComponent(clientID) + "&client_secret=" + encodeURIComponent(secret) + "&code=" + encodeURIComponent(code);
    //localStorage.removeItem("todoistEmail");
    //localStorage.removeItem("todoistPassword");
    //note that xhr request is ASYNCHRONOUS everything after it in this function will get executed
    //before it is even finished the next path of execution HAS to be in the callback function
    xhrRequest(url, 'POST', getToken);
    
    //getProjectsFromToken();
}

function processTodoistDataWithGoogle()
{
    var url = "https://todoist.com/API/loginWithGoogle?email=" +
    encodeURIComponent(localStorage.getItem("googleEmail")) + "&oauth2_token=" + encodeURIComponent(localStorage.getItem("googleToken"));
    //note that xhr request is ASYNCHRONOUS everything after it in this function will get executed
    //before it is even finished the next path of execution HAS to be in the callback function
    xhrRequest(url, 'GET', getToken);
}

function getProjectsFromToken()
{
    var url = "https://api.todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&sync_token=" + encodeURIComponent("'*'") + "&resource_types=" + encodeURIComponent("[\"projects\"]");
    xhrRequest(url, 'GET', getProjects);
}

function getItemsForSelectedProject(projectID)
{
    selectedProjectID = projectID;
    var url = "https://api.todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&sync_token=" + encodeURIComponent("'*'") + "&resource_types=" + encodeURIComponent("[\"items\"]");
    xhrRequest(url, 'GET', getItems);
}

function getItemsForToday()
{
    selectedProjectID = 0;
    var url = "https://api.todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&sync_token=" + encodeURIComponent("'*'") + "&resource_types=" + encodeURIComponent("[\"items\"]");
    xhrRequest(url, 'GET', getItems);
}

function pinTimelineItems()
{
    //var url = "https://api.todoist.com/API/query?queries=" + encodeURIComponent("[\"No Due Date\"]") + "&token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8"));
    var url = "https://api.todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&sync_token=" + encodeURIComponent("'*'") + "&resource_types=" + encodeURIComponent("[\"items\"]");
    xhrRequest(url, 'GET', getAllItemsForTimeline);
}

function addNewItem(itemText, projectID)
{
    var commandsjson = [{
        "type": "item_add", 
        "temp_id": createUUID(), 
        "uuid": createUUID(), 
        "args": 
        {
            "content": itemText, 
            "project_id": projectID
        }
    }];
    var url = "https://todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&commands=" + encodeURIComponent(JSON.stringify(commandsjson));
    xhrRequest(url, 'GET', addItem);
}

function markItemAsCompleted(itemID, duedate)
{
    markCompletedUUID = createUUID();
    markCompletedItemID = itemID;
    var commandsjson = [{
        "type": "item_complete", 
        "uuid": createUUID(), 
        "args": 
        {
            "id": itemID
        }
    }];
    
    var url = "https://todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&commands=" + encodeURIComponent(JSON.stringify(commandsjson));
    var pin = {
                "id": "TodoistMiniItem-" + itemID
              };
    if ((Pebble.getActiveWatchInfo().firmware.major >= 3) && isTimelineEnabled() && (duedate != ""))
    {
        //delete pin from timeline if completed
        deleteUserPin(pin, function(responseText) 
        {
            
        });
    }
    
    xhrRequest(url, 'GET', markItem);
}

function markRecurringItemAsCompleted(itemID)
{
    markCompletedUUID = createUUID();
    markCompletedItemID = itemID;
    var commandsjson = [{
        "type": "item_update_date_complete", 
        "uuid": createUUID(), 
        "args": 
        {
            "id": itemID
        }
    }];
    
    var url = "https://todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&commands=" + encodeURIComponent(JSON.stringify(commandsjson));
    
    xhrRequest(url, 'GET', markRecurringItem);
    if ((Pebble.getActiveWatchInfo().firmware.major >= 3) && isTimelineEnabled() && (duedate != ""))
    {
        //update timeline pins if recurring item completed
        pinTimelineItems();
    }
}

function markItemAsUncompleted(itemID)
{
    markCompletedUUID = createUUID();
    markCompletedItemID = itemID;
    var commandsjson = [{
        "type": "item_uncomplete", 
        "uuid": createUUID(), 
        "args": 
        {
            "id": itemID
        }
    }];
    
    var url = "https://todoist.com/API/v8/sync?token=" + encodeURIComponent(localStorage.getItem("todoistMiniTokenV8")) + "&commands=" + encodeURIComponent(JSON.stringify(commandsjson));
    
    xhrRequest(url, 'GET', uncompleteItem);
    
    if ((Pebble.getActiveWatchInfo().firmware.major >= 3) && isTimelineEnabled() && (duedate != ""))
    {
        //update timeline pins if item uncompleted
        pinTimelineItems();
    }
}


// Listen for when the watchface is opened
Pebble.addEventListener('ready', startup);

//temp settings
function startup()
{
    //localStorage.removeItem("todoistMiniTokenV8");
    //localStorage.setItem("todoistMiniTokenV8", "tokenhere");
    //enables timeline by default if it has never been set.
    if (localStorage.getItem("timelineEnabled") === null)
        localStorage.setItem("timelineEnabled", "false");

    if (localStorage.getItem("todoistMiniTokenV8") === null)
    {
        sendWaitingMessageAndPerformAction(1);
    }
    else
    {
        sendWaitingMessageAndPerformAction(2);
    }
}

// Listen for when an AppMessage is received
Pebble.addEventListener('appmessage',
  function(e) {
    if(e.payload.SELECTED_PROJECT)
    {
        if (e.payload.SELECTED_PROJECT == "0")
            getItemsForToday();
        else
            getItemsForSelectedProject(e.payload.SELECTED_PROJECT);
    }
    if(e.payload.SELECTED_ITEM)
    {
        markItemAsCompleted(e.payload.SELECTED_ITEM, e.payload.ITEM_DATES);
    }
    if(e.payload.SELECTED_ITEM_RECURRING)
    {
        markRecurringItemAsCompleted(e.payload.SELECTED_ITEM_RECURRING, e.payload.ITEM_DATES);
    }
    if(e.payload.SELECTED_ITEM_UNCOMPLETE)
    {
        markItemAsUncompleted(e.payload.SELECTED_ITEM_UNCOMPLETE, e.payload.ITEM_DATES);
    }
    if(e.payload.ADD_NEW_ITEM)
    {
        var array = e.payload.ADD_NEW_ITEM.split("|");
        addNewItem(array[0], array[1]);
    }
  }                     
);

//sets the configuration options from the config page that the user has just saved.
function setConfig(loginData)
{
    try
    {
        localStorage.setItem("ConfigData", JSON.stringify(loginData));
        var configString = loginData.scrollSpeed + '|' + loginData.backgroundColor + '|' + loginData.foregroundColor + '|' + loginData.altBackgroundColor + '|' + loginData.altForegroundColor + '|' + loginData.highlightBackgroundColor + '|' + loginData.highlightForegroundColor + '|' + loginData.timelineEnabled + '|';
        localStorage.setItem("timelineEnabled", loginData.timelineEnabled);
        var dictionary = 
        {
            "CONFIG": configString
        };
    
        // Send to Pebble
        Pebble.sendAppMessage(dictionary,
                              function(e) 
                              {
                                  
                              },
                              function(e) 
                              {
                                  sendErrorString(e.error.message);  
                              }); 
    }
    catch (err)
    {
        sendErrorString(err.message);
    }
}

function openConfig(e) 
{
    if (localStorage.getItem("ConfigData") === null)
    {
        try {xhr.setRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36");} catch(e){}
        Pebble.openURL('https://perogy.github.io/PebbleProject/indexNew.html');
    }
    else
    {
        var configData = JSON.parse(localStorage.getItem("ConfigData"));
        try {xhr.setRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36");} catch(e){}
        Pebble.openURL('https://perogy.github.io/PebbleProject/indexNew.html#' + 'scrollSpeed=' + configData.scrollSpeed + '&backgroundColor=' + configData.backgroundColor + '&foregroundColor=' + configData.foregroundColor + '&altBackgroundColor=' + 
                                                            configData.altBackgroundColor + '&altForegroundColor=' + configData.altForegroundColor + '&highlightBackgroundColor=' + configData.highlightBackgroundColor + '&highlightForegroundColor=' + configData.highlightForegroundColor + '&timelineEnabled=' + configData.timelineEnabled);
        
    }
}

//

function closeConfig(e) {
    try
    {
        //if they pressed back on the settings screen (no save or login), just run the startup function
        if (typeof(e.response) == "undefined")
        {
            startup();
            return;
        }
        var loginData = JSON.parse(decodeURIComponent(e.response));
        
        
        if (loginData.type == "configData")
        {
            setConfig(loginData);
            return;
        }
        
        if (loginData.googleToken)
        {
            //check whether google or normal login and then run appropriate code
            localStorage.setItem("googleToken", loginData.token);
            localStorage.setItem("googleEmail", loginData.email);
            sendWaitingMessageAndPerformAction(3);
            
        }
        else
        {
            code = loginData.code;
            sendWaitingMessageAndPerformAction(4);
        }
    }
    catch (err)
    {
        sendErrorString(err.message);
    }

}
// Listen for a configuration request
Pebble.addEventListener('showConfiguration', openConfig);

//Listen for configuration window closing
Pebble.addEventListener('webviewclosed', closeConfig);






/******************************* timeline lib *********************************/

// The timeline public URL root
var API_URL_ROOT = 'https://timeline-api.rebble.io/';

/**
 * Send a request to the Pebble public web timeline API.
 * @param pin The JSON pin to insert. Must contain 'id' field.
 * @param type The type of request, either PUT or DELETE.
 * @param callback The callback to receive the responseText after the request has completed.
 */
function timelineRequest(pin, type, callback) {
  // User or shared?
  var url = API_URL_ROOT + 'v1/user/pins/' + pin.id;

  // Create XHR
  var xhr = new XMLHttpRequest();
  xhr.onload = function () {
    if(this.status == 200) {
          callback(this.responseText);
    }
    else if(this.status == 429) {
        //this means we received an exceeded rate limit from the server, just return when this happens
        return;
    }
    else
    {
         sendErrorString("Timeline Error:Status " + this.status + " " + this.responseText);     
    }    
  };
  xhr.open(type, url);

  // Get token
  Pebble.getTimelineToken(function(token) {
    // Add headers
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.setRequestHeader('X-User-Token', '' + token);

    // Send
    xhr.send(JSON.stringify(pin));
  }, function(error) { sendErrorString("Timeline Error: " + error); });
}


/**
 * Insert a pin into the timeline for this user.
 * @param pin The JSON pin to insert.
 * @param callback The callback to receive the responseText after the request has completed.
 */
function insertUserPin(pin, callback) {
  timelineRequest(pin, 'PUT', callback);
}

/**
 * Delete a pin from the timeline for this user.
 * @param pin The JSON pin to delete.
 * @param callback The callback to receive the responseText after the request has completed.
 */
function deleteUserPin(pin, callback) {
  timelineRequest(pin, 'DELETE', callback);
}

/***************************** end timeline lib *******************************/


