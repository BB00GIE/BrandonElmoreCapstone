//RegEx formulas add to this when looking for other links
let time = /(When: )[a-zA-Z0-9\–\ \,:]*/
let google = /https?:\/\/(meet)[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,4}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/
let zoom = /(http(s)?:\/\/)?(us05web)[-a-zA-Z0-9@:%._\+~#=]{2,256}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)/
let zoomTime = /(Time:\ )[a-zA-Z0-9:\ (),]*/
//Google Sheet Link
let sheet_id = "https://docs.google.com/spreadsheets/d/1RlhGTQgOkR0SQRxI2_0iNYM_zy3ZmeYJDja15rtIFPk/edit#gid=0";


function main(){
  var list = retrieveMessages();
  saveLinks(list);
  console.log(list);
  makeAppointments(list);
}

function retrieveMessages() {
  var links = []
  //console.log(GmailApp.getInboxUnreadCount())
  var threads = GmailApp.getStarredThreads(0,20); 
  for (var i = 0; i< threads.length; i++){
    var messages = threads[i].getMessages();
    for (var x = 0; x< messages.length;x++) {
      var body = messages[x].getPlainBody();
      var id = messages[x].getId();
      var sender = messages[x].getFrom();
      //console.log(body)
      var hi = body.match(google);
      //console.log(body);
      if (hi != null){
        links.push([sender,hi[0], id, "google"]);
      }
      var bye = body.match(zoom)
      if (bye != null){
        links.push([sender,bye[0], id, "Zoom"])
      }
    }
  }
    
  //console.log(links);
  getDates(links);
  console.log(links)
  return links;
}
    


function saveLinks(links){

  //console.log(links)
  sheet = SpreadsheetApp.openByUrl(sheet_id)
  if (sheet == null){
    sheet = SpreadsheetApp.create("Capstone Sheet").getSheetId();
    console.log("ID is right here" + sheet)
  }else{
    //Do nothing
  }
  
  

  for(var i = 0; i < links.length; i ++){
    var ignore = false;
    for (var j = 0; j< links.length; j++){
      var id_cell = sheet.getRange("C" + (j+1))
      if (links[i][2] == id_cell.getValue()){
        ignore = true
        
      }
    }
    if (ignore){
      //skip me
      links[i].push(1);
    }
    else{
    var cell = sheet.getRange("A" + (i+1));
   cell.setValue(links[i][0]);
   cell = sheet.getRange("B" + (i+1))
   cell.setValue(links[i][1])
   cell = sheet.getRange("C" + (i+1))
   cell.setValue(links[i][2])
   cell = sheet.getRange("D" + (i+1))
   cell.setValue(links[i][3])
   cell = sheet.getRange("E" + (i+1))
   cell.setValue(links[i][4])
   links[i].push(0)
   }
   
  }
}


function makeAppointments(linksList){
  var exists = false;
  var calendars = CalendarApp.getAllCalendars();
  for (var i = 0; i < calendars.length; i ++){
    if (calendars[i].getName() != "CapStone"){
      //Do Nohing
    }else{
      var id = calendars[i].getId();
      exists = true;
    }
  }

  if (!exists){
    var activeCalendar = CalendarApp.createCalendar("CapStone");
  }else{
    var activeCalendar = CalendarApp.getOwnedCalendarById(id);
    }

    for(var i = 0; i < linksList.length; i++){
  var sender = linksList[i][0];
  var startTime = linksList[i][4];
  var endTime = linksList[i][5];
  var meetingLink = linksList[i][1];
  if (linksList[i][6] == 1){
    //do nothing
  }else{
    var event = activeCalendar.createEvent(sender, startTime, endTime);
  event.setDescription("Link: " + meetingLink);


  }
  
}
  

}



function getDates(linksList){

  var defaultDate = new Date("Jan 1, 2022 12:00:00 PM EST");
  var defaultDate2 = new Date("Jan 1, 2022 12:15:00 PM EST");


  for(var i = 0; i < linksList.length; i++){
    var message = GmailApp.getMessageById(linksList[i][2]);
    //console.log(message.getPlainBody().match(time))
    var dates = message.getPlainBody().match(time)
    var zoomDates = message.getPlainBody().match(zoomTime)
    
    //console.log("The Dates are here for message " + (i+1))
    if (dates != null){
      console.log("The Dates are here for message " + (i+1))
      var dates = dates[0].slice(10)
      var dates = dates.split(" ");
      //console.log("testing")
      console.log(dates)
      var extractedDate = new Date(dates[0] + " " + dates[1] + " " + dates[2] + " " + convertTimes(dates[3]))
      var extractedDate2 = new Date(dates[0] + " " + dates[1] + " " + dates[2] + " " + convertTimes(dates[5]))
      console.log(extractedDate)
      console.log(extractedDate2)

      linksList[i].push(extractedDate)
      linksList[i].push(extractedDate2)
    }else if(zoomDates != null){
      zoomDates = zoomDates[0]
      console.log(zoomDates)
      //zoomDates = zoomDates[0].slice(6)
      console.log(zoomDates)
      zoomDates = zoomDates.slice(6,27)
      var extractedDate = new Date(zoomDates)
      var extractedDate2 = new Date(extractedDate.getTime()+ (60*60*1000))
      
      linksList[i].push(extractedDate)
      linksList[i].push(extractedDate2)
    }else{
      linksList[i].push(defaultDate)
      linksList[i].push(defaultDate2)
      
      //Give them the default date for now
    }
  }
}



/*QOL function, converts times in email to js date time format
*/
function convertTimes(time){
  var noon = false
  var minutes = "00";
  if(time.includes("pm")){
    noon = true
  }
  

  if (time.includes(":")){
    newTime = time.split(":");
    minutes = newTime[1].slice(0,2)
    var out = newTime[0]
    //console.log(minutes)
  }else if (time.length > 3){
    var out = time.slice(0,2);
  }else{
    var out = time.slice(0,1)
  }

  out = out + ":" + minutes + ":" + "00"

  if (noon){
    out += " PM"
  }else{
    out+= " AM"
  }
  

  return out;
  //console.log(out)
}



