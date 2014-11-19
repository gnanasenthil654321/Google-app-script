function ret_activesheet(id,sheet_index)
{
  // will return the active sheet whose index is sheet_index
  // id = id of the spreadsheet
  // sheet_index = 0 for firstsheet , 1 for second sheet etc
  var get_spreadsheet = SpreadsheetApp.openById(id) // opens the spreadsheet COL_remainder
  SpreadsheetApp.setActiveSpreadsheet(get_spreadsheet) // set that spreadsheet as the active spreadsheet
  
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet() // active_spreadsheet variable
  
  var sheets = active_spreadsheet.getSheets() // gets all the sheets in current active spreadsheet
  
  SpreadsheetApp.setActiveSheet(sheets[sheet_index]) // the first sheet of the current active spreadsheet is set active
  var active_sheet = SpreadsheetApp.getActiveSheet() // active_sheet variable
  return active_sheet
}





function test()
{
  sort_leave()
}



function sort_leave() 
{
  // sorts the worked days in asceding order of year , then month and day
  
  var get_spreadsheet = SpreadsheetApp.openById("<desired spreadsheet name>") // opens the spreadsheet COL_remainder
  SpreadsheetApp.setActiveSpreadsheet(get_spreadsheet) // set that spreadsheet as the active spreadsheet
  
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet() // active_spreadsheet variable
  
  var sheets = active_spreadsheet.getSheets() // gets all the sheets in current active spreadsheet
  
  SpreadsheetApp.setActiveSheet(sheets[0]) // the first sheet of the current active spreadsheet is set active
  var active_sheet = SpreadsheetApp.getActiveSheet() // active_sheet variable
  
  var last_row = active_sheet.getLastRow() // gets the last row with data in the current active sheet.
  var start_row = 2
  
  var range_to_sort = active_sheet.getRange(start_row, 1, (last_row -1), 6) // the range to be sorted 
  range_to_sort.sort([{column:4,ascending:true},{column:5,ascending:true},{column:6,ascending:true}]) // sorts the range first by column1,then2 and then by3
  Logger.log("baca")
    
}


function expiry(year,month,day)
{
 // calculates when a worked day expires
 // date of expiry approx 6 months
 
 // creation of date of expiry
 var month_expiry = month + 6 // calculates the month of expiry
 
 
 if (month_expiry > 12) // if on adding 6, it the month_expiry goes to the next year, 
 {
   month_expiry = (month_expiry % 12)
   var year_expiry = year + 1
 }  
 if (month_expiry <= 12) // if the expiry month falls in the same year
 {
   var year_expiry = year
 }
 
 var day_expiry = day - 1
 
 var for_return = new Array(year_expiry,month_expiry,day_expiry) // the array that will be returned by this function
 
 if (day_expiry > 0) 
 {
  return for_return 
 }
  
 if (day_expiry == 0)
 {
  if (month_expiry == 1) // if the expiry month was january 1st previously, now it has to become 28th dec of the previous year
  {
    for_return[0] = year_expiry - 1
    for_return[1] = 12
    for_return[2] = 31
    return for_return
  }
  if (month_expiry != 1) // if the expiry month is any thing else
  {
    for_return[1] = month_expiry - 1
    for_return[2] = 28
    return for_return
  }
 }
 // 28th of the previous month was chosen to take into account of non leap febraury 
}





function onEdit(e){
  // trigger when edited
  
  // shall immediately create a date of expiry
  
  // shall arrange according to the date of expiry
  var sheet_obj = e.source // this is sheet object that will be used
  var id_sheet = sheet_obj.getId() // gets the id of the current sheeet
  
  var active_sheet = sheet_obj.getSheets()[0] // gets the active sheet
  
  
  
  Utilities.sleep(9000)
  
  var last_row = active_sheet.getLastRow()
  
  var wanted_range = active_sheet.getRange(last_row,1,1,3)
  Logger.log("a")
  var wanted_values = wanted_range.getValues() // the array containing the worked year,month and day respectively
 
  var expiry_value_array = expiry(Number(wanted_values[0][0]),Number(wanted_values[0][1]),Number(wanted_values[0][2])) // the array containing the expiry date values
  Logger.log(expiry_value_array)
  var to_be_set_range = active_sheet.getRange(last_row,4,1,3)
  to_be_set_range.setValues([expiry_value_array]) // set the expiry date in the corresponding row (the function getRange returns a parameter that is a two dimensional array
  // hence the .setValues function has to be given a two dimensional array
  // expiry_value_array is 1 dimensional array, putting it as the first element in another array makes it a two dimensional array
  var last_row = active_sheet.getLastRow() // gets the last row with data in the current active sheet.
  var start_row = 2
  
  var range_to_sort = active_sheet.getRange(start_row, 1, (last_row -1), 6) // the range to be sorted 
  range_to_sort.sort([{column:4,ascending:true},{column:5,ascending:true},{column:6,ascending:true}]) // sorts the range first by column1,then2 and then by3
  
}








function createDate([year,month,date])
{
  // customized for createTable function
  // given three numbers it will return a string in Indian style
  var date_string = String(Number(date))+ ' - ' + String(Number(month)) + ' - ' + String(Number(year))
  return date_string 
}




function createTable()
{
 // this will return a string that creates a html table 
 var a = ret_activesheet("<name of the active sheet>",0)
 var lastrow = a.getLastRow()
 var wanted_range1 = a.getRange(2,1,(lastrow -1),3)
 var wanted_values1 = wanted_range1.getValues()
 var worked = new Array()
                        
 for (i in wanted_values1)
 {
   worked.push(createDate(wanted_values1[i]))
   
 }
  
 var wanted_range2 = a.getRange(2,4,(lastrow-1),3)
 var wanted_values2 = wanted_range2.getValues()
 var expiry_days = new Array()
 
 for (i in wanted_values2)
 {
   expiry_days.push(createDate(wanted_values2[i]))
 }
 
 var ret_string = "" 
 for (i = 0; i < (worked.length);i = i + 1)
 {
   ret_string = ret_string + "<tr>"+"<td>" + String(worked[i]) + "</td>" + "<td>" + String(expiry_days[i]) + "</td>" +"</tr>"
 }
 
 ret_string = "<tr>" + "<th>" + "You worked on these days" + "</th>" + "<th>" + "COL expiry date" + "</th>" + ret_string
 return ret_string
 Logger.log(ret_string)
  // star here
 Logger.log(worked.length) 
 Logger.log(worked)
 Logger.log(expiry_days)
}





function sendEmail()
{
  // a function send a col summary every month 
  
  var to = "<desired email id>"
  var sub = "Summary of COL"
  var html_body = "<html>"+
                  "<body><b>Please do delete the awailed leaves from the spreadsheet </b></body> <br><br>"+
                  "<table border = '1' style='width:100%'>" + createTable() +
                  "</table>"+
                  "</html>"
  
    
  GmailApp.sendEmail(to, sub, "",{htmlBody:html_body} )
}






function delete_expired_col()
{
  // function to delete dates that are crossed
  var active_sheet = ret_activesheet("<the name of the google sheet>",0)
  var start_row = 2
  var last_row = active_sheet.getLastRow()
  var today_date = new Date()
  var cur_day = Number(today_date.getDate())
  var cur_month = Number(today_date.getMonth() + 1)
  var cur_year = Number(today_date.getYear())
  
  var wanted_range = active_sheet.getRange(2,4,(last_row -1),3)
  var wanted_values = wanted_range.getValues()
  
  for (i in wanted_values)
  {
    var dates = new Array()
    var date1 = today_date
    var date2 = new Date(wanted_values[i][0],(wanted_values[i][1]-1),wanted_values[i][2])
    var date1_timesince = date1.getTime()
    var date2_timesince = date2.getTime()
    
    dates.push(date1_timesince)
    dates.push(date2_timesince)
    Logger.log(dates)
    dates.sort()
    Logger.log(dates)
    
    // dates array is sorted in ascending order
    
    if (dates[0] == date2_timesince) // if the dates array contains col expiry date then the col has expired
    {
      
      active_sheet.deleteRow(i+2) // since the startrow is 2
    }
  }
  
  
}
