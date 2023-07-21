var last_updated_time = '';

var get_active_sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function GetDataRange(){
  
  var all_data_val_arr = get_active_sheet.getRange(1, 1, 20, 100).getValues();

  var column_arr = [];
  for (var j in all_data_val_arr[0]) 
  {
    column_arr.push(all_data_val_arr[0][j])  
  }
  var get_row_index = SpreadsheetApp.getActiveSheet().getActiveCell().getRowIndex();

  var obj = {'col_arr':column_arr,'get_row_index':get_row_index,'values':all_data_val_arr};

  return obj;
}

function setValues(sentiments,orderStatus,returnlabelsts,o_date,track_link,other_tickets){

   var all_arr = GetDataRange();
   var get_row_index = all_arr['get_row_index'];
   var data = all_arr['col_arr'];
   var index =  data.indexOf("Sentiment");
   var index2 = data.indexOf("Order Status");
   var index3 = data.indexOf("Return Label Status Id");
   var index4 = data.indexOf("Order Date");
   var index5 = data.indexOf("Tracking Id");
   console.log(other_tickets);

  var other_ticket_index = data.indexOf("Other Ticket"); 
  
  console.log(all_arr);
  if(index4){
    console.log(o_date);
    var date = new Date(o_date);
    var pst = date.toUTCString();
    get_active_sheet.getRange(get_row_index,index4+1).setValue(pst);
  }
  if(index5){
    get_active_sheet.getRange(get_row_index,index5+1).setValue(track_link);
  }
   if(index3){
    get_active_sheet.getRange(get_row_index,index3+1).setValue(returnlabelsts);
  }
  if(index){
     get_active_sheet.getRange(get_row_index,index+1).setValue(sentiments);
  }
   if(index2){
    get_active_sheet.getRange(get_row_index,index2+1).setValue(orderStatus);
  }
  
  
  

  if(other_ticket_index){
    get_active_sheet.getRange(get_row_index,other_ticket_index+1).setValue(other_tickets);
  }

}
function convert2(str) {
  //var date = new Date(str),
  var date = str,
    mnth = ("0" + (date.getMonth() + 1)).slice(-2),
    day = ("0" + date.getDate()).slice(-2);
  return [date.getFullYear(), mnth, day].join("-");
}

function convert(str) {
  var mnths = {
      Jan: "01",
      Feb: "02",
      Mar: "03",
      Apr: "04",
      May: "05",
      Jun: "06",
      Jul: "07",
      Aug: "08",
      Sep: "09",
      Oct: "10",
      Nov: "11",
      Dec: "12"
    },
    date = str.split(" ");
   // console.log(date);
  return [date[3], mnths[date[2]], date[1]].join("-");
}
function getSecondValues(){
    var all_arr = GetDataRange();
    var get_row_index = all_arr['get_row_index'];
    var data = all_arr['col_arr'];
    var Sentiment_data = data.indexOf("Sentiment");
    var Order_data = data.indexOf("Order Status");
    var Return_data = data.indexOf("Return Label Status Id");
    var Tracking_data = data.indexOf("Tracking Id");

    var Order_date_val_data = data.indexOf("Order Date");
     
    var all_values_data  = all_arr['values'];
    var Order_date_val = '';
    if(Order_date_val_data != -1 ){
      Order_date_val = get_active_sheet.getRange(get_row_index, Order_date_val_data+1).getValue();
      var date = new Date(Order_date_val);
      Order_date_val = date.toUTCString();
      Order_date_val=convert(Order_date_val);
      
    }

    var Other_ticket_data = data.indexOf('Other Ticket');

    
    var  sentiment_val = '';
    
    if(Sentiment_data != -1){
      sentiment_val = get_active_sheet.getRange(get_row_index, Sentiment_data+1).getValue();
    }
    
     var Order_val = '';
    if(Order_data != -1){
      Order_val=get_active_sheet.getRange(get_row_index, Order_data+1).getValue()
    }
    
    var Return_data_val= '';
    
    if(Return_data != -1){
      Return_data_val = get_active_sheet.getRange(get_row_index, Return_data+1).getValue();
    }
    
    var Tracking_data_val = '';
    
    if(Tracking_data != -1){
      var Tracking_data_val = get_active_sheet.getRange(get_row_index, Tracking_data+1).getValue();
    }
    
    var Other_ticket_data_val ='';
    
    if(Other_ticket_data != -1){
      Other_ticket_data_val = get_active_sheet.getRange(get_row_index, Other_ticket_data+1).getValue()
    }
    
    var obj = [sentiment_val ,Order_val,Return_data_val,Tracking_data_val,Order_date_val,Other_ticket_data_val];
    //console.log(obj);
    return obj;
}

function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('HJH Templates')
        .addItem('Show Sidebar', 'showSidebar')
        .addToUi();
}

function onInstall(e) {
    onOpen(e);
    showSidebar();
}

function showSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('Sidebar_new')
        .setTitle('HJH Templates')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showSidebar(html);
}

function check_ticket_avail(check_ticket)
{
    var get_row_index = SpreadsheetApp.getActiveSheet().getActiveCell().getRowIndex();

    var data = get_active_sheet.getDataRange().getValues();

    if(data[get_row_index-1].indexOf(check_ticket) !== -1)
    {
      var obj = ['true',''];
      return obj;
    }
    else
    {
      var sheet_id = get_active_sheet.getSheetId();
      var obj = ['false',sheet_id];    
      return obj;  
    }
}

function check_ticket_val(counter,check_ticket_text_col,check_ticket,get_row_index){

    var range='';
    
    var values='';
    
    var column_arr = [];
    
    if(check_ticket !=undefined){
      for(var i=0;i<3;i++)
      {
         range = get_active_sheet.getRange(1, 1, row_counter, counter);//("A1:ZZ")
         
         values = range.getValues();
        
         for (var j in values[0]) 
         {
            column_arr.push(values[0][j])  
         }
          
         var obj = {'col_arr':column_arr,'values':values};
          
         var target_col_no = column_arr.indexOf(check_ticket_text_col);
          var check_active_cell_row = values[get_row_index-1];
        
          if (check_active_cell_row != undefined && check_active_cell_row.indexOf(check_ticket) !== -1) {

            if(target_col_no == check_active_cell_row.indexOf(check_ticket)){
               return obj;
            }
          }
          
         counter = counter+500;
      }
      
      
    }
    else{
       
       range = get_active_sheet.getRange(1, 1, row_counter, counter);//("A1:ZZ")
         
       values = range.getValues();
       
       for (var j in values[0]) 
       {
          column_arr.push(values[0][j])  
       }
       
       var obj = {'col_arr':column_arr,'values':values};
    
       
    } 
    return obj;
}


var counter = 200;

var row_counter = 500;

function getLastValue(check_ticket,check_ticket_text_col,insert_val_col_name) {

      var get_row_index = SpreadsheetApp.getActiveSheet().getActiveCell().getRowIndex();
      
      row_counter = get_row_index;
      
      var all_columns_arr = check_ticket_val(counter,check_ticket_text_col,check_ticket,get_row_index);
      
      var target_col_no = all_columns_arr['col_arr'].indexOf(check_ticket_text_col);
      
      var data = all_columns_arr['values'];
      
      if(data[get_row_index-1].indexOf(check_ticket) !== -1 && target_col_no == data[get_row_index-1].indexOf(check_ticket))
      {
         // console.log('step 3')
          var sheet_id = get_active_sheet.getSheetId();
          var inserted_id = all_columns_arr['col_arr'].indexOf(insert_val_col_name);
          inserted_id= inserted_id+1;
          var option_last_selected_val = get_active_sheet.getRange(get_row_index, inserted_id).getValue();
          // console.log(option_last_selected_val);
          var second_stages_values = getSecondValues();
          console.log(second_stages_values);
          var object = [option_last_selected_val,get_row_index,sheet_id,second_stages_values];
          return object;
      }
      else
      { 
         var object = ['error'];
         return object;
      }
}

function update_temp_status(status,status_color,ticket_id_row_val,check_ticket,col_temp_status,temp_text,template_insert_col_name,check_ticket_text_col,val_return,insert_val_col_name)
{
    var get_row_index = SpreadsheetApp.getActiveSheet().getActiveCell().getRowIndex();
    //counter = get_row_index;
   // console.log(counter);
   row_counter =get_row_index;
    var all_columns_arr=check_ticket_val(counter);
    var target_col_no = all_columns_arr['col_arr'].indexOf(check_ticket_text_col);
    target_col_no = target_col_no+1;
    get_active_sheet.getRange(ticket_id_row_val, target_col_no).setBackground(status_color);

    

    if(ticket_id_row_val == 0){
      ticket_id_row_val = get_row_index;
    }
    
    if(val_return !='' && insert_val_col_name != undefined){
      var target_col_no = all_columns_arr['col_arr'].indexOf(insert_val_col_name);
      target_col_no = target_col_no+1;
      
        console.log(get_active_sheet.getRange(ticket_id_row_val, target_col_no));
        get_active_sheet.getRange(ticket_id_row_val, target_col_no).setValue(val_return);
      
    }
    

    return check_ticket_avail(check_ticket);
  
}

function saveSuggestion(temp_text,temp_name,ticket_id_row_val,check_ticket,branch_suggestion,suggestion_text){
    
    
     var get_row_index = SpreadsheetApp.getActiveSheet().getActiveCell().getRowIndex();
     row_counter =get_row_index;
   
   
    var all_columns_arr=check_ticket_val(counter);
      
    var branch_sugg_no = all_columns_arr['col_arr'].indexOf(branch_suggestion);
     
    branch_sugg_no = branch_sugg_no+1; 

    var sugg_text_no = all_columns_arr['col_arr'].indexOf(suggestion_text);
     
    sugg_text_no = sugg_text_no+1; 
  
    if(sugg_text_no != 1){
      get_active_sheet.getRange(ticket_id_row_val, sugg_text_no).setValue(temp_text);
    }
   
    if(branch_sugg_no != 1){
      get_active_sheet.getRange(ticket_id_row_val, branch_sugg_no).setValue(temp_name);
    }
   
    return check_ticket_avail(check_ticket);
    
}

function get_sheet_id() {
    var sheet_id = SpreadsheetApp.getActiveSpreadsheet().getId();
    return sheet_id;
}


function paste_selected_val_row(val,check_ticket,ticket_id_row_val,insert_val_col_name) 
{

    var get_row_index = SpreadsheetApp.getActiveSheet().getActiveCell().getRowIndex();
    row_counter= get_row_index;
    if(ticket_id_row_val == 0){
      ticket_id_row_val = get_row_index;
    }
    var all_columns_arr=check_ticket_val(counter);
      
    var target_col_no = all_columns_arr['col_arr'].indexOf(insert_val_col_name);
     
    target_col_no = target_col_no+1;
    
    get_active_sheet.getRange(ticket_id_row_val, target_col_no).setValue(val);

    return check_ticket_avail(check_ticket);
    
    
}

//function changeActiveCell(val,check_ticket,ticket_id_row_val,insert_val_col_name,ticket_id_row_val)
function changeActiveCell(ticket_id_row_val)
{
  // getLastValue(check_ticket,check_ticket_text_col,insert_val_col_name);
  // paste_selected_val_row(val,check_ticket,ticket_id_row_val,insert_val_col_name);
  var selection_val = 'A'+ticket_id_row_val;
  get_active_sheet.setActiveSelection(selection_val);
  SpreadsheetApp.flush(); // Force this update to happen
  // SpreadsheetApp.getActiveSheet().setActiveSelection(range)
  //ss.setActiveSelection(ticket_id_row_val);
}

function request_update_all(insert_new_temp_options_values,update_insert_temp_status,insert_temp_sugg){
    
  paste_selected_val_row(insert_new_temp_options_values[0],insert_new_temp_options_values[1],insert_new_temp_options_values[2],insert_new_temp_options_values[3]);
  
  if(update_insert_temp_status[0] !=""){
    update_temp_status(update_insert_temp_status[0],update_insert_temp_status[1],update_insert_temp_status[2],update_insert_temp_status[3],update_insert_temp_status[4],update_insert_temp_status[5],update_insert_temp_status[6],update_insert_temp_status[7],'','');
  }
  
  if(insert_temp_sugg[0] !=""){
    saveSuggestion(insert_temp_sugg[0],insert_temp_sugg[1],insert_temp_sugg[2],insert_temp_sugg[3],insert_temp_sugg[4],insert_temp_sugg[5]);
  }

  return check_ticket_avail(insert_new_temp_options_values[1]);
  
  //  console.info(update_insert_temp_status);
  
  //  console.info(insert_temp_sugg);
}

