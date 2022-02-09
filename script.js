

function scheduleTrigger(){     
  var builder0 = ScriptApp.newTrigger("clear_send").timeBased();
  builder0.atHour(0).nearMinute(0).everyDays(1);
  builder0.create();

  var builder1 = ScriptApp.newTrigger("module_1").timeBased();
  builder1.atHour(0).nearMinute(10).everyDays(1);
  builder1.create();

  var builder2 = ScriptApp.newTrigger("module_2").timeBased();
  builder2.atHour(0).nearMinute(20).everyDays(1);
  builder2.create();

  var builder3 = ScriptApp.newTrigger("module_3").timeBased();
  builder3.atHour(0).nearMinute(30).everyDays(1);
  builder3.create();

  var builder4 = ScriptApp.newTrigger("module_4").timeBased();
  builder4.atHour(0).nearMinute(40).everyDays(1);
  builder4.create();

  var builder5 = ScriptApp.newTrigger("module_5").timeBased();
  builder5.atHour(0).nearMinute(50).everyDays(1);
  builder5.create();

  var builder6 = ScriptApp.newTrigger("module_6").timeBased();
  builder6.atHour(1).nearMinute(0).everyDays(1);
  builder6.create();

  var builder7 = ScriptApp.newTrigger("module_7").timeBased();
  builder7.atHour(1).nearMinute(10).everyDays(1);
  builder7.create();

  var builder8 = ScriptApp.newTrigger("module_8").timeBased();
  builder8.atHour(1).nearMinute(20).everyDays(1);
  builder8.create();

  var todaylist = ScriptApp.newTrigger("sort_and_get_todaylist").timeBased();
  todaylist.atHour(1).nearMinute(30).everyDays(1);
  todaylist.create();


}





function rewrite(module_number) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var send_sheet = ss.getSheetByName("대기열");

  send_sheet.getRange('A:J').offset(1, 0, send_sheet.getRange('A:J').getNumRows() - 1).sort([{column: 1, ascending: true}]);

  var range2 = send_sheet.getDataRange();
  var last_row = range2.getLastRow();
  var list_module = [];

  for(var i of [module_number]){
    var range2 = send_sheet.getDataRange();
    var last_row = range2.getLastRow();

    for(var j =2; j <= last_row; j++){

          
          if (send_sheet.getRange(j,1).getValue() === "모듈"+String(i)){

              list_module.push(j);

          }// if



    } //for j   

  
      if (list_module.length !== 0){

          var s = list_module[0];  
          var e = list_module[list_module.length-1];
          send_sheet.deleteRows(s,e);


      }
    
      


  } //for i 



}


function module_1_rewrite() {

  rewrite(1);
  module_1();
  sort_and_get_todaylist();

}

function module_2_rewrite() {

  rewrite(2);
  module_2();
  sort_and_get_todaylist();

}

function module_3_rewrite() {

  rewrite(3);
  module_3();
  sort_and_get_todaylist();

}
function module_4_rewrite() {

  rewrite(4);
  module_4();
  sort_and_get_todaylist();

}

function module_5_rewrite() {

  rewrite(5);
  module_5();
  sort_and_get_todaylist();

}

function module_6_rewrite() {

  rewrite(6);
  module_6();
  sort_and_get_todaylist();

}

function module_7_rewrite() {

  rewrite(7);
  module_7();
  sort_and_get_todaylist();

}
function module_8_rewrite() {

  rewrite(8);
  module_8();
  sort_and_get_todaylist();

}




function clear_send(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var send_sheet = ss.getSheetByName("대기열");
  send_sheet.getRange('2:1000000').clear({contentsOnly: true});
}





function sort_and_get_todaylist(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var send_sheet = ss.getSheetByName("대기열");

  send_sheet.getRange('A:J').offset(1, 0, send_sheet.getRange('A:J').getNumRows() - 1).sort([{column: 7, ascending: true}, {column: 8, ascending: true}, {column: 9, ascending: true}, {column: 10, ascending: true}, {column: 1, ascending: true}, {column: 2, ascending: true}]);

  var current_date = new Date();
  var tommorrow = new Date();
  tommorrow.setDate(current_date.getDate()+1);


  var current_Year = current_date.getFullYear();
  var current_Month = current_date.getMonth()+1 ;
  var current_Day = current_date.getDate();
  var current_Hour = current_date.getHours();
  
  console.log(current_date)
  console.log(tommorrow)
  var range2 = send_sheet.getDataRange();
  var last_row = range2.getLastRow();
  console.log(last_row);
  var list_today = [];

      for(var i =2; i <= last_row; i++){

          
          if (send_sheet.getRange(i,7).getValue() === current_Year &&
              send_sheet.getRange(i,8).getValue() === current_Month &&
              send_sheet.getRange(i,9).getValue() === current_Day){

              list_today.push(i);

          }// if
          
          
      } // for i  

      if (list_today.length === 0){

         send_sheet.getRange('2:1000000').clear({contentsOnly: true});


      }
    
      else {

          try{
                var s = list_today[0];  
                var e = list_today[list_today.length-1];

                send_sheet.deleteRows(e+1,last_row-e);
                send_sheet.deleteRows(2,s-2);
              }

          catch{       }


      }
  send_sheet.getRange('D:D').setNumberFormat('00000000000');



}



function data_set(module_number) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name_sheet = ss.getSheetByName("명단");
    var module_sheet = ss.getSheetByName("모듈정보");
    var text_sheet = ss.getSheetByName("문자내용");
    var send_sheet = ss.getSheetByName("대기열");
    var sample_sheet = ss.getSheetByName("문자샘플");
    send_sheet.getRange('D:D').setNumberFormat('00000000000');

  for(var i of [module_number+3]){
                                                
              var module_indicator  = []
              module_indicator.push(i-3)
              console.log(module_indicator)



            
                sample_sheet.getRange('C:C').clear({contentsOnly: true});
                sample_sheet.getRange('D:D').clear({contentsOnly: true});
                sample_sheet.getRange('E:E').clear({contentsOnly: true});
                sample_sheet.getRange('F:F').clear({contentsOnly: true});
                sample_sheet.getRange('G:G').clear({contentsOnly: true});


         
              var range = module_sheet.getRange(2,i-2);
              var date_start = new Date(range.getValue());
              var weekday = date_start.getDay(); 
              var module_name = module_sheet.getRange(3,i-2).getValue();
              var module_prof_name = module_sheet.getRange(4,i-2).getValue();
              var module_prof_phone = module_sheet.getRange(5,i-2).getValue();
              var module_prof_email = module_sheet.getRange(6,i-2).getValue();
              var module_admin_name = module_sheet.getRange(7,i-2).getValue();    
              var module_admin_phone = module_sheet.getRange(8,i-2).getValue();
              var module_admin_email = module_sheet.getRange(9,i-2).getValue();
              var module_off_class = module_sheet.getRange(10,i-2).getValue();
              var module_on_class = module_sheet.getRange(11,i-2).getValue();      
              var module_offattend_survey = module_sheet.getRange(12,i-2).getValue();
    
      
              var date_list_before = []
              var date_list_after = []
              

              for(var d=0; d<15 ; d++){
                var dating1 = new Date(date_start);
                var dating2 = new Date(date_start);  
                dating1.setDate(date_start.getDate()-d);
                dating2.setDate(date_start.getDate()+d);
                date_list_before.push(dating1);
                date_list_after.push(dating2);
                }//for d

              console.log(date_start);
              console.log(date_list_before);
              console.log(date_list_before[4]);
              console.log(date_list_before[4+12/12]);
              console.log(weekday);

            
              if(weekday === 5){
                          for(k of [0,12]){
                              //유형1
                              text_sheet.getRange(1+1+k,50).setValue(date_list_before[4+k/12].getFullYear());
                              text_sheet.getRange(1+1+k,51).setValue(date_list_before[4+k/12].getMonth());
                              text_sheet.getRange(1+1+k,52).setValue(date_list_before[4+k/12].getDate());
                              text_sheet.getRange(1+1+k,53).setValue(7);
                              //유형2-1
                              text_sheet.getRange(2+1+k,50).setValue(date_list_before[4+k/12].getFullYear());
                              text_sheet.getRange(2+1+k,51).setValue(date_list_before[4+k/12].getMonth());
                              text_sheet.getRange(2+1+k,52).setValue(date_list_before[4+k/12].getDate());
                              text_sheet.getRange(2+1+k,53).setValue(7);
                              //유형2-2
                              text_sheet.getRange(3+1+k,50).setValue(date_list_before[1+k/12].getFullYear());
                              text_sheet.getRange(3+1+k,51).setValue(date_list_before[1+k/12].getMonth());
                              text_sheet.getRange(3+1+k,52).setValue(date_list_before[1+k/12].getDate());
                              text_sheet.getRange(3+1+k,53).setValue(7);
                              //유형3
                              text_sheet.getRange(4+1+k,50).setValue(date_list_after[2-k/12].getFullYear());
                              text_sheet.getRange(4+1+k,51).setValue(date_list_after[2-k/12].getMonth());
                              text_sheet.getRange(4+1+k,52).setValue(date_list_after[2-k/12].getDate());
                              text_sheet.getRange(4+1+k,53).setValue(7);
                              //유형4
                              text_sheet.getRange(5+1+k,50).setValue(date_list_before[4+k/12].getFullYear());
                              text_sheet.getRange(5+1+k,51).setValue(date_list_before[4+k/12].getMonth());
                              text_sheet.getRange(5+1+k,52).setValue(date_list_before[4+k/12].getDate());
                              text_sheet.getRange(5+1+k,53).setValue(7);      
                              //유형5
                              text_sheet.getRange(6+1+k,50).setValue(date_list_before[0+k/12].getFullYear());
                              text_sheet.getRange(6+1+k,51).setValue(date_list_before[0+k/12].getMonth());
                              text_sheet.getRange(6+1+k,52).setValue(date_list_before[0+k/12].getDate());
                              text_sheet.getRange(6+1+k,53).setValue(17);   
                              //유형6-1
                              text_sheet.getRange(7+1+k,50).setValue(date_list_before[0+k/12].getFullYear());
                              text_sheet.getRange(7+1+k,51).setValue(date_list_before[0+k/12].getMonth());
                              text_sheet.getRange(7+1+k,52).setValue(date_list_before[0+k/12].getDate());
                              text_sheet.getRange(7+1+k,53).setValue(7);
                              //유형6-2
                              text_sheet.getRange(8+1+k,50).setValue(date_list_after[1-k/12].getFullYear());
                              text_sheet.getRange(8+1+k,51).setValue(date_list_after[1-k/12].getMonth());
                              text_sheet.getRange(8+1+k,52).setValue(date_list_after[1-k/12].getDate());
                              text_sheet.getRange(8+1+k,53).setValue(7);
                              //유형6-3
                              text_sheet.getRange(9+1+k,50).setValue(date_list_after[7-k/12].getFullYear());
                              text_sheet.getRange(9+1+k,51).setValue(date_list_after[7-k/12].getMonth());
                              text_sheet.getRange(9+1+k,52).setValue(date_list_after[7-k/12].getDate());
                              text_sheet.getRange(9+1+k,53).setValue(7);
                              //유형6-4
                              text_sheet.getRange(10+1+k,50).setValue(date_list_after[8-k/12].getFullYear());
                              text_sheet.getRange(10+1+k,51).setValue(date_list_after[8-k/12].getMonth());
                              text_sheet.getRange(10+1+k,52).setValue(date_list_after[8-k/12].getDate());
                              text_sheet.getRange(10+1+k,53).setValue(7);
                              //유형7
                              text_sheet.getRange(11+1+k,50).setValue(date_list_after[1-k/12].getFullYear());
                              text_sheet.getRange(11+1+k,51).setValue(date_list_after[1-k/12].getMonth());
                              text_sheet.getRange(11+1+k,52).setValue(date_list_after[1-k/12].getDate());
                              text_sheet.getRange(11+1+k,53).setValue(17);  
                              //유형8
                              text_sheet.getRange(12+1+k,50).setValue(date_list_after[1-k/12].getFullYear());
                              text_sheet.getRange(12+1+k,51).setValue(date_list_after[1-k/12].getMonth());
                              text_sheet.getRange(12+1+k,52).setValue(date_list_after[1-k/12].getDate());
                              text_sheet.getRange(12+1+k,53).setValue(17);

                                
                          }
                              for(var a=1; a < 25 ; a++){
                              text_sheet.getRange(a+1,51).setValue(text_sheet.getRange(a+1,51).getValue() + 1);
                              }
                              text_sheet.getRange(2,55).setValue(date_start);
                              text_sheet.getRange(2,56).setValue(date_list_after[1]);
                              text_sheet.getRange(2,57).setValue(date_list_after[7]);
                              text_sheet.getRange(2,58).setValue(date_list_after[8]);
                              text_sheet.getRange(2,55).setNumberFormat('m"월 "d"일("ddd")"');
                              text_sheet.getRange(2,56).setNumberFormat('m"월 "d"일("ddd")"');
                              text_sheet.getRange(2,57).setNumberFormat('m"월 "d"일("ddd")"');
                              text_sheet.getRange(2,58).setNumberFormat('m"월 "d"일("ddd")"');
                              text_sheet.getRange(2,55).setValue.toString();
                              text_sheet.getRange(2,56).setValue.toString();
                              text_sheet.getRange(2,57).setValue.toString();
                              text_sheet.getRange(2,58).setValue.toString();
                              sample_sheet.getRange(14,6).setFormula("=text('문자내용'!BC2,\"m월 d일(aaa)\")");  // 일시               
                              sample_sheet.getRange(14,7).setFormula("=text('문자내용'!BE2,\"m월 d일(aaa)\")");  // 일시  
                              sample_sheet.getRange(16,6).setFormula("=text('문자내용'!BD2,\"m월 d일(aaa)\")");  // 일시               
                              sample_sheet.getRange(16,7).setFormula("=text('문자내용'!BF2,\"m월 d일(aaa)\")");  // 일시

              }//if 금
              
              else if(weekday === 6){
                            for(k of [0,12]){

                            //유형1
                            text_sheet.getRange(1+1+k,50).setValue(date_list_before[5+k/12].getFullYear());
                            text_sheet.getRange(1+1+k,51).setValue(date_list_before[5+k/12].getMonth());
                            text_sheet.getRange(1+1+k,52).setValue(date_list_before[5+k/12].getDate());
                            text_sheet.getRange(1+1+k,53).setValue(7)
                            //유형2-1
                            text_sheet.getRange(2+1+k,50).setValue(date_list_before[5+k/12].getFullYear());
                            text_sheet.getRange(2+1+k,51).setValue(date_list_before[5+k/12].getMonth());
                            text_sheet.getRange(2+1+k,52).setValue(date_list_before[5+k/12].getDate());
                            text_sheet.getRange(2+1+k,53).setValue(7);
                            //유형2-2
                            text_sheet.getRange(3+1+k,50).setValue(date_list_before[1+k/12].getFullYear());
                            text_sheet.getRange(3+1+k,51).setValue(date_list_before[1+k/12].getMonth());
                            text_sheet.getRange(3+1+k,52).setValue(date_list_before[1+k/12].getDate());
                            text_sheet.getRange(3+1+k,53).setValue(7);
                            //유형3
                            text_sheet.getRange(4+1+k,50).setValue(date_list_after[1-k/12].getFullYear());
                            text_sheet.getRange(4+1+k,51).setValue(date_list_after[1-k/12].getMonth());
                            text_sheet.getRange(4+1+k,52).setValue(date_list_after[1-k/12].getDate());
                            text_sheet.getRange(4+1+k,53).setValue(7);
                            //유형4
                            text_sheet.getRange(5+1+k,50).setValue(date_list_before[5+k/12].getFullYear());
                            text_sheet.getRange(5+1+k,51).setValue(date_list_before[5+k/12].getMonth());
                            text_sheet.getRange(5+1+k,52).setValue(date_list_before[5+k/12].getDate());
                            text_sheet.getRange(5+1+k,53).setValue(7);
                            //유형5
                            text_sheet.getRange(6+1+k,50).setValue(date_list_before[0+k/12].getFullYear());
                            text_sheet.getRange(6+1+k,51).setValue(date_list_before[0+k/12].getMonth());
                            text_sheet.getRange(6+1+k,52).setValue(date_list_before[0+k/12].getDate());
                            text_sheet.getRange(6+1+k,53).setValue(7);
                            //유형6-1
                            text_sheet.getRange(7+1+k,50).setValue(date_list_before[0+k/12].getFullYear());
                            text_sheet.getRange(7+1+k,51).setValue(date_list_before[0+k/12].getMonth());
                            text_sheet.getRange(7+1+k,52).setValue(date_list_before[0+k/12].getDate());
                            text_sheet.getRange(7+1+k,53).setValue(7);
                            //유형6-2
                            text_sheet.getRange(8+1+k,50).setValue(date_list_after[6-k/12].getFullYear());
                            text_sheet.getRange(8+1+k,51).setValue(date_list_after[6-k/12].getMonth());
                            text_sheet.getRange(8+1+k,52).setValue(date_list_after[6-k/12].getDate());
                            text_sheet.getRange(8+1+k,53).setValue(7);
                            //유형6-3
                            text_sheet.getRange(9+1+k,50).setValue(date_list_after[7-k/12].getFullYear());
                            text_sheet.getRange(9+1+k,51).setValue(date_list_after[7-k/12].getMonth());
                            text_sheet.getRange(9+1+k,52).setValue(date_list_after[7-k/12].getDate());
                            text_sheet.getRange(9+1+k,53).setValue(7);
                            //유형6-4
                            text_sheet.getRange(10+1+k,50).setValue(date_list_after[13-k/12].getFullYear());
                            text_sheet.getRange(10+1+k,51).setValue(date_list_after[13-k/12].getMonth());
                            text_sheet.getRange(10+1+k,52).setValue(date_list_after[13-k/12].getDate());
                            text_sheet.getRange(10+1+k,53).setValue(7);
                            //유형7
                            text_sheet.getRange(11+1+k,50).setValue(date_list_before[0+k/12].getFullYear());
                            text_sheet.getRange(11+1+k,51).setValue(date_list_before[0+k/12].getMonth());
                            text_sheet.getRange(11+1+k,52).setValue(date_list_before[0+k/12].getDate());
                            text_sheet.getRange(11+1+k,53).setValue(22);
                            //유형8
                            text_sheet.getRange(12+1+k,50).setValue(date_list_before[0+k/12].getFullYear());
                            text_sheet.getRange(12+1+k,51).setValue(date_list_before[0+k/12].getMonth());
                            text_sheet.getRange(12+1+k,52).setValue(date_list_before[0+k/12].getDate());
                            text_sheet.getRange(12+1+k,53).setValue(22);
                            

                            }
                  for(var x=1; x < 25 ; x++){
                  text_sheet.getRange(x+1,51).setValue(text_sheet.getRange(x+1,51).getValue() + 1);
                  }
                  text_sheet.getRange(2,55).setValue(date_start);
                  text_sheet.getRange(2,56).setValue(date_list_after[6]);
                  text_sheet.getRange(2,57).setValue(date_list_after[7]);
                  text_sheet.getRange(2,58).setValue(date_list_after[13]);
                  text_sheet.getRange(2,55).setNumberFormat('m"월 "d"일("ddd")"');
                  text_sheet.getRange(2,56).setNumberFormat('m"월 "d"일("ddd")"');
                  text_sheet.getRange(2,57).setNumberFormat('m"월 "d"일("ddd")"');
                  text_sheet.getRange(2,58).setNumberFormat('m"월 "d"일("ddd")"'); 

                  sample_sheet.getRange(14,6).setFormula("=text('문자내용'!BD2,\"m월 d일(aaa)\")");               
                  sample_sheet.getRange(14,7).setFormula("=text('문자내용'!BF2,\"m월 d일(aaa)\")");  
                  sample_sheet.getRange(16,6).setFormula("=text('문자내용'!BC2,\"m월 d일(aaa)\")");              
                  sample_sheet.getRange(16,7).setFormula("=text('문자내용'!BE2,\"m월 d일(aaa)\")"); 

              }//if 토    

                    //2행
                    sample_sheet.getRange(2,3).setValue(module_on_class);
                    //4행
                    text_sheet.getRange('BC3').setFormula('=textjoin(", ",,BC2:BF2)');
                    sample_sheet.getRange(4,3).setFormula("='문자내용'!BC3");
                    //6행
                    //8행
                    text_sheet.getRange(4,55).setValue(module_name);
                    text_sheet.getRange(4,56).setValue(module_prof_name);
                    text_sheet.getRange(5,55).setFormula('=textjoin(,,BC4,"(",BD4," 교수)")');
                    sample_sheet.getRange(8,3).setFormula("='문자내용'!BC5");
                    //10행
                    sample_sheet.getRange(10,3).setFormula("='문자내용'!BC5");
                    //12행
                    text_sheet.getRange(7,55).setFormula('=textjoin(,,BC2," ~ ",BF2)');
                    sample_sheet.getRange(12,3).setFormula("='문자내용'!BC7");
                    sample_sheet.getRange(12,4).setFormula("='문자내용'!BC5");
                    //14행
                    sample_sheet.getRange(14,3).setValue(module_name);  
                    sample_sheet.getRange(14,4).setValue(module_prof_name);  
                    sample_sheet.getRange(14,5).setValue(module_off_class);  
                    //16행
                    sample_sheet.getRange(16,3).setValue(module_name);  
                    sample_sheet.getRange(16,4).setValue(module_prof_name);  
                    sample_sheet.getRange(16,5).setValue(module_off_class);      
                    //18행
                    sample_sheet.getRange(18,3).setValue(module_prof_name);  
                    sample_sheet.getRange(18,4).setFormula("=text('문자내용'!BC2,\"m/d(aaa)\")");  
                    sample_sheet.getRange(18,5).setValue(module_name);  
                    sample_sheet.getRange(18,6).setFormula('=textjoin(", ",,text(\'문자내용\'!BC2,"m/d(aaa)"),text(\'문자내용\'!BD2,"m/d(aaa)"),text(\'문자내용\'!BE2,"m/d(aaa)"),text(\'문자내용\'!BF2,"m/d(aaa)"))');  
                    sample_sheet.getRange(18,7).setValue(module_off_class);  
                    //20행
                        var kor_num_module = ""
                        if(y=4){kor_num_module = "첫"}
                        else if(y=5){kor_num_module = "두"}
                        else if(y=6){kor_num_module = "세"}
                        else if(y=7){kor_num_module = "네"}
                        else if(y=8){kor_num_module = "다섯"}
                        else if(y=9){kor_num_module = "여섯"}
                        else if(y=10){kor_num_module = "일곱"}
                        else if(y=11){kor_num_module = "여덟"}
                    sample_sheet.getRange(20,3).setValue(kor_num_module);  
                    sample_sheet.getRange(20,4).setFormula("=text('문자내용'!BC2,\"m월 d일(aaa)\")");  
                    sample_sheet.getRange(20,5).setValue(module_name);  
                    sample_sheet.getRange(20,6).setFormula('=textjoin(", ",,text(\'문자내용\'!BC2,"m/d(aaa)"),text(\'문자내용\'!BD2,"m/d(aaa)"),text(\'문자내용\'!BE2,"m/d(aaa)"),text(\'문자내용\'!BF2,"m/d(aaa)"))');  
                    sample_sheet.getRange(20,7).setValue(module_offattend_survey);  

             
                    sample_sheet.getRange('\'문자샘플\'!A20').copyTo(text_sheet.getRange(2,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A18').copyTo(text_sheet.getRange(3,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A18').copyTo(text_sheet.getRange(4,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A24').copyTo(text_sheet.getRange(5,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A4').copyTo(text_sheet.getRange(6,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A6').copyTo(text_sheet.getRange(7,i+5), {contentsOnly: true});

                    if(weekday === 5){
                    sample_sheet.getRange('\'문자샘플\'!A14').copyTo(text_sheet.getRange(8,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A16').copyTo(text_sheet.getRange(9,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!B14').copyTo(text_sheet.getRange(10,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!B16').copyTo(text_sheet.getRange(11,i+5), {contentsOnly: true});
                    }
                    else if(weekday === 6){
                    sample_sheet.getRange('\'문자샘플\'!A16').copyTo(text_sheet.getRange(8,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A14').copyTo(text_sheet.getRange(9,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!B16').copyTo(text_sheet.getRange(10,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!B14').copyTo(text_sheet.getRange(11,i+5), {contentsOnly: true});
                    }


                    sample_sheet.getRange('\'문자샘플\'!A10').copyTo(text_sheet.getRange(12,i+5), {contentsOnly: true});
                    sample_sheet.getRange('\'문자샘플\'!A8').copyTo(text_sheet.getRange(13,i+5), {contentsOnly: true});


                                                                                         

                  for(var q=1;q<13;q++){
                            var value1 = "'******************************************* \n담당자 확인용 문자입니다. \n확인하시어 수정사항이 있으실 경우 [모듈정보] 시트 또는 [문자샘플] 시트를 수정하시기 바랍니다. \n기한은 오늘 오후 11시 30분까지입니다. \n*******************************************\n"
                            var value2 = text_sheet.getRange(q+1,i+5).getValue();
                            var value3 = value1 + value2
                            send_sheet.appendRow([module_name,
                                              text_sheet.getRange(q+1,1).getValue(),  
                                              module_admin_name,
                                              module_admin_phone,,
                                              value3,
                                              text_sheet.getRange(q+1+12,50).getValue(),
                                              text_sheet.getRange(q+1+12,51).getValue(),
                                              text_sheet.getRange(q+1+12,52).getValue(),
                                              7])
                          }


                  for(var w of [2,3]){
                            send_sheet.appendRow([module_name,
                                              text_sheet.getRange(w+1,1).getValue(),  
                                              module_prof_name,
                                              module_prof_phone,,
                                              text_sheet.getRange(w+1,i+5).getValue(),
                                              text_sheet.getRange(w+1,50).getValue(),
                                              text_sheet.getRange(w+1,51).getValue(),
                                              text_sheet.getRange(w+1,52).getValue(),
                                              text_sheet.getRange(w+1,53).getValue()])
                          }
              

              
                  for(var j=2; j < 1000 ; j++){
                        var name = name_sheet.getRange(j,16).getValue();
                        var phone = name_sheet.getRange(j,17).getValue();
                            if(name_sheet.getRange(j, i).getValue() === '수강'){      

                                  for(var e of [1,5,7,8,9,10,11]){
                                      if(text_sheet.getRange(e+1,i+5).getValue() !== ""){
                                        send_sheet.appendRow([module_name,
                                                          text_sheet.getRange(e+1,1).getValue(),  
                                                          name,
                                                          phone,,
                                                          text_sheet.getRange(e+1,i+5).getValue(),
                                                          text_sheet.getRange(e+1,50).getValue(),
                                                          text_sheet.getRange(e+1,51).getValue(),
                                                          text_sheet.getRange(e+1,52).getValue(),
                                                          text_sheet.getRange(e+1,53).getValue()])
                                      }                    
                                    }
                                  for(var r of [4]){
                                      if(module_sheet.getRange(13,i-2).getValue() === "있음" && text_sheet.getRange(r+1,i+5).getValue() !== ""){
                                        send_sheet.appendRow([module_name,
                                                          text_sheet.getRange(r+1,1).getValue(),  
                                                          name,
                                                          phone,,
                                                          text_sheet.getRange(r+1,i+5).getValue(),
                                                          text_sheet.getRange(r+1,50).getValue(),
                                                          text_sheet.getRange(r+1,51).getValue(),
                                                          text_sheet.getRange(r+1,52).getValue(),
                                                          text_sheet.getRange(r+1,53).getValue()])
                                      }
                                  }
                                  if(name_sheet.getRange(j, 24).getValue() === '참석'){
                                    for(var z of [6]){
                                        if(text_sheet.getRange(z+1,i+5).getValue() !== ""){
                                          send_sheet.appendRow([module_name,
                                                            text_sheet.getRange(z+1,1).getValue(),  
                                                            name,
                                                            phone,,
                                                            text_sheet.getRange(z+1,i+5).getValue(),
                                                            text_sheet.getRange(z+1,50).getValue(),
                                                            text_sheet.getRange(z+1,51).getValue(),
                                                            text_sheet.getRange(z+1,52).getValue(),
                                                            text_sheet.getRange(z+1,53).getValue()])
                                        }
                                    }
                                  } // if                                 

                              
                            }// if 

                          else if(name_sheet.getRange(j, i).getValue() === '청강'){
                            for(var t of [1,5,7,8,9,10,12]){
                                if(text_sheet.getRange(t+1,i+5).getValue() !== ""){
                                  send_sheet.appendRow([module_name,
                                                    text_sheet.getRange(t+1,1).getValue(),  
                                                    name,
                                                    phone,,
                                                    text_sheet.getRange(t+1,i+5).getValue(),
                                                    text_sheet.getRange(t+1,50).getValue(),
                                                    text_sheet.getRange(t+1,51).getValue(),
                                                    text_sheet.getRange(t+1,52).getValue(),
                                                    text_sheet.getRange(t+1,53).getValue()])
                                }
                            }
                          } //else if                             

 

                  
                  
                  
                  
                  }//for j

            
                

                          
                } // for i
              } // function








function module_1() {

  return data_set(1);
}

function module_2() {

  return data_set(2);
}

function module_3() {

  return data_set(3);
}

function module_4() {

  return data_set(4);
}

function module_5() {

  return data_set(5);
}

function module_6() {

  return data_set(6);

}

function module_7() {

  return data_set(7);

}

function module_8() {

  return data_set(8);

}


