Array.prototype.insert = function ( index, item ) {
    this.splice( index, 0, item );
};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sprawozdanie')
  .addItem('Dodaj nowy miesiąc', 'create_new_month')
  .addItem('Drukuj sprawozdanie', 'print_report')
  .addToUi();
}

function rr(cell_addresses, column_name, sheet_name) {
    var cell_range =  cell_addresses.split('-');
    var start_range = cell_range[0].trim();
    var end_range = cell_range[1].trim();
    return '$' + column_name + '$' + start_range + ':$' + column_name + '$' + end_range;
}

function rd(cell_addresses, column_name, sheet_name) {
    var cell_range =  cell_addresses.split('-');
    var start_range = cell_range[0].trim();
    var end_range = cell_range[1].trim();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    return sheet.getRange(column_name + start_range + ':' + column_name + end_range).getValues();
}

var TEST_DATA = false
var RAND_CELL = false
var HEADER_SIZE = 12

var MINISTRY_GRUOPS = [1,2,3,4,5]
var PUBLISHER_ATTR = 1
var IRREGULAR_ATTR = 2
var DISABLED_ATTR = 4
var AUX_PIONNER_ATTR = 8
var PIONNER_ATTR = 16
var SPECIAL_PIONEER_ATTR = 32
var GROUP_CHIEF_ATTR = 64
var source = SpreadsheetApp.getActiveSpreadsheet();

function create_new_month() {
  var ui = SpreadsheetApp.getUi();
  var sheet_name = ui.prompt("Podaj nazwę miesiąca").getResponseText()
  if (! sheet_name)
    return
  var new_sheet = source.getSheetByName(sheet_name)
  // Creates new month sheet
  if (! new_sheet) {
    new_sheet = create_new_sheet(sheet_name)
  }
  
  new_sheet.getRange("A:Z").clearContent();
  new_sheet.clearFormats();

  // Get publishers groups
  var publishers = source.getSheetByName("LISTA GŁOSICIELI")
  var dataRangeValues = publishers.getDataRange().getValues()

  var row_start = 1;
  var row_end = 0;
  new_sheet.getRange("H:H").setNumberFormat("@STRING@")
  
  for (var nr = 0; nr < MINISTRY_GRUOPS.length; nr++ ) {
    var whole_group = get_group_by_number(dataRangeValues, MINISTRY_GRUOPS[nr])
    var group = whole_group['publishers']
    var group_publishers_attrs = whole_group['publishers_attributes']
    var publishers_flat = []

    group.forEach(function(current_value) {
      publishers_flat.push(current_value[0])
    });
    
    var group_data = group
    var group_size = group_data.length
    row_end += group_size
    
    var nrs = []
    for (var i = 1; i < group_size + 1; i++) {
      nrs.push([i])
    }
    
    // nagłowek
    var header =  [["", "GRUPA " + (nr + 1), "PUBLIKACJE", "FILMY", "GODZINY", "ODWIEDZINY", "STUDIA"]]
    var header_range = new_sheet.getRange("A" + row_start + ":G" + (row_start))
    var summary_range = new_sheet.getRange("A" + row_start + ":G" + (row_start))
    var avg_range = new_sheet.getRange("A" + row_start + ":G" + (row_start))
    header_range.setValues(header).setBackground("#000").setFontColor("#FFF").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(HEADER_SIZE)
    // dane ze służby
    row_start += 1
    row_end += 1
    new_sheet.getRange("A" + row_start + ":A" + row_end).setValues(nrs).setHorizontalAlignment("center").setVerticalAlignment("middle");
    new_sheet.getRange("B" + row_start + ":B" + row_end).setValues(group_data)
    new_sheet.getRange("C" + row_start + ":G" + row_end).setNumberFormat("#0")
    if (TEST_DATA) {
      var fill_values = []
      for(var i=0; i < (row_end-row_start+1);i++){
        var tmp = []
        for(var j=0; j < 5; j++) {
            if (RAND_CELL)
              tmp.push("=rand()*10")
            else
              tmp.push(Number(Math.random()*10).toFixed(0))
        }
        fill_values.push(tmp)
      }
      new_sheet.getRange("C" + row_start + ":G" + row_end).setValues(fill_values)
    }
    new_sheet.getRange("C" + row_start + ":G" + row_end).setNumberFormat("#0")
    new_sheet.getRange("H" + row_start + ":H" + row_end).setValues(group_publishers_attrs)
    var drs = row_start
    var dre = row_end
    // suma 
    row_start += 1
    row_end += 1
    var summary_range = new_sheet.getRange("B" + row_end + ":G" + row_end)
    summary_range.setValues([["SUMA", r2s(c2rng("C",drs, dre)), r2s(c2rng("D",drs, dre)), r2s(c2rng("E",drs, dre)), r2s(c2rng("F",drs, dre)), r2s(c2rng("G",drs, dre))]])
    summary_range.setFontWeight("bold").setBackground("#F2F2F2").setNumberFormat("#0")
    // średnia
    row_start += 1
    row_end += 1
    var avg_range = new_sheet.getRange("B" + row_end + ":G" + row_end)
    avg_range.setValues([["ŚREDNIA", r2a(c2rng("C",drs, dre)), r2a(c2rng("D",drs, dre)), r2a(c2rng("E",drs, dre)), r2a(c2rng("F",drs, dre)), r2a(c2rng("G",drs, dre))]])
    avg_range.setFontWeight("bold").setBackground("#F2F2F2").setNumberFormat("#,##0.00")
    // ustawienie pionierów, nieczynnych itp
    
    set_publishers_special_attr_font(new_sheet, publishers_flat, whole_group['group_chief'], "#FFF", drs)
    set_publishers_special_attr(new_sheet, publishers_flat, whole_group['pioneers'], "#AED6F1", drs)
    set_publishers_special_attr(new_sheet, publishers_flat, whole_group['aux_pioneers'], "#A3E4D7", drs)
    set_publishers_special_attr(new_sheet, publishers_flat, whole_group['irregulars'], "#E6E6E6", drs)
    row_start += group_size
  }
    
    new_sheet.setColumnWidth(1, 30)
    new_sheet.getRange("C:G").setHorizontalAlignment("center")
    
    //statystyki służby
    // PIONIERZY
    var stats_header = new_sheet.getRange("I1:O1")
    var stats_header_values = [["PIONIERZY", "LICZBA", "PUBLIKACJE", "FILMY", "GODZINY", "ODWIEDZINY", "STUDIA" ]]
    stats_header.setValues(stats_header_values).setBackground("#D0ECE7").setFontColor("#000").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(HEADER_SIZE)
    var legend = new_sheet.getRange("I2:I5")
    var legend_values = [["STALI"], ["POMOCNCZY"], ["RAZEM"], ["ŚREDNIA"]]
    legend.setValues(legend_values).setBackground("#FFF").setFontColor("#000").setFontWeight("bold").setHorizontalAlignment("center")
    
    
    var pioneer_sum = new_sheet.getRange("J2:O2")
    var pioneer_sum_values = [find_values_to_count_by_type("H", [PUBLISHER_ATTR+PIONNER_ATTR, GROUP_CHIEF_ATTR]).concat(find_values_to_sum_by_type(["C", "D", "E" , "F", "G"], [PUBLISHER_ATTR+PIONNER_ATTR, GROUP_CHIEF_ATTR], "H"))]
    pioneer_sum.setValues(pioneer_sum_values).setHorizontalAlignment("center").setNumberFormat("#0")
    
    var aux_pioneer_sum = new_sheet.getRange("J3:O3")
    var aux_pioneer_sum_values = [find_values_to_count_by_type("H", [PUBLISHER_ATTR+AUX_PIONNER_ATTR, GROUP_CHIEF_ATTR]).concat(find_values_to_sum_by_type(["C", "D", "E" , "F", "G"], [PUBLISHER_ATTR+AUX_PIONNER_ATTR, GROUP_CHIEF_ATTR], "H"))]
    aux_pioneer_sum.setValues(aux_pioneer_sum_values).setHorizontalAlignment("center").setNumberFormat("#0")
    new_sheet.getRange("J4:O4").setValues([["=sum(J2:J3)", "=sum(K2:K3)", "=sum(L2:L3)", "=sum(M2:M3)", "=sum(N2:N3)", "=sum(O2:O3)"]]).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#F2F2F2").setNumberFormat("#0")

var pioneer_avg = new_sheet.getRange("j5:O5")
    var pioneer_avg_values = [["", "=K4/J4", "=L4/J4", "=M4/J4", "=N4/J4", "=O4/J4"]]
    pioneer_avg.setValues(pioneer_avg_values).setHorizontalAlignment("center").setBackground("#D8D8D8").setNumberFormat("#,##0.00")    
    
    
    // GŁOSICIELE
    var stats_header = new_sheet.getRange("I8:O8")
    var stats_header_values = [["GŁOSICIELE", "LICZBA", "PUBLIKACJE", "FILMY", "GODZINY", "ODWIEDZINY", "STUDIA" ]]
    stats_header.setValues(stats_header_values).setBackground("#AED6F1").setFontColor("#000").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(HEADER_SIZE)
    var legend = new_sheet.getRange("I9:I11")
    var legend_values = [["REGULARNI"], ["NIEREGULARNI"], ["ŚREDNIA"]]
    legend.setValues(legend_values).setBackground("#FFF").setFontColor("#000").setFontWeight("bold").setHorizontalAlignment("center")
    var publishers_sum = new_sheet.getRange("J9:O9")
    var publishers_sum_values = [find_values_to_count_by_type("H", [PUBLISHER_ATTR, GROUP_CHIEF_ATTR]).concat(find_values_to_sum_by_type(["C", "D", "E" , "F", "G"], [PUBLISHER_ATTR, GROUP_CHIEF_ATTR], "H"))]
    publishers_sum.setValues(publishers_sum_values).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#F2F2F2").setNumberFormat("#0")
    var irregulars_sum = new_sheet.getRange("J10:O10")
    var irregulars_sum_values = [find_values_to_count_by_type("H", [PUBLISHER_ATTR+IRREGULAR_ATTR]).concat(["-","-","-","-","-"])]
    irregulars_sum.setValues(irregulars_sum_values).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#F2F2F2").setNumberFormat("#0")

    var publishers_avg = new_sheet.getRange("J11:O11")
    var publishers_avg_values = [["", "=K9/J9", "=L9/J9", "=M9/J9", "=N9/J9", "=O9/J9"]]
    publishers_avg.setValues(publishers_avg_values).setHorizontalAlignment("center").setBackground("#D8D8D8").setNumberFormat("#,##0.00")  


    // CAŁY ZBÓR
    var stats_header = new_sheet.getRange("I13:O13")
    var stats_header_values = [["ZBÓR", "LICZBA", "PUBLIKACJE", "FILMY", "GODZINY", "ODWIEDZINY", "STUDIA" ]]
    stats_header.setValues(stats_header_values).setBackground("#BB8FCE").setFontColor("#000").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(HEADER_SIZE)
    var legend = new_sheet.getRange("I14:I15")
    var legend_values = [["RAZEM"], ["ŚREDNIA"]]
    legend.setValues(legend_values).setBackground("#FFF").setFontColor("#000").setFontWeight("bold").setHorizontalAlignment("center")
    new_sheet.getRange("J14:O14").setValues([["=sum(J4;J9;J10)", "=sum(K4;K9)", "=sum(L4;L9)", "=sum(M4;M9)", "=sum(N4;N9)", "=sum(O4;O9)"]]).setHorizontalAlignment("center").setFontWeight("bold").setBackground("#F2F2F2").setNumberFormat("#0")

    var all_publishers_avg = new_sheet.getRange("J15:O15")
    var all_publishers_avg_values = [["", "=K14/sum(J4;J9)", "=L14/sum(J4;J9)", "=M14/sum(J4;J9)", "=N14/sum(J4;J9)", "=O14/sum(J4;J9)"]]
    all_publishers_avg.setValues(all_publishers_avg_values).setHorizontalAlignment("center").setBackground("#D8D8D8").setNumberFormat("#,##0.00")  
    
    new_sheet.getRange("H:H").setFontColor("#FFF")
    new_sheet.getRange("J:J").setNumberFormat("#")  
    new_sheet.autoResizeColumns(2, 16)
    new_sheet.autoResizeRows(1, 160)
    new_sheet.getRange("A:Z").setVerticalAlignment("middle");
    
}

function find_values_to_count_by_type(columns_name, publisher_types) {
  var column_count_by_type = []

  var publisher_types_formula = []
  var tmp_sum = 0;
  for(var i=0; i < publisher_types.length; i++) {
      tmp_sum += publisher_types[i]
      var tmp = "dec2bin("+ tmp_sum +")"
      publisher_types_formula.push(tmp)
  }
  for(var i=0; i < columns_name.length; i++) {
      var col_letter = columns_name[i]
      var tmp = "=SUM(ARRAYFORMULA(COUNTIF(" + columns_name + ":" + columns_name + "; {" + publisher_types_formula.join(";") + "})))"
      column_count_by_type.push(tmp)
  }
  return column_count_by_type
}

//=SUM(ARRAYFORMULA(COUNTIF(H:H; {dec2bin(1); dec2bin(65)})))
function find_values_to_sum_by_type(columns_name, publisher_types, type_column_name) {
  var column_sum_by_type = [] 
  
  var publisher_types_formula = []
  var tmp_sum = 0;
  for(var i=0; i < publisher_types.length; i++) {
      tmp_sum += publisher_types[i]
      var tmp = "dec2bin("+ tmp_sum +")"
      publisher_types_formula.push(tmp)
  }
  for(var i=0; i < columns_name.length; i++) {
      var col_letter = columns_name[i]
      var tmp = "=SUM(ARRAYFORMULA(SUMIF(" + type_column_name + ":" + type_column_name + "; {" + publisher_types_formula.join(";") + "}; "+ col_letter + ":" + col_letter +")))"
      column_sum_by_type.push(tmp)
  }
  return column_sum_by_type
}


function set_publishers_special_attr(sheet, publishers_flat_list, custom_publishers, color, position_start) {
    for (var i=0; i < custom_publishers.length; i++) {
      var pos = publishers_flat_list.indexOf(custom_publishers[i])
      if ( pos > -1 ) {
           sheet.getRange("B" + (position_start + pos) + ":B" + (position_start + pos)).setBackground(color)
        }
    }
}

function set_publishers_special_attr_font(sheet, publishers_flat_list, custom_publishers, color, position_start) {
    for (var i=0; i < custom_publishers.length; i++) {
      var pos = publishers_flat_list.indexOf(custom_publishers[i])
      if ( pos > -1 ) {
           sheet.getRange("B" + (position_start + pos) + ":B" + (position_start + pos)).setBackground(color).setFontWeight("bold")
        }
    }
}

function create_new_sheet(sheet_name) {
  return source.insertSheet(sheet_name)
}

function get_group_by_number(data, group_nr) {
  var publishers = []
  var aux_pioneers = []
  var pioneers = []
  var irregulars = []
  var disableds = []
  var special_pioneers = []
  var group_chief = []
  var IRREGULAR = 3
  var DISABLED = 4
  var AUX_PIONNER = 5
  var PIONNER = 6
  var SPECIAL_PIONEER = 7
  var GROUP_CHIEF = 8
  var publishers_attrs = []
  
  for (var row_nr=1; row_nr < data.length; row_nr++) {
      if (data[row_nr][1] == group_nr) {
         var dr = data[row_nr]
         var who = dr.slice(0,1)[0]
         
         if (dr[DISABLED] === true) {
             disableds.push(who)
             continue
         }
         if (dr[AUX_PIONNER] === true)
             aux_pioneers.push(who) 
         if (dr[PIONNER] === true)
             pioneers.push(who)
         if (dr[SPECIAL_PIONEER] === true)
             special_pioneers.push(who)
         if (dr[GROUP_CHIEF] === true)
             group_chief.push(who)
         if (dr[IRREGULAR] === true)
           irregulars.push(who)             
         
         var attrs = dr.slice(2, dr.length)
         attrs.forEach(function(element, index, arr) {
           attrs[index] = String(Number(element))
         })
         attrs = attrs.reverse()
         
         if (dr[GROUP_CHIEF] === true) {
            publishers.insert(0, [who]) 
            publishers_attrs.insert(0, [attrs.join("")])
         }
         else {
           publishers.push([who])
           publishers_attrs.push([attrs.join("")])
         }
      }
  }

  return {"publishers": publishers,
          "aux_pioneers": aux_pioneers,
          "pioneers": pioneers,
          "disableds": disableds,
          "irregulars": irregulars,
          "special_pioneers": special_pioneers,
          "group_chief": group_chief,
          "publishers_attributes": publishers_attrs
  }
}

function c2rng(column, range_start, range_stop) {
  return column + range_start + ":" + column + range_stop
}

function r2s(range) {
  return "=SUM(" + range + ")"
}

function r2a(range) {
  return "=IFERROR(AVERAGE(" + range + "); 0)"
}

function test()  {
    generate_monthly_report("Marzec  2020", "download")
}

function generate_monthly_report(act_sheet_name, action_type) {
  var new_sheet = source.getSheetByName(act_sheet_name)
  var publishers_by_groups = new_sheet.getRange("A1:G150")
  var congregation_stats = new_sheet.getRange("I1:O40")
  var report = HtmlService.createTemplateFromFile("report_content")
  report.publishers_by_groups = publishers_by_groups.getValues()
  report.congregation_stats = congregation_stats.getValues()
  report.title = act_sheet_name
  
  if (action_type == "print") {
    return report.evaluate().getContent()
  } 
  
  if (action_type == "download") {
    var report_blob = report.evaluate().getBlob();
    Logger.log(report.evaluate().getContent())
    var pdf = report_blob.getAs("application/pdf");
    return Utilities.base64Encode(pdf.getBytes())
  }
}

function print_report() {
  var active_sheet_name = source.getActiveSheet().getSheetName()
  var html = HtmlService.createTemplateFromFile("print_dialog")
  html.active_sheet_name = active_sheet_name
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(400).setHeight(165), 'Sprawozdanie ' + active_sheet_name)
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).getContent();
};
