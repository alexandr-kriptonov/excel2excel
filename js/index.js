var fs = require("fs");
var path = require('path');
var fileEntry = undefined;
var fileString = undefined;
var inputFileObject = undefined;
var convertedFile = undefined;
var text_log = undefined;
var _current_dir = process.cwd();
var _settings = undefined;
var openButton, convertButton, saveButton, addSettingButton;
var ProgressBar = undefined;
var settings_filename = path.join(_current_dir, "settings.txt");
var outputWB = undefined;
global.window.nwDispatcher.requireNwGui();
var gui = require('nw.gui');
var window = gui.Window.get();
var table_width = '320px'
var global_table_widght=String(parseInt(table_width.replace('px', '')) + 40) + 'px'
var col_width = '150px';
var col_input_width='148px';
var col_but_width='30px';
var XLSX = require('xlsx');

function loadingIndicator(value) {
  if(value == true){
    $('#ajaxBusy').show();
    // NProgress.start();
    document.getElementById('open').disabled = true;
    document.getElementById('convert').disabled = true;
    document.getElementById('save').disabled = true;
  }
  else{
      $('#ajaxBusy').hide();
      // NProgress.done();
      document.getElementById('open').disabled = false;
      document.getElementById('convert').disabled = false;
      document.getElementById('save').disabled = false;
      onload();
  }
}

function toLog(data){
  text_log.append(String(data) + "\r\n");
  console.log(String(data));
}

function addToSettings (argument) {
  var _in = document.getElementById('_in').value;
  var _out = document.getElementById('_out').value;
  if(_in != "" && _out != ""){
    el_to_settings = {}
    el_to_settings['in'] = String(_in);
    el_to_settings['out'] = String(_out);
    _settings.push(el_to_settings);
    toLog('В настройки добавлена новая запись: ' + _in + ' ---> ' + _out);
    $("#_in").val("");
    $("#_out").val("");
    saveSettings();
    readSettings();
  }
  else{
    alert("Значения IN и/или OUT не должны быть пустыми!");
  }
}

function deleteFromSettings(){
  id = this._id;
  toLog('В настройках удалена запись: ' + _settings[id]['in'] + ' ---> ' + _settings[id]['out']);
  _settings.splice(id,1);
  saveSettings();
  readSettings();
}

function updateSettingsTable() {
  var table_settings_content=document.getElementById('table_settings_content');
  table_settings_content.innerHTML = "";
  var tbl1=document.createElement('table');
  tbl1.setAttribute('id','table_header');
  tbl1.style.width='300px';
  tbl1.setAttribute('cellpadding','0');
  tbl1.setAttribute('cellspacing','0');
  var ttr=document.createElement('tr');
  var ttd1=document.createElement('td');
  ttd1.setAttribute('class', 'td_input_header');
  ttd1.style.width='140px';
  ttd1.appendChild(document.createTextNode('Колонки в вых.файле'))
  ttr.appendChild(ttd1)
  var ttd2=document.createElement('td');
  ttd2.setAttribute('class', 'td_input_header');
  ttd2.style.width='140px';
  ttd2.appendChild(document.createTextNode('Колонки в исх.файле'))
  ttr.appendChild(ttd2)
  var ttd3=document.createElement('td');
  ttd3.setAttribute('class', 'td_input_header');
  ttd3.style.width='20px';
  ttd3.appendChild(document.createTextNode('<==>'))
  ttr.appendChild(ttd3)
  tbl1.appendChild(ttr);

  var tr=document.createElement('tr');
  var td1=document.createElement('td');
  td1.setAttribute('class', 'td_input');
  td1.style.width='140px';
  _input = document.createElement('input')
  _input.setAttribute('id', '_out');
  _input.style.width='148px';
  td1.appendChild(_input);
  tr.appendChild(td1);
  var td2=document.createElement('td');
  td2.setAttribute('class', 'td_input');
  td2.style.width='140px';
  _input = document.createElement('input')
  _input.setAttribute('id', '_in');
  _input.style.width='148px';
  td2.appendChild(_input);
  tr.appendChild(td2);
  var td3=document.createElement('td');
  td3.setAttribute('class', 'td_input');
  td3.style.width='20px';
  var _but=document.createElement('button');
  _but.setAttribute('id', 'add_setting');
  var _img=document.createElement('img');
  _img.setAttribute('src','icons/add.png');
  _img.setAttribute('title','Нажмите чтобы добавить настройку.');
  _but.appendChild(_img);
  _but.onclick=addToSettings;
  td3.appendChild(_but);
  tr.appendChild(td3);
  tbl1.appendChild(tr);

  var div=document.createElement('div');
  div.style.width=global_table_widght;
  div.style.overflow='auto';
  div.style.height='100px';
  var tbl2=document.createElement('table');
  tbl2.setAttribute('id','table_settings');
  tbl2.style.width=table_width;
  tbl2.setAttribute('cellpadding','0');
  tbl2.setAttribute('cellspacing','0');
  if(_settings != undefined){
    for(var i=0;i<_settings.length;i++){
      var div_string = document.createElement('div');
      div_string.setAttribute('id','string_'+i);
      div_string.style.width=table_width;
      var tr=document.createElement('tr');
      var td1=document.createElement('td');
      td1.style.width=col_width;
      _input = document.createElement('input')
      _input.style.width=col_input_width;
      _input.value=String(_settings[i]['out'])
      _input.disabled = true;
      td1.appendChild(_input);
      tr.appendChild(td1);
      var td2=document.createElement('td');
      td2.style.width=col_width;
      _input = document.createElement('input')
      _input.style.width=col_input_width;
      _input.value=String(_settings[i]['in'])
      _input.disabled = true;
      td2.appendChild(_input);
      tr.appendChild(td2);
      var td3=document.createElement('td');
      td3.style.width=col_but_width;
      var _but=document.createElement('button');
      var _img=document.createElement('img');
      _img.setAttribute('src','icons/delete.png');
      _img.setAttribute('title','Нажмите чтобы удалить настройку.');
      _but.appendChild(_img);
      _but.onclick=deleteFromSettings;
      _but._id=i;
      td3.appendChild(_but);
      tr.appendChild(td3);
      div_string.appendChild(tr);
      tbl2.appendChild(div_string);
    }
    div.appendChild(tbl2);
    table_settings_content.appendChild(tbl1);
    table_settings_content.appendChild(div);
  }
}

function readSettings(){
  _settings = undefined;
  try {
    data = fs.readFileSync(settings_filename);
    var raw_settings = String(data);
    toLog("Настройки загружены.");
    _settings = JSON.parse(raw_settings);
    updateSettingsTable();
  } catch (e) {
    toLog("Read settings failed: " + e);
    return;
  }
}

function saveSettings(){
  fs.writeFileSync(settings_filename, JSON.stringify(_settings, "", 4), {flag: "w"}, function (err) {
    if (err) {
      toLog("Write failed: " + err);
      return;
    }
    toLog("Настройки сохранены");
  });
}

function handleOpenButton() {
  $("#openFile").trigger("click");
}

function handleSaveButton() {
  $("#saveFile").trigger("click");
}

function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
  for(var R = 0; R != data.length; ++R) {
    for(var C = 0; C != data[R].length; ++C) {
      if(range.s.r > R) range.s.r = R;
      if(range.s.c > C) range.s.c = C;
      if(range.e.r < R) range.e.r = R;
      if(range.e.c < C) range.e.c = C;
      var cell = {v: data[R][C] };
      if(cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
      
      if(typeof cell.v === 'number') cell.t = 'n';
      else if(typeof cell.v === 'boolean') cell.t = 'b';
      else if(cell.v instanceof Date) {
        cell.t = 'n'; cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      }
      else cell.t = 's';
      ws[cell_ref] = cell;
    }
  }
  if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function(sheetName) {
        var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if(roa.length > 0){
            result[sheetName] = roa;
        }
    });
    fs.writeFileSync(path.join(_current_dir, "set.txt"), JSON.stringify(result, "", 4), {flag: "w"}, function (err) {
      if (err) {
        toLog("Write failed: " + err);
      }
    });
    json_to_xls_obj(result["Список призывников"]);
}

function json_to_xls_obj(input_data){
  var sheet_data = input_data;
  var input_header = [];
  var output_header = [];
  for(var k in sheet_data[0]) input_header.push(k);
  for(var k in _settings) output_header.push(_settings[k]["out"]);
  var output_data = [];
  output_data.push(output_header);

  var regex = new RegExp('#"[A-Za-zА-Яа-я0-9_]+"#');

  for(var i=0; i<sheet_data.length-1; i++){
    var line_data = sheet_data[i];
    var output_line = [];
    for(var k=0; k<_settings.length; k++){
      if(input_header.indexOf(_settings[k]["in"]) >= 0){
        output_line.push(line_data[_settings[k]["in"]]);
      }
      else if(_settings[k]["in"] === "#NOTHING#"){
        output_line.push("NOTHING");
      }
      else if(regex.test(_settings[k]["in"])){
        output_line.push(_settings[k]["in"].replace(/#/g, '').replace(/"/g, ''));
      }
      else{output_line.push("");}
      }
    output_data.push(output_line);
    }
  completeConvert(output_data);
  loadingIndicator(false);
}

function completeConvert(data){
  var ws_name = "SheetJS";

  function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
  }

  var wb = new Workbook(), ws = sheet_from_array_of_arrays(data);

  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;
  outputWB = wb;
  toLog(JSON.stringify(outputWB, "", 4));
}

function handleConvertButton() {
  if(fileEntry != undefined){
    toLog("Начинаем конвертирование файла...");
    loadingIndicator(true);
    var workbook = XLSX.readFile(fileEntry);
    to_json(workbook);
    toLog("Конвертирование файла завершено! Сохраните файл, чтобы не потерять изменения.");
  }
  else{
    alert("Для начала конвертирования нужно загрузить файл!");
  }
}

function saveExcelFile (fileName) {
  if(outputWB != undefined){
    XLSX.writeFile(outputWB, fileName);
    alert("Файл: " + String(fileName) +" сохранён!")
  }
  else{
    alert("Сначала конвертируйте файл!");
  }
}

var onChosenFileToOpen = function(theFileEntry) {
  convertedFile = undefined;
  fileEntry = theFileEntry;
  toLog("Файл загружен: " + String(fileEntry));
};

var onChosenFileToSave = function(fileName) {
  saveExcelFile(fileName);
};

function readFileIntoString(theFileEntry){
  fileString = undefined;
  loadingIndicator(true);
  try {
    data = fs.readFileSync(theFileEntry);
    fileString = String(data);
  } catch (e) {
    toLog("Read failed: " + e);
  }
  loadingIndicator(false);
}

window.onload = function() {
  openButton = document.getElementById("open");
  convertButton = document.getElementById("convert");
  saveButton = document.getElementById("save");
  text_log = $("#text_log");
  text_log.attr("disabled", "disabled");
  openButton.addEventListener("click", handleOpenButton);
  convertButton.addEventListener("click", handleConvertButton);
  saveButton.addEventListener("click", handleSaveButton);
  $("#openFile").change(function(evt) {
    onChosenFileToOpen($(this).val());
  });
  $("#saveFile").change(function(evt) {
    onChosenFileToSave($(this).val());
  });
  $('#ajaxBusy').css({
    display:"none",
    margin:"0px",
    paddingLeft:"0px",
    paddingRight:"0px",
    paddingTop:"0px",
    paddingBottom:"0px",
    position:"absolute",
    margin: "0 25% 0 25%",
     width:"50%"
  });
  readSettings();

  require("nw.gui").Window.get().show();
}
