function Main() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("answer form").activate(); //активируем лист с данными пользователей
  var list = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //записываем лист в list
  var numlaststr = list.getLastRow(); //берём числовой порядок последней строки

  for (i = 2; i <= numlaststr; ++i) { // читаем строки начиная со второй и до последней
    if (!list.getRange(i, 8).isBlank() && !list.getRange(i, 6).isBlank()) { // проверяем, что у нас в статусе "+" — соббщение   
      continue                                                                          //уже было отправлено
    } else if (list.getRange(i, 1).isBlank()){
      numlaststr = numlaststr - VoidRemoval(list, numlaststr, i);
    } else {
      var person_name = list.getRange(i, 2).getValue(); //проходимся по строке i и берём имя из столбца(тут: 2-й)
      var person_email = list.getRange(i, 3).getValue(); //берём адресс(в данном случае 3-й столбец)
      SendMessage(person_name, person_email); //отправляем полученное имя и адрес в функцию отправки сообщения
      list.getRange(i, 4).setValue('+');
    }
  } 
}


function ConcatenationOfString(list, firstemptystr, firstfiledstr, numlaststr) { //функция сдвига строк вместе
  range = list.getRange(firstfiledstr, 1, numlaststr, 4);
  result = range.getValues();
  range.offset((firstemptystr - firstfiledstr), 0).setValues(result)
}


function VoidRemoval(list, numlaststr, firstemptystr) {
  endfiledstr = 0;
  flag = false;
  for (j = firstemptystr + 1; j <= numlaststr; ++j) { //мы знаем первую пустую строку из Main, начинаем считать сразу со следующей
    if (!list.getRange(j, 1).isBlank() && !flag) {
      endfiledstr = j;
      flag = true;
      continue
    } else if (list.getRange(j, 1).isBlank() && !flag) { //предполагается, что если пусто в первой колонке, то пусто везде
      continue
    } else if (!list.getRange(j, 1).isBlank() && flag) {
      continue
    } else {
      break;
    }
  } 
ConcatenationOfString(list, firstemptystr, endfiledstr, numlaststr);
return endfiledstr - firstemptystr; //последняя строка у нас сместилась — возвращаем разницу
}


function SendMessage(name, emale) {
  var shublon = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("shublon").getRange(1, 1).getValue(); //берём шаблон сообщения

  var message_text = shublon.replace('{name}', name);   //заменяем в шаблоне "{name}" на то что храниться в переменной name
  MailApp.sendEmail(emale, "Туса у пампуса", message_text); //отправляем сообщение на email с темой "test topic" и текстом message_text
  Logger.log(message_text); //Это логер для отладки скрипта
}



function onEdit(e) { // триггер, если что-то удаляется в первой колонке, то мы считаем, что удалению подлежит вся строка
  var range = e.range;
  if (range.getColumn() == 1 && range.getValue() == '') {
    VoidRemoval(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow(), range.getRow())
  }
}


