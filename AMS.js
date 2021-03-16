function SendMessage(name, emale) {
  var shublon = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("templ").getRange(1, 1).getValue(); //берём шаблон сообщения
  var message_text = shublon.replace('{name}', name);   //заменяем в шаблоне "{name}" на то что храниться в переменной name
  MailApp.sendEmail(emale, "test topic", message_text); //отправляем сообщение на email с темой "test topic" и текстом message_text
  Logger.log(message_text); //Это логер для отладки скрипта
}
 
 
function MySkriptTest() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("answer form").activate(); //активируем лист с данными пользователей
  var list = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //записываем лист в list
  var kolstr = list.getLastRow(); //берём числовой порядок последней строки
 
  for (i = 2; i <= kolstr; ++i) { // читаем строки начиная со второй и до последней
    if (list.getRange(i, 4).getValue() != '') { // проверяем, что у нас в статусе не пустая строка — сообщение уже было отправлено
      continue
    } else {
      var person_name = list.getRange(i, 2).getValue(); //проходимся по строке i и берём имя из столбца(тут: 2-й)
      var person_email = list.getRange(i, 3).getValue(); //берём адресс(в данном случае 3-й столбец)
      SendMessage(person_name, person_email); //отправляем полученное имя и адрес в функцию отправки сообщения
      list.getRange(i, 4).setValue('+');
    }
 
  } 
}
