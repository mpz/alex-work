// Голбальные переменные
var token = "TOKEN-FOR-TEST-ONLY";
var shift_per = 9;

// *** Сервисные скрипты

// Примеры переменных для обозначения листов
function sheet_template() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheet_bal = ss.getSheetByName("Баланс");
    var sheet_data = ss.getSheetByName("Данные");
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_pay = ss.getSheetByName("Платежи");
    var sheet_exp = ss.getSheetByName("Экспорт");
    var sheet_war = ss.getSheetByName("Склад");
    var sheet_set = ss.getSheetByName("Настройки");
    var sheet_his = ss.getSheetByName("История");
}

// Подсчёт количества рядов на листе "Закупка"
function number_row_per(num_r_per) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");

    var a;
    var article, order;
    var num_r_per = sheet_per.getRange(9, 65).getValue();

    num_r_per = num_r_per + shift_per;

    for (a = num_r_per; a < 1001; a++) {
        article = sheet_per.getRange(a, 1).getValue();
        order = sheet_per.getRange(a, 2).getValue();
        if ((article != "") || (order != "")) {
            num_r_per = num_r_per + 1;
        } else {
            break;
        }
    }

    sheet_per.getRange(9, 65).setValue(num_r_per - shift_per);

    return num_r_per;
}

// Подсчёт количества рядов на листе "Платежи"
function number_row_pay(num_r_pay) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_pay = ss.getSheetByName("Платежи");

    var a;
    var payment_id, payment_sum;
    var num_r_pay = sheet_pay.getRange(7, 12).getValue();

    num_r_pay = num_r_pay + shift_per;

    for (a = num_r_pay; a < 101; a++) {
        payment_id = sheet_pay.getRange(a, 1).getValue();
        payment_sum = sheet_pay.getRange(a, 5).getValue();
        if ((payment_id != "") || (payment_sum != "")) {
            num_r_pay = num_r_pay + 1;
        } else {
            break;
        }
    }

    sheet_pay.getRange(7, 12).setValue(num_r_pay - shift_per);

    return num_r_pay;
}

// Выявление уникальных номеров заказов
function order_number() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_data = ss.getSheetByName("Данные");
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_set = ss.getSheetByName("Настройки");

    var num_r_per = number_row_per(num_r_per);
    var last_r_order = sheet_set.getRange(32, 9).getValue() + sheet_set.getRange(32, 10).getValue();
    if (last_r_order == "") {
        last_r_order = 2;
    }

    var a;
    var order_first, order_second_last, order_last;
    var order_start, order_from;

    var order_number = 1;
    var orders = [],
        orders_num = [],
        orders_start = [];

    if (last_r_order < num_r_per) {

        // Проверка на наличие записей в колонке "Заказы" и новых товаров в заказах
        order_first = sheet_per.getRange(9, 67).getValue();
        order_second_last = sheet_per.getRange(last_r_order - 1, 2).getValue();
        order_last = sheet_per.getRange(last_r_order, 2).getValue();

        if ((order_first == "") || (order_second_last == order_last)) {
            order_start = shift_per;
            order_from = shift_per - 1;
            sheet_per.getRange(9, 67, 30, 3).clearContent();
        } else {
            order_start = last_r_order;
            last_r_order = shift_per - 1;
            for (a = 1; a < 51; a++) {
                if (sheet_per.getRange(shift_per - 1 + a, 67).getValue() != "") {
                    last_r_order = last_r_order + 1;
                } else {
                    break;
                }
            }
            order_from = last_r_order;
        }

        for (a = order_start; a < num_r_per; a++) {
            order_second_last = sheet_per.getRange(a - 1, 2).getValue();
            order_last = sheet_per.getRange(a, 2).getValue();

            if (order_second_last != order_last) {
                orders.push(order_last);
                orders_num.push(order_number);
                orders_start.push(a);
                order_number = 1;
            } else {
                order_number = order_number + 1;
            }
        }
        orders_num.push(order_number);

        for (a = 1; a < 101; a++) {
            if (orders[a - 1] != undefined) {
                sheet_per.getRange(order_from + a, 67).setValue(orders[a - 1]);
                sheet_per.getRange(order_from + a, 68).setValue(orders_start[a - 1]);
                sheet_per.getRange(order_from + a, 69).setValue(orders_num[a]);
            } else {
                break;
            }
        }
    }
}

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_pay = ss.getSheetByName("Платежи");

    // Основное меню
    var menu_tao = [];

    menu_tao.push({
        name: "",
        functionName: ""
    });
    menu_tao.push({
        name: "",
        functionName: ""
    });
    ss.addMenu("TaoJet", menu_tao);

    // Меню разработчика
    var menu_dev = [];

    menu_dev.push({
        name: "Скопировать дизайн листа 'Данные'",
        functionName: ""
    });
    ss.addMenu("Разработка", menu_dev);

    var check;

    // Проверка на наличие в таблице информации о товарах
    check = sheet_per.getRange(9, 1).getValue();
    if (check != "<%product_article%>") {
        number_row();
        order_number();
    }

    check = sheet_pay.getRange(9, 1).getValue();
    if (check != "<%transfer_id%>") {
        number_row_pay();
    }
}

function onEdit(event) {
    // Добавление дополнительной информации при создании нового платежа на листе "Платёж"
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_pay = ss.getSheetByName("Платежи");
    var sheet_exp = ss.getSheetByName("Экспорт");
    var sheet_set = ss.getSheetByName("Настройки");
    var sheet_his = ss.getSheetByName("История");

    var cell = event.source.getActiveRange();
    var row = cell.getRow();
    var col = cell.getColumn();
    var sheet_name = sheet.getName();

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var payment_users_sum, payment_users_stat, check_c;
    var val;
    var last_r_exp, last_r_his, last_r_his_f;

    // Добавление новых платежей
    if (sheet_name == "Платежи") {
        payment_users_sum = sheet_pay.getRange(3, 3).getValue();
        check_b = sheet_pay.getRange(3, 4).getValue();
        if (payment_users_sum == "") {
            sheet_pay.getRange(3, 3).setValue("Введите сумму");
        } else {
            if ((col == 3) && (row == 3) && (check_b != "Выберите операцию")) {
                add_new_payment();
            }
        }
        if (check_b == "") {
            sheet_pay.getRange(3, 4).setValue("Выберите операцию");
        } else {
            if ((col == 4) && (row == 3) && (payment_users_sum != "Введите сумму")) {
                add_new_payment();
            }
        }

        // Оповещение о неверном вводе платежа
        if ((col == 5) && (row > 9) && (sheet_pay.getRange(row, 4).getValue() == "")) {
            Browser.msgBox("Оповещение", "Для корректного добавления новых платежей воспользуйтесь специальной формой, расположенной вверху страницы.", Browser.Buttons.OK);
            sheet_pay.getRange(row, col).clearContent();
        }

        // Изменение записи на листе "История" (Финансы) при изменении суммы платежа
        if ((col == 5) && (row > 9) && (sheet_pay.getRange(row, 1).getValue() == "x")) {
            check_c = sheet_pay.getRange(row, 12).getValue();
            sheet_his.getRange(check_c, 16).setValue(sheet_pay.getRange(row, col).getValue());
        }
    }

    if (sheet_name == "Закупка") {
        // Выбор номера заказа на листе "Закупка"
        if ((row == 3) && (col == 3)) {
            order_select();
        }
    }

    if (sheet_name == "Закупка") {
        var position = sheet_per.getRange(3, 5).getValue();

        if (col == 30) {
            // Создаёт пометку на листе "История" (Финансы) при проставлении веса
            last_r_his_f = sheet_his.getRange(7, 17).getValue();
            sheet_his.getRange(last_r_his_f + 1, 14).setValue(today);
            sheet_his.getRange(last_r_his_f + 1, 15).setValue("Начисление за доставку заказа номер " + sheet_per.getRange(row, 2).getValue() + ".");
            sheet_his.getRange(last_r_his_f + 1, 16).setValue(sheet_per.getRange(row, 19).getValue());
            sheet_his.getRange(last_r_his_f + 1, 17).setValue("1");
        }

        // Проверка для обсчёта и переноса данных о статусах
        if (col == 33) {

            // Создаёт пометку при изменении статуса
            val = sheet_per.getRange(row, 33).getValue();
            if (val == "Доставлен клиенту") {
                sheet_per.getRange(row, 34).setValue("1,0099");
            } else if ((val == "Товар отсутствует") || (val == "Отсутствует нужный цвет / размер") || (val == "Возврат товара") || (val == "Деньги возвращены")) {
                sheet_per.getRange(row, 34).setValue("1,01");
                // Создаёт пометку на листе "Экспорт"
                last_r_exp = sheet_exp.getLastRow();
                sheet_exp.getRange(last_r_exp + 1, 1).setValue(sheet_per.getRange(row, 1).getValue());
                sheet_exp.getRange(last_r_exp + 1, 2).setValue('Статус товара изменился на "' + val + '".');
                sheet_exp.getRange(last_r_exp + 1, 5).setValue(row);
                sheet_exp.getRange(last_r_exp + 1, 6).setValue("1");
                // Создаёт поментку на листе "История"
                last_r_his = sheet_his.getRange(7, 8).getValue();
                sheet_his.getRange(last_r_his + 1, 5).setValue(sheet_per.getRange(row, 2).getValue());
                sheet_his.getRange(last_r_his + 1, 6).setValue(today);
                sheet_his.getRange(last_r_his + 1, 7).setValue('Статус товара ' + sheet_per.getRange(row, 1).getValue() + ' изменился на "' + val + '".');
                sheet_his.getRange(last_r_his + 1, 8).setValue("1");

                // Создаёт запись на листе "Данные"
                data_mark_no(row);

                // Создаёт пометку на листе "История" (Финансы)
                if (val == "Деньги возвращены") {
                    last_r_his_f = sheet_his.getRange(7, 17).getValue();
                    sheet_his.getRange(last_r_his_f + 1, 14).setValue(today);
                    sheet_his.getRange(last_r_his_f + 1, 15).setValue("Возврат денег за товар.");
                    sheet_his.getRange(last_r_his_f + 1, 16).setValue(sheet_per.getRange(row, 18).getValue());
                    sheet_his.getRange(last_r_his_f + 1, 17).setValue("1");
                }
            } else if (val == "") {
                sheet_per.getRange(row, 34).clearContent();
            } else {
                sheet_per.getRange(row, 34).setValue("1");
            }

            // Если при изменении ячейки имеется формула суммирования (значит, что выбран заказ), запускается скрипт переброски статусов
            if (position != "") {
                // Если выбран заказ и в нём что-то поменялось, то метка удаляется
                sheet_per.getRange(position, 70).clearContent();
                // Если выбран заказ и в нём что-то поменялось, то обновляется время последнего изменения
                sheet_per.getRange(position, 71).setValue(today);

                status_processing_check();
            }
        }
        // Проставление статусов для выбранного заказа
        if ((row == 3) && (col == 9)) {
            // Проверяет выбран ли заказ
            if (position != "") {
                status_change();
            } else {
                sheet_per.getRange(3, 9).setValue("Выберите статус");
                Browser.msgBox("Оповещение", "Пожалуйста, выберите заказ.", Browser.Buttons.OK);
            }
        }
    }
}

// Добавление новых платежей
function payment_new() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_pay = ss.getSheetByName("Платежи");
    var sheet_his = ss.getSheetByName("История");

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var num_r_pay = sheet_pay.getRange(7, 12).getValue();
    var last_r_fin = sheet_his.getRange(7, 17).getValue();

    if (sheet_pay.getRange(shift_per + num_r_pay - 1, 1).getValue() != "") {
        num_r_pay = number_row_pay(num_r_pay);
    }
    sheet_pay.getRange(num_r_pay, 1).setValue("x");
    sheet_pay.getRange(num_r_pay, 4).setValue(today);
    sheet_pay.getRange(num_r_pay, 5).setValue(sheet_pay.getRange(3, 3).getValue());
    sheet_pay.getRange(num_r_pay, 7).setValue(sheet_pay.getRange(3, 4).getValue() + ".");
    sheet_pay.getRange(num_r_pay, 12).setValue(last_r_fin + 1);

    // Добавляет метку на лист "История" (Финансы)
    sheet_his.getRange(last_r_fin + 1, 14).setValue(today);
    sheet_his.getRange(last_r_fin + 1, 15).setValue(sheet_pay.getRange(3, 4).getValue());
    sheet_his.getRange(last_r_fin + 1, 16).setValue(sheet_pay.getRange(3, 3).getValue());
    sheet_his.getRange(last_r_fin + 1, 17).setValue("1");

    sheet_pay.getRange(3, 3).setValue("Введите сумму");
    sheet_pay.getRange(3, 4).setValue("Выберите операцию");
    sheet_pay.getRange(7, 12).setValue(sheet_pay.getRange(7, 12).getValue() + 1);
}

// Выбор заказа
function order_select() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");

    var last_r = sheet_per.getLastRow() - shift_per + 1;
    var check = sheet_per.getRange(3, 3).getValue();
    var order, position, start, number;
    var status;
    var val;
    var a;

    // Открывает все ряды
    sheet_per.showRows(shift_per, last_r);

    // Проверка на содержимое ячейки, где должны быть номера заказа
    if ((check == "Все") || (check == "")) {
        sheet_per.getRange(3, 3).setValue("Все");
        sheet_per.getRange(3, 5).clearContent();
        sheet_per.getRange(7, 23).clearContent();
        sheet_per.getRange(8, 34).clearContent();
    } else {
        for (a = shift_per; a < 59; a++) {
            order = sheet_per.getRange(a, 67).getValue();
            if (order == check) {
                position = sheet_per.getRange(a, 1).getRowIndex();
                sheet_per.getRange(3, 5).setValue(position);
                start = sheet_per.getRange(position, 68).getValue();
                number = sheet_per.getRange(position, 69).getValue();

                // Если это не первый ряд, прячет предыдущие ряды
                if (position != shift_per) {
                    sheet_per.hideRows(shift_per, start - shift_per);
                }

                sheet_per.hideRows(start + number, last_r + shift_per - start - number);
                val = "=SUM(R" + start + ":R" + (start + number - 1) + ")";
                sheet_per.getRange(7, 23).setFormula(val);
                sheet_per.getRange(start, 23).setFormula(val);
                sheet_per.getRange(8, 34).setFormula("=SUM(AH" + start + ":AH" + (start + number - 1) + ")");

                break;
            }

            // Если нет совпадений по номерам заказов, удаляет формулы
            sheet_per.getRange(3, 5).clearContent();
            sheet_per.getRange(7, 23).clearContent();
            sheet_per.getRange(8, 34).clearContent();
        }
    }

    position = sheet_per.getRange(3, 5).getValue();

    // Проверка, выбран ли заказ
    if (position != "") {
        status_processing_check();
    }
}

// Проверка для переноса данных с листа "Закупка" на лист "Экспорт"
function status_processing_check() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");

    var sum = sheet_per.getRange(8, 34).getValue();
    var sum_round = sum.toFixed(2);
    var position = sheet_per.getRange(3, 5).getValue();
    var start = sheet_per.getRange(position, 68).getValue();
    var number = sheet_per.getRange(position, 69).getValue();
    var status = sheet_per.getRange(position, 70).getValue();
    var result = number * 1.01;
    var result_round = result.toFixed(2);

    var mark_a = 0;
    var mark_b = 0;
    var check_status;
    var a;

    // Закрытие заказа
    if ((status != "0") && (status != "1") && (sum_round == result_round)) {
        sheet_per.getRange(position, 70).setValue("1");
        // Отметка, что после переноса данных не надо удалять метки статусов
        mark_a = 1;
    }

    if ((sum >= number) && (status == "")) {
        // --- Тут должен быть блок по созданию записей об изменении статуса
        status_processing();
        Browser.msgBox("Обсчёт и перенос данных");
        // Проверка, есть ли в заказе отсутствующие товары
        if (sum == number) {
            mark_b = 1;
        }
        // Очистка всех меток
        if (mark_a == 0) {
            sheet_per.getRange(start, 34, number).clearContent();
            if (mark_b == 0) {
                // Проставление меток для отсутствующих товаров
                for (a = start; a < start + number; a++) {
                    check_status = sheet_per.getRange(a, 33).getValue();
                    if ((check_status == "Товар отсутствует") || (check_status == "Отсутствует нужный цвет / размер") || (check_status == "Возврат товара") || (check_status == "Деньги возвращены")) {
                        sheet_per.getRange(a, 34).setValue("1,01");
                    }
                }
            }
        }
    }

    // Смена статуса заказа
    // Если все товары отменены
    if (sum == result) {
        sheet_per.getRange(position, 70).setValue("0");
    }
}

// Автоматическое проставление статусов для выбранного заказа
function status_change() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");

    var position = sheet_per.getRange(3, 5).getValue();
    var start = sheet_per.getRange(position, 68).getValue();
    var number = sheet_per.getRange(position, 69).getValue();
    var status = sheet_per.getRange(3, 9).getValue();

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var mark = "1";
    var check;
    var count = 1;
    var a;

    // В зависимости от выбранного статуса задаёт метку
    if ((status == "Товар отсутствует") || (status == "Отсутствует нужный цвет / размер") || (status == "Возврат товара") || (status == "Деньги возвращены")) {
        mark = "1,01";
    } else if (status == "Доставлен клиенту") {
        mark = "1,0099";
    }

    for (a = start; a < start + number; a++) {
        // Проверка на наличие маркера
        check = sheet_per.getRange(a, 34).getValue();
        if (check == "") {
            // Если маркера нет, статус меняется и добавляется нужный маркер
            sheet_per.getRange(a, 33).setValue(status);
            sheet_per.getRange(a, 34).setValue(mark);
            // Счётчик, проверяющий были ли замены
            count = count + 1;
        }
    }

    // Если замены были, происходит обновление даты последнего изменения
    if (count > 1) {
        sheet_per.getRange(position, 71).setValue(today);
    }

    sheet_per.getRange(3, 9).setValue("Выберите статус");

    // Запуск скрипта, отвечающего за проверку и перенос данных
    status_processing_check();
}

function mail_create() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_exp = ss.getSheetByName("Экспорт");
    var sheet_his = ss.getSheetByName("История");

    var html;
    var content, content_send;
    var greet, no, change, final;

    var name = sheet_per.getRange(10, 37).getValue();
    name = name.split(" ")[0];
    var status, url_tao, photo, article, url_order, num, num_all;
    var link = ss.getId();
    link = '<a href="https://docs.google.com/a/taojet.com/spreadsheet/pub?key=' + DocsList.getFileById(link).getId() + '&single=true&gid=0&output=html">Балансу</a>';

    var last_r = sheet_exp.getLastRow();
    var val;
    var check;
    var mark = 0;
    var a;

    val = sheet_exp.getRange(3, 1).getValue();
    greet = '<p>' + val.split("&")[0] + name + val.split("&")[2] + '</p><br>';

    //no = '<p>' + sheet_exp.getRange(4, 1).getValue() + '</p><table align="center">';
    no = '<p>' + sheet_exp.getRange(4, 1).getValue() + '</p><table bgcolor="fadadd">';
    change = '<ul type="none">';

    for (a = 15; a < last_r + 1; a++) {
        article = sheet_exp.getRange(a, 1).getValue();
        status = sheet_exp.getRange(a, 2).getValue();
        status = status.split('"')[1];

        check = sheet_exp.getRange(a, 5).getValue();
        if (check != "") {
            mark = mark + 1;
            url_tao = sheet_per.getRange(check, 3).getFormula();
            url_tao = url_tao.split('"')[1];
            photo = sheet_per.getRange(check, 8).getFormula();
            photo = photo.split('"')[1];
            photo = '<a href="' + url_tao + '"><img src="' + photo + '" width="50" height="50" alt="На Таобао"></a>';
            url_order = sheet_per.getRange(check, 4).getFormula();
            url_order = url_order.split('"')[1];
            url_order = '<a href="' + url_order + '">В заказе</a>';

            //no = no + '<tr><td>' + photo + '</td><td align="center" width="200">' + status +'</td>';
            no = no + '<tr><td width="50"></td><td>' + photo + '</td><td align="center" width="150">' + article + '</td><td align="center" width="100">' + url_order + '</td><td align="center" width="150">' + status + '</td><td width="50"></td>';
        } else {
            num = sheet_exp.getRange(a, 3).getValue();
            num_all = sheet_exp.getRange(a, 4).getValue();

            val = sheet_exp.getRange(7, 1).getValue();
            change = change + '<li>' + val.split("&")[0] + article + val.split("&")[2] + status + val.split("&")[4] + num + val.split("&")[6] + num_all + val.split("&")[8] + '</li>';
        }
    }

    no = no + "</table><br>";

    if (mark == 0) {
        no = "";
        change = '<p>' + sheet_exp.getRange(6, 3).getValue() + change + '</ul><br>';
    } else {
        change = '<p>' + sheet_exp.getRange(6, 1).getValue() + change + '</ul><br>';
    }

    val = sheet_exp.getRange(8, 1).getValue();
    final = '<p>' + val.split("&")[0] + link + val.split("&")[2] + '</p>';

    content = greet + no + change + final;
    sheet_exp.getRange(3, 4).setValue(content);
    content_send = content + '<p style="text-align: center"><button onclick="google.script.run.mail_send()"><img src="http://cs402128.userapi.com/g44571543/a_edc642ac.jpg" width="15" height="15" alt="" style="vertical-align: middle"> Отправить</button></p>';

    html = HtmlService.createHtmlOutput(content_send);
    ss.show(html);
}

function mail_send(content) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_exp = ss.getSheetByName("Экспорт");
    var sheet_his = ss.getSheetByName("История");

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var last_r_his = sheet_his.getRange(7, 8).getValue();

    var email = "manager@taojet.com";
    var content = sheet_exp.getRange(3, 4).getValue();

    MailApp.sendEmail(email, "Оповещение о смене статуса", content, {
        htmlBody: content
    });

    // Создание пометки на листе "История" об отправке письма
    sheet_his.getRange(last_r_his + 1, 6).setValue(today);
    sheet_his.getRange(last_r_his + 1, 7).setValue("Письмо выслано клиенту.");
    sheet_his.getRange(last_r_his + 1, 8).setValue("1");

    // --- Потом удалить эту часть
    Browser.msgBox("Оповещение", "Письмо отправлено на " + email, Browser.Buttons.OK);
}

function status_processing() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_exp = ss.getSheetByName("Экспорт");
    var sheet_his = ss.getSheetByName("История");

    var position = sheet_per.getRange(3, 5).getValue();
    var order = sheet_per.getRange(position, 67).getValue();
    var start = sheet_per.getRange(position, 68).getValue();
    var number = sheet_per.getRange(position, 69).getValue();

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");
    var last_r_exp, last_r_his;

    var v_obr = 0,
        tov_opl = 0,
        voz_tov = 0,
        den_voz = 0,
        pre_dop = 0,
        ozh_otp = 0,
        otp_kit = 0,
        otp_ros = 0,
        otp_kli = 0,
        dos_kli = 0;
    var val;
    var a;

    var symbol, status;

    for (a = start; a < start + number; a++) {
        val = sheet_per.getRange(a, 33).getValue();
        if (val == "В обработке") {
            v_obr = v_obr + 1;
        } else if (val == "Товар оплачен") {
            tov_opl = tov_opl + 1;
        } else if (val == "Возврат товара") {
            voz_tov = voz_tov + 1;
        } else if (val == "Деньги возвращены") {
            den_voz = den_voz + 1;
        } else if (val == "Предоставьте дополнительную информацию") {
            pre_dop = pre_dop + 1;
        } else if (val == "Ожидание отправки от продавца") {
            ozh_otp = ozh_otp + 1;
        } else if (val == "Отправлен на склад в Китае") {
            otp_kit = otp_kit + 1;
        } else if (val == "Отправлен на склад в России") {
            otp_ros = otp_ros + 1;
        } else if (val == "Отправлен клиенту") {
            otp_kli = otp_kli + 1;
        } else if (val == "Доставлен клиенту") {
            dos_kli = dos_kli + 1;
        }
    }

    if (v_obr != 0) {
        symbol = v_obr;
        status = "В обработке";
        record_history(symbol, status);
    }

    if (tov_opl != 0) {
        symbol = tov_opl;
        status = "Товар оплачен";
        record_export(symbol, status);
        record_history(symbol, status);
        record_finance();
    }

    if (voz_tov != 0) {
        symbol = voz_tov;
        status = "Возврат товара";
        record_export(symbol, status);
        record_history(symbol, status);
    }

    if (den_voz != 0) {
        symbol = den_voz;
        status = "Деньги возвращены";
        record_export(symbol, status);
        record_history(symbol, status);
    }

    if (pre_dop != 0) {
        symbol = pre_dop;
        status = "Предоставьте дополнительную информацию";
        record_export(symbol, status);
        record_history(symbol, status);
    }

    if (ozh_otp != 0) {
        symbol = ozh_otp;
        status = "Ожидание отправки от продавца";
        record_export(symbol, status);
        record_history(symbol, status);
    }

    if (otp_kit != 0) {
        symbol = otp_kit;
        status = "Отправлен на склад в Китае";
        record_export(symbol, status);
        record_history(symbol, status);
    }

    if (otp_ros != 0) {
        symbol = otp_ros;
        status = "Отправлен на склад в России";
        record_export(symbol, status);
        record_history(symbol, status);
    }

    if (otp_kli != 0) {
        symbol = otp_kli;
        status = "Отправлен клиенту";
        record_export(symbol, status);
        record_history(symbol, status);
    }

    if (dos_kli != 0) {
        symbol = dos_kli;
        status = "Доставлен клиенту";
        record_history(symbol, status);
    }
}

function record_export(symbol, status) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_exp = ss.getSheetByName("Экспорт");

    var position = sheet_per.getRange(3, 5).getValue();
    var order = sheet_per.getRange(position, 67).getValue();
    var number = sheet_per.getRange(position, 69).getValue();

    var last_r_exp = sheet_exp.getRange(14, 6).getValue();
    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    sheet_exp.getRange(last_r_exp + 1, 1).setValue(order);
    sheet_exp.getRange(last_r_exp + 1, 2).setValue('Изменение статуса на "' + status + '".');
    sheet_exp.getRange(last_r_exp + 1, 3).setValue(symbol);
    sheet_exp.getRange(last_r_exp + 1, 4).setValue(number);
    sheet_exp.getRange(last_r_exp + 1, 6).setValue("1");
}

function record_history(symbol, status) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_his = ss.getSheetByName("История");

    var position = sheet_per.getRange(3, 5).getValue();
    var order = sheet_per.getRange(position, 67).getValue();
    var number = sheet_per.getRange(position, 69).getValue();

    var last_r_his = sheet_his.getRange(7, 8).getValue();
    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    sheet_his.getRange(last_r_his + 1, 5).setValue(order);
    sheet_his.getRange(last_r_his + 1, 6).setValue(today);
    sheet_his.getRange(last_r_his + 1, 7).setValue('Изменение статуса на "' + status + '" (' + symbol + ' из ' + number + ').');
    sheet_his.getRange(last_r_his + 1, 8).setValue("1");
}

function record_finance() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_per = ss.getSheetByName("Закупка");
    var sheet_his = ss.getSheetByName("История");

    var position = sheet_per.getRange(3, 5).getValue();
    var order = sheet_per.getRange(position, 67).getValue();
    var start = sheet_per.getRange(position, 68).getValue();
    var sum = sheet_per.getRange(start, 23).getValue();
    sum = "-" + sum;

    var last_r_his_f = sheet_his.getRange(7, 17).getValue();
    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    sheet_his.getRange(last_r_his_f + 1, 14).setValue(today);
    sheet_his.getRange(last_r_his_f + 1, 15).setValue("Закупка товаров заказа номер " + order + ".");
    sheet_his.getRange(last_r_his_f + 1, 16).setValue(sum);
    sheet_his.getRange(last_r_his_f + 1, 17).setValue("1");
}

// Добавляет пометку на лист "Данные" об отсутствующих товарах
function data_mark_no(row) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_data = ss.getSheetByName("Данные");
    var sheet_per = ss.getSheetByName("Закупка");

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var last_r_no = sheet_data.getRange(11, 1).getValue();

    var row;
    var check;
    var mark = 0;
    var a;

    // Проверка, есть ли уже запись об этом товаре
    for (a = 2; a < last_r_no + 2; a++) {
        check = sheet_data.getRange(10, a).getValue();
        if (row == check) {
            sheet_data.getRange(9, a).setValue(today);
            mark = mark + 1;
            break;
        }
    }

    if (mark == 0) {
        sheet_data.getRange(8, last_r_no + 1).setFormula(sheet_per.getRange(row, 8).getFormula());
        sheet_data.getRange(9, last_r_no + 1).setValue(today);
        sheet_data.getRange(10, last_r_no + 1).setValue(row);
        sheet_data.getRange(11, last_r_no + 1).setValue("1");
    }
}

// Создание на листе "Баланс" зоны с заголовками
function balance_header(row) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_bal = ss.getSheetByName("Баланс");
    var sheet_data = ss.getSheetByName("Данные");

    sheet_bal.getRange(row, 5).setValue("Артикул");
    sheet_bal.getRange(row, 6).setValue("Номер заказа");
    sheet_bal.getRange(row, 7).setValue("Ссылка на товар на Таобао");
    sheet_bal.getRange(row, 8).setValue("Ссылка на товар в заказе");
    sheet_bal.getRange(row, 9).setValue("Фото");
    sheet_bal.getRange(row, 10).setValue("Размер");
    sheet_bal.getRange(row, 11).setValue("Цвет");
    sheet_bal.getRange(row, 12).setValue("Количество");
    sheet_bal.getRange(row, 13).setValue("Сумма с комиссией");
    sheet_bal.getRange(row, 14).setValue("Вес, кг");
    sheet_bal.getRange(row, 15).setValue("Стоимость доставки");
    sheet_bal.getRange(row, 16).setValue("Статус");
    sheet_bal.getRange(row, 17).setValue("Примечание <агента>");
    sheet_bal.getRange(row, 5, 1, 13).setBackground("#ffe599");
    sheet_bal.getRange(row, 5, 1, 13).setFontSize(10);
}

// Создание на листе "Баланс" зоны для отсутствующих товаров
function balance_no() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_bal = ss.getSheetByName("Баланс");
    var sheet_data = ss.getSheetByName("Данные");

    var last_r_data = sheet_data.getRange(11, 1).getValue();
    var shift = 8;

    var position;
    var row;
    var a;

    balance_clear();

    row = 1;
    balance_header(row);

    // Объединяет ячейки - Пока не добаят возможность выбирать цвета границ, не использовать!
    //sheet_bal.getRange(2, 5, 1, 3).merge();
    sheet_bal.getRange(2, 5).setFontSize(10);
    sheet_bal.getRange(2, 5).setValue("Отсутствующие товары");

    for (a = 2; a < last_r_data + 1; a++) {
        position = sheet_data.getRange(10, a).getValue() + shift;
        sheet_bal.getRange(a + 1, 5, 1, 4).setValues(sheet_data.getRange(position, 1, 1, 4).getValues());
        sheet_bal.getRange(a + 1, 9).setFormula("='Данные'!E" + position);
        sheet_bal.getRange(a + 1, 10).setFormula("='Данные'!F" + position);
        sheet_bal.getRange(a + 1, 11).setFormula("='Данные'!G" + position);
        sheet_bal.getRange(a + 1, 12).setFormula("='Данные'!H" + position);
        sheet_bal.getRange(a + 1, 13).setFormula("='Данные'!I" + position);
        sheet_bal.getRange(a + 1, 14).setFormula("='Данные'!J" + position);
        sheet_bal.getRange(a + 1, 15).setFormula("='Данные'!K" + position);
        sheet_bal.getRange(a + 1, 16).setFormula("='Данные'!L" + position);
        sheet_bal.getRange(a + 1, 17).setFormula("='Данные'!M" + position);
        sheet_bal.getRange(a + 1, 5, 1, 13).setBackground("#fadadd");
    }

    sheet_data.getRange(16, 19).setValue(1 + last_r_data);

    balance_last_order();
}

function balance_last_order() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_bal = ss.getSheetByName("Баланс");
    var sheet_data = ss.getSheetByName("Данные");
    var sheet_set = ss.getSheetByName("Настройки");

    var shift = 8;
    var begin = sheet_data.getRange(16, 19).getValue() + 2;
    var start = sheet_set.getRange(32, 9).getValue() + shift;
    var number = sheet_set.getRange(32, 10).getValue();
    var finish = start + number - 1;

    sheet_bal.getRange(begin, 5).setFontSize(10);
    sheet_bal.getRange(begin, 5).setValue("Последний заказ");
    sheet_bal.getRange(begin + 1, 5).setFormula("=SORT('Данные'!A" + start + ":M" + finish + ";2;0)");

    sheet_data.getRange(16, 19).setValue(begin + number);

    if ((start - shift) != shift_per) {
        balance_other();
    }
}

function balance_other() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_bal = ss.getSheetByName("Баланс");
    var sheet_data = ss.getSheetByName("Данные");
    var sheet_set = ss.getSheetByName("Настройки");

    var shift = 8;
    var begin = sheet_data.getRange(16, 19).getValue() + 3;
    var row = begin - 1;

    var start = shift_per + shift;
    var finish = sheet_set.getRange(32, 9).getValue() + shift - 1;

    balance_header(row);

    sheet_bal.getRange(begin, 5).setFontSize(10);
    sheet_bal.getRange(begin, 5).setValue("Остальные заказы");
    sheet_bal.getRange(begin + 1, 5).setFormula("=SORT('Данные'!A" + start + ":M" + finish + ";2;0)");
}

function balance_clear() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_bal = ss.getSheetByName("Баланс");
    var sheet_data = ss.getSheetByName("Данные");

    var last_r = sheet_bal.getLastRow();

    // Разделяет ячейки
    //sheet_bal.getRange(1, 5, last_r, 13).breakApart();
    // Закрашивает границы белым
    sheet_bal.getRange(1, 5, last_r, 13).setBorder(top, left, bottom, right, vertical, horizontal)
    // Сброс цвета ячеек
    sheet_bal.getRange(1, 5, last_r, 13).setBackground("white");
    // Установка размера шрифтов
    sheet_bal.getRange(1, 5, last_r, 1).setFontSize(8);
    sheet_bal.getRange(1, 10, last_r, 2).setFontSize(8);
    sheet_bal.getRange(1, 17, last_r, 1).setFontSize(8);
    // Очищает ячейку с количеством рядов на листе "Данные"
    sheet_data.getRange(16, 19).clearContent();
    // Очищает ячейки
    sheet_bal.getRange(1, 5, last_r, 13).clearContent();
}
