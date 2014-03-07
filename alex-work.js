// Голбальные переменные
var token = "TOKEN-FOR-TEST-ONLY";
var shift_per = 9;
var shift_pay = 9;
var shift_exp = 15;
var shift_data = 8;

// Переменные для листов

//
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheets = {
    act: ss.getActiveSheet(), // Текущий лист
    bal: ss.getSheetByName("Баланс"), // Лист "Баланс"
    data: ss.getSheetByName("Данные"), // Лист "Данные"
    per: ss.getSheetByName("Закупка"), // Лист "Закупка"
    pay: ss.getSheetByName("Платежи"), // Лист "Платежи"
    exp: ss.getSheetByName("Экспорт"), // Лист "Зкспорт"
    war: ss.getSheetByName("Склад"), // Лист "Склад"
    set: ss.getSheetByName("Настройки"), // Лист "Настройка"
    his: ss.getSheetByName("История")
}; // Лист "История"
//

// Переменные на листе "Настройка"
//
var set = {
    last_pos: sheets.set.getRange(32, 9), // Начальная позиция (ряд) последнего заказа
    last_num: sheets.set.getRange(32, 10)
}; // Количество товаров в последнем заказе
//

// Переменные на листе "Закупка"
//
var per = {
    pos: sheets.per.getRange(3, 5), // Ячейка с номером ряда, с которого начинается заказ
    ord_user: sheets.per.getRange(3, 3), // Ячейка для выбора заказа
    stat_user: sheets.per.getRange(3, 9), // Ячейка для выбора статуса заказа
    art_first: sheets.per.getRange(shift_per, 1), // Первый артикул товара
    ord_cost: sheets.per.getRange(7, 23), // Ячейка со стоимостью выбранного заказа
    stat_check: sheets.per.getRange(8, 34), // Ячейка для проверки изменений статусов в заказе
    cli_name: sheets.per.getRange(9, 37), // Ячейка с именем клиента
    cki_mail: sheets.per.getRange(9, 41), // Ячейка с электронной почтой клиента
    num_r: sheets.per.getRange(9, 65), // Ячейка с количеством заполненных рядов
    ord_all: sheets.per.getRange(9, 67, 30, 3), // Диапазон с данными по заказам (уникальные номера, начальные позиции, количество товаров)
    // Колонки
    article: 1, // Колонка с артикулами
    order: 2, // Колонка с номерами заказов
    url_tao: 3, // Колонка с ссылкой на товар на Таобао
    url_order: 4, // Колонка с ссылкой на заказ
    photo: 8, // Колонка с фотографией товара
    summ_com: 18, // Колонка с суммой с комиссией
    delivery: 19, // Колонка со стоймостью доставки
    order_cost: 23, // Колонка со стоимостью заказов
    weight: 30, // Колонка с весом
    status: 33, // Колонка со статусами
    status_mark: 34, // Колонка с маркерами статусов
    order_list: 67, // Колонка с уникальными номерами заказов
    order_pos: 68, // Колонка с начальными позициями (рядами) заказов
    order_num: 69, // Колонка с количеством товаров в заказе 
    order_mark: 70, // Колонка с маркером статуса заказа
    order_date: 71
}; // Колонка с датой последнего изменения заказа
//

// Переменные на листе "Платёж"
//
var pay = {
    id_first: sheets.pay.getRange(shift_pay, 1), // Первая айдишка платежа
    num_r: sheets.pay.getRange(7, 12), // Ячейка с количеством заполненных рядов
    sum_user: sheets.pay.getRange(3, 3), // Ячейка для внесения суммы в новый пользовательский платёж
    oper_user: sheets.pay.getRange(3, 4), // Ячейка для выбора операции в новом пользовательском платеже
    // Колонки
    id: 1, // Колонка с айдишкой платежа
    date: 4, // Дата платежа
    sum: 5, // Сумма платежа
    operation: 7, // Операция
    history_row: 12
}; // Пометка о положении записи об этом платеже на листе "История" (Финансы)
//

//
var his = {
    last_his: sheets.his.getRange(7, 8), // Количество записей на листе "История"
    last_fin: sheets.his.getRange(7, 17), // Количество записей финансовом блоке на листе "История"
    // Колонки
    whose: 5, // Колонка с информацией о заказе или артикле
    date_his: 6, // Колонка с датами
    text_his: 7, // Колонка с основным текстом
    counter_fin: 8, // Колонка с счётчиком рядов
    date_fin: 14, // Колонка с датами (финансовый блок)
    text_fin: 15, // Колонка с основным текстом (финансовый блок)
    summ: 16, // Колонка с суммами 
    counter_fin: 17
}; // Колонка с счётчиком рядов (финансовый блок)
//

//
var exp = {
    last_r: sheets.exp.getRange(14, 6), // Ячейка с количеством записей на листе "Экспорт"
    // Колонки
    whose: 1, // Колонка с информацией о заказе или артикуле
    text: 2, // Колонка с текстовым описанием события
    number: 3, // Колонка с количеством товаров с текущим статусом
    number_all: 4, // Колонка с общим количеством товаров в заказе
    mark: 5
}; // Колонка с маркерами (для разделения записей об отсутствующих товарах и смене статусов)
//

//
var data = {
    num_no: sheets.data.getRange(11, 1)
}; // Количество записей об отсутствующих товарах
//

// *** Сервисные скрипты

// Подсчёт количества заполненных рядов на листах "Закупка" и "Платежи" -------------------------------------------- Работает
function number_row(sheet) {
    var a;

    //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Платежи"); // ------------------------------ Проверка

    if (sheet.getName() == sheets.per.getName()) {
        num_r = per.num_r.getValue() + shift_per;
    } else {
        num_r = pay.num_r.getValue() + shift_per;
    }

    for (a = num_r; a < 1001; a++) {
        if ((sheet.getRange(a, 1).getValue() != "") || (sheet.getRange(a, 2).getValue() != "")) {
            num_r = num_r + 1;
        } else {
            break;
        }
    }

    if (sheet.getName() == sheets.per.getName()) {
        per.num_r.setValue(num_r - shift_per);
    } else {
        pay.num_r.setValue(num_r - shift_per);
    }

    return num_r;
}

// Выявление уникальных номеров заказов на листе "Закупка" ------------------------------- Работает
function order_number() {
    var sheet = sheets.per;
    var num_r = number_row(sheet);

    var a;
    var order_second_last, order_last; // Номера заказов у предпоследнего и последнего товаров
    var orders = []; // Массив для учёта данных по уникальным заказам
    var order_position; // Позиция (ряд), к которого начинается заказ
    var order_number = 1; // Количество товаров в заказе

    var orders_len = 0; // Длинна массива с номерами уникальных заказов
    var last_r_order = 0; // Последний учтённый ряд с информацией по товарам

    // Вычисление длинны массива с номерами уникальных заказов
    for (a = shift_per; a < 60; a++) {
        if (sheets.per.getRange(a, per.order_list).getValue() != "") {
            orders_len++;
        } else {
            break;
        }
    }

    if (orders_len != 0) {
        last_r_order = sheets.per.getRange(shift_per + orders_len - 1, per.order_pos).getValue() + sheets.per.getRange(shift_per + orders_len - 1, per.order_num).getValue();
    }

    // Проверка на наличие записи о выявленных заказах
    if (last_r_order == 0) {
        last_r_order = shift_per;
        per.ord_all.clearContent();
    }

    if (last_r_order < num_r) {
        // Проверка на наличие новых товаров в последнем заказе
        order_second_last = sheets.per.getRange(last_r_order - 1, per.order).getValue();
        order_last = sheets.per.getRange(last_r_order, per.order).getValue();

        if ((order_second_last == order_last) || (orders_len == 0)) {
            per.ord_all.clearContent();
            last_r_order = shift_per;
        } else {
            orders = sheets.per.getRange(shift_per, per.order_list, orders_len, 3).getValues();
        }

        // Заполнение массива данными
        for (a = last_r_order; a < num_r; a++) {
            order_second_last = sheets.per.getRange(a, per.order).getValue();
            order_last = sheets.per.getRange(a + 1, per.order).getValue();

            if (order_second_last == order_last) {
                if (order_number == 1) {
                    order_position = a;
                }
                order_number++;
            } else {
                if (order_number == 1) {
                    order_position = a;
                }
                orders.push([order_second_last, order_position, order_number]);
                order_number = 1;
            }
        }

        // Выгрузка данных об уникальных заказах из массива в таблицу
        orders_len = orders.length;
        sheets.per.getRange(shift_per, per.order_list, orders_len, 3).setValues(orders);
    }
}

// Срабатывает при открытии или обновлении таблицы
function onOpen() {
    // Основное меню
    var menu_tao = [];
    menu_tao.push({
        name: "Загрузить данные на лист 'Баланс'",
        functionName: "balance_no_product"
    });
    menu_tao.push({
        name: "",
        functionName: ""
    });
    ss.addMenu("TaoJet", menu_tao);

    // Меню разработчика
    /*var menu_dev = [];
  menu_dev.push({ name: "Загрузить данные на лист 'Баланс'", functionName: "balance_no_product" });
  ss.addMenu("Разработка", menu_dev);*/

    var sheet;

    // Проверка на наличие в таблице информации о товарах
    if (per.art_first.getValue() != "<%product_article%>") {
        sheet = sheets.per;
        number_row(sheet);
        order_number();
    }

    // Проверка на наличие в таблице информации о платежах
    if (pay.id_first.getValue() != "<%transfer_id%>") {
        sheet = sheets.pay;
        number_row(sheet);
    }
}

// Срабатывает при изменении содержимого ячеек таблицы
function onEdit(event) {
    // Добавление дополнительной информации при создании нового платежа на листе "Платёж"


    var cell = event.source.getActiveRange();
    var row = cell.getRow();
    var col = cell.getColumn();
    var sheet_name = sheets.act.getName();

    var last_r_exp;

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var check;
    var arr, text, summ, whose; // Переменные для работы с рядами
    var counter = 1;

    // Скрипты на листе "Платежи"
    if (sheet_name == "Платежи") {
        // Добавление новых платежей
        if (pay.sum_user.getValue() == "") {
            pay.sum_user.setValue("Введите сумму");
        } else {
            if ((col == 3) && (row == 3) && (pay.oper_user.getValue() != "Выберите операцию")) {
                payment_new();
            }
        }
        if (pay.oper_user.getValue() == "") {
            pay.oper_user.setValue("Выберите операцию");
        } else {
            if ((col == 4) && (row == 3) && (pay.sum_user.getValue() != "Введите сумму")) {
                payment_new();
            }
        }

        // Оповещение о неверном вводе платежа
        if ((col == pay.sum) && (row >= shift_pay) && (sheets.pay.getRange(row, pay.date).getValue() == "")) {
            Browser.msgBox("Оповещение", "Для корректного добавления новых платежей воспользуйтесь специальной формой, расположенной вверху страницы.", Browser.Buttons.OK);
            sheets.pay.getRange(row, col).clearContent();
        }

        // Изменение записи на листе "История" (Финансы) при изменении суммы платежа
        if ((col == pay.sum) && (row >= shift_pay) && (sheets.pay.getRange(row, pay.id).getValue() == "x")) {
            check = sheets.pay.getRange(row, pay.history_row).getValue();
            sheets.his.getRange(check, his.summ).setValue(sheets.pay.getRange(row, col).getValue());
        }
    }

    // Скрипты на листе "Закупка"
    if (sheet_name == "Закупка") {
        // Выбор номера заказа на листе "Закупка"
        if ((row == 3) && (col == 3)) {
            order_select();
        }

        if (col == per.weight) {
            // Создаёт запись на листе "История" (Финансы) при проставлении веса
            text = "Начисление за доставку заказа номер " + sheets.per.getRange(row, per.order).getValue() + ".";
            summ = sheets.per.getRange(row, per.delivery).getValue();

            arr = [];
            arr.push([today, text, summ, counter]);
            sheets.his.getRange(his.last_fin.getValue() + 1, 14, 1, 4).setValues(arr);
        }

        // Проверка для обсчёта и переноса данных о статусах
        if (col == per.status) {
            // Создаёт пометку при изменении статуса
            check = sheets.per.getRange(row, per.status).getValue();
            if (check == "Доставлен клиенту") {
                sheets.per.getRange(row, per.status_mark).setValue("1,0099");
            } else if ((check == "Товар отсутствует") || (check == "Отсутствует нужный цвет / размер") || (check == "Возврат товара") || (check == "Деньги возвращены")) {
                sheets.per.getRange(row, per.status_mark).setValue("1,01");
                // Создаёт пометку на листе "Экспорт"
                last_r_exp = sheets.exp.getLastRow();

                whose = sheets.per.getRange(row, per.article).getValue();
                text = 'Статус товара изменился на "' + check + '".';

                arr = [];
                arr.push([whose, text]);
                sheets.exp.getRange(exp.last_r.getValue() + 1, 1, 1, 2).setValues(arr);

                arr = [];
                arr.push([row, counter]);
                sheets.exp.getRange(exp.last_r.getValue() + 1, 5, 1, 2).setValues(arr);

                // Создаёт запись об отсутсвующем товаре на листе "Данные"
                data_mark_no(row);

                // Создаёт поментку на листе "История"
                whose = sheets.per.getRange(row, per.order).getValue();
                text = 'Статус товара ' + sheets.per.getRange(row, per.article).getValue() + ' изменился на "' + check + '".';

                arr = [];
                arr.push([whose, today, text, counter]);
                sheets.his.getRange(his.last_his.getValue() + 1, 5, 1, 4).setValues(arr);

                // Создаёт пометку на листе "История" (Финансы)
                if (check == "Деньги возвращены") {
                    text = "Возврат денег за товар.";
                    summ = sheets.per.getRange(row, per.summ_com).getValue();

                    arr = [];
                    arr.push([today, text, summ, counter]);
                    sheets.his.getRange(his.last_fin.getValue() + 1, 14, 1, 4).setValues(arr);
                }
            } else if (check == "") {
                sheets.per.getRange(row, per.status_mark).clearContent();
            } else {
                sheets.per.getRange(row, per.status_mark).setValue("1");
            }

            // Если при изменении ячейки имеется формула суммирования (значит, что выбран заказ), запускается скрипт переброски статусов
            if (per.pos.getValue() != "") {
                // Если выбран заказ и в нём что-то поменялось, то метка удаляется
                sheets.per.getRange(per.pos.getValue(), per.order_mark).clearContent();
                // Если выбран заказ и в нём что-то поменялось, то обновляется время последнего изменения
                sheets.per.getRange(per.pos.getValue(), per.order_date).setValue(today);

                status_processing_check();
            }
        }
        // Проставление статусов для выбранного заказа
        if ((row == 3) && (col == 9)) {
            // Проверяет выбран ли заказ
            if (per.pos.getValue() != "") {
                status_change();
            } else {
                per.stat_user.setValue("Выберите статус");
                Browser.msgBox("Оповещение", "Пожалуйста, выберите заказ.", Browser.Buttons.OK);
            }
        }
    }
}

// Добавление новых платежей ------------------------------------------------------------ Рабочий
function payment_new() {
    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var sheet = sheets.pay;
    var num_r = number_row(sheet);

    var arr, text, summ; // Переменные для работы с рядами
    var counter = 1;

    // Добавляет запись о пользовательском платеже на лист "Платежи"
    sheets.pay.getRange(num_r, pay.id).setValue("x");
    sheets.pay.getRange(num_r, pay.date).setValue(today);
    sheets.pay.getRange(num_r, pay.sum).setValue(pay.sum_user.getValue());
    sheets.pay.getRange(num_r, pay.operation).setValue(pay.oper_user.getValue() + ".");
    sheets.pay.getRange(num_r, pay.history_row).setValue(his.last_fin.getValue() + 1);

    // Добавляет запись на лист "История" (Финансы)
    text = pay.oper_user.getValue();
    summ = pay.sum_user.getValue();

    arr = [];
    arr.push([today, text, summ, counter]);
    sheets.his.getRange(his.last_fin.getValue() + 1, 14, 1, 4).setValues(arr);

    pay.sum_user.setValue("Введите сумму");
    pay.oper_user.setValue("Выберите операцию");
    pay.num_r.setValue(pay.num_r.getValue() + 1);
}

// Выбор заказа на листе "Закупка" --------------------------------------- Работает
function order_select() {
    var last_r = sheets.per.getLastRow() - shift_per + 1; // Общее количество рядов без учёта "шапки"
    var order_check = per.ord_user.getValue(); // Выбранный пользователем заказ

    var order, position, first, number;
    var formula;
    var a;

    // Открывает все ряды
    sheets.per.showRows(shift_per, last_r);

    // Проверка на содержимое ячейки, где должны быть номера заказа
    if ((order_check == "Все") || (order_check == "")) {
        // Удаляет все временные записи
        per.ord_user.setValue("Все");
        per.pos.clearContent();
        per.ord_cost.clearContent();
        per.stat_check.clearContent();
    } else {
        for (a = shift_per; a < 59; a++) {
            order = sheets.per.getRange(a, per.order_list).getValue(); // Текущий заказ из списка
            if (order == order_check) {
                position = sheets.per.getRange(a, per.article).getRowIndex(); // Позиция (ряд) на которой находится информация о заказе
                first = sheets.per.getRange(position, per.order_pos).getValue(); // Ряд, с которого начинается заказ
                per.pos.setValue(position);
                number = sheets.per.getRange(position, per.order_num).getValue(); // Количество товаров в заказе

                // Если это не первый ряд, прячет предыдущие ряды
                if (position != shift_per) {
                    sheets.per.hideRows(shift_per, first - shift_per);
                }

                sheets.per.hideRows(first + number, last_r + shift_per - first - number);
                formula = "=SUM(R" + first + ":R" + (first + number - 1) + ")"; // Формула для рассчёта стоимости заказа
                per.ord_cost.setFormula(formula);
                sheets.per.getRange(first, per.order_cost).setFormula(formula);
                per.stat_check.setFormula("=SUM(AH" + first + ":AH" + (first + number - 1) + ")");

                break;
            }

            // Если нет совпадений по номерам заказов, удаляет формулы
            per.pos.clearContent();
            per.ord_cost.clearContent();
            per.stat_check.clearContent();
        }
    }

    // Проверка, выбран ли заказ
    if (first != "") {
        status_processing_check();
    }
}

// Проверка для переноса данных с листа "Закупка" на лист "Экспорт"
function status_processing_check() {
    var summ = per.stat_check.getValue(); // Сумма всех числовых меток статусов заказа
    var summ_round = summ.toFixed(2); // Округлённая сумма всех числовых меток статусов заказа
    var position = per.pos.getValue(); // Позиция (ряд) на которой находится информация о заказе
    var first = sheets.per.getRange(position, per.order_pos).getValue(); // Стартовая позиция (ряд) заказа
    var number = sheets.per.getRange(position, per.order_num).getValue(); // Количество товаров в заказе
    var order_status = sheets.per.getRange(position, per.order_mark).getValue(); // Статус заказа
    var result = number * 1.01;
    var result_round = result.toFixed(2);

    var mark_a = 0,
        mark_b = 0; // Метки, необходимые для определения действий со статусами заказов
    var status_check; // Переменная для проверки статуса
    var a;

    // Закрытие заказа
    if ((order_status != "0") && (order_status != "1") && (summ_round == result_round)) {
        sheets.per.getRange(position, per.order_mark).setValue("1");
        // Отметка, что после переноса данных не надо удалять метки статусов
        mark_a = 1;
    }

    if ((summ >= number) && (order_status == "")) {
        status_processing();
        // ------------------------------ Тут должно быть всплывающее окошко о том, что данные перенесены
        //Browser.msgBox("Обсчёт и перенос данных");

        // Проверка, есть ли в заказе отсутствующие товары
        if (summ == number) {
            mark_b = 1;
        }

        // Удаление всех меток
        if (mark_a == 0) {
            sheets.per.getRange(first, per.status_mark, number).clearContent();
            if (mark_b == 0) {
                // Проставление меток для отсутствующих товаров
                for (a = first; a < first + number; a++) {
                    status_check = sheets.per.getRange(a, per.status).getValue();
                    if ((status_check == "Товар отсутствует") || (status_check == "Отсутствует нужный цвет / размер") || (status_check == "Возврат товара") || (status_check == "Деньги возвращены")) {
                        sheets.per.getRange(a, per.status_mark).setValue("1,01");
                    }
                }
            }
        }
    }

    // Смена статуса заказа
    // Если все товары отменены
    if (summ == result) {
        sheets.per.getRange(position, per.order_mark).setValue("0");
    }
}

// Автоматическое проставление статусов для выбранного заказа ------------------------------------ Работает
function status_change() {
    var position = per.pos.getValue(); // Позиция (ряд) на которой находится информация о заказе
    var first = sheets.per.getRange(position, per.order_pos).getValue(); // Стартовая позиция заказа
    var number = sheets.per.getRange(position, per.order_num).getValue(); // Количество товаров в заказе
    var status_user = per.stat_user.getValue(); // Статус для заказа, выбранный пользователем

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var mark = "1"; // Маркер статуса по умолчанию (для обычных статусов)
    var marker_check; // Проверка маркера
    var count = 1;
    var a;

    // В зависимости от выбранного статуса задаёт метку
    if ((status_user == "Товар отсутствует") || (status_user == "Отсутствует нужный цвет / размер") || (status_user == "Возврат товара") || (status_user == "Деньги возвращены")) {
        mark = "1,01";
    } else if (status_user == "Доставлен клиенту") {
        mark = "1,0099";
    }

    for (a = first; a < first + number; a++) {
        // Проверка на наличие маркера
        marker_check = sheets.per.getRange(a, per.status_mark).getValue();
        if (marker_check == "") {
            // Если маркера нет, статус меняется и добавляется нужный маркер
            sheets.per.getRange(a, per.status).setValue(status_user);
            sheets.per.getRange(a, per.status_mark).setValue(mark);
            // Счётчик, проверяющий были ли замены
            count++;
        }
    }

    // Если замены были, происходит обновление даты последнего изменения
    if (count > 1) {
        sheets.per.getRange(position, per.order_date).setValue(today);
    }

    per.stat_user.setValue("Выберите статус");

    // Запуск скрипта, отвечающего за проверку и перенос данных
    status_processing_check();
}

// Создание текста письма и его проверка
function mail_create() {
    var html;
    var content, content_send; // Содержимое письма
    var greet, no_prod, change, final; // Текстовые блоки: приветствие, отсутствующие товары, смена статусов, окончание

    var name = per.cli_name.getValue(); // Имя и фамилия клиента
    name = name.split(" ")[0]; // Имя клиента
    var status, url_tao, photo, article, url_order, num, num_all;
    var link = ss.getId();
    link = '<a href="https://docs.google.com/a/taojet.com/spreadsheet/pub?key=' + DocsList.getFileById(link).getId() + '&single=true&gid=0&output=html" class="underline">Балансу</a>'; // Ссылка на лист "Баланс" для этой таблички

    var last_r = sheets.exp.getLastRow();
    var val; // Переменная для временного хранения разных данных
    var mark_check; // Проверка маркеров (для разделения записей об отсутствующих товарах и смене статусов)
    var mark = 0; // Переменная для определения наличия информации об отсутствующих товарах в письме
    var a;

    val = sheets.exp.getRange(3, 1).getValue();
    greet = '<p>' + val.split("&")[0] + name + val.split("&")[2] + '</p><br>';

    //no_prod = '<p>' + sheets.exp.getRange(4, 1).getValue() + '</p><table align="center">';
    no_prod = '<p>' + sheets.exp.getRange(4, 1).getValue() + '</p><table bgcolor="fadadd">';
    change = '<ul type="none">';

    for (a = shift_exp; a < last_r + 1; a++) {
        article = sheets.exp.getRange(a, 1).getValue();
        status = sheets.exp.getRange(a, exp.text).getValue();
        status = status.split('"')[1];

        mark_check = sheets.exp.getRange(a, exp.mark).getValue();
        if (mark_check != "") {
            mark++;
            url_tao = sheets.per.getRange(mark_check, per.url_tao).getFormula();
            url_tao = url_tao.split('"')[1]; // Ссылка на товар на Таобао
            photo = sheets.per.getRange(mark_check, per.photo).getFormula();
            photo = photo.split('"')[1]; // Ссылка на фото товара
            photo = '<a href="' + url_tao + '"><img src="' + photo + '" width="50" height="50" alt="На Таобао"></a>';
            url_order = sheets.per.getRange(mark_check, per.url_order).getFormula();
            url_order = url_order.split('"')[1]; // Ссылка на заказ
            url_order = '<a href="' + url_order + '" class="underline">В заказе</a>';

            //no_prod = no_prod + '<tr><td>' + photo + '</td><td align="center" width="200">' + status +'</td>';
            no_prod = no_prod + '<tr><td width="50"></td><td>' + photo + '</td><td align="center" width="150">' + article + '</td><td align="center" width="100">' + url_order + '</td><td align="center" width="150">' + status + '</td><td width="50"></td>';
        } else {
            num = sheets.exp.getRange(a, exp.number).getValue(); // Количество товаров с текущим статусом
            num_all = sheets.exp.getRange(a, exp.number_all).getValue(); // Общее количество товаров в заказе

            val = sheets.exp.getRange(7, 1).getValue();
            change = change + '<li>' + val.split("&")[0] + article + val.split("&")[2] + status + val.split("&")[4] + num + val.split("&")[6] + num_all + val.split("&")[8] + '</li>';
        }
    }

    no_prod = no_prod + "</table><br>";

    if (mark == 0) {
        no_prod = "";
        change = '<p>' + sheets.exp.getRange(6, 3).getValue() + change + '</ul><br>';
    } else {
        change = '<p>' + sheets.exp.getRange(6, 1).getValue() + change + '</ul><br>';
    }

    val = sheets.exp.getRange(8, 1).getValue();
    final = '<p>' + val.split("&")[0] + link + val.split("&")[2] + '</p>';

    content = greet + no_prod + change + final;
    sheets.exp.getRange(3, 4).setValue(content); // Выгрузка содержимого письма в ячейку для последующей отправки
    content_send = content + '<p style="text-align: center"><button onclick="google.script.run.mail_send()"><img src="http://cs402128.userapi.com/g44571543/a_edc642ac.jpg" width="15" height="15" alt="" style="vertical-align: middle"> Отправить</button></p>';

    html = HtmlService.createHtmlOutput(content_send);
    ss.show(html); // Показывает будущее письмо для проверки
}

function mail_send() {
    var arr, whose, text;
    var counter = 1;

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var email = "manager@taojet.com";
    //var email = per.cli_mail.getValue(); // Электронная почта, на которую будет отправленно письмо
    var content = sheets.exp.getRange(3, 4).getValue(); // Текст письма

    // Отправка письма
    MailApp.sendEmail(email, "Оповещение о смене статуса", content, {
        htmlBody: content
    });

    // Создание пометки на листе "История" об отправке письма
    whose = "";
    text = "Письмо выслано клиенту.";

    arr = [];
    arr.push([whose, today, text, counter]);
    heets.his.getRange(his.last_his + 1, 5, 1, 4).setValues(arr);

    // --------------------------------- Заменить эту часть на всплывающе сообщение об отправке
    Browser.msgBox("Оповещение", "Письмо отправлено на " + email, Browser.Buttons.OK);
}

function status_processing() {
    var position = per.pos.getValue(); // Позиция (ряд) на которой находится информация о заказе
    var first = sheets.per.getRange(position, per.order_pos).getValue(); // Стартовая позиция (ряд) заказа
    var order = sheets.per.getRange(position, per.order_list).getValue(); // Номер заказа
    var number = sheets.per.getRange(position, per.order_num).getValue(); // Количество товаров в заказе

    var a, b;

    var status_arr = ["v_obr", 0, "tov_opl", 0, "voz_tov", 0, "den_voz", 0, "pre_dop", 0, "ozh_otp", 0, "otp_kit", 0, "otp_ros", 0, "otp_kli", 0, "dos_kli", 0]; // Массив с данными о статусах в текущем заказе
    var status_all = ["В обработке", "Товар оплачен", "Возврат товара", "Деньги возвращены", "Предоставьте дополнительную информацию", "Ожидание отправки от продавца",
        "Отправлен на склад в Китае", "Отправлен на склад в России", "Отправлен клиенту", "Доставлен клиенту"
    ]; // Список всех статусов

    var symbol, status;

    for (a = first; a < first + number; a++) {
        status = sheets.per.getRange(a, per.status).getValue();

        for (b = 1; b < 11; b++) {
            if (status == status_all[b - 1]) {
                status_arr[b * 2 - 1] = status_arr[b * 2 - 1] + 1;
                break;
            }
        }
    }

    // Создание записей об изменениях статусов
    for (a = 1; a < 11; a++) {
        if (status_arr[a * 2 - 1] != 0) {
            symbol = status_arr[a * 2 - 1]; // Количество товаров с текущим статусом
            status = status_all[a - 1]; // Статус

            record_history(symbol, status);

            if ((a != 1) && (a != 10)) {
                record_export(symbol, status);
            }

            if (a == 2) {
                record_finance();
            }
        }
    }
}

// Создание записи на листе "Экспорт"
function record_export(symbol, status) {
    var position = per.pos.getValue(); // Стартовая позиция (ряд) заказа
    var order = sheets.per.getRange(position, per.order_list).getValue(); // Номер заказа
    var number = sheets.per.getRange(position, per.order_num).getValue(); // Количество товаров в заказе

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var arr, whose, text, number;
    var counter = 1;

    text = 'Изменение статуса на "' + status + '".';

    arr = [];
    arr.push([order, text, symbol, number, "", counter]);
    sheets.exp.getRange(exp.last_r.getValue() + 1, 1, 1, 6).setValues(arr);
}

// Создание записи на листе "История"
function record_history(symbol, status) {
    var position = per.pos.getValue();
    var order = sheets.per.getRange(position, per.order_list).getValue();
    var number = sheets.per.getRange(position, per.order_num).getValue();

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var arr, text;
    var counter = 1;

    text = 'Изменение статуса на "' + status + '" (' + symbol + ' из ' + number + ').';

    arr = [];
    arr.push([order, today, text, counter]);
    sheets.his.getRange(his.last_his.getValue() + 1, 5, 1, 4).setValues(arr);
}

// Создание записи на листе "История" (Финансы)
function record_finance() {
    var position = per.pos.getValue();
    var order = sheets.per.getRange(position, per.order_list).getValue();
    var order_cost = sheets.per.getRange(position, per.order_cost).getValue();
    order_cost = "-" + order_cost;

    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var arr, text;
    var counter = 1;

    text = "Закупка товаров заказа номер " + order + ".";

    arr = [];
    arr.push([today, text, order_cost, counter]);
    sheets.his.getRange(his.last_fin.getValue() + 1, 14, 1, 4).setValues(arr);
}

// Создание на листе "Баланс" зоны с заголовками
function balance_header(row) {
    sheets.bal.getRange(row, 5).setValue("Артикул");
    sheets.bal.getRange(row, 6).setValue("Номер заказа");
    sheets.bal.getRange(row, 7).setValue("Ссылка на товар на Таобао");
    sheets.bal.getRange(row, 8).setValue("Ссылка на товар в заказе");
    sheets.bal.getRange(row, 9).setValue("Фото");
    sheets.bal.getRange(row, 10).setValue("Размер");
    sheets.bal.getRange(row, 11).setValue("Цвет");
    sheets.bal.getRange(row, 12).setValue("Количество");
    sheets.bal.getRange(row, 13).setValue("Сумма с комиссией");
    sheets.bal.getRange(row, 14).setValue("Вес, кг");
    sheets.bal.getRange(row, 15).setValue("Стоимость доставки");
    sheets.bal.getRange(row, 16).setValue("Статус");
    sheets.bal.getRange(row, 17).setValue("Примечание <агента>");
    sheets.bal.getRange(row, 5, 1, 13).setBackground("#ffe599");
    sheets.bal.getRange(row, 5, 1, 13).setFontSize(10);
}

// Добавляет пометку на лист "Данные" об отсутствующих товарах
function data_mark_no(row) {
    var today = new Date();
    today = Utilities.formatDate(today, Session.getTimeZone(), "dd.MM.yyyy");

    var product_check; // Переменная для проверки наличия товара по его позиции (ряду)
    var mark = 0; // Переменная для определения наличия записи об отсутсвующем товаре на листе "Данные"
    var a;

    var arr, photo;
    var counter = 1;

    // Проверка, есть ли уже запись об этом товаре
    for (a = 2; a < data.num_no.getValue() + 2; a++) {
        product_check = sheets.data.getRange(10, a).getValue();
        if (row == product_check) {
            sheets.data.getRange(9, a).setValue(today);
            mark = mark + 1;
            break;
        }
    }

    if (mark == 0) {
        sheets.data.getRange(8, data.num_no.getValue() + 1).setFormula(sheets.per.getRange(row, 8).getFormula());

        /*arr = [];
    arr.push([today, row, counter]);
    sheets.data.getRange(9, data.num_no + 1, 3, 1).setValues(arr);*/

        sheets.data.getRange(9, data.num_no.getValue() + 1).setValue(today);
        sheets.data.getRange(10, data.num_no.getValue() + 1).setValue(row);
        sheets.data.getRange(11, data.num_no.getValue() + 1).setValue("1");
    }
}

// Создание на листе "Баланс" зоны для отсутствующих товаров
function balance_no_product() {
    var position_no; // Позиция (ряд) отсутствующего товара
    var row = 1;
    var cal_r; // Рассчёт рядов
    var a;

    balance_clear();
    balance_header(row);

    // Объединяет ячейки - Пока не добаят возможность выбирать цвета границ, не использовать!
    //sheets.bal.getRange(2, 5, 1, 3).merge();
    sheets.bal.getRange(2, 5).setFontSize(10);
    sheets.bal.getRange(2, 5).setValue("Отсутствующие товары");

    for (a = 2; a < data.num_no.getValue() + 1; a++) {
        position_no = sheets.data.getRange(10, a).getValue() + shift_data;
        sheets.bal.getRange(a + 1, 5, 1, 4).setValues(sheets.data.getRange(position_no, 1, 1, 4).getValues());
        sheets.bal.getRange(a + 1, 9).setFormula("='Данные'!E" + position_no);
        sheets.bal.getRange(a + 1, 10).setFormula("='Данные'!F" + position_no);
        sheets.bal.getRange(a + 1, 11).setFormula("='Данные'!G" + position_no);
        sheets.bal.getRange(a + 1, 12).setFormula("='Данные'!H" + position_no);
        sheets.bal.getRange(a + 1, 13).setFormula("='Данные'!I" + position_no);
        sheets.bal.getRange(a + 1, 14).setFormula("='Данные'!J" + position_no);
        sheets.bal.getRange(a + 1, 15).setFormula("='Данные'!K" + position_no);
        sheets.bal.getRange(a + 1, 16).setFormula("='Данные'!L" + position_no);
        sheets.bal.getRange(a + 1, 17).setFormula("='Данные'!M" + position_no);
        sheets.bal.getRange(a + 1, 5, 1, 13).setBackground("#fadadd");
    }

    //sheets.data.getRange(16, 19).setValue(1 + data.num_no.getRange());
    cal_r = data.num_no.getValue() + 1;

    balance_order_last(cal_r);
}

function balance_order_last(cal_r) {
    var begin = cal_r + 2; // Начало блока с информацией о последнем заказе
    var position_last = set.last_pos.getValue() + shift_data;
    var number_last = set.last_num.getValue();
    var end = position_last + number_last - 1;

    sheets.bal.getRange(begin, 5).setFontSize(10);
    sheets.bal.getRange(begin, 5).setValue("Последний заказ");
    sheets.bal.getRange(begin + 1, 5).setFormula("=SORT('Данные'!A" + position_last + ":M" + end + ";2;0)");

    //sheets.data.getRange(16, 19).setValue(begin + number);
    cal_r = begin + number_last;

    if ((position_last - shift_data) != shift_per) {
        balance_order_other(cal_r);
    }
}

function balance_order_other(cal_r) {
    var begin = cal_r + 3;
    var row = begin - 1;

    var start = shift_per + shift_data;
    var end = set.last_pos.getValue() + shift_data - 1;

    balance_header(row);

    sheets.bal.getRange(begin, 5).setFontSize(10);
    sheets.bal.getRange(begin, 5).setValue("Остальные заказы");
    sheets.bal.getRange(begin + 1, 5).setFormula("=SORT('Данные'!A" + start + ":M" + end + ";2;0)");
}

function balance_clear() {
    var last_r = sheets.bal.getLastRow();

    // Разделяет ячейки - Пока не добаят возможность выбирать цвета границ, не использовать!
    //sheets.bal.getRange(1, 5, last_r, 13).breakApart();
    //heets.bal.getRange(1, 5, last_r, 13).setBorder(top, left, bottom, right, vertical, horizontal);

    sheets.bal.getRange(1, 5, last_r, 13).setBackground("white"); // Сброс цвета ячеек
    // Установка размера шрифтов
    sheets.bal.getRange(1, 5, last_r, 1).setFontSize(8);
    sheets.bal.getRange(1, 10, last_r, 2).setFontSize(8);
    sheets.bal.getRange(1, 17, last_r, 1).setFontSize(8);
    // Очищает ячейки
    sheets.bal.getRange(1, 5, last_r, 13).clearContent();
}

function test() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var a = sheets.pay.getRange(10, pay.date.getValue());

}