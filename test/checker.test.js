// Запуск: node test/checker.test.js
'use strict';
var assert = require('assert');
var path = require('path');
var C = require(path.join(__dirname, '..', 'withdrawal_bot_checker.user.js'));
var HIST = require('./fixtures/history_26081917.json').data;
var REPORT = require('./fixtures/withdrawal_5550119957.json');

var BOT = C.RULES.botLogin;
var passed = 0;
function test(name, fn){ fn(); passed++; console.log('  ✓ ' + name); }

// ── Хелперы ──────────────────────────────────────────
function clone(x){ return JSON.parse(JSON.stringify(x)); }
// Настенное время +03:00 → строка отчёта "/Date(ms+0300)/"
function net(wall){ return '/Date(' + (C.parseWallDate(wall) - 180 * 60000) + '+0300)/'; }
function txRow(over){
  return Object.assign({id: 5550119957, summa: 1000, status: 3, Stat: 'Rejected', currency: 'DZDM',
    date: net('2026-11-17 01:45:10'), date_edit: net('2026-11-17 05:24:54'), date_payment: net('2026-11-17 08:12:00')}, over || {});
}
// Строка истории: новые сверху, как в API
function h(date, admin, extra){
  return Object.assign({dateEdit: date, adminProcessedLogin: admin, idTopic: '2', nameTopic: "I didn't receive my withdrawal",
    requestTypeId: '4', requestTypeName: 'BT M', nameInternalStatus: 'New', nameExternalStatus: 'New',
    nameUserStatus: 'New', amount: '1000', currencyKode: 'DZDM', idTransaction: 5550119957, answer: ''}, extra || {});
}
function botRow(date, answer, extra){
  return h(date, BOT, Object.assign({nameInternalStatus: 'Response to customer', nameExternalStatus: 'Auto approved (M)',
    nameUserStatus: 'Reviewed', answer: answer}, extra || {}));
}
// Простой тикет: создан → в работе → (ответ)
function ticket(opts){
  opts = opts || {};
  var rows = [];
  if(opts.response) rows.push(opts.response);
  rows.push(h('2026-11-17 05:29:49', 'Agent_1', Object.assign({nameInternalStatus: 'Customer provided additional info', nameUserStatus: 'In progress'}, opts.before || {})));
  rows.push(h(opts.created || '2026-11-17 02:03:22', '', opts.first || {}));
  return rows;
}
function failed(r){ return Object.keys(r.checks).filter(function(k){ return r.checks[k].state === 'fail'; }); }

// ── Разбор форматов ──────────────────────────────────
test('даты отчёта /Date(ms+0300)/ переводятся в настенное время +03', function(){
  var tx = C.normalizeTx(C.parseReport(REPORT)[0]);
  assert.strictEqual(C.fmt(tx.createdAt), '2026-11-17 01:45:10');   // = dtPayTransaction в истории
  assert.strictEqual(C.fmt(tx.editedAt), '2026-11-17 05:24:54');
  assert.strictEqual(C.fmt(tx.finalAt), '2026-11-17 08:12:00');
  assert.strictEqual(tx.status, 3);
  assert.strictEqual(tx.amount, 1000);
  assert.strictEqual(tx.currency, 'DZDM');
  assert.strictEqual(C.txOutcome(tx), 'reject');
});

test('номер шаблона берётся из начала ответа целиком (4 ≠ 407)', function(){
  assert.strictEqual(C.templateNumber('412 This withdrawal was unsuccessful'), 412);
  assert.strictEqual(C.templateNumber('4 Your withdrawal was declined'), 4);
  assert.strictEqual(C.templateNumber('407 Dear customer'), 407);
  assert.strictEqual(C.templateNumber('Dear customer'), null);
});

test('суммы: "1000", "1 000,50", 1000.00', function(){
  assert.strictEqual(C.parseAmount('1000'), 1000);
  assert.strictEqual(C.parseAmount('1 000,50'), 1000.5);
  assert.strictEqual(C.parseAmount(1000.00), 1000);
  assert.strictEqual(C.parseAmount(''), null);
});

// ── Реальный тикет ───────────────────────────────────
test('реальный тикет 26081917: бот ответил верно (BT M, отказ → 412)', function(){
  var r = C.analyze(HIST, C.parseReport(REPORT)[0]);
  assert.strictEqual(r.category, 'bot_ok', JSON.stringify(r.checks, null, 1));
  assert.strictEqual(r.expected, 412);
  assert.strictEqual(r.actual, 412);
  assert.strictEqual(r.before.userStatus, 'In progress');
  assert.strictEqual(r.before.typeName, 'BT M');  // тип на момент ответа, а не при создании (API)
  assert.strictEqual(C.fmt(r.created), '2026-11-17 02:03:22');
  assert.strictEqual(r.checks.c10.state, 'na');
});

test('transactionNeeded: из истории берётся Transaction ID', function(){
  assert.strictEqual(C.transactionNeeded(HIST), '5550119957');
});

// ── Ошибки бота по пунктам ───────────────────────────
test('п.11: BT M отказ, а бот ответил 411 → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '411 Approved')}), txRow());
  assert.strictEqual(r.category, 'bot_error');
  assert.deepStrictEqual(failed(r), ['c11']);
});

test('п.11: бот поставил не Response to customer → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x', {nameInternalStatus: 'Closed'})}), txRow());
  // Reviewed в Customer status есть, но внутренний статус не тот
  assert.ok(failed(r).indexOf('c11') !== -1);
});

test('п.1: тема не про выводы → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x'), before: {idTopic: '1', nameTopic: 'Unsuccessful deposit'}}), txRow());
  assert.deepStrictEqual(failed(r), ['c1']);
});

test('п.2: тип вне правил → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x'), before: {requestTypeId: '99', requestTypeName: 'XX'}}), txRow());
  assert.ok(failed(r).indexOf('c2') !== -1);
  assert.strictEqual(r.category, 'bot_error');
});

test('п.3: Customer status перед ботом Closed → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x'), before: {nameUserStatus: 'Closed'}}), txRow());
  assert.deepStrictEqual(failed(r), ['c3']);
});

test('п.3: разрешённые статусы (Clarification data, Payment on its way, Approval of compensation, Contact customer support)', function(){
  ['Clarification data', 'Payment on its way', 'Approval of compensation', 'Contact customer support', 'New'].forEach(function(s){
    var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x'), before: {nameUserStatus: s}}), txRow());
    assert.strictEqual(r.checks.c3.state, 'ok', s);
  });
});

test('п.4: тикету больше 30 дней → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x'), created: '2026-10-10 02:03:22'}),
    txRow({date_payment: net('2026-11-17 08:12:00')}));
  assert.deepStrictEqual(failed(r), ['c4']);
});

test('п.5: валюта тикета не совпадает с валютой транзакции → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x')}), txRow({currency: 'EGP'}));
  assert.deepStrictEqual(failed(r), ['c5']);
});

test('п.6/8а: транзакция стала финальной раньше создания тикета → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x')}), txRow({date_payment: net('2026-11-17 01:59:00')}));
  assert.deepStrictEqual(failed(r), ['c6']);
});

test('п.6/8а: одна и та же минута → ручная проверка, а не ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x')}), txRow({date_payment: net('2026-11-17 02:03:00')}));
  assert.strictEqual(r.category, 'manual');
  assert.strictEqual(r.checks.c6.state, 'unknown');
});

test('п.7: «Ответ игроку» уже был раньше → ошибка', function(){
  var rows = ticket({response: botRow('2026-11-17 08:13:08', '412 x')});
  rows.splice(1, 0, h('2026-11-17 04:00:00', 'pay_mena_1', {nameInternalStatus: 'Response to customer', nameUserStatus: 'Reviewed', answer: '1 x'}));
  // после прошлого ответа тикет снова в работе
  rows.splice(1, 0, h('2026-11-17 05:00:00', 'Agent_1', {nameInternalStatus: 'New', nameUserStatus: 'In progress'}));
  var r = C.analyze(rows, txRow());
  assert.deepStrictEqual(failed(r), ['c7']);
});

test('п.8б: транзакция ещё не финальная (На рассмотрении) → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x')}), txRow({status: 1, Stat: 'Pending'}));
  assert.ok(failed(r).indexOf('c8') !== -1);
});

test('п.8б: транзакция стала финальной уже после ответа бота → ошибка', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x')}), txRow({date_payment: net('2026-11-17 09:00:00')}));
  assert.deepStrictEqual(failed(r), ['c8']);
});

test('п.9: BT M, суммы не совпадают → бот не должен был закрывать', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x')}), txRow({summa: 900}));
  assert.deepStrictEqual(failed(r), ['c9']);
});

test('п.9: BT M, бот изменил сумму → ошибка п.10', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '412 x', {amount: '900'})}), txRow());
  assert.deepStrictEqual(failed(r), ['c10']);
});

// ── Типы A: шаблоны 579 / 5 / 4 и правка суммы ───────
function apiTicket(answer, botAmount, beforeAmount){
  var api = {requestTypeId: '5', requestTypeName: 'API'};
  return ticket({
    response: botRow('2026-11-17 08:13:08', answer, Object.assign({amount: botAmount}, api)),
    before: Object.assign({amount: beforeAmount}, api), first: api
  });
}
test('API, одобрено, суммы совпадают → 579', function(){
  var r = C.analyze(apiTicket('579 ok', '1000', '1000'), txRow({status: 2, Stat: 'Approved'}));
  assert.strictEqual(r.category, 'bot_ok', JSON.stringify(r.checks));
  assert.strictEqual(r.expected, 579);
});
test('API, одобрено, сумма отличалась, бот поправил → 5', function(){
  var r = C.analyze(apiTicket('5 ok', '1000', '800'), txRow({status: 2, Stat: 'Approved'}));
  assert.strictEqual(r.category, 'bot_ok', JSON.stringify(r.checks));
  assert.strictEqual(r.expected, 5);
  assert.strictEqual(r.checks.c10.state, 'ok');
});
test('API, одобрено, сумма отличалась, а бот ответил 579 → ошибка п.11', function(){
  var r = C.analyze(apiTicket('579 ok', '1000', '800'), txRow({status: 2, Stat: 'Approved'}));
  assert.deepStrictEqual(failed(r), ['c11']);
});
test('API, сумма отличалась, бот не поправил сумму → ошибка п.10', function(){
  var r = C.analyze(apiTicket('5 ok', '800', '800'), txRow({status: 2, Stat: 'Approved'}));
  assert.deepStrictEqual(failed(r), ['c10']);
});
test('API, отказ → 4', function(){
  var r = C.analyze(apiTicket('4 declined', '1000', '1000'), txRow());
  assert.strictEqual(r.category, 'bot_ok');
  assert.strictEqual(r.expected, 4);
});
test('BT M, одобрено → 411', function(){
  var r = C.analyze(ticket({response: botRow('2026-11-17 08:13:08', '411 ok')}), txRow({status: 2, Stat: 'Approved'}));
  assert.strictEqual(r.category, 'bot_ok');
  assert.strictEqual(r.expected, 411);
});

// ── Пропуски ─────────────────────────────────────────
test('пропуск: условия выполнены, а ответил человек', function(){
  var resp = h('2026-11-17 09:00:00', 'pay_mena_201', {nameInternalStatus: 'Response to customer', nameUserStatus: 'Reviewed', answer: '412 x'});
  var r = C.analyze(ticket({response: resp}), txRow());
  assert.strictEqual(r.category, 'missed', JSON.stringify(r.checks));
  assert.strictEqual(r.who, 'pay_mena_201');
  assert.strictEqual(r.expected, 412);
  assert.strictEqual(Math.round(r.waitedMin), 48);
});

test('не пропуск: человек ответил до того, как транзакция стала финальной', function(){
  var resp = h('2026-11-17 07:00:00', 'pay_mena_201', {nameInternalStatus: 'Response to customer', nameUserStatus: 'Reviewed', answer: '420 x'});
  var r = C.analyze(ticket({response: resp}), txRow());
  assert.strictEqual(r.category, 'not_applicable');
});

test('пропуск: ответа нет, транзакция давно финальная', function(){
  var r = C.analyze(ticket(), txRow(), {now: C.parseWallDate('2026-11-17 12:00:00')});
  assert.strictEqual(r.category, 'missed');
});

test('ответа нет, транзакция финальная 10 минут → ручная (бот мог не дойти)', function(){
  var r = C.analyze(ticket(), txRow(), {now: C.parseWallDate('2026-11-17 08:22:00')});
  assert.strictEqual(r.category, 'manual');
});

test('ответа нет, транзакция не финальная → не под правила', function(){
  var r = C.analyze(ticket(), txRow({status: 1, Stat: 'Pending'}), {now: C.parseWallDate('2026-11-17 12:00:00')});
  assert.strictEqual(r.category, 'not_applicable');
});

test('ответа нет, BT M, суммы не совпадают → не под правила (бот прав)', function(){
  var r = C.analyze(ticket(), txRow({summa: 900}), {now: C.parseWallDate('2026-11-17 12:00:00')});
  assert.strictEqual(r.category, 'not_applicable');
});

test('тема не про выводы и бот не отвечал → транзакция даже не нужна', function(){
  var rows = ticket({before: {idTopic: '1', nameTopic: 'Unsuccessful deposit'}});
  assert.strictEqual(C.transactionNeeded(rows), null);
  assert.strictEqual(C.analyze(rows, null).category, 'not_applicable');
});

test('нет транзакции в отчёте → не под правила; отчёт не ответил → ручная', function(){
  assert.strictEqual(C.analyze(ticket(), null, {txError: 'не найдена'}).category, 'not_applicable');
  assert.strictEqual(C.analyze(ticket(), null, {txError: 'отчёт не ответил', txLoadFailed: true}).category, 'manual');
});

test('бот ответил, а отчёт не ответил → ручная проверка', function(){
  var r = C.analyze(HIST, null, {txError: 'отчёт не ответил', txLoadFailed: true});
  assert.strictEqual(r.category, 'manual');
});


// ── Выгрузка транзакций из базы ──────────────────────
// Заголовки как в выгрузке; даты — числом Excel (как отдаёт SheetJS с raw:true) или строкой
function serial(wall){ return C.parseWallDate(wall) / 86400000 + 25569; }
function dbRow(over){
  return Object.assign({'Номер заказа': 5550861141, 'Статус': 'Одобрено', 'Сумма': 5000, 'Валюта': 'TRY',
    'Сумма платежа': 5000, 'Валюта платежа': 'TRY', 'Дата создания': serial('2026-11-17 08:03:38'),
    'Дата совершения платежа': serial('2026-11-17 13:42:31'), 'Комментарий': 'Выплата (код: 2)', 'Агент': 'x'}, over || {});
}
function pspTicket(answer, over){
  var psp = Object.assign({requestTypeId: '8', requestTypeName: 'PSP', currencyKode: 'TRY', amount: '5000', idTransaction: 5550861141}, over || {});
  return [botRow('2026-11-17 13:43:01', answer, psp), h('2026-11-17 12:00:00', 'Agent_1', Object.assign({nameUserStatus: 'In progress'}, psp)),
          h('2026-11-17 11:40:21', '', psp)];
}

test('выгрузка из базы: колонки по названию, даты Excel и строкой', function(){
  var idx = C.buildDbIndex([dbRow(), dbRow({'Номер заказа': '5539888149', 'Статус': 'Отказ', 'Дата совершения платежа': '17.11.2026 03:12:00'})]);
  assert.strictEqual(idx.count, 2);
  assert.deepStrictEqual(idx.missing, []);
  var t = idx.map['5550861141'];
  assert.strictEqual(C.fmt(t.finalAt), '2026-11-17 13:42:31');
  assert.strictEqual(C.fmt(t.createdAt), '2026-11-17 08:03:38');
  assert.strictEqual(C.txOutcome(t), 'approve');
  assert.strictEqual(C.txOutcome(idx.map['5539888149']), 'reject');
  assert.strictEqual(C.fmt(idx.map['5539888149'].finalAt), '2026-11-17 03:12:00');
});

test('выгрузка из базы: без «Номер заказа» ничего не загружается', function(){
  var idx = C.buildDbIndex([{'ID': 1, 'Статус': 'Одобрено'}]);
  assert.strictEqual(idx.count, 0);
  assert.ok(idx.missing.indexOf('номер заказа') !== -1);
});

test('выгрузка из базы: статусы кроме Одобрено/Отказ не финальные', function(){
  ['На рассмотрении', 'Отправлено', 'Получение данных', 'Заблокировано'].forEach(function(st){
    var t = C.buildDbIndex([dbRow({'Статус': st})]).map['5550861141'];
    assert.strictEqual(C.txOutcome(t), null, st);
  });
});

test('PSP по выгрузке из базы: одобрено, суммы совпадают, ответ 579 → ОК', function(){
  var tx = C.buildDbIndex([dbRow()]).map['5550861141'];
  var r = C.analyze(pspTicket('579 ok'), tx);
  assert.strictEqual(r.category, 'bot_ok', JSON.stringify(r.checks));
  assert.strictEqual(r.tx.source, 'база');
});

test('PSP по выгрузке из базы: ответ 28 вместо 579 → ошибка п.11', function(){
  var tx = C.buildDbIndex([dbRow()]).map['5550861141'];
  var r = C.analyze(pspTicket('28 ok'), tx);
  assert.deepStrictEqual(failed(r), ['c11']);
});

test('валюта тикета совпадает только с валютой платежа (114 USD / 194 AZN) → п.5 с пояснением', function(){
  var tx = C.buildDbIndex([dbRow({'Сумма': 114, 'Валюта': 'USD', 'Сумма платежа': 194, 'Валюта платежа': 'AZN'})]).map['5550861141'];
  var r = C.analyze(pspTicket('579 ok', {currencyKode: 'AZN', amount: '114'}), tx);
  assert.deepStrictEqual(failed(r), ['c5']);
  assert.ok(/только валюта платежа: 194 AZN/.test(r.checks.c5.detail), r.checks.c5.detail);
  assert.ok(/сумма платежа 194 AZN/.test(r.checks.c9.detail), r.checks.c9.detail);
});

console.log('\n' + passed + ' тестов прошло');
