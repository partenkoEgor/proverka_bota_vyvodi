// ==UserScript==
// @name         Проверка бота автозакрытия выводов
// @namespace    proverka-bota-vyvodi
// @version      0.3.0
// @description  Проверяет работу «Бот_ПК закрытия тикетов выводов» по правилам: для каждого тикета грузит историю и транзакцию из отчёта Withdrawal requests, сверяет 11 условий, находит ошибки бота и тикеты, которые бот должен был закрыть, но не закрыл. Ничего не меняет в тикетах — только читает и выгружает Excel.
// @match        *://th-managment.com/*paymentsupport*
// @match        *://my-managment.com/*paymentsupport*
// @match        *://managment.io/*paymentsupport*
// @match        *://webmanegment.com/*paymentsupport*
// @match        *://th-managment.com/*ExtendedPaymentRequestList*
// @match        *://my-managment.com/*ExtendedPaymentRequestList*
// @match        *://managment.io/*ExtendedPaymentRequestList*
// @match        *://webmanegment.com/*ExtendedPaymentRequestList*
// @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
// @grant        GM_addStyle
// @run-at       document-idle
// @homepageURL  https://github.com/partenkoEgor/proverka_bota_vyvodi
// @updateURL    https://raw.githubusercontent.com/partenkoEgor/proverka_bota_vyvodi/main/withdrawal_bot_checker.user.js
// @downloadURL  https://raw.githubusercontent.com/partenkoEgor/proverka_bota_vyvodi/main/withdrawal_bot_checker.user.js
// ==/UserScript==
(function(){'use strict';

// ════════════════════════════════════════════════════
// ПРАВИЛА БОТА — всё, что может поменяться, собрано здесь
// ════════════════════════════════════════════════════
var RULES = {
  // Логин бота в колонке Admin username истории тикета
  botLogin: 'Бот_ПК закрытия тикетов выводов',

  // п.1 — тема тикета (idTopic в истории)
  topics: { 2: "I didn't receive my withdrawal", 4: 'Error while withdrawing funds' },

  // п.2 — тип тикета (requestTypeId в истории). A — автоматические, M — ручные
  typesA: { 1: 'BT A', 2: 'SMP A', 5: 'API', 8: 'PSP' },
  typesM: { 4: 'BT M', 6: 'SMP M' },

  // п.3 — Customer status перед ответом бота (nameUserStatus, сравнение без регистра).
  // 1 New, 2 In progress, 3 Clarification data, 7 Payment on its way,
  // 8 Approval of compensation, 9 Contact customer support
  allowedUserStatuses: ['new', 'in progress', 'clarification data', 'payment on its way',
                        'approval of compensation', 'contact customer support'],

  // п.7 и п.11 — «Ответ игроку»: внутренний статус и Customer status, который ставит бот
  responseInternalStatus: 'response to customer',
  responseUserStatus: 'reviewed',

  // п.4 — тикет не старше N дней на момент ответа
  maxAgeDays: 30,

  // п.8 — финальные статусы транзакции вывода (поле status отчёта; Stat — текст)
  txApprovedCodes: [2],   // Одобрено
  txRejectedCodes: [3],   // Отказ (в отчёте Stat = "Rejected")

  // п.11 — номер шаблона в начале Response to customer
  templates: {
    A: { approve: 579, approveChanged: 5, reject: 4 },
    M: { approve: 411, reject: 412 }
  },

  // Расхождение сумм меньше этого считается совпадением
  amountTolerance: 0.005,

  // Пропуск: тикет ещё без ответа, транзакция финальная меньше N минут —
  // бот мог просто не дойти до него. Такие тикеты идут в «Ручную проверку»
  graceMinutes: 30,

  // Часовой пояс дат истории (минуты от UTC). Даты отчёта приходят со своим
  // смещением (+0300), dtPayTransaction в истории совпадает с ними — значит,
  // история тоже в +03:00
  siteTzOffsetMin: 180
};

// ════════════════════════════════════════════════════
// CHECKER — чистая логика без DOM (тестируется в Node)
// Все даты — «настенное время сайта» в мс (как будто UTC), чтобы
// история и отчёт сравнивались напрямую
// ════════════════════════════════════════════════════
var Checker = (function(){
  var MIN = 60000, DAY = 24 * 60 * MIN;

  function str(v){ return v == null ? '' : String(v).trim(); }
  function low(v){ return str(v).toLowerCase().replace(/\s+/g, ' '); }

  // "2026-11-17 08:13:08" / "2026-11-17T01:45:10" → мс
  function parseWallDate(s){
    var m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/.exec(str(s));
    if(!m) return null;
    return Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
  }
  // "/Date(1794892320000+0300)/" → мс настенного времени в указанном смещении
  function parseNetDate(s){
    var m = /Date\((-?\d+)([+-])(\d{2})(\d{2})\)/.exec(str(s));
    if(!m){
      var m2 = /Date\((-?\d+)\)/.exec(str(s));
      return m2 ? +m2[1] + RULES.siteTzOffsetMin * MIN : parseWallDate(s);
    }
    var off = (+m[3] * 60 + +m[4]) * (m[2] === '-' ? -1 : 1);
    return +m[1] + off * MIN;
  }
  function fmt(ms){
    if(ms == null) return '';
    var d = new Date(ms), p = function(n){ return (n < 10 ? '0' : '') + n; };
    return d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate()) + ' ' +
      p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) + ':' + p(d.getUTCSeconds());
  }
  function floorMin(ms){ return Math.floor(ms / MIN) * MIN; }
  function nowWall(){ return Date.now() + RULES.siteTzOffsetMin * MIN; }

  // "1 000,50" / "1000" / 1000 → число или null
  function parseAmount(v){
    if(typeof v === 'number') return isFinite(v) ? v : null;
    var s = str(v).replace(/[\s ]/g, '');
    if(!s) return null;
    if(s.indexOf(',') !== -1 && s.indexOf('.') === -1) s = s.replace(',', '.');
    else s = s.replace(/,/g, '');
    var n = Number(s);
    return isFinite(n) ? n : null;
  }
  // "412 This withdrawal…" → 412
  function templateNumber(answer){
    var m = /^\s*(\d{1,6})(?!\d)/.exec(str(answer));
    return m ? +m[1] : null;
  }

  // ── История ─────────────────────────────────────────
  // API отдаёт новые записи сверху; при одинаковом времени нижняя запись — более ранняя
  function normalizeHistory(data){
    var rows = (data || []).map(function(r, idx){
      return {
        idx: idx,
        date: parseWallDate(r.dateEdit),
        dateStr: str(r.dateEdit),
        admin: str(r.adminProcessedLogin),
        internal: str(r.nameInternalStatus),
        external: str(r.nameExternalStatus),
        userStatus: str(r.nameUserStatus),
        answer: str(r.answer),
        amountRaw: str(r.amount),
        amount: parseAmount(r.amount),
        currency: str(r.currencyKode),
        topicId: str(r.idTopic),
        topicName: str(r.nameTopic),
        typeId: str(r.requestTypeId),
        typeName: str(r.requestTypeName),
        txId: str(r.idTransaction)
      };
    }).filter(function(r){ return r.date != null; });
    rows.sort(function(a, b){ return (a.date - b.date) || (b.idx - a.idx); });
    return rows;
  }
  function isResponse(row){
    return low(row.internal) === RULES.responseInternalStatus || low(row.userStatus) === RULES.responseUserStatus;
  }
  function isBot(row){ return low(row.admin) === low(RULES.botLogin); }
  function topicIdOf(row){
    if(row.topicId) return row.topicId;
    for(var id in RULES.topics) if(low(RULES.topics[id]) === low(row.topicName)) return id;
    return '';
  }
  function kindOf(row){
    if(Object.prototype.hasOwnProperty.call(RULES.typesA, row.typeId)) return 'A';
    if(Object.prototype.hasOwnProperty.call(RULES.typesM, row.typeId)) return 'M';
    return null;
  }

  // ── Транзакция из отчёта Withdrawal requests ────────
  // resp — весь ответ отчёта; data — строка JSON {Success, Value:{Columns, Data}}
  function parseReport(resp){
    var data = resp && resp.data;
    if(typeof data === 'string'){ try{ data = JSON.parse(data); }catch(e){ data = null; } }
    var value = data && data.Value;
    if(!value || !Array.isArray(value.Columns) || !Array.isArray(value.Data)){
      throw new Error((value && value.Error) || (data && data.Error) || 'неожиданный ответ отчёта');
    }
    return value.Data.map(function(row){
      var o = {};
      value.Columns.forEach(function(c, i){ o[c] = row[i]; });
      return o;
    });
  }
  function normalizeTx(o){
    if(!o) return null;
    return {
      id: str(o.id),
      status: o.status == null || o.status === '' ? null : Number(o.status),
      statText: str(o.Stat),
      amount: parseAmount(o.summa),
      currency: str(o.currency),
      createdAt: parseNetDate(o.date),
      editedAt: parseNetDate(o.date_edit),
      finalAt: parseNetDate(o.date_payment),
      comment: str(o.commet),
      source: 'отчёт',
      _n: true
    };
  }

  // ── Транзакции из выгрузки базы (xlsx/csv) ──────────
  // Колонки ищутся по названию (без регистра). Обязательна только «Номер заказа» —
  // это тот же номер, что Transaction ID в тикете
  var DB_COLS = {
    id: 'номер заказа', status: 'статус', amount: 'сумма', currency: 'валюта',
    payAmount: 'сумма платежа', payCurrency: 'валюта платежа',
    createdAt: 'дата создания', finalAt: 'дата совершения платежа', comment: 'комментарий'
  };
  // Дата Excel (число дней с 1899-12-30) → мс «настенного» времени, без часовых поясов
  function parseAnyDate(v){
    if(v == null || v === '') return null;
    if(typeof v === 'number') return v > 20000 && v < 80000 ? Math.round((v - 25569) * 86400) * 1000 : null;
    if(v instanceof Date) return isNaN(v) ? null : Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes(), v.getSeconds());
    var s = str(v);
    var m = /^(\d{2})\.(\d{2})\.(\d{4})[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(s);
    if(m) return Date.UTC(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0));
    if(/Date\(/.test(s)) return parseNetDate(s);
    return parseWallDate(s);
  }
  // rows — массив объектов {заголовок: значение} (как sheet_to_json).
  // → {map: {номер: tx}, count, missing: [колонки, которых нет]}
  function buildDbIndex(rows){
    var map = {}, count = 0, keys = null, missing = [];
    (rows || []).forEach(function(o){
      if(!keys){
        keys = {};
        Object.keys(o).forEach(function(h){ keys[low(h)] = h; });
        Object.keys(DB_COLS).forEach(function(k){ if(!keys[DB_COLS[k]]) missing.push(DB_COLS[k]); });
      }
      function get(k){ var h = keys[DB_COLS[k]]; return h == null ? null : o[h]; }
      var id = str(get('id')).replace(/\.0+$/, '');
      if(!/^\d+$/.test(id)) return;
      if(!map[id]) count++;
      map[id] = {
        id: id, status: null, statText: str(get('status')),
        amount: parseAmount(get('amount')), currency: str(get('currency')),
        payAmount: parseAmount(get('payAmount')), payCurrency: str(get('payCurrency')),
        createdAt: parseAnyDate(get('createdAt')), editedAt: null,
        finalAt: parseAnyDate(get('finalAt')), comment: str(get('comment')),
        source: 'база', _n: true
      };
    });
    return {map: map, count: count, missing: keys ? missing : Object.keys(DB_COLS).map(function(k){ return DB_COLS[k]; })};
  }

  function txOutcome(tx){
    if(!tx) return null;
    if(RULES.txApprovedCodes.indexOf(tx.status) !== -1) return 'approve';
    if(RULES.txRejectedCodes.indexOf(tx.status) !== -1) return 'reject';
    // Выгрузка из базы: статус только текстом
    if(tx.status == null){
      var t = low(tx.statText);
      if(/^(одобрен|approved)/.test(t)) return 'approve';
      if(/^(отказ|rejected)/.test(t)) return 'reject';
    }
    return null;
  }
  function txStatusLabel(tx){
    if(!tx) return '';
    return (tx.status == null ? '' : tx.status + ' ') + (tx.statText || '');
  }

  // ── Условия ─────────────────────────────────────────
  // Результат одной проверки: ok | fail | unknown | na (не применимо)
  function R(state, detail){ return {state: state, detail: detail || ''}; }

  // ctx: {before, at, ageAt, created, prior, tx}
  //   before — состояние тикета перед ответом (строка истории)
  //   at     — момент ответа (или момент проверки для тикетов без ответа)
  //   ageAt  — от какого момента считать возраст тикета (п.4)
  //   prior  — записи истории до ответа
  function checkConditions(ctx){
    var b = ctx.before, tx = ctx.tx, c = {};
    var topic = topicIdOf(b);
    c.c1 = RULES.topics[topic]
      ? R('ok', topic + ' ' + RULES.topics[topic])
      : R('fail', 'тема ' + (topic || '?') + ' ' + (b.topicName || ''));

    var kind = kindOf(b);
    c.c2 = kind ? R('ok', b.typeId + ' ' + b.typeName) : R('fail', 'тип ' + (b.typeId || '?') + ' ' + b.typeName);

    c.c3 = RULES.allowedUserStatuses.indexOf(low(b.userStatus)) !== -1
      ? R('ok', b.userStatus) : R('fail', 'Customer status «' + (b.userStatus || 'пусто') + '»');

    var ageDays = (ctx.ageAt - ctx.created) / DAY;
    c.c4 = ageDays <= RULES.maxAgeDays
      ? R('ok', ageDays.toFixed(1) + ' дн.')
      : R('fail', 'тикету ' + ageDays.toFixed(1) + ' дн. (больше ' + RULES.maxAgeDays + ')');

    var outcome = txOutcome(tx);
    if(!tx){
      c.c5 = c.c6 = c.c8 = c.c9 = R('unknown', ctx.txError || 'нет данных транзакции');
    }else{
      c.c5 = !b.currency || !tx.currency ? R('unknown', 'валюта не указана')
        : low(b.currency) === low(tx.currency) ? R('ok', b.currency)
        : R('fail', 'тикет ' + b.currency + ', транзакция ' + tx.currency +
            (tx.payCurrency && low(tx.payCurrency) === low(b.currency)
              ? ' (с валютой тикета совпадает только валюта платежа: ' + tx.payAmount + ' ' + tx.payCurrency + ')' : ''));

      // п.6 и п.8а: тикет создан раньше, чем транзакция стала финальной.
      // date_payment отчёта — с точностью до минуты, поэтому одна и та же минута — «не понять»
      if(tx.finalAt == null) c.c6 = R('unknown', 'в отчёте нет даты проведения (date_payment)');
      else{
        var cm = floorMin(ctx.created), fm = floorMin(tx.finalAt);
        c.c6 = fm > cm ? R('ok', 'создан ' + fmt(ctx.created) + ' < проведена ' + fmt(tx.finalAt))
          : fm < cm ? R('fail', 'транзакция проведена ' + fmt(tx.finalAt) + ' — раньше создания тикета ' + fmt(ctx.created))
          : R('unknown', 'тикет создан в ту же минуту, когда проведена транзакция (' + fmt(tx.finalAt) + ')');
      }

      // п.8б: на момент ответа транзакция в финальном статусе
      if(!outcome) c.c8 = R('fail', 'статус транзакции не финальный: ' + txStatusLabel(tx));
      else if(tx.finalAt == null) c.c8 = R('unknown', txStatusLabel(tx) + ', но нет даты проведения');
      else if(floorMin(tx.finalAt) > ctx.at) c.c8 = R('fail', 'транзакция стала финальной ' + fmt(tx.finalAt) + ' — позже ответа ' + fmt(ctx.at));
      else c.c8 = R('ok', txStatusLabel(tx));

      // п.9: суммы. Для A расхождение допустимо — бот должен поправить сумму (п.10)
      var ta = b.amount, xa = tx.amount;
      var payNote = tx.payAmount != null && tx.payCurrency && low(tx.payCurrency) !== low(tx.currency)
        ? '; сумма платежа ' + tx.payAmount + ' ' + tx.payCurrency : '';
      if(xa == null) c.c9 = R('unknown', 'нет суммы транзакции');
      else if(ta != null && Math.abs(ta - xa) < RULES.amountTolerance) c.c9 = R('ok', 'совпадает: ' + xa + payNote);
      else if(kind === 'A') c.c9 = R('ok', 'отличается (тикет ' + (b.amountRaw || 'пусто') + ', транзакция ' + xa + payNote + ') — для ' + b.typeName + ' бот меняет сумму');
      else if(kind === 'M') c.c9 = R('fail', 'отличается (тикет ' + (b.amountRaw || 'пусто') + ', транзакция ' + xa + payNote + ') — для ' + b.typeName + ' закрывать нельзя');
      else c.c9 = R('unknown', 'отличается, тип тикета не из правил');
    }

    var prev = (ctx.prior || []).filter(isResponse)[0];
    c.c7 = prev ? R('fail', 'уже был «Ответ игроку» ' + prev.dateStr + ' (' + (prev.admin || 'без логина') + ')') : R('ok', '');

    var amountsEqual = tx && tx.amount != null && b.amount != null && Math.abs(b.amount - tx.amount) < RULES.amountTolerance;
    var expected = null;
    if(outcome && kind){
      var t = RULES.templates[kind];
      expected = outcome === 'reject' ? t.reject
        : kind === 'A' && !amountsEqual ? t.approveChanged : t.approve;
    }
    return {checks: c, kind: kind, outcome: outcome, amountsEqual: amountsEqual, expected: expected};
  }

  function worst(checks, keys){
    var s = 'ok';
    keys.forEach(function(k){
      var st = checks[k] && checks[k].state;
      if(st === 'fail') s = 'fail';
      else if(st === 'unknown' && s !== 'fail') s = 'unknown';
    });
    return s;
  }
  var COND_KEYS = ['c1', 'c2', 'c3', 'c4', 'c5', 'c6', 'c7', 'c8', 'c9'];

  // Что нужно от транзакции, чтобы разобрать тикет: id транзакции или null,
  // если тикет и без неё не наш (бот не отвечал, тема/тип не из правил)
  function transactionNeeded(historyData){
    var rows = normalizeHistory(historyData);
    if(!rows.length) return null;
    var botIdx = findBotResponse(rows);
    var row = botIdx >= 0 ? rows[botIdx] : rows[rows.length - 1];
    if(botIdx < 0){
      var last = rows[rows.length - 1];
      if(!RULES.topics[topicIdOf(last)] || !kindOf(last)) return null;
    }
    for(var i = rows.length - 1; i >= 0; i--) if(rows[i].txId && rows[i].txId !== '0') return rows[i].txId;
    return row.txId || null;
  }
  function findBotResponse(rows){
    for(var i = 0; i < rows.length; i++) if(isBot(rows[i]) && isResponse(rows[i])) return i;
    return -1;
  }

  // Главная функция. historyData — resp.data истории; txRow — строка отчёта
  // (объект с колонками) или null; opts: {now, txError}
  //
  // category:
  //   bot_ok        — бот ответил, все условия соблюдены
  //   bot_error     — бот ответил, но какое-то условие нарушено
  //   missed        — условия выполнены, а бот не ответил (ответил человек или ответа нет)
  //   manual        — не хватает данных, чтобы решить однозначно
  //   not_applicable — тикет не подходит под правила, бот его не трогал — всё верно
  function analyze(historyData, txRow, opts){
    opts = opts || {};
    var rows = normalizeHistory(historyData);
    var res = {category: 'not_applicable', checks: {}, notes: []};
    if(!rows.length){ res.category = 'manual'; res.notes.push('пустая история'); return res; }

    var created = rows[0].date;
    var tx = txRow && txRow._n ? txRow : normalizeTx(txRow);
    var botIdx = findBotResponse(rows);
    var respIdx = -1;
    for(var i = 0; i < rows.length; i++) if(isResponse(rows[i])){ respIdx = i; break; }
    var last = rows[rows.length - 1];

    res.created = created;
    res.tx = tx;
    res.txId = (botIdx >= 0 ? rows[botIdx].txId : last.txId) || '';
    if(tx && tx.id) res.txId = tx.id;
    res.botRowsWithoutAnswer = rows.filter(function(r){ return isBot(r) && !isResponse(r); }).length;

    // ── 1. Бот ответил: проверяем, имел ли он право и правильно ли ответил
    if(botIdx >= 0){
      if(botIdx === 0){ res.category = 'manual'; res.notes.push('ответ бота — первая запись истории, состояния «до» нет'); return res; }
      var bot = rows[botIdx], before = rows[botIdx - 1];
      var ev = checkConditions({before: before, at: bot.date, ageAt: bot.date, created: created,
                                prior: rows.slice(0, botIdx), tx: tx, txError: opts.txError});
      var c = ev.checks;

      // п.10: сумма после ответа бота
      if(ev.kind === 'M' || !tx || tx.amount == null) c.c10 = R('na', ev.kind === 'M' ? 'для M не меняется' : '');
      else if(bot.amount != null && Math.abs(bot.amount - tx.amount) < RULES.amountTolerance)
        c.c10 = R('ok', ev.amountsEqual ? 'сумма не менялась' : 'бот поставил ' + bot.amountRaw + ' (было ' + (before.amountRaw || 'пусто') + ')');
      else c.c10 = R('fail', 'после бота сумма ' + (bot.amountRaw || 'пусто') + ', а в транзакции ' + tx.amount);
      if(ev.kind === 'M' && bot.amount != null && before.amount != null && Math.abs(bot.amount - before.amount) >= RULES.amountTolerance)
        c.c10 = R('fail', 'бот изменил сумму M-тикета: ' + before.amountRaw + ' → ' + bot.amountRaw);

      // п.11: статус и номер шаблона
      var actual = templateNumber(bot.answer);
      if(low(bot.internal) !== RULES.responseInternalStatus)
        c.c11 = R('fail', 'внутренний статус «' + bot.internal + '» вместо Response to customer');
      else if(ev.expected == null)
        c.c11 = R('unknown', 'не понять, какой шаблон нужен (тип или статус транзакции вне правил)');
      else if(actual === ev.expected) c.c11 = R('ok', String(actual));
      else c.c11 = R('fail', 'ответ ' + (actual == null ? 'без номера' : actual) + ', нужен ' + ev.expected);

      res.who = bot.admin; res.at = bot.date; res.before = before; res.response = bot;
      res.checks = c; res.kind = ev.kind; res.outcome = ev.outcome;
      res.expected = ev.expected; res.actual = actual;
      var w = worst(c, COND_KEYS.concat(['c10', 'c11']));
      res.category = w === 'fail' ? 'bot_error' : w === 'unknown' ? 'manual' : 'bot_ok';
      return res;
    }

    // ── 2. Бот не ответил. Подходил ли тикет под правила?
    // Нет транзакции (нет Transaction ID или её нет в отчёте) — боту не с чем
    // сверять, он правильно не ответил. Если же отчёт просто не ответил — решить нельзя
    if(!tx){
      if(opts.txLoadFailed){ res.category = 'manual'; res.notes.push(opts.txError || 'не загрузилась транзакция'); }
      else res.notes.push(opts.txError || 'нет транзакции');
      return res;
    }
    var ctx;
    if(respIdx > 0){
      // Ответил кто-то другой — смотрим на момент этого ответа
      var resp = rows[respIdx];
      ctx = {before: rows[respIdx - 1], at: resp.date, ageAt: resp.date, created: created,
             prior: rows.slice(0, respIdx), tx: tx, txError: opts.txError};
      res.who = resp.admin || '(без логина)'; res.at = resp.date; res.response = resp;
      res.actual = templateNumber(resp.answer);
    }else if(respIdx === 0){
      res.notes.push('ответ — первая запись истории');
      return res;
    }else{
      // Ответа нет вовсе — смотрим на текущее состояние. Возраст — на момент,
      // когда транзакция стала финальной: тогда бот и должен был сработать
      var now = opts.now != null ? opts.now : nowWall();
      var ageAt = tx && tx.finalAt != null ? Math.max(tx.finalAt, created) : now;
      ctx = {before: last, at: now, ageAt: ageAt, created: created, prior: rows, tx: tx, txError: opts.txError};
      res.who = ''; res.at = null;
    }
    var ev2 = checkConditions(ctx);
    res.checks = ev2.checks; res.kind = ev2.kind; res.outcome = ev2.outcome;
    res.expected = ev2.expected; res.before = ctx.before;
    var w2 = worst(ev2.checks, COND_KEYS);
    if(w2 === 'fail'){ res.category = 'not_applicable'; return res; }
    if(w2 === 'unknown'){
      // Если тема/тип не наши — это не «ручная проверка», а просто чужой тикет
      res.category = ev2.checks.c1.state === 'ok' && ev2.checks.c2.state === 'ok' ? 'manual' : 'not_applicable';
      return res;
    }
    // Все условия выполнены, а бот не ответил
    var waited = tx.finalAt != null ? (ctx.at - tx.finalAt) / MIN : null;
    res.waitedMin = waited;
    if(respIdx < 0 && waited != null && waited < RULES.graceMinutes){
      res.category = 'manual';
      res.notes.push('транзакция финальная только ' + Math.round(waited) + ' мин — бот мог ещё не дойти');
      return res;
    }
    res.category = 'missed';
    res.notes.push(respIdx > 0 ? 'ответил ' + res.who + ' через ' + Math.round(waited) + ' мин после финального статуса транзакции'
                               : 'ответа нет, транзакция финальная ' + Math.round(waited) + ' мин');
    return res;
  }

  return {
    RULES: RULES, COND_KEYS: COND_KEYS,
    parseWallDate: parseWallDate, parseNetDate: parseNetDate, fmt: fmt, parseAmount: parseAmount,
    templateNumber: templateNumber, normalizeHistory: normalizeHistory, parseReport: parseReport,
    normalizeTx: normalizeTx, txOutcome: txOutcome, txStatusLabel: txStatusLabel,
    transactionNeeded: transactionNeeded, analyze: analyze, nowWall: nowWall,
    buildDbIndex: buildDbIndex, parseAnyDate: parseAnyDate, DB_COLS: DB_COLS
  };
})();

// Node (тесты): отдаём Checker и не трогаем DOM
if(typeof module !== 'undefined' && module.exports){ module.exports = Checker; return; }

// не запускаться во фреймах и не инжектиться дважды
if(window.top !== window.self) return;
if(window.__pbv_loaded) return;
window.__pbv_loaded = true;

// ════════════════════════════════════════════════════
// КОНФИГ ЗАПРОСОВ
// ════════════════════════════════════════════════════
var VERSION = '0.3.0';
var HISTORY_URL = '/admin/backoffice/paymentsupporthistory';
var REPORT_URL = '/en/admin/report/requestwithdrawal';
// Поля формы отчёта Withdrawal requests (как шлёт сама страница); заполнен только id.
// Даты при заполненном id сервер не учитывает
var REPORT_FORM = {filter_by_date: 1, status: null, user_id: null, bank_id: null,
                   countryIds: [], currencyIds: [], ref_id: null, ref_ids: [], subagent_id: []};
var DELAY_MIN = 150, DELAY_MAX = 800, DELAY_STEP = 50;
var MAX_RETRIES = 5;
var POLL_INTERVAL = 800, POLL_MAX = 8;
var LIST_RETRIES = 8, LIST_RETRY_BASE_MS = 3000;
var DEFAULT_THREADS = 3;
var TX_FILE_LIMIT = 5000;       // сколько номеров в одной части (лимит выгрузки из базы)
var COLLECT_KEY = 'pbv_collected'; // собранный список — в localStorage, чтобы пережил перезагрузку

var CATS = {
  bot_error:      {title: 'Ошибки бота',      color: '#ff6e5a'},
  missed:         {title: 'Пропуски',         color: '#f0a500'},
  manual:         {title: 'Ручная проверка',  color: '#7aa7ff'},
  bot_ok:         {title: 'Бот ОК',           color: '#5fd876'},
  not_applicable: {title: 'Не под правила',   color: '#8899aa'},
  load_error:     {title: 'Не загрузилось',   color: '#c080ff'}
};

// ════════════════════════════════════════════════════
// STATE
// ════════════════════════════════════════════════════
var state = {running: false, source: '', total: 0, done: 0, currentPage: 1, expectedTotal: 0,
             processedIds: {}, results: [], counts: {}, startTime: 0, _delay: DELAY_MIN,
             stopReason: '', pagesDone: 0, finishing: false};
var _listUrl = null, _listPayload = null, _keepaliveTimer = null;
var txCache = {};
var dbIndex = null;   // транзакции из выгрузок базы: {map, count, files: [{name, count}], missing}
// Шаг 1: {ticketIds, txIds, byTicket: {ticket: tx}, skipped, noTx, loadErrors, source, savedAt}
// и истории тикетов (только в памяти — для шага 2 без повторной загрузки)
var collected = null, histCache = {};

// ════════════════════════════════════════════════════
// УТИЛИТЫ
// ════════════════════════════════════════════════════
function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
function headers(json){
  var h = {'X-Requested-With': 'XMLHttpRequest', 'Accept': 'application/json, text/plain, */*'};
  h['Content-Type'] = json ? 'application/json' : 'application/x-www-form-urlencoded';
  var csrf = document.head.querySelector('meta[name="csrf-token"]');
  if(csrf) h['X-CSRF-TOKEN'] = csrf.getAttribute('content');
  var w = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
  if(typeof w._TIMEZONE === 'string') h['X-Time-Zone'] = w._TIMEZONE;
  return h;
}
function adaptDelay(ok){
  state._delay = ok ? Math.max(DELAY_MIN, state._delay - DELAY_STEP) : Math.min(DELAY_MAX, state._delay + DELAY_STEP * 2);
}
function parseIds(text){
  var seen = {}, out = [];
  (String(text || '').match(/\d{5,}/g) || []).forEach(function(id){ if(!seen[id]){ seen[id] = 1; out.push(id); } });
  return out;
}
function isHtml(text){ var t = String(text || '').replace(/^\s+/, ''); return t.indexOf('<!DOCTYPE') === 0 || t.indexOf('<html') === 0; }
function todayAt(time){
  var d = new Date(), p = function(n){ return (n < 10 ? '0' : '') + n; };
  return p(d.getDate()) + '.' + p(d.getMonth() + 1) + '.' + d.getFullYear() + ' ' + time;
}
function startKeepalive(){
  stopKeepalive();
  _keepaliveTimer = setInterval(function(){
    if(state.running) fetch('/admin/backoffice', {credentials: 'include', cache: 'no-store'}).catch(function(){});
  }, 3 * 60 * 1000);
}
function stopKeepalive(){ if(_keepaliveTimer){ clearInterval(_keepaliveTimer); _keepaliveTimer = null; } }

// ════════════════════════════════════════════════════
// ПЕРЕХВАТ ЗАПРОСА СПИСКА ПОСЛЕ APPLY
// ════════════════════════════════════════════════════
function _isListBody(b){
  return b && typeof b === 'object' && typeof b.dateFrom !== 'undefined' &&
    (typeof b.page !== 'undefined' || typeof b.numberPage !== 'undefined');
}
function _isListResp(r){
  return r && (typeof r.is_polling !== 'undefined' || (Array.isArray(r.data) && (r.otherData || r.success)));
}
function _pageKey(p){ return typeof p.page !== 'undefined' ? 'page' : 'numberPage'; }
function _applyCapture(url, payload){
  if(state.running) return;
  _listUrl = url;
  _listPayload = JSON.parse(JSON.stringify(payload));
  console.log('%c[PBV ' + VERSION + '] ✅ Фильтр перехвачен', 'color:#3fa;font-weight:bold');
  updateApplyState();
}
(function(){
  var _open = XMLHttpRequest.prototype.open, _send = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(m, u){ this._pbv_m = m; this._pbv_u = u; return _open.apply(this, arguments); };
  XMLHttpRequest.prototype.send = function(body){
    var self = this;
    if(this._pbv_m === 'POST' && body){
      try{
        var p = typeof body === 'string' ? JSON.parse(body) : body;
        if(_isListBody(p)) self.addEventListener('load', function(){
          try{ var r = JSON.parse(self.responseText); if(_isListResp(r)) _applyCapture(self.responseURL || self._pbv_u, p); }catch(e){}
        });
      }catch(e){}
    }
    return _send.apply(this, arguments);
  };
  var _fetch = window.fetch;
  window.fetch = function(url, opts){
    var res = _fetch.apply(this, arguments);
    if(opts && opts.method === 'POST' && opts.body && !state.running){
      try{
        var p = typeof opts.body === 'string' ? JSON.parse(opts.body) : opts.body;
        if(_isListBody(p)) res.then(function(r){
          r.clone().json().then(function(d){ if(_isListResp(d)) _applyCapture(r.url || (typeof url === 'string' ? url : ''), p); }).catch(function(){});
        }).catch(function(){});
      }catch(e){}
    }
    return res;
  };
})();

// ════════════════════════════════════════════════════
// CSS + HTML
// ════════════════════════════════════════════════════
GM_addStyle(`
#pbv-fab{position:fixed;left:16px;bottom:174px;z-index:2147483647;display:flex;align-items:center;gap:7px;padding:8px 14px;border-radius:999px;background:rgba(12,18,35,.9);border:1px solid rgba(90,170,255,.5);color:#a9d1ff;font:600 13px system-ui,Arial,sans-serif;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,.25);user-select:none;}
#pbv-fab .dot{width:7px;height:7px;border-radius:50%;background:#5aa9ff;}
#pbv-fab.running .dot{animation:pbvpulse 1.2s ease-in-out infinite;}
@keyframes pbvpulse{0%,100%{opacity:1;}50%{opacity:.3;}}
#pbv-wrap{position:fixed;inset:0;z-index:2147483646;background:rgba(4,8,18,.6);display:none;align-items:center;justify-content:center;}
#pbv-wrap.open{display:flex;}
#pbv-panel{width:92vw;max-width:920px;max-height:88vh;background:rgba(12,18,32,.97);border:1px solid rgba(90,170,255,.4);border-radius:16px;display:flex;flex-direction:column;overflow:hidden;font:13px Arial,sans-serif;color:#e8f0ff;box-shadow:0 24px 64px rgba(0,0,0,.7);}
#pbv-hdr{display:flex;align-items:center;gap:10px;padding:12px 18px;border-bottom:1px solid rgba(90,170,255,.3);}
#pbv-hdr .t{font-weight:800;color:#a9d1ff;}
#pbv-hdr .v{font-size:10px;color:#6a8ab0;border:1px solid rgba(90,170,255,.3);border-radius:4px;padding:1px 6px;}
#pbv-close{margin-left:auto;background:none;border:none;color:rgba(255,255,255,.4);font-size:18px;cursor:pointer;}
#pbv-body{display:flex;gap:16px;padding:16px 18px;overflow:hidden;flex:1;min-height:0;}
#pbv-side{width:300px;flex-shrink:0;display:flex;flex-direction:column;gap:8px;overflow-y:auto;padding-right:4px;}
#pbv-side > *{flex-shrink:0;}
.pbv-row{display:flex;gap:6px;}
.pbv-row .pbv-btn{flex:1;padding:8px 4px;}
.pbv-step{border:1px solid rgba(90,170,255,.25);border-radius:10px;padding:9px;display:flex;flex-direction:column;gap:7px;}
.pbv-step .pbv-sec{color:#8fb6e6;}
.pbv-mini{padding:5px 8px;border-radius:6px;border:1px solid rgba(90,170,255,.4);background:rgba(90,170,255,.1);color:#cfe4ff;font:700 11px Arial,sans-serif;cursor:pointer;margin:3px 4px 0 0;}
.pbv-mini:disabled{opacity:.35;cursor:not-allowed;}
.pbv-note{font-size:10px;color:rgba(200,220,255,.45);line-height:1.45;}
#pbv-main{flex:1;display:flex;flex-direction:column;gap:10px;min-width:0;min-height:0;}
.pbv-sec{font-size:9px;font-weight:700;color:#6a8ab0;text-transform:uppercase;letter-spacing:.7px;}
.pbv-field{border-radius:8px;padding:8px 10px;font-size:12px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.04);}
.pbv-field.ok{border-color:rgba(63,250,130,.4);color:#5fd876;background:rgba(63,250,130,.07);}
.pbv-field.warn{border-color:rgba(240,165,0,.35);color:#f0a500;background:rgba(240,165,0,.06);}
#pbv-ids{width:100%;box-sizing:border-box;height:70px;resize:vertical;padding:8px;border-radius:8px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.05);color:#e0ecff;font:12px monospace;outline:none;}
.pbv-btn{width:100%;padding:9px;border-radius:8px;border:1px solid rgba(90,170,255,.6);background:rgba(90,170,255,.16);color:#cfe4ff;font:800 12px Arial,sans-serif;cursor:pointer;}
.pbv-btn:disabled{opacity:.35;cursor:not-allowed;}
#pbv-db-state{font-size:11px;}
#pbv-db-state.ok{border-color:rgba(63,250,130,.4);color:#5fd876;}
#pbv-db-state.warn{border-color:rgba(240,165,0,.35);color:#f0a500;}
#pbv-stop{display:none;border-color:rgba(250,80,80,.4);background:rgba(250,80,80,.08);color:#f06060;}
.pbv-metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;}
.pbv-metric{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:9px;padding:7px 4px;text-align:center;}
.pbv-metric .val{font-size:18px;font-weight:800;color:#fff;}
.pbv-metric .lbl{font-size:8px;color:rgba(180,210,255,.5);text-transform:uppercase;letter-spacing:.5px;margin-top:3px;}
#pbv-found{flex:1;overflow-y:auto;min-height:120px;}
.pbv-item{padding:6px 9px;margin-bottom:4px;border-radius:7px;background:rgba(255,255,255,.04);border-left:3px solid #888;font-size:11px;}
.pbv-item .sub{color:rgba(200,220,255,.6);margin-top:2px;}
#pbv-pb-wrap{height:3px;background:rgba(255,255,255,.06);border-radius:2px;overflow:hidden;}
#pbv-pb{height:100%;width:0;background:linear-gradient(90deg,#3a6ab0,#8fc4ff);transition:width .3s;}
#pbv-status{font-size:11px;color:rgba(180,210,255,.65);min-height:16px;}
#pbv-toast{position:fixed;bottom:220px;left:16px;z-index:2147483647;padding:12px 18px;border-radius:12px;color:#e0ecff;font:600 12px Arial,sans-serif;max-width:360px;box-shadow:0 8px 24px rgba(0,0,0,.5);pointer-events:none;}
`);

var fab = document.createElement('button');
fab.id = 'pbv-fab';
fab.title = 'Alt+B — открыть/закрыть';
fab.innerHTML = '<span class="dot"></span><span>Проверка бота</span>';
document.body.appendChild(fab);

function metric(id, label, color){
  return '<div class="pbv-metric"><div class="val" id="' + id + '"' + (color ? ' style="color:' + color + '"' : '') + '>0</div><div class="lbl">' + label + '</div></div>';
}
var wrap = document.createElement('div');
wrap.id = 'pbv-wrap';
wrap.innerHTML =
  '<div id="pbv-panel">' +
    '<div id="pbv-hdr"><span class="t">🤖 Проверка бота автозакрытия выводов</span><span class="v">' + VERSION + '</span><button id="pbv-close">✕</button></div>' +
    '<div id="pbv-body">' +
      '<div id="pbv-side">' +
        '<div class="pbv-sec">Тикеты</div>' +
        '<div class="pbv-field warn" id="pbv-apply">⚠ Выставь фильтр и нажми APPLY — перехвачу список</div>' +
        '<textarea id="pbv-ids" placeholder="или свои Ticket ID: 26081917 26104522 … (любой разделитель)"></textarea>' +
        '<div class="pbv-step">' +
          '<div class="pbv-sec">Шаг 1 — собрать номера транзакций</div>' +
          '<div class="pbv-row"><button class="pbv-btn" id="pbv-col-apply" disabled>① По APPLY</button><button class="pbv-btn" id="pbv-col-ids">① По списку</button></div>' +
          '<div class="pbv-field" id="pbv-col-state">Грузит только истории тикетов и собирает Transaction ID в один столбец, частями по ' + TX_FILE_LIMIT + '</div>' +
        '</div>' +
        '<div class="pbv-step">' +
          '<div class="pbv-sec">Шаг 2 — выгрузки из базы и проверка</div>' +
          '<label class="pbv-btn" style="text-align:center;box-sizing:border-box">📄 Загрузить выгрузку (можно несколько)<input id="pbv-db-file" type="file" accept=".xlsx,.xls,.csv" multiple style="display:none"></label>' +
          '<div class="pbv-field" id="pbv-db-state">Не загружена — транзакции ищутся в Withdrawal requests</div>' +
          '<button class="pbv-btn" id="pbv-check-collected" disabled>② Проверить собранные</button>' +
        '</div>' +
        '<div class="pbv-sec" style="margin-top:4px">Или сразу проверить, без шага 1</div>' +
        '<div class="pbv-row"><button class="pbv-btn" id="pbv-run-apply" disabled>▶ По APPLY</button><button class="pbv-btn" id="pbv-run-ids">▶ По списку</button></div>' +
        '<button class="pbv-btn" id="pbv-stop">⏹ Стоп</button>' +
        '<label style="font-size:11px;color:rgba(200,220,255,.6);display:flex;align-items:center;gap:8px">Потоков (1–5): <input id="pbv-threads" type="number" min="1" max="5" value="' + DEFAULT_THREADS + '" style="width:48px;padding:4px;border-radius:6px;border:1px solid rgba(255,255,255,.15);background:rgba(255,255,255,.05);color:#e0ecff"></label>' +
        '<div class="pbv-note">Только чтение. Бот: <b>' + esc(RULES.botLogin) + '</b>. Удобно сузить фильтр темами «I didn\'t receive my withdrawal» и «Error while withdrawing funds».</div>' +
      '</div>' +
      '<div id="pbv-main">' +
        '<div class="pbv-metrics">' +
          metric('pbv-m-total', 'Тикетов') + metric('pbv-m-done', 'Проверено') +
          metric('pbv-m-bot_error', 'Ошибки бота', CATS.bot_error.color) + metric('pbv-m-missed', 'Пропуски', CATS.missed.color) +
          metric('pbv-m-manual', 'Ручная', CATS.manual.color) + metric('pbv-m-bot_ok', 'Бот ОК', CATS.bot_ok.color) +
          metric('pbv-m-not_applicable', 'Не под правила', CATS.not_applicable.color) + metric('pbv-m-load_error', 'Не загрузилось', CATS.load_error.color) +
        '</div>' +
        '<div class="pbv-sec">Ошибки, пропуски и ручная проверка</div>' +
        '<div id="pbv-found"></div>' +
        '<div id="pbv-pb-wrap"><div id="pbv-pb"></div></div>' +
        '<div id="pbv-status">Готов к запуску</div>' +
      '</div>' +
    '</div>' +
  '</div>';
document.body.appendChild(wrap);

function $(id){ return document.getElementById(id); }
function togglePanel(){ wrap.classList.toggle('open'); }
fab.addEventListener('click', togglePanel);
$('pbv-close').addEventListener('click', function(){ wrap.classList.remove('open'); });
wrap.addEventListener('click', function(e){ if(e.target === wrap) wrap.classList.remove('open'); });
document.addEventListener('keydown', function(e){
  if(e.altKey && (e.code === 'KeyB')){ e.preventDefault(); e.stopPropagation(); togglePanel(); }
}, true);

// ════════════════════════════════════════════════════
// UI
// ════════════════════════════════════════════════════
function setStatus(m){ $('pbv-status').textContent = m; }
function showToast(msg, type){
  var old = $('pbv-toast'); if(old) old.remove();
  var t = document.createElement('div');
  t.id = 'pbv-toast';
  t.style.background = type === 'ok' ? 'rgba(0,80,30,.92)' : type === 'err' ? 'rgba(90,15,15,.92)' : 'rgba(70,50,10,.92)';
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(function(){ t.remove(); }, 5000);
}
function updateApplyState(){
  var el = $('pbv-apply'); if(!el) return;
  if(_listUrl){
    el.className = 'pbv-field ok';
    el.textContent = '✅ Фильтр пойман' + (_listPayload && _listPayload.dateFrom ? ' (с ' + _listPayload.dateFrom + ')' : '');
  }
  $('pbv-run-apply').disabled = !_listUrl || state.running;
  $('pbv-col-apply').disabled = !_listUrl || state.running;
}
function updateCollectedButton(){
  var b = $('pbv-check-collected'); if(!b) return;
  var n = collected ? collected.ticketIds.length + (collected.loadErrors || []).length : 0;
  b.disabled = !n || state.running;
  b.textContent = collected ? '② Проверить собранные — тикетов: ' + n : '② Проверить собранные';
}
function setProgress(){
  var total = Math.max(state.total, state.expectedTotal || 0);
  $('pbv-m-total').textContent = total || '—';
  $('pbv-m-done').textContent = state.done;
  Object.keys(CATS).forEach(function(k){ $('pbv-m-' + k).textContent = state.counts[k] || 0; });
  $('pbv-pb').style.width = (total ? Math.min(100, Math.round(state.done / total * 100)) : 0) + '%';
}
function problemsText(r){
  var out = [];
  Object.keys(r.checks || {}).forEach(function(k){
    var c = r.checks[k];
    if(c.state === 'fail' || c.state === 'unknown') out.push('п.' + k.slice(1) + ': ' + c.detail);
  });
  return out.concat(r.notes || []).join('; ');
}
function addFoundItem(r){
  if(['bot_error', 'missed', 'manual', 'load_error'].indexOf(r.category) === -1) return;
  var div = document.createElement('div');
  div.className = 'pbv-item';
  div.style.borderLeftColor = CATS[r.category].color;
  div.innerHTML = '<b>#' + esc(r.ticketId) + '</b> — ' + esc(CATS[r.category].title) +
    (r.expected != null || r.actual != null ? ' · шаблон ' + esc(r.actual == null ? '—' : r.actual) + ' / нужен ' + esc(r.expected == null ? '?' : r.expected) : '') +
    '<div class="sub">' + esc(problemsText(r) || r.error || '') + '</div>';
  var box = $('pbv-found');
  box.insertBefore(div, box.firstChild);
}
function setRunning(on){
  state.running = on;
  fab.classList.toggle('running', on);
  ['pbv-run-ids', 'pbv-col-ids', 'pbv-db-file'].forEach(function(id){ $(id).disabled = on; });
  $('pbv-stop').style.display = on ? 'block' : 'none';
  updateApplyState();
  updateCollectedButton();
}

// ════════════════════════════════════════════════════
// ЗАПРОСЫ
// ════════════════════════════════════════════════════
function fetchListPage(pageNum, pollAttempt){
  pollAttempt = pollAttempt || 0;
  var payload = JSON.parse(JSON.stringify(_listPayload));
  payload[_pageKey(payload)] = pageNum;
  var pageSize = Number(payload.countOfItemPage || payload.perPage) || 0;
  function fail(reason){ return {ok: false, reason: reason, tickets: []}; }
  return fetch(_listUrl, {method: 'POST', headers: headers(true), credentials: 'include', body: JSON.stringify(payload)})
    .then(function(r){ return r.text().then(function(text){ return {status: r.status, ok: r.ok, text: text}; }); })
    .then(function(res){
      if(isHtml(res.text)) throw new Error('SESSION_EXPIRED');
      if(!res.ok) return fail('HTTP ' + res.status);
      var resp; try{ resp = JSON.parse(res.text); }catch(e){ return fail('ответ не JSON'); }
      if(resp.is_polling && resp.count < 5 && pollAttempt < POLL_MAX){
        setStatus('⏳ Загружаю стр.' + pageNum + '… (' + (pollAttempt + 1) + '/' + POLL_MAX + ')');
        return sleep(POLL_INTERVAL).then(function(){ return fetchListPage(pageNum, pollAttempt + 1); });
      }
      var total = null;
      (Array.isArray(resp.data) ? resp.data : []).forEach(function(x){ if(x && !Array.isArray(x) && typeof x.total === 'number') total = x.total; });
      var raw = Array.isArray(resp.data) ? resp.data : [];
      if(Array.isArray(raw[0])) raw = raw[0];
      var tickets = raw.map(function(item){
        var t = Array.isArray(item) ? item[0] : item;
        return t ? {ticketId: String(t.idTicket || t.idTkt || '')} : null;
      }).filter(function(t){ return t && /^\d{5,}$/.test(t.ticketId); });
      var size = pageSize || tickets.length || 1;
      var known = Math.max(state.expectedTotal || 0, total > 0 ? total : 0) || null;
      var serverHasMore = resp.otherData ? resp.otherData.hasMore : undefined;
      if(!tickets.length){
        if(serverHasMore === false) return {ok: true, tickets: [], hasMore: false, total: total};
        if(known != null && (pageNum - 1) * size >= known) return {ok: true, tickets: [], hasMore: false, total: total};
        if(known == null && !resp.is_polling) return {ok: true, tickets: [], hasMore: false, total: total};
        return fail(resp.is_polling ? 'сервер ещё готовит страницу' : 'пустая страница, хотя total = ' + known);
      }
      var hasMore = serverHasMore === false ? false : known != null ? pageNum * size < known : true;
      return {ok: true, tickets: tickets, hasMore: hasMore, total: total};
    })
    .catch(function(e){ if(e.message === 'SESSION_EXPIRED') throw e; return fail('сеть: ' + (e && e.message)); });
}

// Общая обёртка: повторы на 429/5xx и сбоях сети. Возвращает разобранный JSON или
// бросает Error (SESSION_EXPIRED — сессия кончилась)
function requestJson(url, init, label, attempt){
  attempt = attempt || 1;
  function retry(reason){
    adaptDelay(false);
    if(attempt > MAX_RETRIES) throw new Error(reason);
    setStatus('⏳ ' + label + ': ' + reason + ', повтор ' + attempt + '/' + MAX_RETRIES);
    return sleep(attempt * 2000).then(function(){ return requestJson(url, init, label, attempt + 1); });
  }
  return fetch(url, init).then(function(r){
    if([429, 500, 502, 503, 504, 529].indexOf(r.status) !== -1) return retry('HTTP ' + r.status);
    return r.text().then(function(text){
      if(isHtml(text)) throw new Error('SESSION_EXPIRED');
      if(!r.ok) throw new Error('HTTP ' + r.status);
      adaptDelay(true);
      try{ return JSON.parse(text); }catch(e){ return retry('ответ не JSON'); }
    });
  }, function(e){ return retry('сеть: ' + (e && e.message)); });
}

function fetchHistory(ticketId){
  return requestJson(HISTORY_URL, {
    method: 'POST', headers: headers(false), credentials: 'include',
    body: 'ticketId=' + encodeURIComponent(ticketId) + '&is_iframe=1'
  }, 'история #' + ticketId).then(function(resp){
    if(!resp || resp.success === false || !Array.isArray(resp.data)) throw new Error('история: неожиданный ответ');
    return resp.data;
  });
}

// Транзакция из отчёта Withdrawal requests: {row} | {row:null, error}. Кэш по id
function fetchTransaction(txId){
  if(txCache[txId]) return txCache[txId];
  var body = Object.assign({date_from: todayAt('00:00'), date_to: todayAt('23:59')}, REPORT_FORM, {id: Number(txId)});
  var p = requestJson(REPORT_URL, {method: 'POST', headers: headers(true), credentials: 'include', body: JSON.stringify(body)},
    'транзакция ' + txId)
    .then(function(resp){
      var rows = Checker.parseReport(resp).filter(function(r){ return String(r.id) === String(txId); });
      return rows.length ? {row: rows[0]} : {row: null, error: 'транзакция ' + txId + ' не найдена в Withdrawal requests'};
    })
    .catch(function(e){
      if(e.message === 'SESSION_EXPIRED') throw e;
      delete txCache[txId];  // не кэшируем сбой — следующий тикет спросит снова
      return {row: null, failed: true, error: 'отчёт не ответил: ' + e.message};
    });
  txCache[txId] = p;
  return p;
}

// ════════════════════════════════════════════════════
// ОБРАБОТКА ТИКЕТА
// ════════════════════════════════════════════════════
function processTicket(ticketId, useCache){
  var histP = useCache && histCache[ticketId] ? Promise.resolve(histCache[ticketId]) : fetchHistory(ticketId);
  return histP.then(function(hist){
    var txId = Checker.transactionNeeded(hist);
    var txP = !txId ? Promise.resolve({row: null, error: 'в тикете нет Transaction ID'})
      : dbIndex && dbIndex.map[txId] ? Promise.resolve({row: dbIndex.map[txId]})
      : fetchTransaction(txId).then(function(t){
          if(!t.row && dbIndex) t.error = 'транзакции ' + txId + ' нет ни в выгрузке из базы, ни в Withdrawal requests';
          return t;
        });
    return txP.then(function(tx){
      var r = Checker.analyze(hist, tx.row, {txError: tx.error, txLoadFailed: !!tx.failed});
      r.ticketId = ticketId;
      return r;
    });
  }).catch(function(e){
    if(e && e.message === 'SESSION_EXPIRED') throw e;
    return {ticketId: ticketId, category: 'load_error', error: e && e.message, checks: {}, notes: []};
  });
}
function record(r){
  state.results.push(r);
  state.counts[r.category] = (state.counts[r.category] || 0) + 1;
  addFoundItem(r);
}

function getThreads(){
  var n = parseInt(($('pbv-threads') || {}).value, 10);
  return Math.min(5, Math.max(1, n || DEFAULT_THREADS));
}
function fmtEta(){
  var total = Math.max(state.total, state.expectedTotal || 0);
  if(!state.done || !total || state.done >= total) return '';
  var min = Math.round((Date.now() - state.startTime) / state.done * (total - state.done) / 60000);
  return min >= 60 ? ' | ~' + Math.floor(min / 60) + ' ч ' + (min % 60) + ' мин' : ' | ~' + min + ' мин';
}
function runPool(ids){
  var idx = 0;
  function worker(){
    if(!state.running || idx >= ids.length) return Promise.resolve();
    var id = ids[idx++];
    return state.task(id).then(function(){
      state.done++;
      setProgress();
      if(state.mode === 'collect' && (state.done % 10 === 0)) renderCollected();
      setStatus((state.source === 'apply' ? 'Стр.' + state.currentPage + ' | ' : '') + state.done + '/' +
        Math.max(state.total, state.expectedTotal || 0) + ' → #' + id + fmtEta());
      return sleep(state._delay);
    }).then(worker);
  }
  var ws = [];
  for(var w = 0; w < Math.min(getThreads(), ids.length); w++) ws.push(worker());
  return Promise.all(ws);
}

// mode: 'check' — проверка, 'collect' — шаг 1 (только истории и номера транзакций)
function resetRun(source, mode, task){
  state.mode = mode || 'check';
  state.task = task;
  state.source = source; state.total = 0; state.done = 0; state.currentPage = 1; state.expectedTotal = 0;
  state.processedIds = {}; state.results = []; state.counts = {}; state._delay = DELAY_MIN;
  state.stopReason = ''; state.pagesDone = 0; state.finishing = false; state.startTime = Date.now();
  txCache = {};
  $('pbv-found').innerHTML = '';
  setProgress();
  setRunning(true);
  startKeepalive();
}
function onError(e){
  if(e && e.message === 'SESSION_EXPIRED'){
    showToast('🔐 Сессия истекла — войди заново. Выгружаю то, что успел проверить.', 'err');
    finish('🔐 Сессия истекла');
    return;
  }
  console.error('[PBV]', e);
  finish('⚠ Ошибка: ' + (e && e.message));
}

function collectPage(attempt){
  attempt = attempt || 0;
  fetchListPage(state.currentPage).then(function(page){
    if(!state.running){ finish(); return; }
    if(page.total > 0) state.expectedTotal = Math.max(state.expectedTotal || 0, page.total);
    if(!page.ok){
      if(attempt < LIST_RETRIES){
        var wait = Math.min(LIST_RETRY_BASE_MS * (attempt + 1), 20000);
        setStatus('⚠ Стр.' + state.currentPage + ': ' + page.reason + '. Повтор ' + (attempt + 1) + '/' + LIST_RETRIES);
        return sleep(wait).then(function(){ if(state.running) collectPage(attempt + 1); else finish(); });
      }
      state.stopReason = 'стр.' + state.currentPage + ' не загрузилась: ' + page.reason;
      finish('⚠ ' + state.stopReason);
      return;
    }
    if(!page.tickets.length){ finish(); return; }
    var fresh = page.tickets.map(function(t){ return t.ticketId; }).filter(function(id){
      if(state.processedIds[id]) return false;
      state.processedIds[id] = 1;
      return true;
    });
    state.total += fresh.length;
    state.pagesDone++;
    setProgress();
    return runPool(fresh).then(function(){
      if(!state.running){ finish(); return; }
      if(!page.hasMore){ finish(); return; }
      state.currentPage++;
      return sleep(400).then(function(){ collectPage(0); });
    });
  }).catch(onError);
}

// ════════════════════════════════════════════════════
// ШАГ 1 — СБОР НОМЕРОВ ТРАНЗАКЦИЙ
// ════════════════════════════════════════════════════
function checkTask(id){ return processTicket(id, false).then(record); }
function checkCollectedTask(id){ return processTicket(id, true).then(record); }
function collectTask(id){
  return fetchHistory(id).then(function(hist){
    histCache[id] = hist;
    var txId = Checker.transactionNeeded(hist);
    collected.ticketIds.push(id);
    if(txId){
      collected.byTicket[id] = txId;
      if(!collected._seen[txId]){ collected._seen[txId] = 1; collected.txIds.push(txId); }
      state.counts.collected = (state.counts.collected || 0) + 1;
    }else{
      // тема/тип не наши и бот не отвечал — или в тикете нет Transaction ID
      collected.skipped++;
      state.counts.not_applicable = (state.counts.not_applicable || 0) + 1;
    }
  }).catch(function(e){
    if(e && e.message === 'SESSION_EXPIRED') throw e;
    collected.loadErrors.push({ticketId: id, error: e && e.message});
    state.counts.load_error = (state.counts.load_error || 0) + 1;
  });
}
function startCollect(source){
  histCache = {};
  collected = {ticketIds: [], txIds: [], byTicket: {}, _seen: {}, skipped: 0, loadErrors: [], source: '', savedAt: ''};
  resetRun(source, 'collect', collectTask);
  collected.source = source === 'apply'
    ? 'фильтр' + (_listPayload && _listPayload.dateFrom ? ' с ' + _listPayload.dateFrom + ' по ' + (_listPayload.dateTo || '') : '')
    : 'список Ticket ID';
  renderCollected();
}
function chunks(arr, n){ var out = []; for(var i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n)); return out; }
function copyText(text, what){
  function ok(){ showToast('📋 Скопировано: ' + what, 'ok'); }
  function fallback(){
    var ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    try{ document.execCommand('copy'); ok(); }catch(e){ showToast('Не удалось скопировать', 'err'); }
    ta.remove();
  }
  if(navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(ok, fallback);
  else fallback();
}
// Кнопки «скопировать часть N» для списка номеров
function partButtons(list, prefix){
  return chunks(list, TX_FILE_LIMIT).map(function(part, k){
    var from = k * TX_FILE_LIMIT + 1, to = from + part.length - 1;
    return '<button class="pbv-mini" data-copy="' + prefix + ':' + k + '">📋 ' +
      (list.length > TX_FILE_LIMIT ? 'Часть ' + (k + 1) + ' (' + from + '–' + to + ')' : 'Копировать ' + part.length) + '</button>';
  }).join('');
}
function renderCollected(){
  var el = $('pbv-col-state');
  if(!collected){ updateCollectedButton(); return; }
  var n = collected.txIds.length, parts = Math.ceil(n / TX_FILE_LIMIT);
  var html = (state.running && state.mode === 'collect' ? '⏳ ' : '✅ ') + 'Тикетов: <b>' + collected.ticketIds.length + '</b>, транзакций: <b>' + n + '</b>' +
    (parts > 1 ? ' (' + parts + ' части по ≤' + TX_FILE_LIMIT + ')' : '') +
    (collected.skipped ? '<br><span style="opacity:.7">не нужны: ' + collected.skipped + ' (тема/тип не те или нет Transaction ID)</span>' : '') +
    (collected.loadErrors.length ? '<br><span style="color:#c080ff">не загрузилось: ' + collected.loadErrors.length + '</span>' : '') +
    (collected.savedAt ? '<br><span style="opacity:.6">собрано ' + esc(collected.savedAt) + ', ' + esc(collected.source) + '</span>' : '');
  if(n && !(state.running && state.mode === 'collect')){
    html += '<div>' + partButtons(collected.txIds, 'all') + '<button class="pbv-mini" id="pbv-col-xlsx">⬇ Excel</button></div>';
  }
  el.className = 'pbv-field' + (n ? ' ok' : '');
  el.innerHTML = html;
  var x = $('pbv-col-xlsx');
  if(x) x.addEventListener('click', function(){ var fn = buildCollectedXLSX(); showToast('⬇ ' + fn, 'ok'); });
  updateCollectedButton();
  renderDbState();
}
function missingTx(){
  if(!collected) return [];
  return collected.txIds.filter(function(id){ return !(dbIndex && dbIndex.map[id]); });
}
// Клики по кнопкам «📋» (делегирование: кнопки перерисовываются)
wrap.addEventListener('click', function(e){
  var b = e.target.closest && e.target.closest('[data-copy]');
  if(!b) return;
  var a = b.getAttribute('data-copy').split(':');
  var list = a[0] === 'missing' ? missingTx() : (collected ? collected.txIds : []);
  var part = chunks(list, TX_FILE_LIMIT)[+a[1]] || [];
  copyText(part.join('\n'), part.length + ' номеров' + (a[0] === 'missing' ? ' (недостающие)' : ''));
});

function saveCollected(){
  try{
    localStorage.setItem(COLLECT_KEY, JSON.stringify({ticketIds: collected.ticketIds, txIds: collected.txIds,
      byTicket: collected.byTicket, skipped: collected.skipped, loadErrors: collected.loadErrors,
      source: collected.source, savedAt: collected.savedAt}));
  }catch(e){ console.warn('[PBV] не удалось сохранить список', e); }
}
(function restoreCollected(){
  try{
    var c = JSON.parse(localStorage.getItem(COLLECT_KEY) || 'null');
    if(c && Array.isArray(c.ticketIds) && Array.isArray(c.txIds)){
      c._seen = {}; c.txIds.forEach(function(id){ c._seen[id] = 1; });
      collected = c;
      setTimeout(renderCollected, 0);
    }
  }catch(e){}
})();

function finishCollect(reason){
  collected.savedAt = new Date().toLocaleString('ru-RU');
  saveCollected();
  renderCollected();
  var n = collected.txIds.length;
  if(!n){ setStatus((reason ? reason + ' | ' : '') + 'нужных транзакций не нашлось'); return; }
  var fn = buildCollectedXLSX();
  setStatus((reason ? reason + ' | ' : '✅ ') + fn + ' | транзакций ' + n + (n > TX_FILE_LIMIT ? ', частей ' + Math.ceil(n / TX_FILE_LIMIT) : '') +
    '. Выгрузи их из базы и загрузи в шаге 2');
  showToast('✅ Собрано ' + n + ' транзакций — ' + fn, 'ok');
}

// ════════════════════════════════════════════════════
// ШАГ 2 — ВЫГРУЗКИ ИЗ БАЗЫ
// ════════════════════════════════════════════════════
function renderDbState(){
  var el = $('pbv-db-state');
  if(!dbIndex){
    el.className = 'pbv-field';
    el.textContent = 'Не загружена — транзакции ищутся в Withdrawal requests';
    return;
  }
  var need = ['статус', 'сумма', 'валюта', 'дата совершения платежа'].filter(function(c){ return dbIndex.missing.indexOf(c) !== -1; });
  var miss = missingTx();
  var html = (need.length || miss.length ? '⚠ ' : '✅ ') + 'Файлов: ' + dbIndex.files.length + ', транзакций: <b>' + dbIndex.count + '</b>' +
    '<br><span style="opacity:.7">' + dbIndex.files.map(function(f){ return esc(f.name) + ' (' + f.count + ')'; }).join(', ') + '</span>';
  if(need.length) html += '<br>Нет колонок: ' + need.join(', ') + ' — эти пункты будут «❓»';
  if(collected && collected.txIds.length){
    html += '<br>Из собранных ' + collected.txIds.length + ' найдено ' + (collected.txIds.length - miss.length) +
      (miss.length ? ', <b>не хватает ' + miss.length + '</b> — их поищу в Withdrawal requests<div>' + partButtons(miss, 'missing') + '</div>' : ' — все');
  }
  html += '<div><button class="pbv-mini" id="pbv-db-reset">✕ Сбросить выгрузки</button></div>';
  el.className = 'pbv-field ' + (need.length || miss.length ? 'warn' : 'ok');
  el.innerHTML = html;
  $('pbv-db-reset').addEventListener('click', function(){ dbIndex = null; renderDbState(); });
}
function readDbFile(file){
  return file.arrayBuffer().then(function(buf){
    var wb = XLSX.read(buf, {type: 'array', raw: true});
    var rows = [];
    wb.SheetNames.forEach(function(n){ rows = rows.concat(XLSX.utils.sheet_to_json(wb.Sheets[n], {raw: true, defval: ''})); });
    return Checker.buildDbIndex(rows);
  });
}
$('pbv-db-file').addEventListener('change', function(e){
  var files = Array.prototype.slice.call(e.target.files || []);
  if(!files.length) return;
  var el = $('pbv-db-state');
  el.className = 'pbv-field'; el.textContent = 'Читаю ' + files.map(function(f){ return f.name; }).join(', ') + '…';
  var bad = [];
  files.reduce(function(p, file){
    return p.then(function(){ return readDbFile(file); }).then(function(idx){
      if(!idx.count){ bad.push(file.name + ': нет колонки «Номер заказа» или она пустая'); return; }
      if(!dbIndex) dbIndex = {map: {}, count: 0, files: [], missing: idx.missing};
      else dbIndex.missing = dbIndex.missing.filter(function(c){ return idx.missing.indexOf(c) !== -1; });
      Object.keys(idx.map).forEach(function(id){ if(!dbIndex.map[id]) dbIndex.count++; dbIndex.map[id] = idx.map[id]; });
      dbIndex.files.push({name: file.name, count: idx.count});
    }, function(err){ bad.push(file.name + ': ' + err.message); });
  }, Promise.resolve()).then(function(){
    e.target.value = '';
    renderDbState();
    if(bad.length) showToast('⚠ Не загружено — ' + bad.join('; '), 'warn');
  });
});

// ════════════════════════════════════════════════════
// ЗАПУСК
// ════════════════════════════════════════════════════
function idsFromBox(){
  var ids = parseIds($('pbv-ids').value);
  if(!ids.length) showToast('Вставь хотя бы один Ticket ID', 'warn');
  return ids;
}
function runIds(ids){
  state.total = ids.length;
  setProgress();
  runPool(ids).then(function(){ finish(); }).catch(onError);
}
$('pbv-col-apply').addEventListener('click', function(){
  if(state.running || !_listUrl) return;
  startCollect('apply');
  setStatus('Шаг 1: загружаю список по фильтру…');
  collectPage(0);
});
$('pbv-col-ids').addEventListener('click', function(){
  if(state.running) return;
  var ids = idsFromBox(); if(!ids.length) return;
  startCollect('ids');
  runIds(ids);
});
$('pbv-check-collected').addEventListener('click', function(){
  if(state.running || !collected) return;
  resetRun('collected', 'check', checkCollectedTask);
  // тикеты, чья история не загрузилась на шаге 1, пробуем ещё раз
  runIds(collected.ticketIds.concat((collected.loadErrors || []).map(function(x){ return x.ticketId; })));
});
$('pbv-run-apply').addEventListener('click', function(){
  if(state.running || !_listUrl) return;
  resetRun('apply', 'check', checkTask);
  setStatus('Загружаю список по фильтру…');
  collectPage(0);
});
$('pbv-run-ids').addEventListener('click', function(){
  if(state.running) return;
  var ids = idsFromBox(); if(!ids.length) return;
  resetRun('ids', 'check', checkTask);
  runIds(ids);
});
$('pbv-stop').addEventListener('click', function(){
  if(!state.running) return;
  state.stopReason = 'остановлено вручную';
  state.running = false;
  setStatus('⏹ Останавливаю — дождусь текущих запросов…');
});

function finish(reason){
  if(state.finishing) return;
  state.finishing = true;
  setRunning(false);
  stopKeepalive();
  if(state.mode === 'collect'){ finishCollect(reason); return; }
  if(!state.results.length){ setStatus((reason ? reason + ' | ' : '') + 'нет проверенных тикетов'); return; }
  try{
    var fn = buildXLSX();
    var c = state.counts;
    setStatus((reason ? reason + ' | ' : '✅ ') + fn + ' | ошибок бота ' + (c.bot_error || 0) + ', пропусков ' + (c.missed || 0) + ', ручных ' + (c.manual || 0));
    showToast('✅ ' + fn + ' | ошибок ' + (c.bot_error || 0) + ', пропусков ' + (c.missed || 0) + ' из ' + state.done, 'ok');
  }catch(e){
    console.error('[PBV] Excel', e);
    setStatus('⚠ Ошибка Excel: ' + e.message);
  }
}

// ════════════════════════════════════════════════════
// EXCEL
// ════════════════════════════════════════════════════
var COND_TITLES = {
  c1: 'п.1 Тема', c2: 'п.2 Тип', c3: 'п.3 Customer status', c4: 'п.4 ≤30 дней', c5: 'п.5 Валюта',
  c6: 'п.6/8а Тикет раньше проведения', c7: 'п.7 Не было ответа', c8: 'п.8б Транзакция финальная',
  c9: 'п.9 Сумма', c10: 'п.10 Сумма после ответа', c11: 'п.11 Шаблон'
};
var MARK = {ok: '✅', fail: '❌', unknown: '❓', na: '—'};
var XC = {W: 'FFFFFFFF', alt: 'FFF4F8FB', DT: 'FF1F1F1F', red: 'FFFFC7CE', yel: 'FFFFEB9C', grn: 'FFC6EFCE'};
function xhdr(){ return {font: {bold: true, name: 'Arial', sz: 10}, fill: {patternType: 'solid', fgColor: {rgb: 'FFD9E5F5'}}, alignment: {horizontal: 'center', vertical: 'center', wrapText: true}, border: {bottom: {style: 'thin', color: {rgb: 'FF999999'}}, right: {style: 'thin', color: {rgb: 'FF999999'}}}}; }
function xcs(bg){ return {font: {name: 'Arial', sz: 10, color: {rgb: XC.DT}}, fill: {patternType: 'solid', fgColor: {rgb: bg}}, alignment: {horizontal: 'left', vertical: 'top', wrapText: true}, border: {right: {style: 'thin', color: {rgb: 'FFD9D9D9'}}, bottom: {style: 'thin', color: {rgb: 'FFD9D9D9'}}}}; }

function makeSheet(wb, name, hdr, rows, widths, cellBg){
  var ws = XLSX.utils.aoa_to_sheet([hdr].concat(rows));
  ws['!cols'] = widths.map(function(w){ return {wch: w}; });
  var rng = XLSX.utils.decode_range(ws['!ref']);
  for(var r = 0; r <= rng.e.r; r++){
    for(var c = 0; c <= rng.e.c; c++){
      var a = XLSX.utils.encode_cell({r: r, c: c});
      if(!ws[a]) ws[a] = {t: 's', v: ''};
      if(r === 0){ ws[a].s = xhdr(); continue; }
      var bg = (cellBg && cellBg(r - 1, c, ws[a].v)) || (r % 2 === 0 ? XC.alt : XC.W);
      ws[a].s = xcs(bg);
    }
  }
  ws['!freeze'] = {xSplit: 1, ySplit: 1};
  ws['!autofilter'] = {ref: XLSX.utils.encode_range({s: {r: 0, c: 0}, e: {r: 0, c: hdr.length - 1}})};
  XLSX.utils.book_append_sheet(wb, ws, name);
}

var CHECK_KEYS = Object.keys(COND_TITLES);
var MAIN_HDR = ['Ticket ID', 'Тип', 'Тема', 'Создан', 'Кто ответил', 'Когда ответил', 'Шаблон', 'Нужен шаблон',
  'Transaction ID', 'Статус транзакции', 'Транзакция проведена', 'Сумма тикета', 'Сумма транзакции',
  'Валюта тикета', 'Валюта транзакции', 'Customer status до ответа', 'Источник транзакции', 'Текст ответа'].concat(CHECK_KEYS.map(function(k){ return COND_TITLES[k]; }))
  .concat(['Что не так / пояснение']);
var MAIN_W = [11, 8, 22, 17, 24, 17, 8, 8, 13, 14, 17, 10, 10, 9, 9, 16, 10, 40].concat(CHECK_KEYS.map(function(){ return 9; })).concat([70]);
var FIRST_CHECK_COL = 18;

function mainRow(r){
  var b = r.before || {}, tx = r.tx || {};
  return [
    r.ticketId, b.typeName || '', b.topicName || '', Checker.fmt(r.created), r.who || '', Checker.fmt(r.at),
    r.actual == null ? '' : r.actual, r.expected == null ? '' : r.expected,
    r.txId || '', Checker.txStatusLabel(r.tx), Checker.fmt(tx.finalAt), b.amountRaw || '',
    tx.amount == null ? '' : tx.amount, b.currency || '', tx.currency || '', b.userStatus || '',
    tx.source || '', r.response ? r.response.answer : ''
  ].concat(CHECK_KEYS.map(function(k){
    var c = r.checks && r.checks[k];
    return c ? MARK[c.state] + (c.state === 'ok' || c.state === 'na' ? '' : ' ' + c.detail) : '';
  })).concat([problemsText(r) || r.error || '']);
}
function checkBg(ri, c, v){
  if(c < FIRST_CHECK_COL || c >= FIRST_CHECK_COL + CHECK_KEYS.length) return null;
  var s = String(v || '');
  return s.indexOf('❌') === 0 ? XC.red : s.indexOf('❓') === 0 ? XC.yel : null;
}

function buildXLSX(){
  var wb = XLSX.utils.book_new();
  ['bot_error', 'missed', 'manual', 'bot_ok'].forEach(function(cat){
    var rows = state.results.filter(function(r){ return r.category === cat; }).map(mainRow);
    makeSheet(wb, CATS[cat].title, MAIN_HDR, rows, MAIN_W, checkBg);
  });
  var errs = state.results.filter(function(r){ return r.category === 'load_error'; })
    .map(function(r){ return [r.ticketId, r.error || '']; });
  makeSheet(wb, 'Не загрузилось', ['Ticket ID', 'Причина'], errs, [12, 80]);

  // Сводка: итоги и какие пункты чаще всего нарушены
  var byCheck = {};
  state.results.filter(function(r){ return r.category === 'bot_error'; }).forEach(function(r){
    CHECK_KEYS.forEach(function(k){ if(r.checks[k] && r.checks[k].state === 'fail') byCheck[k] = (byCheck[k] || 0) + 1; });
  });
  var sum = [
    ['Источник', state.source === 'collected' ? 'Собранный список (шаг 1): ' + (collected ? collected.source : '') : state.source === 'apply' ? 'Фильтр (APPLY)' + (_listPayload && _listPayload.dateFrom ? ', с ' + _listPayload.dateFrom + ' по ' + (_listPayload.dateTo || '') : '') : 'Список Ticket ID'],
    ['Бот', RULES.botLogin],
    ['Транзакции', dbIndex ? 'выгрузки из базы: ' + dbIndex.files.map(function(f){ return '«' + f.name + '» (' + f.count + ')'; }).join(', ') + '; чего в них нет — Withdrawal requests' : 'Withdrawal requests'],
    ['Проверено тикетов', state.done],
    ['Тикетов по фильтру (total)', state.expectedTotal || '—'],
    ['Сбор завершён', state.stopReason ? 'РАНЬШЕ ВРЕМЕНИ: ' + state.stopReason : 'полностью'],
    ['', '']
  ];
  Object.keys(CATS).forEach(function(k){ sum.push([CATS[k].title, state.counts[k] || 0]); });
  sum.push(['', ''], ['Ошибки бота по пунктам', 'тикетов']);
  CHECK_KEYS.forEach(function(k){ if(byCheck[k]) sum.push([COND_TITLES[k], byCheck[k]]); });
  sum.push(['', ''], ['Дата выгрузки', new Date().toLocaleString('ru-RU')], ['Версия скрипта', VERSION]);
  makeSheet(wb, 'Сводка', ['Показатель', 'Значение'], sum, [40, 50]);

  var d = new Date(), p = function(n){ return (n < 10 ? '0' : '') + n; };
  var fn = 'proverka_bota_vyvodi_' + d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + '_' + p(d.getHours()) + '-' + p(d.getMinutes()) + '.xlsx';
  XLSX.writeFile(wb, fn);
  return fn;
}

// Шаг 1: номера транзакций одним столбцом, без шапки — сразу вставлять в запрос к базе.
// Больше TX_FILE_LIMIT — несколько листов «Часть N»
function buildCollectedXLSX(){
  var wb = XLSX.utils.book_new();
  var parts = chunks(collected.txIds, TX_FILE_LIMIT);
  parts.forEach(function(part, k){
    var ws = XLSX.utils.aoa_to_sheet(part.map(function(id){ return [Number(id)]; }));
    part.forEach(function(id, r){ var a = XLSX.utils.encode_cell({r: r, c: 0}); if(ws[a]) ws[a].z = '0'; });
    ws['!cols'] = [{wch: 14}];
    XLSX.utils.book_append_sheet(wb, ws, parts.length > 1 ? 'Часть ' + (k + 1) : 'Транзакции');
  });
  var map = collected.ticketIds.map(function(t){ return [t, collected.byTicket[t] || '']; });
  makeSheet(wb, 'Тикеты', ['Ticket ID', 'Transaction ID'], map, [12, 14]);
  if(collected.loadErrors.length)
    makeSheet(wb, 'Не загрузилось', ['Ticket ID', 'Причина'], collected.loadErrors.map(function(x){ return [x.ticketId, x.error || '']; }), [12, 80]);
  var d = new Date(), p = function(n){ return (n < 10 ? '0' : '') + n; };
  var fn = 'tranzakcii_dlya_bazy_' + d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + '_' + p(d.getHours()) + '-' + p(d.getMinutes()) + '.xlsx';
  XLSX.writeFile(wb, fn);
  return fn;
}

setTimeout(updateApplyState, 1500);
console.log('%c[PBV ' + VERSION + '] загружен — Alt+B', 'color:#8fc4ff;font-weight:bold');
})();
