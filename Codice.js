/**
 * PAPIRO DINAMICA — Backend Google Apps Script
 * ---------------------------------------------
 * Foglio "Configurazione" (colonne A→J):
 *   A id · B tipo · C titolo · D testo · E mediaUrl
 *   F opzioni (separate da | ) · G rispostaCorretta · H feedbackEsatto · I feedbackErrato
 *   J config  ← NUOVA: JSON con media (immagine/video) e struttura dei tipi complessi (es. completamento)
 *
 * Tipi card: narrazione · quiz · aperta · completamento · (testo · video · mappa_neurale legacy)
 *
 * I fogli "Risposte" e "Studenti" vengono creati da soli al primo utilizzo.
 */

/* ── MENU NEL FOGLIO ────────────────────────────────── */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 Corso IA')
    .addItem('Gestione Card', 'mostraGestione')
    .addSeparator()
    .addItem('📊 Aggiorna riepilogo studenti', 'generaRiepilogoStudenti')
    .addItem('📊 Aggiorna riepilogo per domanda', 'generaRiepilogoDomande')
    .addItem('📈 Dashboard grafica', 'mostraDashboard')
    .addToUi();
}

function mostraDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1100)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard — Corso IA');
}

function mostraGestione() {
  var html = HtmlService.createHtmlOutputFromFile('Gestione')
    .setWidth(1200)
    .setHeight(770);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gestione Card — Corso IA');
}

/* ── ENTRY POINT (app studente) ─────────────────────── */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(getTitoloModulo())
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ── LETTURA CONFIGURAZIONE (legge anche colonna config) ── */
function getCardConfig() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configurazione');
  if (!sheet) throw new Error('Foglio "Configurazione" non trovato.');

  var data  = sheet.getDataRange().getValues();
  var cards = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] == null || row[0] === '') continue;

    var opzioni = [];
    if (row[5]) {
      opzioni = row[5].toString().split('|').map(function (s) { return s.trim(); })
                      .filter(function (s) { return s !== ''; });
    }

    // colonna J: config JSON (media, parts, blanks…)
    var cfg = {};
    if (row[9]) { try { cfg = JSON.parse(row[9]); } catch (e) { cfg = {}; } }

    cards.push({
      id:               parseInt(row[0], 10),
      tipo:             row[1] ? row[1].toString().trim().toLowerCase() : 'narrazione',
      titolo:           row[2] ? row[2].toString() : '',
      testo:            row[3] ? row[3].toString() : '',
      mediaUrl:         row[4] ? row[4].toString().trim() : '',
      opzioni:          opzioni,
      rispostaCorretta: row[6] ? row[6].toString().trim().toLowerCase() : '',
      feedbackEsatto:   row[7] ? row[7].toString() : '',
      feedbackErrato:   row[8] ? row[8].toString() : '',
      media:            cfg.media   || '',     // 'immagine' | 'video' | ''
      parts:            cfg.parts   || null,   // per completamento
      blanks:           cfg.blanks  || null    // per completamento
    });
  }
  return cards;
}

/* ══════════════════════════════════════════════════════
   GESTIONE CARD — funzioni chiamate dall'interfaccia
   ══════════════════════════════════════════════════════ */

/* Assicura che la colonna J si chiami "config" */
function assicuraHeaderConfig_(sheet) {
  var n = Math.max(sheet.getLastColumn(), 10);
  var header = sheet.getRange(1, 1, 1, n).getValues()[0];
  if (!header[9] || header[9].toString().toLowerCase() !== 'config') {
    sheet.getRange(1, 10).setValue('config');
  }
  sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
}

/* Trasforma il testo di una riga di opzioni in array + risposta corretta.
   Convenzione: una opzione per riga, l'opzione giusta inizia con * */
function parseOpzioni_(raw) {
  var opzioni = [], correct = '';
  String(raw || '').split('\n').forEach(function (line) {
    var s = line.trim();
    if (!s) return;
    if (s.charAt(0) === '*') { s = s.substring(1).trim(); correct = s; }
    opzioni.push(s);
  });
  return { opzioni: opzioni, correct: correct };
}

/* Analizza una frase di completamento.
   Sintassi: ogni buco è [opzione1 / opzione2 / *opzioneGiusta]
   Es: "Il sole tramonta a [est / *ovest]." */
function parseDC_(testo) {
  var parts = [], blanks = [];
  var re = /\[([^\]]+)\]/g, last = 0, m;
  while ((m = re.exec(testo)) !== null) {
    if (m.index > last) parts.push({ t: 'text', v: testo.substring(last, m.index) });
    var correct = '';
    var opts = m[1].split('/').map(function (o) {
      o = o.trim();
      if (o.charAt(0) === '*') { o = o.substring(1).trim(); correct = o; }
      return o;
    }).filter(function (o) { return o !== ''; });
    blanks.push({ options: opts, correct: correct });
    parts.push({ t: 'blank', i: blanks.length - 1 });
    last = re.lastIndex;
  }
  if (last < testo.length) parts.push({ t: 'text', v: testo.substring(last) });
  return { parts: parts, blanks: blanks };
}

/* Salva (o aggiorna) una card. Riceve un oggetto dal form. */
function salvaCard(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Configurazione');
  if (!sheet) throw new Error('Foglio "Configurazione" non trovato.');
  assicuraHeaderConfig_(sheet);

  var tipo     = String(p.tipo || 'narrazione').trim().toLowerCase();
  var testo    = String(p.testo || '');
  var mediaUrl = String(p.mediaUrl || '').trim();
  var media    = mediaUrl ? (p.media || 'immagine') : '';

  var opzioniStr = '', rispostaCorretta = '';
  var feA = String(p.feedbackEsatto || ''), feE = String(p.feedbackErrato || '');
  var config = {};
  if (media) config.media = media;

  if (tipo === 'quiz') {
    var po = parseOpzioni_(p.opzioniRaw);
    opzioniStr = po.opzioni.join('|');
    rispostaCorretta = po.correct;
  } else if (tipo === 'completamento') {
    var dc = parseDC_(testo);
    config.parts = dc.parts;
    config.blanks = dc.blanks;
  }
  // 'aperta' e 'narrazione' non hanno opzioni né risposta corretta

  var configStr = Object.keys(config).length ? JSON.stringify(config) : '';

  // calcola id / trova riga esistente
  var data = sheet.getDataRange().getValues();
  var id = p.id ? parseInt(p.id, 10) : 0;
  var rowIndex = -1;
  if (id) {
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0], 10) === id) { rowIndex = i + 1; break; }
    }
  } else {
    var maxId = 0;
    for (var j = 1; j < data.length; j++) {
      var v = parseInt(data[j][0], 10);
      if (!isNaN(v) && v > maxId) maxId = v;
    }
    id = maxId + 1;
  }

  var values = [id, tipo, '', testo, mediaUrl, opzioniStr, rispostaCorretta, feA, feE, configStr];
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  else             sheet.appendRow(values);

  return elencaCard();
}

/* Restituisce l'elenco compatto delle card (per la lista nel form). */
function elencaCard() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configurazione');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var nomi = { narrazione: 'Narrazione', quiz: 'Risposta multipla', aperta: 'Risposta aperta',
               completamento: 'Completamento', capitolo: 'Capitolo', testo: 'Testo', video: 'Video', mappa_neurale: 'Mappa' };
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r[0] === '' || r[0] == null) continue;
    var t = String(r[1] || '').toLowerCase();
    var grezzo = String(r[3] || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g,' ').replace(/\s+/g,' ').trim();
    var anteprima = grezzo.replace(/\[[^\]]*\]/g, '___').slice(0, 80);
    out.push({
      id: parseInt(r[0], 10),
      tipo: t,
      tipoLabel: nomi[t] || r[1],
      anteprima: anteprima
    });
  }
  return out;
}

/* Carica una singola card per modificarla nel form. */
function caricaCard(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configurazione');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (parseInt(data[i][0], 10) === parseInt(id, 10)) {
      var r = data[i];
      var cfg = {};
      if (r[9]) { try { cfg = JSON.parse(r[9]); } catch (e) {} }

      // ricostruisce il testo delle opzioni (giusta con *)
      var opzioniRaw = '';
      if (r[5]) {
        var corr = String(r[6] || '').toLowerCase().trim();
        opzioniRaw = String(r[5]).split('|').map(function (o) {
          o = o.trim();
          return (o.toLowerCase() === corr ? '*' : '') + o;
        }).join('\n');
      }
      return {
        id: parseInt(r[0], 10),
        tipo: String(r[1] || '').toLowerCase(),
        testo: String(r[3] || ''),
        mediaUrl: String(r[4] || ''),
        media: cfg.media || '',
        opzioniRaw: opzioniRaw,
        feedbackEsatto: String(r[7] || ''),
        feedbackErrato: String(r[8] || '')
      };
    }
  }
  return null;
}

/* Sposta una card su (-1) o giù (+1) scambiando la riga con quella adiacente. */
function spostaCard(id, dir) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configurazione');
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][0] != null) rows.push(i + 1); // righe reali (1-based)
  }
  var pos = -1;
  for (var k = 0; k < rows.length; k++) {
    if (parseInt(data[rows[k] - 1][0], 10) === parseInt(id, 10)) { pos = k; break; }
  }
  if (pos === -1) return elencaCard();
  var target = pos + (dir < 0 ? -1 : 1);
  if (target < 0 || target >= rows.length) return elencaCard(); // già in cima/fondo
  var rA = rows[pos], rB = rows[target];
  var n = 10;
  var valsA = sheet.getRange(rA, 1, 1, n).getValues()[0];
  var valsB = sheet.getRange(rB, 1, 1, n).getValues()[0];
  sheet.getRange(rA, 1, 1, n).setValues([valsB]);
  sheet.getRange(rB, 1, 1, n).setValues([valsA]);
  return elencaCard();
}

/* Elimina una card per id. */
function eliminaCard(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configurazione');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (parseInt(data[i][0], 10) === parseInt(id, 10)) { sheet.deleteRow(i + 1); break; }
  }
  return elencaCard();
}

/* ── TITOLO DEL MODULO (tab "Impostazioni") ─────────── */
function getTitoloModulo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Impostazioni');
  if (!sheet) return 'Corso IA';
  var v = sheet.getRange('B1').getValue();
  return v ? v.toString() : 'Corso IA';
}

function salvaTitoloModulo(titolo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Impostazioni');
  if (!sheet) {
    sheet = ss.insertSheet('Impostazioni');
    sheet.getRange('A1').setValue('Titolo del modulo').setFontWeight('bold');
  }
  sheet.getRange('B1').setValue(titolo || '');
  return 'OK';
}

/* ── ESPORTA TUTTO IL CONTENUTO IN TESTO ───────────── */
function esportaTestoCorso() {
  var cards = getCardConfig();
  var titolo = getTitoloModulo();
  function pulito(html) {
    return String(html || '')
      .replace(/<\/(p|div|h4|li)>/gi, '\n')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/\n{3,}/g, '\n\n').trim();
  }
  var nomi = { narrazione: 'Card', quiz: 'Domanda a risposta multipla', aperta: 'Domanda aperta', completamento: 'Completamento' };
  var out = titolo.toUpperCase() + '\n' + '='.repeat(titolo.length) + '\n\n';
  cards.forEach(function (c, idx) {
    if (c.tipo === 'capitolo') {
      var tit = String(c.testo || '').replace(/<[^>]+>/g, '').trim();
      out += '\n\n########  ' + tit.toUpperCase() + '  ########\n\n';
      return;
    }
    out += '── Card ' + c.id + ' · ' + (nomi[c.tipo] || c.tipo) + ' ──\n';
    out += pulito(c.testo) + '\n';
    if (c.mediaUrl) out += '[media: ' + c.mediaUrl + ']\n';
    if (c.opzioni && c.opzioni.length) {
      out += 'Opzioni: ' + c.opzioni.join(' / ') + '\n';
      if (c.rispostaCorretta) out += 'Risposta corretta: ' + c.rispostaCorretta + '\n';
    }
    if (c.feedbackEsatto) out += 'Feedback (corretto): ' + pulito(c.feedbackEsatto) + '\n';
    if (c.feedbackErrato) out += 'Feedback (errato): ' + pulito(c.feedbackErrato) + '\n';
    out += '\n';
  });
  return out;
}

/* ── SALVA STUDENTE (primo accesso) ────────────────── */
function salvaStudente(nome, email) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Studenti');

  if (!sheet) {
    sheet = ss.insertSheet('Studenti');
    sheet.appendRow(['Timestamp', 'Nome', 'Email']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }

  var righe = sheet.getDataRange().getValues();
  for (var i = 1; i < righe.length; i++) {
    if (righe[i][2] && righe[i][2].toString().toLowerCase() === String(email).toLowerCase()) {
      return 'GIA_PRESENTE';
    }
  }
  sheet.appendRow([new Date(), nome, email]);
  return 'OK';
}

/* Dati aggregati per la dashboard grafica */
function getStatistiche() {
  var cards = getCardConfig();
  var risp = _risposte();
  var byId = {};
  risp.forEach(function (r) { (byId[r.id] = byId[r.id] || []).push(r); });

  var studentiSet = {};
  risp.forEach(function (r) { studentiSet[r.studente] = true; });

  var domande = [];
  cards.forEach(function (c) {
    if (['quiz', 'aperta', 'completamento'].indexOf(c.tipo) === -1) return;
    var lista = byId[c.id] || [];
    var tot = lista.length;
    var ok = lista.filter(function (r) { return r.esito === 'CORRETTO'; }).length;
    var d = {
      id: c.id, tipo: c.tipo, testo: _plain(c.testo),
      totale: tot, corrette: ok, pctCorrette: tot ? Math.round(ok / tot * 100) : 0,
      opzioni: [], aperte: []
    };
    if (c.tipo === 'quiz') {
      (c.opzioni || []).forEach(function (opz) {
        var n = lista.filter(function (r) { return r.risposta.toLowerCase().trim() === opz.toLowerCase().trim(); }).length;
        d.opzioni.push({ testo: opz, count: n, pct: tot ? Math.round(n / tot * 100) : 0,
          corretta: opz.toLowerCase().trim() === String(c.rispostaCorretta).toLowerCase().trim() });
      });
    } else if (c.tipo === 'aperta') {
      d.aperte = lista.slice(0, 100).map(function (r) { return { studente: r.studente, risposta: r.risposta }; });
    }
    domande.push(d);
  });

  var per = {};
  risp.forEach(function (r) {
    if (!per[r.studente]) per[r.studente] = { ok: 0, ko: 0, ap: 0 };
    if (r.esito === 'CORRETTO') per[r.studente].ok++;
    else if (r.esito === 'ERRATO') per[r.studente].ko++;
    else per[r.studente].ap++;
  });
  var studenti = Object.keys(per).sort().map(function (s) {
    var p = per[s], val = p.ok + p.ko;
    return { nome: s, ok: p.ok, ko: p.ko, ap: p.ap, pct: val ? Math.round(p.ok / val * 100) : 0 };
  });

  var totOk = 0, totVal = 0;
  studenti.forEach(function (s) { totOk += s.ok; totVal += s.ok + s.ko; });

  return {
    studentiTotali: Object.keys(studentiSet).length,
    risposteTotali: risp.length,
    pctGlobale: totVal ? Math.round(totOk / totVal * 100) : 0,
    titolo: getTitoloModulo(),
    domande: domande,
    studenti: studenti
  };
}

/* ── SALVA RISPOSTA ─────────────────────────────────── */
function salvaRispostaInSheet(studente, idCard, titoloCard, rispostaData, esito) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Risposte');

  if (!sheet) {
    sheet = ss.insertSheet('Risposte');
    sheet.appendRow(['Timestamp', 'Studente', 'ID Card', 'Domanda', 'Risposta Data', 'Esito']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  sheet.appendRow([new Date(), studente, idCard, titoloCard, rispostaData, esito]);
  // aggiorna i riepiloghi con i nuovi dati (senza bloccare il salvataggio)
  try { generaRiepiloghi(); } catch (e) {}
  return 'OK';
}

/* ══════════════════════════════════════════════════════
   RIEPILOGHI (per l'educatore) — menu "📚 Corso IA"
   ══════════════════════════════════════════════════════ */

/* Chiamata a fine modulo dall'app studente.
   I risultati del singolo studente sono già nel foglio "Risposte" e nei riepiloghi. */
function inviaRisultatiFinali(nome, email) {
  try { generaRiepiloghi(); } catch (e) {}
  return 'OK';
}

function generaRiepiloghi() {
  generaRiepilogoStudenti();
  generaRiepilogoDomande();
}

function _plain(html) {
  return String(html || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&').replace(/\[[^\]]*\]/g, '___').replace(/\s+/g, ' ').trim();
}
function _risposte() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Risposte');
  if (!sheet) return [];
  var d = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < d.length; i++) {
    if (d[i][1] === '' || d[i][1] == null) continue;
    out.push({ studente: String(d[i][1]), id: parseInt(d[i][2], 10), risposta: String(d[i][4] || ''), esito: String(d[i][5] || '') });
  }
  return out;
}
function _mappaCard() {
  var m = {};
  getCardConfig().forEach(function (c) { m[c.id] = c; });
  return m;
}

/* Tabella: una riga per studente con corrette / sbagliate / aperte / % */
function generaRiepilogoStudenti() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Riepilogo Studenti') || ss.insertSheet('Riepilogo Studenti');
  sh.clear();
  var risp = _risposte();
  var per = {}; // studente -> {ok, ko, ap}
  risp.forEach(function (r) {
    if (!per[r.studente]) per[r.studente] = { ok: 0, ko: 0, ap: 0 };
    if (r.esito === 'CORRETTO') per[r.studente].ok++;
    else if (r.esito === 'ERRATO') per[r.studente].ko++;
    else per[r.studente].ap++;
  });
  var rows = [['Studente', 'Corrette', 'Sbagliate', 'Aperte', 'Totale valutate', '% corrette']];
  Object.keys(per).sort().forEach(function (s) {
    var p = per[s], val = p.ok + p.ko;
    var pct = val ? Math.round(p.ok / val * 100) + '%' : '—';
    rows.push([s, p.ok, p.ko, p.ap, val, pct]);
  });
  if (rows.length === 1) rows.push(['(nessuna risposta ancora)', '', '', '', '', '']);
  sh.getRange(1, 1, rows.length, 6).setValues(rows);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#EDEBFB');
  sh.setColumnWidth(1, 220);
  sh.setFrozenRows(1);
}

/* Stile Moduli Google: per ogni domanda, conteggio risposte per opzione + % corrette */
function generaRiepilogoDomande() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Riepilogo Domande') || ss.insertSheet('Riepilogo Domande');
  sh.clear();
  var cards = getCardConfig();
  var risp = _risposte();
  var byId = {};
  risp.forEach(function (r) { (byId[r.id] = byId[r.id] || []).push(r); });

  var out = [['RIEPILOGO RISPOSTE PER DOMANDA', '', '']];
  var bold = [0];
  cards.forEach(function (c) {
    if (['quiz', 'aperta', 'completamento'].indexOf(c.tipo) === -1) return;
    var lista = byId[c.id] || [];
    out.push(['', '', '']);
    bold.push(out.length);
    out.push(['D' + c.id + ' · ' + _plain(c.testo), '', '']);

    if (c.tipo === 'quiz') {
      var tot = lista.length;
      (c.opzioni || []).forEach(function (opz) {
        var n = lista.filter(function (r) { return r.risposta.toLowerCase().trim() === opz.toLowerCase().trim(); }).length;
        var pct = tot ? Math.round(n / tot * 100) + '%' : '0%';
        var segno = (opz.toLowerCase().trim() === String(c.rispostaCorretta).toLowerCase().trim()) ? '  ✓ corretta' : '';
        out.push(['   ' + opz + segno, n, pct]);
      });
      var ok = lista.filter(function (r) { return r.esito === 'CORRETTO'; }).length;
      out.push(['   → risposte totali: ' + tot + ' · corrette: ' + ok + (tot ? ' (' + Math.round(ok / tot * 100) + '%)' : ''), '', '']);
    } else if (c.tipo === 'completamento') {
      var ok2 = lista.filter(function (r) { return r.esito === 'CORRETTO'; }).length;
      out.push(['   completate correttamente: ' + ok2 + ' su ' + lista.length, '', '']);
    } else { // aperta
      out.push(['   risposte ricevute: ' + lista.length, '', '']);
      lista.slice(0, 50).forEach(function (r) { out.push(['   • ' + r.studente + ': ' + r.risposta, '', '']); });
    }
  });
  if (out.length === 1) out.push(['(nessuna domanda con risposte ancora)', '', '']);
  sh.getRange(1, 1, out.length, 3).setValues(out);
  bold.forEach(function (i) { sh.getRange(i + 1, 1, 1, 3).setFontWeight('bold'); });
  sh.getRange(1, 1, 1, 3).setBackground('#EDEBFB');
  sh.setColumnWidth(1, 520); sh.setColumnWidth(2, 90); sh.setColumnWidth(3, 90);
  sh.setFrozenRows(1);
}