// ============================================================
//  hso-daten.js  –  Clientseitige Datenschicht für HSO-Portal + intern
//
//  Ersetzt die live-Backend-Aufrufe für die Aktionen "projekte",
//  "programm" und "kalender" durch Auswertung zweier statischer
//  JSON-Dateien (uebersetzer.json, projekte.json), die nachts per
//  Apps-Script-Trigger (uebersetzerAktualisierenUndPushen) neu
//  nach GitHub gepusht werden. Grund: das gemeinsame Apps-Script-
//  Web-App-Deployment hat ein Limit für gleichzeitige Ausführungen
//  und antwortete bei mehreren gleichzeitigen Nutzer:innen zeitweise
//  mit 404/Timeout. GitHub Pages hat dieses Limit nicht.
//
//  Andere Aktionen (Login, Tasks, Probendispositionen, hfmtProxy,
//  Programmheft, Werkvertrag, ...) bleiben unverändert live gegen
//  das Apps-Script-Backend.
//
//  Portiert aus Projectsync.gs: studienjahr(), studienjahrgrenzen(),
//  aktuellesStudienjahr(), projektNameAusTitel(), _sortiergewicht(),
//  _getProjekte(), _getProgramm(), _getKalender(), _eventGehoertZuProjekt(),
//  _parseDatumStr() — Logik 1:1 übernommen, nur die Datenquelle ist
//  jetzt das geladene JSON statt Sheet/Kalender direkt.
//
//  Einbindung: <script>window.HSO_DATEN_BASE='./';</script> (oder '../'
//  für Unterordner) VOR <script src=".../hso-daten.js"></script>.
// ============================================================

(function (global) {
  'use strict';

  var BASE = global.HSO_DATEN_BASE || './';

  // Aus Config.gs (Apps-Script-Projekt "Abfrage Google Kalender HFMT") übernommen.
  var KEIN_EIGENES_PROJEKT = ['Aufbau', 'Abbau', 'Anspielprobe', 'Workshop'];
  var PROJEKT_REIHENFOLGE = [
    'HSO 1', 'HSO 2', 'HSO 3', 'HSO 4',
    '13/14',
    'HSO Blattspiel', 'HSO Repertoire', 'HSO Repertoireprobe',
    'WK 1', 'WK 2', 'WK 3', 'WK 4', 'WK 5', 'WK 6', 'WK 7', 'WK 8',
    'WK Lü 1', 'WK Lü 2', 'WK Lü 3', 'WK Lü 4'
  ];

  // ── STUDIENJAHR-HILFSFUNKTIONEN (aus Projectsync.gs) ──────────

  function studienjahr(datum) {
    var m = datum.getMonth();
    var y = datum.getFullYear();
    if (m >= 8) {
      return String(y).slice(2) + '_' + String(y + 1).slice(2);
    } else {
      return String(y - 1).slice(2) + '_' + String(y).slice(2);
    }
  }

  function studienjahrgrenzen(sjStr) {
    var t = sjStr.split('_');
    return {
      von: new Date(2000 + parseInt(t[0], 10), 8, 1),
      bis: new Date(2000 + parseInt(t[1], 10), 6, 31, 23, 59, 59)
    };
  }

  function aktuellesStudienjahr() {
    return studienjahr(new Date());
  }

  function projektNameAusTitel(titel) {
    for (var i = 0; i < KEIN_EIGENES_PROJEKT.length; i++) {
      if (titel.indexOf(KEIN_EIGENES_PROJEKT[i]) === 0) return null;
    }
    var mZusatz = titel.match(/^\*(.+)/);
    if (mZusatz) return '*' + mZusatz[1].trim();

    var mLue = titel.match(/^(WK\s+L[üu]\s+\d+)/);
    if (mLue) return mLue[1];
    var mNum = titel.match(/^((?:HSO|WK)\s+\d+)/);
    if (mNum) return mNum[1];
    var m1314 = titel.match(/^(13\/14)/);
    if (m1314) return m1314[1];
    var mHso = titel.match(/^(HSO\s+\S+)/);
    if (mHso) return mHso[1];
    return null;
  }

  function _sortiergewicht(projektName) {
    var idx = PROJEKT_REIHENFOLGE.indexOf(projektName);
    if (idx !== -1) return idx;
    return 1000 + projektName.charCodeAt(0);
  }

  function _eventGehoertZuProjekt(eventTitel, projektName) {
    if (!eventTitel || !projektName) return false;
    var titel = eventTitel.replace(/^\*/, '');
    var pName = projektName.replace(/^\*/, '');
    var re = new RegExp(
      '^' + pName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\s|\\(|$)'
    );
    return re.test(titel);
  }

  function _parseDatumStr(s) {
    if (!s) return null;
    var m = String(s).match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  // ── LADEN + CACHEN DER STATISCHEN JSON-DATEIEN ────────────────

  var _ladenPromise = null;
  var _events = null;   // uebersetzer.json → events[]
  var _projekte = null; // projekte.json → projekte[] ({name, docId, sj})

  var _ladenUnterrichtPromise = null;
  var _unterrichtTermine = null; // unterricht-termine.json → data[]

  function ladeJson(datei) {
    var url = BASE + datei + '?t=' + Date.now();
    return fetch(url, { cache: 'no-store' }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' bei ' + datei);
      return r.json();
    });
  }

  function laden() {
    if (_ladenPromise) return _ladenPromise;
    _ladenPromise = Promise.all([
      ladeJson('uebersetzer.json'),
      ladeJson('projekte.json')
    ]).then(function (res) {
      _events = (res[0] && res[0].events) || [];
      _projekte = (res[1] && res[1].projekte) || [];
    }).catch(function (e) {
      // Beim nächsten Aufruf erneut versuchen statt dauerhaft zu scheitern.
      _ladenPromise = null;
      throw e;
    });
    return _ladenPromise;
  }

  function ladenUnterricht() {
    if (_ladenUnterrichtPromise) return _ladenUnterrichtPromise;
    _ladenUnterrichtPromise = ladeJson('unterricht-termine.json').then(function (res) {
      _unterrichtTermine = (res && res.data) || [];
    }).catch(function (e) {
      // Beim nächsten Aufruf erneut versuchen statt dauerhaft zu scheitern.
      _ladenUnterrichtPromise = null;
      throw e;
    });
    return _ladenUnterrichtPromise;
  }

  // ── PORTIERTE PROJEKTIONEN (aus Projectsync.gs) ────────────────

  function _getProjekte(sjFilter) {
    var heute = new Date();
    heute.setHours(0, 0, 0, 0);
    var sjFuerFilter = sjFilter || aktuellesStudienjahr();
    var grenzenFilter = studienjahrgrenzen(sjFuerFilter);

    var projMitEvent = {};
    var projLetzterEvent = {};
    _events.forEach(function (ev) {
      var pName = projektNameAusTitel(ev._titel || '');
      if (!pName) return;
      var d = _parseDatumStr(ev.datum);
      if (!d) return;
      if (d < grenzenFilter.von || d > grenzenFilter.bis) return;
      if (!projLetzterEvent[pName] || d > projLetzterEvent[pName]) {
        projLetzterEvent[pName] = d;
      }
      projMitEvent[pName] = true;
    });

    var result = [];
    for (var i = 0; i < _projekte.length; i++) {
      var name = String(_projekte[i].name || '').trim();
      var docId = String(_projekte[i].docId || '').trim();
      var sj = String(_projekte[i].sj || '').trim();
      if (!name || !docId) continue;
      if (sjFilter && sj !== sjFilter) continue;
      if (!projMitEvent[name]) continue;
      var letzter = projLetzterEvent[name];
      if (letzter) {
        var diffTage = (heute.getTime() - letzter.getTime()) / 86400000;
        if (diffTage > 14) continue;
      }
      result.push({ name: name, docId: docId, sj: sj, zusatz: name.indexOf('*') === 0 });
    }

    result.sort(function (a, b) {
      return _sortiergewicht(a.name) - _sortiergewicht(b.name);
    });

    return result;
  }

  function _getProgramm(sjFilter) {
    var sj = sjFilter || aktuellesStudienjahr();
    var grenzen = studienjahrgrenzen(sj);

    var eventsImSj = _events.filter(function (ev) {
      var d = _parseDatumStr(ev.datum);
      if (!d) return false;
      return d >= grenzen.von && d <= grenzen.bis;
    });

    var projekte = _getProjekte(sj);
    var result = {};

    for (var p = 0; p < projekte.length; p++) {
      var pName = projekte[p].name;
      var istBR = pName.indexOf('Blattspiel') !== -1 || pName.indexOf('Repertoire') !== -1;

      var zugehoerig = eventsImSj.filter(function (ev) {
        return _eventGehoertZuProjekt(ev._titel, pName);
      });

      if (zugehoerig.length === 0) continue;

      if (istBR) {
        result[pName] = zugehoerig.map(function (ev) {
          return {
            _eventId: ev._eventId,
            _titel: ev._titel,
            datum: ev.datum,
            uhrzeit: ev.uhrzeit,
            ort: ev.ort,
            marker: ev.marker,
            werke: ev.werke || [],
            solisten: ev.solisten || [],
            leitung: ev.leitung || [],
            probenplan: ev.probenplan || []
          };
        });
      } else {
        var vorstellungen = zugehoerig.filter(function (ev) {
          return /\(\d+\)/.test(ev._titel || '') ||
            (ev.marker && ev.marker.toLowerCase() === 'vorstellung');
        });

        if (vorstellungen.length === 0) {
          vorstellungen = zugehoerig.filter(function (ev) {
            return ev.werke && ev.werke.length > 0;
          });
        }

        if (vorstellungen.length === 0) continue;

        result[pName] = vorstellungen.map(function (ev) {
          return {
            _eventId: ev._eventId,
            _titel: ev._titel,
            _titelNotiz: ev._titelNotiz || null,
            datum: ev.datum,
            uhrzeit: ev.uhrzeit,
            ort: ev.ort,
            konzertart: ev.konzertart,
            marker: ev.marker,
            werke: ev.werke || [],
            solisten: ev.solisten || [],
            leitung: ev.leitung || [],
            ensemble: ev.ensemble || null,
            probenplan: ev.probenplan || []
          };
        });
      }
    }

    return result;
  }

  function _getKalender(sjFilter) {
    var sj = sjFilter || aktuellesStudienjahr();
    var grenzen = studienjahrgrenzen(sj);
    var result = [];
    _events.forEach(function (ev) {
      var d = _parseDatumStr(ev.datum);
      if (!d || d < grenzen.von || d > grenzen.bis) return;
      result.push({
        _eventId: ev._eventId,
        _titel: ev._titel,
        datum: ev.datum,
        uhrzeit: ev.uhrzeit,
        ort: ev.ort,
        konzertart: ev.konzertart,
        marker: ev.marker,
        werke: ev.werke || [],
        solisten: ev.solisten || [],
        leitung: ev.leitung || [],
        ensemble: ev.ensemble || null,
        abwesend: ev.abwesend || null
      });
    });
    return result;
  }

  // ── ÖFFENTLICHE API ────────────────────────────────────────────
  // holen(action, sj) liefert dasselbe Format wie bisher
  // `fetchMitRetry(...).then(r=>r.json())` gegen das Live-Backend:
  // { ok: true, data: ... } bzw. { ok: false, error: '...' }.

  function holen(action, sj) {
    if (action === 'unterrichtTermine') {
      return ladenUnterricht().then(function () {
        return { ok: true, data: _unterrichtTermine };
      }).catch(function (e) {
        return { ok: false, error: e.message || String(e) };
      });
    }
    return laden().then(function () {
      if (action === 'projekte') return { ok: true, data: _getProjekte(sj || null) };
      if (action === 'programm') return { ok: true, data: _getProgramm(sj || null) };
      if (action === 'kalender') return { ok: true, data: _getKalender(sj || null) };
      return { ok: false, error: 'HsoDaten: unbekannte action ' + action };
    }).catch(function (e) {
      return { ok: false, error: e.message || String(e) };
    });
  }

  global.HsoDaten = {
    holen: holen,
    // Für den Fall, dass eine Seite die Hilfsfunktionen selbst braucht:
    studienjahr: studienjahr,
    studienjahrgrenzen: studienjahrgrenzen,
    aktuellesStudienjahr: aktuellesStudienjahr
  };

})(window);

