// ============================================================================
// Master_Cache.gs — Crescentia HealthTech / CresRx
// Reference lists read once every few minutes, not on every screen.
// ----------------------------------------------------------------------------
// The drug master, the lab test list, the clinical phrase bank, the package
// rates and the doctor list are read in full by nearly every clinical screen:
// opening a consult reads all five, and a busy morning opens a consult every
// few minutes on every desk. They change a handful of times a month.
//
// Reading a sheet is the slowest thing Apps Script does (hundreds of ms), and
// the script cache answers in a few. So each list is built once and kept for
// MCACHE_TTL_S, shared by every user.
//
// STALENESS. Three things end a cached copy early:
//   * the writers in this project that change a list call crescMasterBust_();
//   * an edit made by hand in the spreadsheet fires onEdit (below), which
//     busts the list for that sheet;
//   * Admin Dashboard -> Operations -> "Reload reference lists" clears all.
// Anything else is at most MCACHE_TTL_S old.
//
// The cache holds reference data only — no patient appears in any of these
// lists — so sharing one copy between users discloses nothing.
// ============================================================================

var MCACHE_TTL_S = 600;

/** Cache value limit is 100 KB; parts stay under it with room for UTF-8. */
var MCACHE_PART = 90000;

/** Sheet -> the cached lists built from it. */
var MCACHE_SOURCES = {
  'Drug_Master_Universal': ['drugs'],
  'Lab_Test_Master':       ['labtests'],
  'Clinical_Templates':    ['phrases'],
  'Package_Master':        ['packages'],
  'Doctors':               ['doctors']
};

function mcache_ver_(cache, name) {
  return cache.get('MCV_' + name) || '0';
}

/**
 * The cached value of one list, or build() it and cache the result.
 * Never throws for a cache fault: a cache that cannot be read is a slow
 * read, not a failure.
 *
 * @param {string} name     one of the names in MCACHE_SOURCES
 * @param {function} build  returns a JSON-serialisable value
 */
function crescMasterGet_(name, build) {
  var cache = null, base = '';
  try {
    cache = CacheService.getScriptCache();
    base = 'MC1_' + name + '_' + mcache_ver_(cache, name);
    var head = cache.get(base);
    if (head) {
      var n = parseInt(head, 10);
      var keys = [];
      for (var i = 0; i < n; i++) keys.push(base + '_' + i);
      var parts = cache.getAll(keys);
      var s = '';
      for (var j = 0; j < n; j++) {
        if (parts[keys[j]] === undefined || parts[keys[j]] === null) { s = null; break; }
        s += parts[keys[j]];
      }
      if (s !== null) return JSON.parse(s);
    }
  } catch (e) { /* fall through to a fresh read */ }

  var value = build();
  // An empty list or a failed read is not cached: it is far more likely to
  // be a transient fault than a master that has really been emptied, and
  // caching it would blank every desk's suggestions for ten minutes.
  var empty = (Array.isArray(value) && !value.length) ||
              (value && typeof value === 'object' && value.success === false);
  try {
    if (cache && !empty) {
      var json = JSON.stringify(value);
      var put = {}, count = 0;
      for (var k = 0; k < json.length; k += MCACHE_PART) {
        put[base + '_' + count] = json.substr(k, MCACHE_PART);
        count++;
      }
      // Parts first, then the head, so a reader never finds a head whose
      // parts are not there yet.
      if (count <= 20) {
        cache.putAll(put, MCACHE_TTL_S);
        cache.put(base, String(count), MCACHE_TTL_S);
      }
    }
  } catch (e) {}
  return value;
}

/** Forget a list (or every list built from a sheet). Never throws. */
function crescMasterBust_(nameOrSheet) {
  try {
    var cache = CacheService.getScriptCache();
    var names = MCACHE_SOURCES[nameOrSheet] || [nameOrSheet];
    names.forEach(function (n) { cache.put('MCV_' + n, Utilities.getUuid().slice(0, 8), 21600); });
  } catch (e) {}
}

/**
 * FRONTEND ENTRY. Reload every reference list now — for after a batch of
 * edits made straight into the spreadsheet.
 */
function crescMasterReload(sessionToken) {
  try {
    crescRequire_(sessionToken, 'admin.config');
    var all = [];
    Object.keys(MCACHE_SOURCES).forEach(function (s) { all = all.concat(MCACHE_SOURCES[s]); });
    all.forEach(crescMasterBust_);
    return { success: true, message: 'Reference lists will be read fresh on next use (' + all.join(', ') + ').' };
  } catch (err) {
    return { success: false, message: cresc_reason_(err) };
  }
}

/**
 * SIMPLE TRIGGER. A hand edit in the spreadsheet ends the cached copy of
 * whatever list that sheet feeds.
 *
 * Reachable from google.script.run like every public function, which is
 * harmless: the browser cannot pass a Range, so a call from there busts
 * nothing, and busting is only ever a slower next read.
 */
function onEdit(e) {
  try {
    if (!e || !e.range || typeof e.range.getSheet !== 'function') return;
    var name = e.range.getSheet().getName();
    if (MCACHE_SOURCES[name]) crescMasterBust_(name);
  } catch (err) {}
}
