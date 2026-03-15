// ── Unrack — Parser Module ────────────────────────────────────────────
// Extracted from index.html for standalone testing and modularity.
// This file is the single source of truth for all parser logic.
// In production, include via <script src="parsers.js"></script> before the main script.

// ── LIFT GROUPS & CLASSIFICATION ──────────────────────────────────────────────
const LIFT_GROUPS = {
  squat: ['squat','front squat','zercher','belt squat','box squat','high bar squat','high bar','low bar squat','low bar','pin squat','pause squat','safety bar','ssb','1/4 squat','quarter squat','partial squat',
    'hatfield','walkout'],
  bench: ['bench','close grip bench','close grip press','close grip','larson','larsen','slingshot','incline press','incline bench','decline press','decline bench','floor press','feet up bench','pause bench','touch and go bench','touch and go press','tng bench','tng press','comp bench','board press',
    'spoto','spoto press','dead bench','pin press','cgbp','wgbp','rgbp'],
  deadlift: ['deadlift','sumo deadlift','sumo pull','sumo dead','conventional deadlift','conventional pull','romanian','rdl','rack pull','block pull','deficit deadlift','deficit pull','stiff leg','trap bar','hex bar','snatch grip','pause dead','semi sumo',
    'halting','mat pull','pin pull','sldl','sgdl'],
};

const COMP_KEYWORDS = {
  squat:    ['squat','back squat','low bar squat','barbell squat','comp squat','competition squat','barbell back squat','barbell low bar squat','barbell comp squat','barbell competition squat'],
  bench:    ['bench','bench press','barbell bench','barbell bench press','comp bench','comp bench press','barbell comp bench','barbell comp bench press','competition bench','competition bench press','paused bench','paused bench press','pause bench','pause bench press'],
  deadlift: ['deadlift','barbell deadlift','conventional deadlift','comp deadlift','competition deadlift','conventional pull','comp pull','competition style deadlift','barbell conventional deadlift','barbell comp deadlift'],
};

const VARIATION_MODIFIERS = [
  'pin','close grip','wide grip','tempo','incline','decline','deficit',
  'rack pull','block pull','snatch grip','stiff leg','romanian','rdl',
  'floor press','board press','slingshot','sling shot','feet up',
  'touch and go','tng','safety bar','ssb','zercher','front squat',
  'goblet','hack','belt squat','box squat','semi sumo',
  '1/4','quarter','partial',
  'pause squat','paused squat',
  'pause deadlift','paused deadlift','pause dead','paused dead',
  // Named bench variations
  'spoto','larsen','larson','jm press','california','guillotine','dead bench','reverse grip',
  // Specialty bars
  'cambered','buffalo','duffalo','axle','earthquake','bamboo','swiss bar','football bar','kadillac','spider bar','transformer bar',
  // Deadlift variations
  'dimel','halting','mat pull','pin pull','clean pull','snatch pull','reeves','suitcase',
  // Accommodating resistance
  'banded','chain','reverse band','weight releaser',
  // Squat variations
  'hatfield','cyclist','overhead squat','ohs','zombie','walkout','jefferson',
  // Effort/touch modifiers
  'dead stop','speed','dynamic effort','eccentric','isometric','floating','1.5 rep',
  // Stance/position
  'wide stance','narrow stance','sumo stance','sumo deadlift','sumo pull','sumo dead','heel elevated','beltless',
  // High bar squat — variation, not comp (some lifters use as comp but default to variation)
  'high bar',
];

// ── EXERCISE ALIAS DICTIONARY ────────────────────────────────────────────────
const EXERCISE_ALIASES = {
  // Exact shorthand aliases — full string match (case-insensitive)
  exact: {
    'comp squat': 'Competition Squat',
    'comp bench': 'Competition Bench Press',
    'comp deadlift': 'Competition Deadlift',
    'comp dl': 'Competition Deadlift',
    'pause squat': 'Paused Squat',
    'pause bench': 'Paused Bench Press',
    'pin squat': 'Pin Squat',
    'spoto press': 'Spoto Press',
    'sumo': 'Sumo Deadlift',
    'conv': 'Conventional Deadlift',
    'front sq': 'Front Squat',
    'rear delt fly': 'Rear Delt Fly',
    'face pull': 'Face Pull',
    'pull up': 'Pull-Up',
    'chin up': 'Chin-Up',
    'dip': 'Dip',
    't-bar row': 'T-Bar Row',
    'pend row': 'Pendlay Row',
    'lat pull': 'Lat Pulldown',
    'lat pd': 'Lat Pulldown',
  },
  // Word-level abbreviation expansions (case-insensitive)
  abbrevs: {
    'bb': 'Barbell',
    'db': 'Dumbbell',
    'kb': 'Kettlebell',
    'ohp': 'Overhead Press',
    'rdl': 'Romanian Deadlift',
    'sldl': 'Stiff-Leg Deadlift',
    'dl': 'Deadlift',
    'sq': 'Squat',
    'bp': 'Bench Press',
    'cg': 'Close Grip',
    'cgbp': 'Close Grip Bench Press',
    'tri': 'Tricep',
    'bi': 'Bicep',
    'inc': 'Incline',
    'dec': 'Decline',
    'ext': 'Extension',
    'ham': 'Hamstring',
    'pd': 'Pulldown',
  },
};

function isCompLift(name){
  // Strip dedup bracket suffix: "Barbell Deadlift [1x1(352)]" → "barbell deadlift"
  const n = name.toLowerCase().trim().replace(/\s*\[.*\]$/, '');
  if (VARIATION_MODIFIERS.some(v => n.includes(v))) return false;
  for (const keywords of Object.values(COMP_KEYWORDS)) {
    if (keywords.includes(n)) return true;
  }
  return false;
}
// Helper: extract base name from potentially deduplicated exercise name
function _getBaseName(exName) {
  const m = exName.match(/^(.+?)\s*\[.*\]$/);
  return m ? m[1] : exName;
}

// Keywords that disqualify a LIFT_GROUPS match — exercise is an accessory, not a variation
const LIFT_GROUP_DISQUALIFIERS = [
  // Accessory movement patterns (prevent false positives from base keywords like 'bench', 'squat')
  'pulldown','pull-down','pull down','lat pull','cable row','dip','kickback',
  'curl','fly','flye','raise','extension','pushdown','push-down','push down',
  'pullover','pull-over','shrug','row',
  // Squat-pattern accessories (contain 'squat' but aren't barbell squat variations)
  'goblet','hack squat','cyclist','overhead squat','ohs','jefferson squat','zombie',
  // Bench-pattern accessories (contain 'bench' or 'press' but aren't bench variations)
  'jm press','california press','guillotine',
  // Deadlift-pattern accessories (contain 'deadlift' but aren't comp deadlift variations)
  'clean pull','snatch pull','hack lift','dimel','reeves deadlift','suitcase deadlift',
  'zercher deadlift','jefferson deadlift',
];
// Base classifier — returns group name or null. No access to runtime state.
// index.html layers on custom lifts and 'other' fallback via classifyLift().
function _classifyLiftBase(name){
  const n=name.toLowerCase();
  // If name contains a disqualifying word, it's an accessory regardless of keyword match
  const disqualified = LIFT_GROUP_DISQUALIFIERS.some(d=>n.includes(d));
  if(!disqualified){
    for(const [group,keywords] of Object.entries(LIFT_GROUPS)){
      if(keywords.some(k=>n.includes(k))) return group;
    }
  }
  return null;
}

// ── EXERCISE NAME SYNONYMS ───────────────────────────────────────────────────
// Maps alternate names → one canonical display name.
// Used for metrics grouping, var-pills, tracked lifts, and e1RM aggregation.
// Lookup is case-insensitive; keys must be lowercase.
const SYNONYM_MAP = {
  // ── Comp Squat synonyms ──
  'squat':                'Squat',
  'back squat':           'Squat',
  'barbell squat':        'Squat',
  'competition squat':    'Squat',
  'comp squat':           'Squat',
  'low bar squat':        'Squat',
  'squats':               'Squat',

  // ── Comp Bench synonyms ──
  'bench press':          'Bench Press',
  'bench':                'Bench Press',
  'barbell bench':        'Bench Press',
  'barbell bench press':  'Bench Press',
  'comp bench':           'Bench Press',
  'comp bench press':     'Bench Press',
  'barbell comp bench':   'Bench Press',
  'barbell comp bench press': 'Bench Press',
  'competition bench':    'Bench Press',
  'competition bench press': 'Bench Press',
  'competition pause bench': 'Bench Press',
  'paused bench':         'Bench Press',
  'paused bench press':   'Bench Press',
  'pause bench':          'Bench Press',
  'pause bench press':    'Bench Press',
  'raw bench':            'Bench Press',

  // ── Comp Deadlift synonyms ──
  'deadlift':             'Deadlift',
  'barbell deadlift':     'Deadlift',
  'comp deadlift':        'Deadlift',
  'competition deadlift': 'Deadlift',
  'competition style deadlift': 'Deadlift',
  'conv deadlift':        'Deadlift',
  'conv. deadlift':       'Deadlift',
  'conventional deadlift':'Deadlift',
  'deadlifts':            'Deadlift',

  // ── High Bar Squat (variation, not comp) ──
  'high bar':             'High Bar Squat',
  'high bar squats':      'High Bar Squat',
  'hb squat':             'High Bar Squat',

  // ── Low Bar Squat (comp variant) ──
  'low bar':              'Squat',

  // ── Pause Squat synonyms ──
  'paused squat':         'Pause Squat',
  'squat 2ct pause':      'Pause Squat',
  'paused squat, 2ct':    'Pause Squat',
  'paused squat, 2 count':'Pause Squat',
  'squat, 2ct paused':    'Pause Squat',
  'double paused squat':  'Pause Squat',
  'dbl paused squat':     'Pause Squat',
  'short pause squat':    'Pause Squat',

  // ── TNG Bench synonyms ──
  'tng bench':            'Touch and Go Bench',
  'bench tng':            'Touch and Go Bench',
  'barbell tng bench':    'Touch and Go Bench',
  'tng bench press':      'Touch and Go Bench',
  't-shirt bench press':  'Touch and Go Bench',

  // ── Close Grip Bench synonyms ──
  'close grip':              'Close Grip Bench',
  'close grip bench press':  'Close Grip Bench',
  'close-grip bench':        'Close Grip Bench',
  'close-grip bench press':  'Close Grip Bench',
  'bench close grip':        'Close Grip Bench',
  'cgbp':                    'Close Grip Bench',

  // ── Pause Deadlift synonyms ──
  'paused deadlift':         'Pause Deadlift',
  'deadlift w/ pause':       'Pause Deadlift',
  'deadlift w/pause':        'Pause Deadlift',
  'paused deadlift, 2ct':    'Pause Deadlift',

  // ── Larsen Press synonyms ──
  'larson press':             'Larsen Press',

  // ── Stiff Leg Deadlift synonyms ──
  'stiff legged deadlift':   'Stiff Leg Deadlift',
  'stiff legged deadlifts':  'Stiff Leg Deadlift',
  'stiff-legged deadlift':   'Stiff Leg Deadlift',
  'sldl':                    'Stiff Leg Deadlift',
  'sl deadlift':             'Stiff Leg Deadlift',

  // ── SSB Squat synonyms ──
  'safety bar squat':         'SSB Squat',

  // ── Snatch Grip Deadlift synonyms ──
  'snatch grip conv. deadlift': 'Snatch Grip Deadlift',
  'sgdl':                       'Snatch Grip Deadlift',

  // ── Pin Squat synonyms ──
  'pin squats':               'Pin Squat',
  'squat pin':                'Pin Squat',
  'squat to pins':            'Pin Squat',

  // ── Sumo Deadlift synonyms (variation by default, user can toggle on) ──
  'sumo pull':                'Sumo Deadlift',
  'sumo dead':                'Sumo Deadlift',

  // ── Paused Sumo Deadlift synonyms ──
  'paused sumo deadlift, 2ct':'Paused Sumo Deadlift',

  // ── RDL synonyms ──
  'romanian deadlift':        'RDL',

  // ── Rack Pull synonyms ──
  'rack pulls':               'Rack Pull',

  // ── Deficit Deadlift synonyms ──
  'deficit deadlifts':        'Deficit Deadlift',
  'deficit pull':             'Deficit Deadlift',
  'defecit deadlift':         'Deficit Deadlift',

  // ── Incline Bench synonyms ──
  'incline bench press':      'Incline Bench',
  'incline press':            'Incline Bench',

  // ── Decline Bench synonyms ──
  'decline bench press':      'Decline Bench',
  'decline press':            'Decline Bench',

  // ── Floor Press synonyms ──
  'close grip floor press':   'Floor Press',

  // ── Abbreviations ──
  'dl':   'Deadlift',
  'rdl':  'RDL',
  'cgb':  'Close Grip Bench',
  'wb':   'Wide Grip Bench',
  'wgb':  'Wide Grip Bench',
  'wgbp': 'Wide Grip Bench Press',
  'tng':  'Touch and Go Bench',
  'tag':  'Touch and Go Bench',
  'ssb':  'SSB Squat',

  // ── Verb-Adjective Variations ──
  'paused bench press':    'Bench Press',
  'paused deadlift':       'Pause Deadlift',

  // ── Compound Synonyms ──
  'conventional pull':     'Deadlift',
  'rack deadlift':         'Rack Pull',

  // ── Bench Variations ──
  'spoto':                 'Spoto Press',
  'spoto press':           'Spoto Press',
  'jm press':              'JM Press',
  'comp dl':               'Deadlift',

  // ── Equipment prefix abbreviations ──
  'bb squat':              'Squat',
  'bb bench':              'Bench Press',
  'bb bench press':        'Bench Press',
  'bb deadlift':           'Deadlift',
  'bb row':                'Barbell Row',
  'bb ohp':                'Overhead Press',
  'db bench':              'Dumbbell Bench Press',
  'db bench press':        'Dumbbell Bench Press',
  'db row':                'Dumbbell Row',
  'db curl':               'Dumbbell Curl',
  'db ohp':                'Dumbbell Overhead Press',
  'db press':              'Dumbbell Press',
  'db fly':                'Dumbbell Fly',
  'db flye':               'Dumbbell Fly',
  'db lateral raise':      'Dumbbell Lateral Raise',
  'db rdl':                'Dumbbell RDL',
  'kb swing':              'Kettlebell Swing',
  'kb goblet squat':       'Goblet Squat',

  // ── Common abbreviations ──
  'ohp':                   'Overhead Press',
  'overhead press':        'Overhead Press',
  'bb overhead press':     'Overhead Press',
  'military press':        'Overhead Press',
  'strict press':          'Overhead Press',
  'seated ohp':            'Seated Overhead Press',
  'btn press':             'Behind the Neck Press',
  'fs':                    'Front Squat',
  'front squat':           'Front Squat',
  'bss':                   'Bulgarian Split Squat',
  'bulgarian split squat': 'Bulgarian Split Squat',
  'hip thrust':            'Hip Thrust',
  'ht':                    'Hip Thrust',
  'bb hip thrust':         'Barbell Hip Thrust',
  'ghd':                   'GHD',
  'ghr':                   'Glute Ham Raise',
  'rev hyper':             'Reverse Hyper',
  'reverse hyper':         'Reverse Hyper',
  'good morning':          'Good Morning',
  'gm':                    'Good Morning',
  'bb good morning':       'Barbell Good Morning',
  'face pull':             'Face Pull',
  'lat pulldown':          'Lat Pulldown',
  'seated row':            'Seated Cable Row',
  'cable row':             'Seated Cable Row',
  't-bar row':             'T-Bar Row',
  'pendlay row':           'Pendlay Row',
  'barbell row':           'Barbell Row',
  'bent over row':         'Barbell Row',
  'bent row':              'Barbell Row',
  'leg press':             'Leg Press',
  'lp':                    'Leg Press',
  'leg curl':              'Leg Curl',
  'leg ext':               'Leg Extension',
  'leg extension':         'Leg Extension',
  'ham curl':              'Leg Curl',
  'hamstring curl':        'Leg Curl',
  'calf raise':            'Calf Raise',
  'tricep pushdown':       'Tricep Pushdown',
  'tri pushdown':          'Tricep Pushdown',
  'skull crusher':         'Skull Crusher',
  'skullcrusher':          'Skull Crusher',
  'bb curl':               'Barbell Curl',
  'ez curl':               'EZ Bar Curl',
  'ez bar curl':           'EZ Bar Curl',
  'preacher curl':         'Preacher Curl',
  'hammer curl':           'Hammer Curl',
  'lat raise':             'Lateral Raise',
  'lateral raise':         'Lateral Raise',
  'rear delt fly':         'Rear Delt Fly',
  'rear delt flye':        'Rear Delt Fly',
  'chest fly':             'Chest Fly',
  'pec fly':               'Chest Fly',
  'pec dec':               'Pec Dec',
};

/**
 * Canonicalize an exercise name: strips bracket suffixes, trims,
 * and maps known synonyms to a single canonical display name.
 * Returns the canonical name (proper-cased) or the original name if no synonym found.
 */
function canonicalizeLift(name) {
  const base = _getBaseName(name).trim();
  const key = base.toLowerCase().replace(/[,.:;!?]+$/, '').trim();
  if (SYNONYM_MAP[key]) return SYNONYM_MAP[key];
  return base;
}

// ── EXERCISE SPELL CORRECTION ─────────────────────────────────────────────────
const EXERCISE_DICT = [
  'Incline','Decline','Bench','Press','Squat','Deadlift','Row','Pull','Push',
  'Curl','Extension','Raise','Fly','Flye','Pulldown','Pullover','Pushdown',
  'Romanian','Reverse','Lateral','Front','Overhead','Close','Wide','Grip',
  'Barbell','Dumbbell','Cable','Machine','Smith','Hack','Leg','Hip',
  'Hamstring','Quadricep','Glute','Back','Chest','Shoulder','Tricep','Bicep',
  'Preacher','Spider','Concentration','Hammer','Pronated','Supinated','Neutral',
  'Prone','Supine','Standing','Seated','Lying','Bent','Straight','Single','Arm',
  'Unilateral','Bilateral','Compound','Isolation','Bulgarian','Split','Lunge',
  'Step','Box','Pause','Tempo','Deficit','Rack','Block','Pin','Board','Floor',
  'Sling','Shot','Larson','Larsen','Comp','Competition','Sumo','Conventional',
  'Stiff','Leg','Trap','Hex','Snatch','Clean','Jerk','Thruster','Turkish',
  'Get','Up','Swing','Windmill','Plank','Hold','Crunch','Situp','Sit',
  'Nautilus','Leverage','Preacher','Incline','Delt','Deltoid','Pec','Dec',
  'Crossover','Kickback','Skull','Crusher','Nose','Breaker','Tate','JM',
  'Zottman','Drag','Pinwheel','Reverse','Wrist','Forearm','Calf','Raise',
  'Nordic','Glute','Ham','Hyper','Extension','Good','Morning','Ab','Wheel',
  'Plank','Hollow','Body','Dragon','Flag','Hanging','Knee','Raise','Tuck',
  'Cuban','Arnold','Bradford','Landmine','Meadows','Pendlay','Yates','Kroc',
  'XPLoad','Rope','Band','Chain','Plate','Weight','Weighted','Bodyweight',
  'Dip','Chinup','Pullup','Pushup','Row','Inverted','Ring','TRX','Suspension',
  // Named barbell variations (from exercise report)
  'Spoto','Zercher','Hatfield','Dimel','Guillotine','California',
  'Duffalo','Buffalo','Kadillac','Cambered','Marrs','Transformer',
  'Earthquake','Bamboo','Axle','Tsunami',
  // Specialty bar names
  'SSB','Safety',
  // Effort/modifier words
  'Beltless','Eccentric','Isometric','Slingshot','Accommodating',
  'Dynamic','Banded','Lockout','Walkout','Floating',
  // Abbreviations (>3 chars to qualify for spell check)
  'CGBP','SGDL','SLDL','WGBP','RGBP','AMRAP',
  // Additional proper names
  'Reeves','Jefferson','Gironda','Anderson','Hackenschmidt',
];

function editDistance(a,b){
  const m=a.length,n=b.length;
  const dp=Array.from({length:m+1},(_,i)=>Array.from({length:n+1},(_,j)=>i===0?j:j===0?i:0));
  for(let i=1;i<=m;i++) for(let j=1;j<=n;j++)
    dp[i][j]=a[i-1]===b[j-1]?dp[i-1][j-1]:1+Math.min(dp[i-1][j],dp[i][j-1],dp[i-1][j-1]);
  return dp[m][n];
}

/**
 * Smart exercise name matching with hierarchical fallback.
 * 1. Exact match (zero false positives)
 * 2. Synonym mapping via canonicalizeLift (handles known variations)
 * 3. Levenshtein distance with common-word safety (handles typos)
 */
function _smartExerciseMatch(a, b, opts = {}) {
  if (!a || !b) return false;
  const normA = a.toLowerCase().trim();
  const normB = b.toLowerCase().trim();
  // 1. EXACT
  if (normA === normB) return true;
  // 2. SYNONYM MAP
  const canonA = canonicalizeLift(a);
  const canonB = canonicalizeLift(b);
  if (canonA && canonB && canonA.toLowerCase() === canonB.toLowerCase()) return true;
  // 3. EDIT DISTANCE (typos only — require shared base word for safety)
  const maxDist = opts.maxEditDistance || 2;
  const dist = editDistance(normA, normB);
  if (dist <= maxDist) {
    const wordsA = normA.split(/\s+/);
    const wordsB = normB.split(/\s+/);
    const hasCommon = wordsA.some(w => wordsB.some(v => editDistance(w, v) <= 1));
    if (hasCommon) return true;
  }
  return false;
}

/**
 * Extract the primary lift keyword for a name in a given category.
 * Used for category-aware matching to find core lift.
 */
function _extractPrimaryKeyword(name, group) {
  if (!LIFT_GROUPS[group]) return null;
  const lower = name.toLowerCase();
  const keywords = LIFT_GROUPS[group].slice().sort((a, b) => b.length - a.length);
  for (const keyword of keywords) {
    if (lower.includes(keyword)) return keyword;
  }
  return null;
}

function spellCorrectWord(word){
  if(word.length<=3) return word;
  const lower=word.toLowerCase();
  const match=EXERCISE_DICT.find(d=>d.toLowerCase()===lower);
  if(match) return match;
  let best=null; let bestDist=Infinity;
  for(const d of EXERCISE_DICT){
    const dist=editDistance(lower,d.toLowerCase());
    const threshold=word.length<=5?1:2;
    if(dist<bestDist&&dist<=threshold){ bestDist=dist; best=d; }
  }
  return best||word;
}

function spellCorrectExerciseName(name){
  return name.split(/\s+/).map(spellCorrectWord).join(' ');
}

// ── CONSTANTS ─────────────────────────────────────────────────────────────────
const DAYS = ["monday","tuesday","wednesday","thursday","friday","saturday","sunday"];

// ── UTILITIES ─────────────────────────────────────────────────────────────────

// Shared cell normalization — used across 20+ scoring/parsing functions
function _normalizeCell(val) { return String(val || '').toLowerCase().trim(); }

// Excel date serial threshold — numbers above this in cells are likely date serials, not data
const EXCEL_DATE_SERIAL_MIN = 40000;

// Common regex patterns for week/day detection in scoring functions
const RE_WEEK_NUM   = /^week\s*\d+/i;
const RE_DAY_NUM    = /^day\s*\d/i;
const RE_DAY_NAME   = /^(mon|tue|wed|thu|fri|sat|sun)/i;

// Fix Excel date serialization: when a cell like "6-8" is stored as a date,
// cell.v becomes a serial number (e.g. 45816). Use cell.w (formatted text) instead.
function fixDateSerials(ws, rows) {
  if (!ws || !ws['!ref']) return rows;
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let r = 0; r < rows.length; r++) {
    if (!rows[r]) continue;
    for (let c = 0; c < rows[r].length; c++) {
      const val = rows[r][c];
      // With cellDates:true, date-encoded cells become Date objects in the rows array
      if (val instanceof Date) {
        const cellAddr = XLSX.utils.encode_cell({r: r + range.s.r, c: c + range.s.c});
        const cell = ws[cellAddr];
        // Use cell.w (formatted text like "6-8") if available, otherwise stringify
        rows[r][c] = (cell && cell.w) ? cell.w : String(val);
      }
      // Keep the old number check as a safety net (in case cellDates misses something)
      else if (typeof val === 'number' && val > EXCEL_DATE_SERIAL_MIN) {
        const cellAddr = XLSX.utils.encode_cell({r: r + range.s.r, c: c + range.s.c});
        const cell = ws[cellAddr];
        if (cell && cell.w && /^\d+[-\/]\d+$/.test(cell.w)) {
          rows[r][c] = cell.w;
        }
      }
    }
  }
  return rows;
}

// ── CONFIDENCE SCORE FUNCTIONS ────────────────────────────────────────────────
// Each returns { id, score (0.0–1.0), signals: { matched[], missing[], negative[] } }
// Pure functions — no side effects. Read wb directly.

function _sheetRows(wb, si) {
  const ws = wb.Sheets[wb.SheetNames[si]];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
}

function scoreA(wb) {
  const id = 'A';
  const matched = [], missing = [], negative = [];
  let score = 0;

  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return { id, score: 0, signals: { matched, missing, negative } };
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

  let dayColFound = false;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    const hits = row.filter(c => c && DAYS.some(d => String(c).toLowerCase().trim().startsWith(d)));
    if (hits.length >= 2) { dayColFound = true; score += 0.7; matched.push('2+ day columns in header'); break; }
  }
  if (!dayColFound) missing.push('2+ day columns in first 10 rows');

  // Positive: first sheet only, not multi-sheet program
  if (wb.SheetNames.length >= 2 && wb.SheetNames.length <= 15) { score += 0.1; matched.push('reasonable sheet count'); }

  // Negative: GZCL T1/T2/T3 markers
  for (let si = 0; si < Math.min(3, wb.SheetNames.length); si++) {
    const r = _sheetRows(wb, si);
    for (let i = 0; i < Math.min(20, r.length); i++) {
      if (!r[i]) continue;
      if (r[i].some(c => c != null && /^T[123]$/i.test(String(c).trim()))) {
        score -= 0.3; negative.push('T1/T2/T3 tier markers (→U)'); break;
      }
    }
  }
  // Negative: Hepburn Power Phase / Pump Phase sheet
  const sheetNames = wb.SheetNames.map(s => s.toLowerCase());
  if (sheetNames.some(n => /power\s*phase|pump\s*phase/i.test(n))) {
    score -= 0.3; negative.push('Power/Pump Phase sheets (→J)');
  }

  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreB(wb) {
  // B is the fallback — never wins via scoring; wins only when everything else scores < 0.2
  return { id: 'B', score: 0.0, signals: { matched: ['fallback'], missing: [], negative: [] } };
}

function scoreC(wb) {
  const id = 'C';
  const matched = [], missing = [], negative = [];
  let score = 0;

  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return { id, score: 0, signals: { matched, missing, negative } };
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

  let hasSplitHeaders = false, hasDayLabel = false;
  for (let i = 0; i < Math.min(40, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    const strs = row.map(_normalizeCell);
    if (strs.includes('sets') && strs.includes('reps') && (strs.includes('load') || strs.includes('weight'))) {
      hasSplitHeaders = true;
    }
    if (row[0] && /^day\s*\d/i.test(String(row[0]).trim())) hasDayLabel = true;
  }

  if (hasSplitHeaders) { score += 0.4; matched.push('Sets/Reps/Load column headers'); }
  else missing.push('Sets/Reps/Load headers');
  if (hasDayLabel) { score += 0.4; matched.push('Day N label in col A'); }
  else missing.push('Day N label in col A');

  // Negative: multiple day columns → A
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    const hits = row.filter(c => c && DAYS.some(d => String(c).toLowerCase().trim().startsWith(d)));
    if (hits.length >= 2) { score -= 0.3; negative.push('2+ day columns (→A)'); break; }
  }

  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreD(wb) {
  const id = 'D';
  const matched = [], missing = [], negative = [];
  let score = 0;

  for (let si = 0; si < Math.min(6, wb.SheetNames.length); si++) {
    const rows = _sheetRows(wb, si);
    let dSetsCount = 0, dHasDaySection = false, dSetsHeaderRow = -1;
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      const sc = strs.filter(s => s === 'sets').length;
      if (sc >= 2 && dSetsHeaderRow === -1) {
        const exRepeat = strs.filter(s => s === 'exercise').length;
        if (exRepeat >= 2) { negative.push('Exercise repeating header (→F)'); continue; }
        dSetsCount = sc; dSetsHeaderRow = i;
      }
      if (row.some(c => c != null && /^WEEK\s*\d/i.test(String(c).trim()))) dHasDaySection = true;
    }
    if (dSetsCount >= 2) { score += 0.3; matched.push('2+ Sets columns'); }
    else missing.push('2+ Sets columns');
    if (dHasDaySection) { score += 0.3; matched.push('WEEK N section header'); }
    else missing.push('WEEK N header');

    if (dSetsCount >= 2 && dHasDaySection && dSetsHeaderRow >= 0) {
      let hasTextExercise = false;
      for (let i = dSetsHeaderRow + 1; i < Math.min(dSetsHeaderRow + 6, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        const colA = row[0]; if (colA == null) continue;
        const colAStr = String(colA).trim();
        if (!colAStr || /^\d+\.?\d*$/.test(colAStr) || /^WEEK\s*\d/i.test(colAStr)) continue;
        if (/[a-zA-Z]{2,}/.test(colAStr)) { hasTextExercise = true; break; }
      }
      if (hasTextExercise) { score += 0.3; matched.push('text exercise names in col A'); }
      else { score -= 0.2; negative.push('no text exercise names in col A (→G)'); }
    }
    if (score > 0) break; // found signals in this sheet
  }

  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreE(wb) {
  const id = 'E';
  const matched = [], missing = [], negative = [];
  let score = 0;
  const SKIP_SHEET = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|output)$/i;

  for (let si = 0; si < Math.min(6, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP_SHEET.test(sn.trim())) continue;
    const rows = _sheetRows(wb, si);
    if (!rows || rows.length < 5) continue;

    // E-nSuns: day label + numeric data
    for (let i = 0; i < Math.min(50, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let col = 0; col <= 2; col++) {
        const val = row[col]; if (!val) continue;
        const str = String(val).trim().toLowerCase();
        const isDayName = DAYS.some(d => str === d || str.startsWith(d + ' '));
        const isDayLabel = /^day\s*\d/i.test(String(val).trim());
        if (isDayName || isDayLabel) {
          const otherDays = row.filter((c2, ci) => ci !== col && c2 && DAYS.some(d => String(c2).trim().toLowerCase().startsWith(d)));
          if (otherDays.length > 0) continue; // multi-day header → A
          let hasColHeaders = false;
          for (let j = i; j < Math.min(i + 3, rows.length); j++) {
            const hr = rows[j]; if (!hr) continue;
            const hstrs = hr.map(_normalizeCell);
            if ((hstrs.includes('sets') || hstrs.includes('set')) && (hstrs.includes('reps') || hstrs.includes('rep goal'))) hasColHeaders = true;
            if (hstrs.includes('movement') || hstrs.includes('exercise movement')) hasColHeaders = true;
          }
          if (hasColHeaders) continue;
          for (let j = i + 1; j < Math.min(i + 6, rows.length); j++) {
            const dr = rows[j]; if (!dr) continue;
            let numCount = 0;
            for (let c2 = col + 1; c2 < dr.length; c2++) { if (typeof dr[c2] === 'number') numCount++; }
            if (numCount >= 6) { score += 0.6; matched.push('day label + 6+ numerics below'); break; }
          }
        }
      }
      if (score >= 0.6) break;
    }

    // E-weekText: Week N + SxR without Sets header
    if (score < 0.6) {
      let hasWeekLabel = false, hasSxR = false, hasSetsColHeader = false, sxrCount = 0, hasBFormatMarker = false;
      for (let i = 0; i < Math.min(30, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        for (let c = 0; c < Math.min(row.length, 25); c++) {
          const val = row[c]; if (val == null) continue;
          const str = String(val).trim();
          if (/^Week\s*\d/i.test(str)) hasWeekLabel = true;
          if (/^\d+\s*x\s*\d/i.test(str) || /^\d+x\d/i.test(str)) { hasSxR = true; sxrCount++; }
          if (/^sets$/i.test(str) || /^sets\s*x/i.test(str)) hasSetsColHeader = true;
          const lo = str.toLowerCase();
          if (lo.includes('record') || lo.includes('coach') || (lo.includes('sets') && lo.includes('rpe'))) hasBFormatMarker = true;
        }
      }
      if ((hasWeekLabel && hasSxR && !hasSetsColHeader && !hasBFormatMarker) || (sxrCount >= 3 && !hasSetsColHeader && !hasBFormatMarker)) {
        score += 0.6;
        matched.push(hasWeekLabel ? 'Week N + SxR patterns' : 'many SxR patterns');
      }
    }

    // E-tabular: Week + exercise + Sets/Reps headers
    if (score < 0.6) {
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        const strs = row.map(_normalizeCell);
        if (strs.includes('week') && strs.some(s => s.includes('primary lift') || s === 'exercise') && (strs.includes('sets') || strs.includes('reps'))) {
          score += 0.65; matched.push('Week+Exercise+Sets/Reps tabular headers'); break;
        }
      }
    }

    // E-cycleWeek: Cycle/Phase markers + exercise names
    if (score < 0.6) {
      let hasCycle = false, hasExNames = false, hasWeightRepsGroups = false;
      for (let i = 0; i < Math.min(20, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        for (let c = 0; c < Math.min(row.length, 15); c++) {
          const val = row[c]; if (val == null) continue;
          const str = String(val).trim();
          if (/^Cycle\s*\d/i.test(str) || /^Week\s*\d.*\d+x\d/i.test(str) || /Phase\s*[-–]\s*Week\s*\d/i.test(str) || /^\d+\s*Rep\s*Wave/i.test(str)) hasCycle = true;
          if (/^(Bench|Squat|Deadlift|Press|Military|Overhead|OHP|Core\s*Lift|OH\s*Press)/i.test(str)) hasExNames = true;
        }
        const strs = row.map(_normalizeCell);
        if (strs.filter(s => s === 'weight').length >= 2 && strs.filter(s => s === 'reps').length >= 2) hasWeightRepsGroups = true;
      }
      if ((hasCycle && hasExNames) || (hasWeightRepsGroups && hasExNames)) {
        score += 0.6; matched.push('Cycle/Phase markers + exercise names');
      }
    }

    if (score >= 0.5) break;
  }

  if (score === 0) missing.push('day labels or Week+SxR or tabular Week/Exercise/Sets headers');

  // Negative signals
  const sheetNamesLo = wb.SheetNames.map(s => s.toLowerCase());
  if (sheetNamesLo.some(n => /1rm\s*input/i.test(n))) { score -= 0.3; negative.push('1RM Input sheet (→J)'); }
  if (sheetNamesLo.some(n => n.includes('maxes')) && wb.SheetNames.some(s => /^(bench|squat|dl)\s+\dx/i.test(s))) {
    score -= 0.3; negative.push('Maxes+training sheets (→M)');
  }

  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreF(wb) {
  const id = 'F';
  const matched = [], missing = [], negative = [];
  let score = 0;
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|output|pr\s*sheet|notes?|start\s*here|1\.\s*start|predicted)$/i;

  for (let si = 0; si < Math.min(8, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP.test(sn.trim())) continue;
    const rows = _sheetRows(wb, si);
    if (!rows || rows.length < 5) continue;

    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const weekCols = [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        if (/^Week\s*\d/i.test(String(val).trim())) weekCols.push(c);
      }
      if (weekCols.length >= 2) {
        score += 0.4; matched.push('2+ Week N columns horizontal');
        for (let j = i + 1; j < Math.min(i + 3, rows.length); j++) {
          const sr = rows[j]; if (!sr) continue;
          const strs = sr.map(_normalizeCell);
          const setsRepeat = strs.filter(s => s === 'sets').length;
          const hasIntensity = strs.some(s => s === 'intensity');
          if (setsRepeat >= 2 && hasIntensity) { score -= 0.2; negative.push('Sets+Intensity repeating (→D)'); continue; }
          const subMatch = (strs.filter(s => s === 'weight').length >= 2 && strs.filter(s => s === 'reps').length >= 2)
            || strs.filter(s => s === '%' || s === 'percentage').length >= 2
            || strs.filter(s => s === 'exercise').length >= 2;
          if (subMatch) { score += 0.3; matched.push('F sub-headers (Weight/Reps or % or Exercise)'); }
          // Sheiko exclusion
          let sheikoPat = false;
          for (let k = j + 1; k < Math.min(j + 5, rows.length); k++) {
            const dr = rows[k]; if (!dr) continue;
            const a = dr[0], b = dr[1], cc = dr[2];
            if (typeof a === 'number' && a >= 1 && a <= 20 && b != null && /[a-zA-Z]{2,}/.test(String(b)) && typeof cc === 'number' && cc > 0 && cc < 1) {
              sheikoPat = true; break;
            }
          }
          if (sheikoPat) { score -= 0.3; negative.push('Sheiko numbered exercise pattern (→G)'); }
        }
        break;
      }
      // Madcow: Day + Exercise + W1/W2 columns
      const strs = row.map(_normalizeCell);
      if (strs.some(s => s === 'day') && strs.some(s => s === 'exercise')) {
        let wLabels = strs.filter(s => /^w\d+/i.test(s)).length;
        let seqCount = row.filter(v => typeof v === 'number' && v === Math.floor(v) && v >= 1 && v <= 52).length;
        if (wLabels >= 3 || seqCount >= 4) { score += 0.65; matched.push('Madcow Day+Exercise+week cols'); break; }
      }
    }
    if (score > 0.5) break;
  }

  if (score === 0) missing.push('2+ horizontal Week N columns with F sub-headers');
  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreG(wb) {
  const id = 'G';
  const matched = [], missing = [], negative = [];
  let score = 0;
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|max|output|pr\s*sheet|notes?|start\s*here|1\.\s*start|predicted|volume)/i;

  for (let si = 0; si < Math.min(8, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP.test(sn.trim())) continue;
    const rows = _sheetRows(wb, si);
    if (!rows || rows.length < 10) continue;

    let hasWeekLabel = false, hasNumberedExercise = false;
    for (let i = 0; i < Math.min(40, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const colA = row[0], colB_val = row[1];
      if (colA != null && /^week\s*\d/i.test(String(colA).trim())) hasWeekLabel = true;
      if (colB_val != null && /^week\s*\d/i.test(String(colB_val).trim()) && colA == null) hasWeekLabel = true;
      if (typeof colA === 'number' && colA >= 1 && colA <= 20 && colA === Math.floor(colA)) {
        const colB = row[1], colC = row[2];
        if (colB != null && /[a-zA-Z]{2,}/.test(String(colB)) && typeof colC === 'number' && colC > 0 && colC < 1) {
          hasNumberedExercise = true;
        }
      }
      if (hasWeekLabel && hasNumberedExercise) break;
    }
    if (hasWeekLabel) { score += 0.35; matched.push('Week N label'); }
    else missing.push('Week N label');
    if (hasNumberedExercise) { score += 0.55; matched.push('numbered exercise + percentage (col A int, col C decimal)'); }
    else missing.push('numbered exercise (col A int 1-20, col B text, col C decimal 0-1)');
    break;
  }

  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreH(wb) {
  const id = 'H';
  const matched = [], missing = [], negative = [];
  let score = 0;

  for (const sn of wb.SheetNames) {
    if (/disbrow/i.test(sn) || /\b10x3\b/i.test(sn)) {
      score = 0.9; matched.push('DeathBench sheet name (Disbrows/10x3)'); break;
    }
  }
  if (score === 0) {
    for (const sn of wb.SheetNames) {
      if (/hatfield/i.test(sn)) {
        const ws = wb.Sheets[sn];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
        for (let i = 0; i < Math.min(6, rows.length); i++) {
          if ((rows[i]||[]).some(c => c && /one\s*rep\s*max/i.test(String(c)))) {
            score = 0.9; matched.push('Hatfield sheet + One Rep Max'); break;
          }
        }
        if (score > 0) break;
      }
    }
  }
  if (score === 0) {
    for (const sn of wb.SheetNames) {
      const ws = wb.Sheets[sn];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
      for (let i = 0; i < Math.min(15, rows.length); i++) {
        const joined = (rows[i]||[]).map(c => String(c||'')).join(' ').toLowerCase();
        if (joined.includes('ed coan') && joined.includes('peaking')) { score = 0.9; matched.push('Ed Coan Peaking program'); break; }
        if (joined.includes('projected new 1rm')) { score = 0.9; matched.push('Projected new 1RM label'); break; }
      }
      if (score > 0) break;
    }
  }
  if (score === 0) {
    for (const sn of wb.SheetNames) {
      const ws = wb.Sheets[sn];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const vals = (rows[i]||[]).map(c => String(c||'').toLowerCase());
        if (vals.some(v => v.includes('back') && v.includes('sqt')) && vals.some(v => v.includes('front') && v.includes('sqt'))) {
          score = 0.9; matched.push('back sqt + front sqt in same header row (Hatch)'); break;
        }
      }
      if (score > 0) break;
    }
  }
  if (score === 0) {
    for (const sn of wb.SheetNames) {
      const ws = wb.Sheets[sn];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
      let hasDesired = false, hasCurrent = false;
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const joined = (rows[i]||[]).map(c => String(c||'')).join(' ').toLowerCase();
        if (joined.includes('desired max')) hasDesired = true;
        if (joined.includes('current max')) hasCurrent = true;
        if (hasDesired && hasCurrent) { score = 0.9; matched.push('Desired max + Current Max (Coan-Phillipi)'); break; }
      }
      if (score > 0) break;
    }
  }

  if (score === 0) missing.push('none of H sub-format patterns (DeathBench/Hatfield/EdCoan/Hatch/Coan-Phillipi)');
  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreI(wb) {
  const id = 'I';
  const matched = [], missing = [], negative = [];
  let score = 0;

  if (wb.SheetNames.length > 2) { missing.push('≤2 sheets required'); return { id, score: 0, signals: { matched, missing, negative } }; }
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return { id, score: 0, signals: { matched, missing, negative } };
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    const hasTM = row.some(c => c && /^texas\s*method$/i.test(String(c).trim()));
    if (!hasTM) continue;
    matched.push('Texas Method standalone cell');
    const dayHits = row.filter(c => c && /^(mon|wed|fri)/i.test(String(c).trim()));
    if (dayHits.length >= 3) { score = 0.95; matched.push('Mon/Wed/Fri day columns'); break; }
    else { score = 0.5; missing.push('Mon/Wed/Fri columns in same row'); }
  }
  if (score === 0) missing.push('Texas Method standalone cell');
  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

function scoreJ(wb) {
  const id = 'J';
  const matched = [], missing = [], negative = [];
  const names = wb.SheetNames.map(s => s.toLowerCase());

  const has1RM = names.some(n => /1rm\s*input/i.test(n));
  const hasPhase = names.some(n => /power\s*phase/i.test(n) || /pump\s*phase/i.test(n));

  if (has1RM) matched.push('1RM Input sheet'); else missing.push('1RM Input sheet');
  if (hasPhase) matched.push('Power Phase / Pump Phase sheet'); else missing.push('Power/Pump Phase sheet');

  const score = (has1RM && hasPhase) ? 0.95 : (has1RM || hasPhase) ? 0.35 : 0;
  return { id, score, signals: { matched, missing, negative } };
}

function scoreK(wb) {
  const id = 'K';
  const matched = [], missing = [], negative = [];
  const names = wb.SheetNames.map(s => s.toLowerCase());

  if (!names.some(n => n.includes('start here'))) {
    missing.push('Start Here sheet');
    return { id, score: 0, signals: { matched, missing, negative } };
  }
  matched.push('Start Here sheet');

  let hasGroupMarker = false;
  for (const sn of wb.SheetNames) {
    if (/start\s*here|bare\s*minimum/i.test(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      if (row[0] && /^\(\s*[A-E]\s*\)$/i.test(String(row[0]).trim())) { hasGroupMarker = true; break; }
    }
    if (hasGroupMarker) break;
  }

  if (hasGroupMarker) { matched.push('(A)/(B)/(C) group markers in col A'); }
  else missing.push('(A)-(E) group markers in col A of training sheets');

  const score = (hasGroupMarker) ? 0.9 : 0.3;
  return { id, score, signals: { matched, missing, negative } };
}

function scoreL(wb) {
  const id = 'L';
  const matched = [], missing = [], negative = [];
  let score = 0;

  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const joined = (rows[i]||[]).map(c => String(c||'')).join(' ');
      if (/candito/i.test(joined)) { score = 0.85; matched.push('Candito name in first 5 rows'); break; }
    }
    if (score > 0) break;
  }

  if (score === 0) {
    const snLower = wb.SheetNames.map(s => s.toLowerCase());
    if (snLower.some(s => s.includes('strength hypertrophy') || s.includes('strength control') || s.includes('strength power'))) {
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const a = String((rows[i]||[])[0]||'').toLowerCase();
        if (/^(monday|tuesday|thursday|friday)$/.test(a)) { score = 0.8; matched.push('Strength Hypertrophy/Control/Power sheets + day names'); break; }
      }
    }
  }

  if (score === 0) missing.push('Candito name or Strength Hypertrophy/Control/Power sheet names');
  return { id, score, signals: { matched, missing, negative } };
}

function scoreM(wb) {
  const id = 'M';
  const matched = [], missing = [], negative = [];

  const hasMaxes = wb.SheetNames.some(s => /^maxes$/i.test(s.trim()));
  const hasTraining = wb.SheetNames.some(s => /^(bench|squat|dl)\s+\dx/i.test(s.trim()));

  if (hasMaxes) matched.push('Maxes sheet'); else missing.push('Maxes sheet');
  if (hasTraining) matched.push('bench/squat/dl training sheet (NxX)'); else missing.push('training sheet (bench/squat/dl NxX)');

  if (!hasMaxes || !hasTraining) return { id, score: 0, signals: { matched, missing, negative } };

  let hasWeek1 = false;
  for (const sn of wb.SheetNames) {
    if (/^maxes$/i.test(sn.trim())) continue;
    if (!/^(bench|squat|dl)\s+\dx/i.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      if (/^week\s*1$/i.test(String((rows[i]||[])[0]||'').trim())) { hasWeek1 = true; break; }
    }
    break;
  }

  if (hasWeek1) matched.push('Week 1 label in training sheet col A');
  else missing.push('Week 1 in training sheet col A');

  const score = (hasMaxes && hasTraining && hasWeek1) ? 0.95 : 0.4;
  return { id, score, signals: { matched, missing, negative } };
}

function scoreN(wb) {
  const id = 'N';
  const matched = [], missing = [], negative = [];
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());

  // Greyskull LP
  if (names.some(s => s === 'lift') && names.some(s => s === 'setup')) {
    const ws = wb.Sheets[wb.SheetNames[names.indexOf('lift')]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    if (rows[0] && String(rows[0][4]||'').match(/week\s*1/i)) {
      return { id, score: 0.95, signals: { matched: ['Greyskull: Lift+Setup sheets + Week 1 in col E'], missing, negative } };
    }
    return { id, score: 0.5, signals: { matched: ['Lift+Setup sheets'], missing: ['Week 1 in col E'], negative } };
  }

  // Ivysaur
  if (names.some(s => s === 'template')) {
    const ws = wb.Sheets[wb.SheetNames[names.indexOf('template')]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const txt = (rows[i]||[]).map(c => String(c||'')).join(' ').toLowerCase();
      if (txt.includes('4-4-8') || txt.includes('ivysaur') || txt.includes('4.4.8')) {
        return { id, score: 0.95, signals: { matched: ['Ivysaur: Template sheet + 4-4-8/ivysaur text'], missing, negative } };
      }
    }
  }

  // Starting Strength
  if (names.some(s => s.includes('novice program') || s.includes('onus wunsler') || s.includes('wichita falls'))) {
    return { id, score: 0.95, signals: { matched: ['Starting Strength: novice program/onus wunsler/wichita falls sheet'], missing, negative } };
  }

  // StrongLifts
  if (names.some(s => s.includes('stronglifts')) && names.some(s => s.includes('beginner') || s.includes('experienced'))) {
    return { id, score: 0.95, signals: { matched: ['StrongLifts: stronglifts + beginner/experienced sheets'], missing, negative } };
  }

  missing.push('none of N sub-format patterns (Greyskull/Ivysaur/SS/StrongLifts)');
  return { id, score: 0, signals: { matched, missing, negative } };
}

function scoreO(wb) {
  const id = 'O';
  const matched = [], missing = [], negative = [];
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());

  if (!names.includes('training')) {
    missing.push('Training sheet');
    return { id, score: 0, signals: { matched, missing, negative } };
  }
  matched.push('Training sheet');

  const ws = wb.Sheets[wb.SheetNames[names.indexOf('training')]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
  if (!rows || rows.length < 20) {
    missing.push('Training sheet has < 20 rows');
    return { id, score: 0.2, signals: { matched, missing, negative } };
  }

  const r8 = rows[8] || [];
  const h4 = String(r8[4]||'').trim().toLowerCase();
  const h5 = String(r8[5]||'').trim().toLowerCase();
  const h6 = String(r8[6]||'').trim().toLowerCase();
  const h7 = String(r8[7]||'').trim().toLowerCase();
  const d1 = String(r8[3]||'').trim().toLowerCase();

  const hasExactHeaders = (h4 === 'sets' && h5 === 'reps' && h6 === 'intensity' && h7 === 'load');
  const hasDayCol = d1.startsWith('day');

  if (hasExactHeaders) matched.push('row 8 cols 4-7: sets/reps/intensity/load'); else missing.push('row 8 exact headers');
  if (hasDayCol) matched.push('row 8 col 3: Day label'); else missing.push('row 8 col 3 Day label');

  let hasSQ = false, hasBN = false;
  for (let r = 9; r < Math.min(40, rows.length); r++) {
    const c0 = String((rows[r]||[])[0]||'').trim().toLowerCase();
    if (c0.startsWith('sq ')) hasSQ = true;
    if (c0.startsWith('bn ')) hasBN = true;
  }
  if (hasSQ && hasBN) matched.push('SQ/BN exercise prefix labels'); else missing.push('SQ/BN prefix labels');

  const score = (hasExactHeaders && hasDayCol && hasSQ && hasBN) ? 0.95 : (hasExactHeaders || hasDayCol) ? 0.35 : 0.1;
  return { id, score, signals: { matched, missing, negative } };
}

function scoreP(wb) {
  const id = 'P';
  const matched = [], missing = [], negative = [];

  const mainSheet = _pGetSheet(wb, 'Main');
  const inputsSheet = _pGetSheet(wb, 'Inputs');

  if (!mainSheet) { missing.push('Main sheet'); return { id, score: 0, signals: { matched, missing, negative } }; }
  if (!inputsSheet) { missing.push('Inputs sheet'); return { id, score: 0.1, signals: { matched, missing, negative } }; }
  matched.push('Main + Inputs sheets');

  const mainData = _pSheetToArray(mainSheet, 20);
  let headerRow = -1;
  for (let r = 0; r < Math.min(5, mainData.length); r++) {
    const row = mainData[r]; if (!row) continue;
    if (row[0] === 'Date' && row[1] === 'Exercise') { headerRow = r; break; }
  }

  if (headerRow === -1) { missing.push('Date+Exercise headers in col 0-1'); return { id, score: 0.2, signals: { matched, missing, negative } }; }
  matched.push('Date + Exercise headers');

  const hRow = mainData[headerRow];
  const h5 = hRow[5] ? hRow[5].toString() : '';
  const h6 = hRow[6] ? hRow[6].toString() : '';
  const hasFatigue = h5.includes('Fatigue') || h6.includes('Fatigue');
  if (hasFatigue) matched.push('Fatigue column'); else missing.push('Fatigue column in col 5-6');

  let foundWeek1 = false;
  for (let r = 0; r < Math.min(15, mainData.length); r++) {
    if (mainData[r] && mainData[r][1] && /^Week\s+1$/i.test(mainData[r][1].toString())) { foundWeek1 = true; break; }
  }
  if (foundWeek1) matched.push('Week 1 marker'); else missing.push('Week 1 in col 1');

  const score = (hasFatigue && foundWeek1) ? 0.9 : (hasFatigue || foundWeek1) ? 0.4 : 0.2;
  return { id, score, signals: { matched, missing, negative } };
}

function scoreQ(wb) {
  const id = 'Q';
  const matched = [], missing = [], negative = [];

  // V1
  for (const sheetName of wb.SheetNames) {
    if (!sheetName.toLowerCase().includes('bridge')) continue;
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const cell1 = _qGetCell(ws, 1, 1), cell2 = _qGetCell(ws, 1, 2);
    if (cell1 === 'Exercise' && cell2 === 'Prescription') {
      return { id, score: 0.95, signals: { matched: ['Q-v1: Bridge sheet + Exercise/Prescription headers'], missing, negative } };
    }
  }

  // V2
  if (wb.Sheets['The Bridge 1.0']) {
    const ws = wb.Sheets['The Bridge 1.0'];
    let foundWeek = false, foundPattern = false;
    for (let r = 5; r <= 10; r++) { const cell = _qGetCell(ws, r, 0); if (cell && cell.toString().includes('Week')) { foundWeek = true; break; } }
    for (let r = 7; r <= 20; r++) { const cell = _qGetCell(ws, r, 2); if (cell && (cell.toString().includes('Weight') || cell.toString().includes('Reps'))) { foundPattern = true; break; } }
    if (foundWeek && foundPattern) return { id, score: 0.95, signals: { matched: ['Q-v2: The Bridge 1.0 sheet + Week marker + Weight/Reps pattern'], missing, negative } };
    if (foundWeek || foundPattern) return { id, score: 0.5, signals: { matched: ['Q-v2: The Bridge 1.0 partial'], missing, negative } };
  }

  // V3
  const weekPattern = /^Week\s+\d+\s*-\s*\w+\s*Stress/i;
  const weekCount = wb.SheetNames.filter(sn => weekPattern.test(sn)).length;
  if (weekCount >= 3) return { id, score: 0.95, signals: { matched: ['Q-v3: 3+ Week N - *Stress sheets'], missing, negative } };
  if (weekCount > 0) return { id, score: 0.3, signals: { matched: [`Q-v3: ${weekCount} stress sheet(s)`], missing: ['need 3+'], negative } };

  missing.push('none of Q sub-format patterns (v1 Bridge/v2 Bridge 1.0/v3 Stress sheets)');
  return { id, score: 0, signals: { matched, missing, negative } };
}

function scoreR(wb) {
  const id = 'R';
  const matched = [], missing = [], negative = [];
  const dayNames = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  const daySheets = wb.SheetNames.filter(name => dayNames.includes(name));

  if (daySheets.length < 3) { missing.push('3+ day-of-week sheet names'); return { id, score: 0, signals: { matched, missing, negative } }; }
  matched.push(`${daySheets.length} day-of-week sheets`);

  const sheet = wb.Sheets[daySheets[0]];
  if (!sheet) return { id, score: 0.2, signals: { matched, missing, negative } };

  const A0 = _rGetCell(sheet, 0, 0);
  if (A0 !== 'Week') { missing.push('col A row 0 = "Week"'); return { id, score: 0.3, signals: { matched, missing, negative } }; }
  matched.push('col A row 0 = Week');

  let hasTarget = false, hasSetLabel = false, hasRepLabel = false;
  for (let r = 0; r < 10; r++) {
    const cellA = _rGetCell(sheet, r, 0);
    if (cellA === 'Target') hasTarget = true;
    if (cellA === 'Set(s)') hasSetLabel = true;
    if (cellA === 'Rep(s)') hasRepLabel = true;
  }
  if (hasTarget) matched.push('Target label'); else missing.push('Target label');
  if (hasSetLabel) matched.push('Set(s) label'); else missing.push('Set(s) label');
  if (hasRepLabel) matched.push('Rep(s) label'); else missing.push('Rep(s) label');

  let hasExerciseInTop4 = false;
  for (let r = 0; r < 4; r++) { if (_rGetCell(sheet, r, 0) === 'Exercise') { hasExerciseInTop4 = true; break; } }
  if (hasExerciseInTop4) { score -= 0.2; negative.push('Exercise in rows 0-3 col A (not PHUL)'); }

  const score = (hasTarget && hasSetLabel && hasRepLabel && !hasExerciseInTop4) ? 0.9 : 0.3;
  return { id, score: Math.max(0, score), signals: { matched, missing, negative } };
}

function scoreS(wb) {
  const id = 'S';
  const matched = [], missing = [], negative = [];

  if (!wb.SheetNames.includes('Inputs')) { missing.push('Inputs sheet'); return { id, score: 0, signals: { matched, missing, negative } }; }
  matched.push('Inputs sheet');

  const weekSheets = wb.SheetNames.filter(name => /^Week\s+\d+/.test(name));
  if (weekSheets.length === 0) { missing.push('Week N sheets'); return { id, score: 0.1, signals: { matched, missing, negative } }; }
  matched.push(`${weekSheets.length} Week N sheet(s)`);

  const ws = wb.Sheets[weekSheets[0]];
  if (!ws) return { id, score: 0.2, signals: { matched, missing, negative } };

  let hasSetScheme = false;
  for (let row = 1; row <= 3; row++) {
    const cell = ws[`C${row}`];
    if (cell && cell.v && typeof cell.v === 'string' && cell.v.includes('Set Scheme')) { hasSetScheme = true; break; }
  }
  if (hasSetScheme) matched.push('Set Scheme in col C row 1-3'); else missing.push('Set Scheme in col C');

  let hasSectionHeader = false;
  const sectionPattern = /^\d+:\s+.*(Power|HT|Hypertrophy)/;
  for (const cellRef in ws) {
    if (cellRef.startsWith('A') && ws[cellRef] && ws[cellRef].v && sectionPattern.test(ws[cellRef].v)) { hasSectionHeader = true; break; }
  }
  if (hasSectionHeader) matched.push('section headers (digit: ...Power/HT)'); else missing.push('section headers in col A');

  const score = (hasSetScheme && hasSectionHeader) ? 0.9 : (hasSetScheme || hasSectionHeader) ? 0.4 : 0.15;
  return { id, score, signals: { matched, missing, negative } };
}

function scoreT(wb) {
  const id = 'T';
  const matched = [], missing = [], negative = [];
  const sheets = wb.SheetNames;

  const weekSheetCount = sheets.filter(s => /^Week \d+$/.test(s)).length;
  const hasMetaSheet = sheets.includes('General Overview') || sheets.includes('Exercise Table');

  if (weekSheetCount < 3) { missing.push('3+ Week N sheets (exact "Week N" format)'); return { id, score: 0, signals: { matched, missing, negative } }; }
  matched.push(`${weekSheetCount} Week N sheets`);
  if (!hasMetaSheet) { missing.push('General Overview or Exercise Table sheet'); return { id, score: 0.2, signals: { matched, missing, negative } }; }
  matched.push('General Overview / Exercise Table sheet');

  const firstWeekSheet = sheets.find(s => /^Week \d+$/.test(s));
  const ws = wb.Sheets[firstWeekSheet];
  if (!ws) return { id, score: 0.3, signals: { matched, missing, negative } };

  let headerRow = null;
  const r4A = ws['A4'], r1A = ws['A1'];
  if (r4A && r4A.v === 'Day') headerRow = 4;
  if (r1A && r1A.v === 'Day') headerRow = 1;

  if (headerRow === null) { missing.push('Day header at row 4 or row 1'); return { id, score: 0.3, signals: { matched, missing, negative } }; }

  const cols = {};
  for (const col of ['A','B','C','D','E']) {
    const cell = ws[`${col}${headerRow}`];
    if (cell) cols[col] = (cell.v instanceof Date) ? (cell.w || String(cell.v)) : cell.v;
  }
  const headersMatch = cols['A'] === 'Day' && cols['B'] === 'Muscle Group' && cols['C'] === 'Exercise' && cols['D'] === 'Sets' && cols['E'] === 'Reps';
  if (headersMatch) { matched.push('Day/Muscle Group/Exercise/Sets/Reps headers'); }
  else { missing.push('exact headers: Day, Muscle Group, Exercise, Sets, Reps'); }

  const score = headersMatch ? 0.95 : 0.3;
  return { id, score, signals: { matched, missing, negative } };
}

function scoreU(wb) {
  const id = 'U';
  const matched = [], missing = [], negative = [];
  let score = 0;
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|maxes?|inputs?|output|calculator|1rm)$/i;

  for (let si = 0; si < Math.min(8, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP.test(sn.trim())) continue;
    const rows = _sheetRows(wb, si);
    if (!rows || rows.length < 3) continue;

    let hasTier = false, hasSetsRepsWeight = false, hasTierHeader = false;
    for (let i = 0; i < Math.min(25, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c <= Math.min(1, row.length - 1); c++) {
        const val = row[c]; if (val == null) continue;
        if (/^T[123]$/i.test(String(val).trim())) hasTier = true;
      }
      for (let c = 0; c < Math.min(6, row.length); c++) {
        if (String(row[c]||'').trim().toLowerCase() === 'tier') hasTierHeader = true;
      }
      const strs = row.map(_normalizeCell);
      const hasSets = strs.some(s => s === 'sets' || s === 'set');
      const hasReps = strs.some(s => s === 'reps' || s === 'rep');
      const hasWeight = strs.some(s => s === 'weight' || s === 'load' || s === 'lbs' || s === 'kg');
      if (hasSets && hasReps && hasWeight) hasSetsRepsWeight = true;
    }
    if ((hasTier || hasTierHeader) && hasSetsRepsWeight) {
      score = 0.85;
      matched.push((hasTier ? 'T1/T2/T3 tier markers' : 'Tier column header') + ' + Sets/Reps/Weight columns');
      break;
    }
  }

  // Variant B: T1:/T2:/T3: prefixes + Week N columns
  if (score === 0) {
    for (let si = 0; si < Math.min(4, wb.SheetNames.length); si++) {
      const sn = wb.SheetNames[si];
      if (SKIP.test(sn.trim())) continue;
      const rows = _sheetRows(wb, si);
      if (!rows || rows.length < 5) continue;
      let tierPrefixCount = 0, hasWeekCols = false;
      for (let i = 0; i < Math.min(40, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        if (row[0] != null && /^T[123]\s*[:\-]/i.test(String(row[0]).trim())) tierPrefixCount++;
        for (let c = 1; c < Math.min(row.length, 30); c++) {
          if (row[c] != null && /^Week\s*\d/i.test(String(row[c]).trim())) hasWeekCols = true;
        }
      }
      if (tierPrefixCount >= 3 && hasWeekCols) {
        score = 0.85; matched.push('T1:/T2:/T3: prefixes + Week N column headers'); break;
      }
    }
  }

  if (score === 0) missing.push('T1/T2/T3 tier markers + Sets/Reps/Weight columns (or T1:/T2:/T3: prefixes + Week cols)');
  return { id, score, signals: { matched, missing, negative } };
}

function scoreV(wb) {
  const id = 'V';
  const matched = [], missing = [], negative = [];
  let score = 0;

  let hasCalcSheet = false;
  for (const sn of wb.SheetNames) {
    const lo = sn.toLowerCase().trim();
    if (lo.includes('calculator') || lo.includes('1rm') || lo === 'calc' || lo.includes('max weight') || lo.includes('maxes')) {
      const ws = wb.Sheets[sn];
      if (!ws) continue;
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
      for (let i = 0; i < Math.min(15, rows.length); i++) {
        const joined = (rows[i]||[]).map(c => String(c||'')).join(' ').toLowerCase();
        if (joined.includes('1rm') || joined.includes('max') || joined.includes('squat') || joined.includes('bench') || joined.includes('deadlift')) {
          hasCalcSheet = true; break;
        }
      }
      if (hasCalcSheet) break;
    }
  }
  if (hasCalcSheet) { matched.push('calculator/1RM sheet with relevant content'); score += 0.25; }
  else { missing.push('calculator or 1RM sheet'); return { id, score: 0, signals: { matched, missing, negative } }; }

  const weekCount = wb.SheetNames.filter(sn => /^Week\s*\d/i.test(sn.trim())).length;
  if (weekCount >= 3) { matched.push(`${weekCount} Week N sheets`); score += 0.3; }
  else { missing.push('3+ Week N sheets'); return { id, score: score * 0.3, signals: { matched, missing, negative } }; }

  for (const sn of wb.SheetNames) {
    if (!/^Week\s*\d/i.test(sn.trim())) continue;
    const rows = _sheetRows(wb, wb.SheetNames.indexOf(sn));
    let hasRIR = false, hasPercentLoad = false, hasDayLabel = false, hasExerciseHeader = false;
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      if (strs.some(s => s === 'rir' || s === 'rpe' || s === 'rir/rpe')) hasRIR = true;
      if (strs.some(s => s.includes('%') && /\d/.test(s))) hasPercentLoad = true;
      if (row[0] && /^Day\s*\d/i.test(String(row[0]).trim())) hasDayLabel = true;
      if (strs.some(s => s === 'exercise' || s === 'movement' || s === 'lift')) hasExerciseHeader = true;
    }
    if ((hasRIR || hasPercentLoad) && hasDayLabel && hasExerciseHeader) {
      score += 0.4; matched.push('RIR/RPE or % load + Day N + Exercise header in week sheet'); break;
    }
  }

  if (score < 0.8) missing.push('RIR/RPE or % load + Day N + Exercise header in week sheets');
  return { id, score: Math.min(1.0, score), signals: { matched, missing, negative } };
}

// ── FORMAT W SCORE (Simple Specific Scientific coaching template) ─────────────
function scoreW(wb) {
  const id = 'W';
  const matched = [], missing = [], negative = [];
  let score = 0;

  // Signal 1: "SIMPLE SPECIFIC SCIENTIFIC" title in any sheet's first 5 rows
  let foundTitle = false;
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (const cell of row) {
        if (cell && /simple\s+specific\s+scientific/i.test(String(cell).trim())) {
          score += 0.5; matched.push('"SIMPLE SPECIFIC SCIENTIFIC" title found');
          foundTitle = true; break;
        }
      }
      if (foundTitle) break;
    }
    if (foundTitle) break;
  }
  if (!foundTitle) missing.push('"SIMPLE SPECIFIC SCIENTIFIC" title');

  // Signals 2 & 3: PROGRAM sheet with horizontal week groups + repeating EXERCISE sub-headers
  let foundProgram = false;
  for (const sn of wb.SheetNames) {
    if (!/^program$/i.test(sn.trim())) continue;
    foundProgram = true;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

    let weekHeaderRow = -1;
    const weekPositions = [];
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const wpos = [];
      for (let c = 0; c < row.length; c++) {
        if (row[c] && /^WEEK\s*\d/i.test(String(row[c]).trim())) wpos.push(c);
      }
      if (wpos.length >= 2) { weekHeaderRow = i; weekPositions.push(...wpos); break; }
    }

    if (weekPositions.length >= 4) {
      score += 0.2; matched.push('4 horizontal WEEK N headers in PROGRAM sheet');
    } else if (weekPositions.length >= 2) {
      score += 0.1; matched.push('2+ horizontal WEEK N headers in PROGRAM sheet');
    } else {
      missing.push('WEEK N horizontal headers in PROGRAM sheet');
    }

    if (weekHeaderRow >= 0 && weekPositions.length >= 2) {
      let foundExHeaders = false;
      for (let i = weekHeaderRow + 1; i < Math.min(weekHeaderRow + 5, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        let exCount = 0;
        for (const wc of weekPositions) {
          if (row[wc] && /^exercise$/i.test(String(row[wc]).trim())) exCount++;
        }
        if (exCount >= 2) {
          score += 0.2; matched.push('EXERCISE sub-headers repeating at week positions');
          foundExHeaders = true; break;
        }
      }
      if (!foundExHeaders) missing.push('EXERCISE sub-headers at week column positions');
    }
    break;
  }
  if (!foundProgram) missing.push('PROGRAM sheet');

  // Signal 4: HOMEPAGE_FAQ sheet with ESTIMATED 1 REP MAX
  for (const sn of wb.SheetNames) {
    if (!/homepage|faq/i.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    let found1RM = false;
    for (let i = 0; i < Math.min(12, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      if (row.some(c => c && /estimated\s*1\s*rep\s*max/i.test(String(c)))) {
        score += 0.1; matched.push('ESTIMATED 1 REP MAX in HOMEPAGE_FAQ');
        found1RM = true; break;
      }
    }
    if (!found1RM) missing.push('ESTIMATED 1 REP MAX in HOMEPAGE_FAQ');
    break;
  }

  if (score === 0) missing.push('"SIMPLE SPECIFIC SCIENTIFIC" title, WEEK N horizontal headers');
  return { id, score: Math.min(1.0, Math.max(0.0, score)), signals: { matched, missing, negative } };
}

// ── ADAPTIVE PARSER SCORE ─────────────────────────────────────────────────────
// Returns max 0.3 so dedicated parsers always win when they score > 0.3.
// Fires as fallback only when no dedicated parser recognises the file.
function scoreAdaptive(wb) {
  const id = 'ADAPTIVE';
  try {
    if (!wb || !wb.SheetNames || !wb.SheetNames.length) {
      return { id, score: 0.0, signals: { matched: [], missing: ['no sheets'], negative: [] } };
    }

    // Quick workout-region scan across all non-empty sheets.
    // Threshold ≥ 0.5 means the region has at least an exercise-name column (+0.50),
    // which filters out pure-text FAQ/instruction blocks.
    for (const sn of wb.SheetNames) {
      const ws = wb.Sheets[sn];
      if (!ws || !ws['!ref']) continue;
      const matrix = _buildOccupancyMatrix(ws);
      const regions = _detectRegions(matrix, ws);
      const workoutRegions = regions.filter(r => r.workoutScore >= 0.5);
      if (workoutRegions.length > 0) {
        return {
          id,
          score: 0.3,
          signals: {
            matched: [workoutRegions.length + ' workout-like region(s) in sheet "' + sn + '"'],
            missing: [],
            negative: []
          }
        };
      }
    }

    return { id, score: 0.0, signals: { matched: [], missing: ['no workout-like regions found'], negative: [] } };
  } catch (e) {
    console.error('[liftlog] scoreAdaptive error:', e);
    return { id, score: 0.0, signals: { matched: [], missing: [], negative: ['error: ' + e.message] } };
  }
}

// ── TIE-BREAKING HELPERS ──────────────────────────────────────────────────────

function _tryParse(wb, formatId) {
  try {
    switch(formatId) {
      case 'A': return parseA(wb);
      case 'B': return parseB(wb);
      case 'C': return parseCAutoFormat(wb);
      case 'D': return parseD(wb);
      case 'E': return parseE(wb);
      case 'F': return parseF(wb);
      case 'G': return parseG(wb);
      case 'H': return parseH(wb);
      case 'I': return parseTexasMethod(wb);
      case 'J': return parseHepburn(wb);
      case 'K': return parseBulgarianMethod(wb);
      case 'L': return parseL(wb);
      case 'M': return parseM(wb);
      case 'N': return parseN(wb);
      case 'O': return parseO(wb);
      case 'P': return parseP(wb);
      case 'Q': return parseQ(wb);
      case 'R': return parseR(wb);
      case 'S': return parseS(wb);
      case 'T': return parseT(wb);
      case 'U': return parseU(wb);
      case 'V': return parseV(wb);
      case 'W': return parseW(wb);
      case 'ADAPTIVE': return parseAdaptive(wb);
      default: return null;
    }
  } catch(e) { return null; }
}

function _scoreOutputQuality(blocks) {
  if (!blocks || !Array.isArray(blocks) || blocks.length === 0) return 0;
  let score = 0;
  for (const block of blocks) {
    const weeks = block.weeks || [];
    score += Math.min(weeks.length, 20) * 0.05;
    for (const week of weeks) {
      for (const day of (week.days || [])) {
        const exes = day.exercises || [];
        score += exes.length * 0.02;
        for (const ex of exes) {
          if (ex.name && /[a-zA-Z]{3,}/.test(ex.name)) score += 0.01;
          if (ex.prescription) score += 0.01;
        }
      }
    }
  }
  return score;
}

// ── FORMAT DETECTION ──────────────────────────────────────────────────────────
function detectFormat(wb){
  try {
    const scores = [
      scoreA(wb), scoreC(wb), scoreD(wb), scoreE(wb), scoreF(wb),
      scoreG(wb), scoreH(wb), scoreI(wb), scoreJ(wb), scoreK(wb),
      scoreL(wb), scoreM(wb), scoreN(wb), scoreO(wb), scoreP(wb),
      scoreQ(wb), scoreR(wb), scoreS(wb), scoreT(wb), scoreU(wb),
      scoreV(wb), scoreW(wb),
      scoreAdaptive(wb)  // must be last — max 0.3, dedicated parsers always win at >0.3; stable sort preserves order on ties
    ].sort((a, b) => b.score - a.score);

    const top = scores[0];
    const runner = scores[1];

    // Clear winner: high confidence with clear margin
    if (top.score >= 0.6 && (top.score - runner.score) >= 0.15) {
      return top.id;
    }

    // Ambiguous: top two are close — try both parsers, pick better output
    if (top.score >= 0.4 && runner.score >= 0.4 && (top.score - runner.score) < 0.15) {
      try {
        const resultTop = _tryParse(wb, top.id);
        const resultRunner = _tryParse(wb, runner.id);
        const qTop = _scoreOutputQuality(resultTop);
        const qRunner = _scoreOutputQuality(resultRunner);
        return qTop >= qRunner ? top.id : runner.id;
      } catch(e) {
        return top.id;
      }
    }

    // Low confidence but something matched
    if (top.score > 0.2) {
      return top.id;
    }

    // Nothing matched — Column Mapper fallback
    return 'B';
  } catch(e) {
    console.error('[liftlog] detectFormat error:', e);
    return 'B';
  }
}

// ── FORMAT A PARSER ───────────────────────────────────────────────────────────
function parseA(wb){
  const blocks=[];
  for(const sn of wb.SheetNames){
    try{
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null});
      const clean=rows.map(r=>r?r.map(c=>{
        if(typeof c==='string'&&c.startsWith("'")) return c.slice(1);
        return c;
      }):r);
      const b=parseASheet(sn,clean);
      if(b) blocks.push(b);
    }catch(e){ console.error(`[parseA] Error parsing sheet "${sn}":`,e); }
  }
  return blocks;
}

function parseASheet(sn,rows){
  if(rows.length<4) return null;
  let hdr=-1;
  for(let i=0;i<Math.min(10,rows.length);i++){
    const r=rows[i]; if(!r) continue;
    if(r.filter(c=>c&&DAYS.some(d=>String(c).toLowerCase().trim().startsWith(d))).length>=2){hdr=i;break;}
  }
  if(hdr===-1) return null;

  const athleteName=rows[0]?.[0]?String(rows[0][0]).trim():'Athlete';
  const row1val=rows[1]?.[0]?String(rows[1][0]).trim():'';
  const isDateLike=/^\d{1,2}[\/\-]/.test(row1val)||/\d{4}/.test(row1val);
  const name=(!isDateLike&&row1val)?row1val:sn;
  const dateRange=isDateLike?row1val:(rows[2]?.[0]?String(rows[2][0]).trim():'');
  let maxes={};
  for(let i=0;i<Math.min(5,rows.length);i++){
    const r=rows[i]; if(!r) continue;
    for(let j=0;j<r.length;j++){
      const c=_normalizeCell(r[j]);
      if(c==='squat'&&typeof rows[i+1]?.[j]==='number') maxes.squat=rows[i+1][j];
      if(c==='bench'&&typeof rows[i+1]?.[j]==='number') maxes.bench=rows[i+1][j];
      if(c==='deadlift'&&typeof rows[i+1]?.[j]==='number') maxes.deadlift=rows[i+1][j];
    }
  }
  const hrRow=rows[hdr];
  const dayCols=[];
  for(let col=0;col<hrRow.length;col++){
    const c=hrRow[col];
    if(c&&DAYS.some(d=>String(c).toLowerCase().trim().startsWith(d)))
      dayCols.push({col,name:String(c).trim(),noteCol:col+1});
  }
  const rawDays=dayCols.map(({col,name,noteCol})=>{
    const exs=[]; let cur=null; let curSupersetGroup=null;
    for(let r=hdr+1;r<rows.length;r++){
      const row=rows[r]; if(!row) continue;
      const cell=row[col]; const note=row[noteCol];
      if(cell===null||cell===undefined||!String(cell).trim()){ curSupersetGroup=null; continue; }
      const cs=String(cell).trim();
      const wm=cs.match(/^W(\d+):\s*(.+)/i);
      if(wm){ if(cur){ const wTrack=(note!==null&&note!==undefined)?String(note).trim():null; cur.weeks.push({week:parseInt(wm[1]),prescription:wm[2].trim(),note:null,trackingRaw:wTrack}); } }
      else{
        const lp=(/\(/.test(cs)&&(/\d/.test(cs)||/[LMH]/.test(cs)));
        const lp2=/^\d+x\d+/i.test(cs);
        const ln=cs.startsWith('(')||cs.toLowerCase().includes('increase');
        const supersetMatch=
          cs.match(/^(\d+)\s*rounds?[\s:]*/i) ||
          (cs.match(/^superset[\s:]*/i) && ['1']) ||
          (cs.match(/^SS[\s:]*/i) && ['1']) ||
          (cs.match(/^giant\s*set[\s:]*/i) && ['1']) ||
          (cs.match(/^circuit[\s:]*/i) && ['1']);
        const isSectionHeader=
          supersetMatch!=null ||
          /^round\s*\d/i.test(cs) ||
          /^\d+\s*rep\s/i.test(cs) ||
          /^\d+\.?\d*\s*$/.test(cs);
        const wordCount=cs.split(/\s+/).length;
        const isCoachNoteLine=
          /^-.*-$/.test(cs) ||
          /^goal\b/i.test(cs) ||
          /^continue\b/i.test(cs) ||
          /^protocol\b/i.test(cs) ||
          (cs.endsWith('.')&&wordCount>=4&&!lp&&!lp2&&!ln);

        if(supersetMatch){
          const rounds=parseInt(supersetMatch[1])||1;
          const label=rounds>1?rounds+' rounds':'superset';
          curSupersetGroup={label,rounds,startIdx:exs.length};
        }

        if(!lp&&!lp2&&!ln&&!isSectionHeader&&!isCoachNoteLine){
          let normalizedCs=cs.replace(/\s+/g,' ').trim();
          let embeddedPres=null;
          const leadingRepsMatch=normalizedCs.match(/^(\d+(?:[:\.]\d+)?)\s+([A-Za-z].+)$/);
          if(leadingRepsMatch){
            embeddedPres=leadingRepsMatch[1];
            normalizedCs=leadingRepsMatch[2];
          }
          const correctedName=_hCapitalizeName(normalizedCs);
          cur={name:correctedName,weeks:[],staticPrescription:embeddedPres||null,staticTrackingRaw:null,note:note?String(note).trim():null,coachNotes:[],supersetGroup:curSupersetGroup||null};
          exs.push(cur);
        } else if(cur&&(isCoachNoteLine||ln)){
          const noteText=cs.replace(/^-|-$/g,'').trim();
          if(noteText&&!cur.coachNotes.includes(noteText)) cur.coachNotes.push(noteText);
        } else if(cur&&(lp||lp2)&&!cur.staticPrescription){
          cur.staticPrescription=cs;
          if(note){
            const noteStr=String(note).trim();
            const looksLikeTracking=/^[xX✓✔️]$/.test(noteStr)||/^\d/.test(noteStr)||noteStr.toLowerCase()==='done';
            cur.staticTrackingRaw=noteStr;
            if(!looksLikeTracking) cur.note=noteStr;
          }
        }
      }
    }
    return{name,exercises:exs};
  });

  let nw=0;
  for(const d of rawDays) for(const e of d.exercises) if(e.weeks.length>nw) nw=e.weeks.length;
  if(nw===0) nw=4;

  const weeks=[];
  for(let w=1;w<=nw;w++){
    weeks.push({label:`W${w}`,days:rawDays.map(d=>({name:d.name,exercises:d.exercises.map(ex=>{
      const wd=ex.weeks.find(x=>x.week===w);
      const isStatic=!wd;
      return{name:ex.name,prescription:wd?.prescription||ex.staticPrescription||null,note:wd?.note||ex.note||null,coachNotes:ex.coachNotes||[],supersetGroup:ex.supersetGroup||null,trackingRaw:wd?wd.trackingRaw:(isStatic&&w===1?ex.staticTrackingRaw:null),isStaticPrescription:isStatic};
    })}))});
  }

  return{id:sn+'_'+Date.now(),name,dateRange,athleteName,maxes,weeks,format:'A'};
}

// ── FORMAT B PARSER ───────────────────────────────────────────────────────────
function parseB(wb){
  const weeks=[];
  let athleteName='Athlete';
  for(const sn of wb.SheetNames){
    try{
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null});
      const week=parseBSheet(sn,rows);
      if(week) weeks.push(week);
    }catch(e){ console.error(`[parseB] Error parsing sheet "${sn}":`,e); }
  }
  if(weeks.length===0) return [];
  const bName = wb._plFilename ? wb._plFilename.replace(/\.xlsx?$/i,'').trim() : wb.SheetNames[0];
  const bId = bName + '_' + Date.now();
  const blocks=[{id:bId,name:bName||'Training Program',dateRange:'',athleteName,maxes:{},weeks,format:'B'}];
  return blocks;
}

function parseBSheet(sn,rows){
  if(!rows||rows.length===0) return null;
  const days=[]; let curDay=null;
  let colPres=1, colLogged=2;
  const seenDayNames={};
  let weekOffset=0;
  let curSupersetGroup=null;

  for(let i=0;i<rows.length;i++){
    const row=rows[i]; if(!row) continue;
    const a=row[0]; if(!a) continue;
    const aStr=String(a).trim(); if(!aStr) continue;
    if(/^\d+\.?\d*\s*$/.test(aStr)) continue;

    const rowStrs=row.map(c=>String(c||'').toLowerCase());
    const hasSetsHeader=rowStrs.some(s=>s.includes('sets x') || s.includes('sets/') || (s.includes('sets') && s.includes('rpe')));
    const hasRecordHeader=rowStrs.some(s=>s.includes('record weight') || s.includes('record'));
    const isDayHeader=(hasSetsHeader||hasRecordHeader) && aStr.length < 40;

    if(isDayHeader){
      const prescIdx=rowStrs.findIndex(s=>s.includes('sets'));
      const loggedIdx=rowStrs.findIndex(s=>s.includes('record')||s.includes('weight'));
      const coachNoteIdx=rowStrs.findIndex(s=>s.includes('coach'));
      const lifterNoteIdx=rowStrs.findIndex(s=>s.includes('lifter')||s.includes('athlete'));
      if(prescIdx>0) colPres=prescIdx;
      if(loggedIdx>colPres) colLogged=loggedIdx;
      else colLogged=colPres+1;

      if(seenDayNames[aStr]!==undefined){
        weekOffset++;
        Object.keys(seenDayNames).forEach(k=>delete seenDayNames[k]);
      }
      seenDayNames[aStr]=true;

      const uniqueDayName = weekOffset > 0 ? `${aStr} (W${weekOffset+1})` : aStr;
      curDay={name:uniqueDayName,exercises:[],_colCoachNote:coachNoteIdx>0?coachNoteIdx:null,_colLifterNote:lifterNoteIdx>0?lifterNoteIdx:null};
      curSupersetGroup=null;
      days.push(curDay);
    } else if(curDay){
      const pres=row[colPres]!=null?String(row[colPres]).trim():null;
      const rawLogged=row[colLogged];

      let cleanLogged=null;
      if(rawLogged!=null){
        const ls=String(rawLogged).trim();
        if(ls&&!/^[✔️\s]+$/.test(ls)) cleanLogged=ls;
      }

      let note=null;
      if(curDay._colCoachNote!=null&&row[curDay._colCoachNote]!=null){
        const vs=String(row[curDay._colCoachNote]).trim();
        if(vs&&!/^[✔️\s]+$/.test(vs)) note=vs;
      } else {
        for(let c=colLogged+1;c<row.length;c++){
          const v=row[c];
          if(v==null) continue;
          const vs=String(v).trim();
          if(vs&&!/^[✔️\s]+$/.test(vs)){ note=vs; break; }
        }
      }

      let lifterNote=null;
      if(curDay._colLifterNote!=null&&row[curDay._colLifterNote]!=null){
        const vs=String(row[curDay._colLifterNote]).trim();
        if(vs&&!/^[✔️\s]+$/.test(vs)) lifterNote=vs;
      }

      const normName=aStr.replace(/\s+/g,' ').trim();

      const supersetMatch=
        normName.match(/^(\d+)\s*rounds?[\s:]*/i) ||
        (normName.match(/^superset[\s:]*/i) && ['1']) ||
        (normName.match(/^super\s*set[\s:]*/i) && ['1']) ||
        (normName.match(/^ss[\s:]*/i) && ['1']) ||
        (normName.match(/^giant\s*set[\s:]*/i) && ['1']) ||
        (normName.match(/^circuit[\s:]*/i) && ['1']);

      if(supersetMatch){
        const rounds=parseInt(supersetMatch[1])||1;
        const label=rounds>1?rounds+' rounds':'superset';
        curSupersetGroup={label,rounds,startIdx:curDay.exercises.length};
      }

      const isCircuitRow = /^circuit[\s\-:]/i.test(normName) || (normName.match(/,/g)||[]).length >= 2;
      if(isCircuitRow){
        const stripped = normName.replace(/^circuit[\s\-:]+/i,'').trim();
        const parts = stripped.split(/,\s*/);
        if(parts.length >= 2){
          const _circuitIdx = curDay.exercises.length;
          const supersetGroup = curSupersetGroup || { label:`circuit_${_circuitIdx}`, rounds:1, startIdx:_circuitIdx };
          const presStr = pres || '';
          const presParts = presStr.split(/,\s*/);
          if (presParts.length >= parts.length) {
            parts.forEach((exName, idx) => {
              curDay.exercises.push({name:exName.trim(), prescription:(presParts[idx]||'').trim()||null, note, lifterNote, loggedWeight:cleanLogged, supersetGroup});
            });
          } else {
            const exNames = parts.map(p => p.trim().toLowerCase());
            const assigned = new Array(parts.length).fill(null);
            let defaultPres = null;
            presParts.forEach(seg => {
              const segLower = seg.toLowerCase().trim();
              const presMatch = segLower.match(/(\d[\dx\-\s@.secminrep]*\S*)/i);
              const presOnly = presMatch ? presMatch[0].trim() : seg.trim();
              const matched = [];
              exNames.forEach((en, i) => {
                const words = en.split(/\s+/).filter(w => w.length >= 3);
                if (words.some(w => segLower.includes(w))) matched.push(i);
              });
              if (matched.length > 0) {
                matched.forEach(i => { if (!assigned[i]) assigned[i] = presOnly; });
              } else {
                if (!defaultPres) defaultPres = presOnly;
              }
            });
            const fallback = defaultPres || assigned.filter(Boolean)[0] || presStr;
            parts.forEach((exName, idx) => {
              curDay.exercises.push({name:exName.trim(), prescription:assigned[idx]||fallback||null, note, lifterNote, loggedWeight:cleanLogged, supersetGroup});
            });
          }
          continue;
        }
      }

      curDay.exercises.push({name:normName,prescription:pres,note,lifterNote,loggedWeight:cleanLogged,supersetGroup:curSupersetGroup||null});
    }
  }
  // Fallback: Candito-style — dates as day separators, "Set N" column headers, weight/"xN" pairs
  if(days.length===0){
    let curCDay = null;
    for(let i=0;i<rows.length;i++){
      const row=rows[i]; if(!row) continue;
      const a=row[0];
      // Excel serial date as day separator
      if(typeof a === 'number' && a > EXCEL_DATE_SERIAL_MIN && a < 50000){
        // Check if next row has "Set 1"/"Set 2" headers
        const nextRow = rows[i+1];
        if(nextRow && nextRow.some(c => c != null && /^set\s*\d/i.test(String(c).trim()))){
          const dateObj = XLSX.SSF.parse_date_code(a);
          const dayLabel = dateObj ? `Day ${days.length+1}` : `Day ${days.length+1}`;
          curCDay = {name: dayLabel, exercises:[]};
          days.push(curCDay);
          i++; // skip "Set N" header row
          continue;
        }
      }
      if(!curCDay) continue;
      const aStr = a != null ? String(a).trim() : '';
      if(!aStr || /^\d+\.?\d*\s*$/.test(aStr)) continue;
      // Skip instructions/notes (long text)
      if(aStr.length > 80) continue;
      if(/^(no\s|note|if\s|skip|still|extra|back\s*off|reduce|enter|take)/i.test(aStr)) continue;
      // Build prescription from weight/"xN" pairs in columns
      const parts = [];
      for(let c=2; c < row.length; c++){
        const v = row[c]; if(v == null) continue;
        const vs = String(v).trim(); if(!vs) continue;
        if(/^x\s*\d/i.test(vs) || /^x\s*MR/i.test(vs)){
          // rep notation like "x6", "xMR10", "x4-6"
          const reps = vs.replace(/^x\s*/i, '');
          const wt = row[c-1];
          if(typeof wt === 'number' && wt > 0){
            parts.push(`1x${reps}(${Math.round(wt)})`);
          } else {
            parts.push(`1x${reps}`);
          }
        }
      }
      if(parts.length > 0 || /warm\s*up/i.test(String(row[1]||''))){
        const exName = aStr.replace(/\s+/g,' ').trim();
        curCDay.exercises.push({name: exName, prescription: parts.join(', ') || null, note:'', lifterNote:'', loggedWeight:''});
      }
    }
  }
  if(days.length===0) return null;
  return{label:sn,days};
}



// ── POST-PROCESSING UTILITIES ─────────────────────────────────────────────────

function _fixParenTypos(pres) {
  if (!pres) return pres;
  return pres.replace(/(\d+)(L|MH|M|H)\)/gi, '$1($2)');
}

function _isCircuitPrefix(name) {
  if (!name) return null;
  const s = name.trim();
  const m = s.match(/^(\d+)\s*rounds?[\s:]*$/i) ||
            s.match(/^(\d+)\s*circuits?[\s:]*$/i) ||
            s.match(/^(\d+)\s*sets?\s+of[\s:]*$/i);
  if (m) return s.replace(/[\s:]+$/, '').trim();
  if (/^superset[\s:]*$/i.test(s)) return 'superset';
  if (/^giant\s*set[\s:]*$/i.test(s)) return 'giant set';
  if (/^circuit[\s:]*$/i.test(s)) return 'circuit';
  return null;
}

function _isCoachInstruction(name) {
  if (!name) return false;
  const s = name.trim();
  if (/^\(.*\)$/.test(s)) return true;
  if (/^(goal|note:|remember|focus on)\b/i.test(s)) return true;
  const words = s.split(/\s+/);
  if (words.length >= 4 && s.endsWith('.') && !/\d+x\d+/i.test(s) && !/\d+\([^)]+\)/.test(s)) return true;
  return false;
}

function _isSectionHeader(name, pres) {
  if (!name) return false;
  const s = name.trim();
  if (pres && pres.trim()) return false;
  if (/^-+.+-+$/.test(s)) return true;
  if (/^=+.+=+$/.test(s)) return true;
  if (/^\*+.+\*+$/.test(s)) return true;
  return false;
}

const HEADER_BLOCKLIST = new Set([
  'sets', 'reps', 'weight', 'load', 'intensity',
  'rpe', 'rir', 'tempo', 'rest', 'notes', 'note',
  'exercise', 'exercises', 'movement', 'movements',
  'day', 'week', 'date', 'session',
  'warm up', 'warmup', 'warm-up',
  'cool down', 'cooldown', 'cool-down',
  'volume', 'tonnage', 'total',
  'set', 'rep', 'wt', 'wt.',
  '%', 'percent', 'percentage',
  '#', 'no.', 'number'
]);

function _isHeaderTerm(name) {
  if (!name || typeof name !== 'string') return false;
  const cleaned = name.trim().toLowerCase();
  if (HEADER_BLOCKLIST.has(cleaned)) return true;
  if (/^(week|day|session|phase|block|cycle)\s*\d+$/i.test(cleaned)) return true;
  return false;
}

function _normalizeExerciseName(name) {
  if (!name) return name;
  let s = name.trim();
  if (!s) return s;

  // 1. Check exact match first (case-insensitive) against full shorthand aliases
  const lower = s.toLowerCase();
  if (EXERCISE_ALIASES.exact[lower]) {
    return EXERCISE_ALIASES.exact[lower];
  }

  // 2. Word-level abbreviation expansion
  let changed = false;
  const words = s.split(/\s+/);
  const expanded = words.map(word => {
    const canonical = EXERCISE_ALIASES.abbrevs[word.toLowerCase()];
    if (canonical) {
      changed = true;
      return canonical;
    }
    return word;
  });

  if (!changed) {
    return s.replace(/\s+/g, ' ');
  }

  // 3. Rejoin and normalize spacing
  s = expanded.join(' ').replace(/\s+/g, ' ');

  // 4. Title-case the result
  s = s.replace(/(^|[\s-])(\w)/g, (m, pre, ch) => pre + ch.toUpperCase());

  return s;
}

// ── POST-PARSE VALIDATION ─────────────────────────────────────────────────────
/**
 * Post-parse structural validation. Returns { valid, reason, score }.
 * Called after every successful parse BEFORE returning to the app.
 * Rejects output that would produce a broken/confusing user experience.
 * Format B (Column Mapper) is exempt — user explicitly chose that path.
 */
function _validateParseOutput(blocks, formatId) {
  // Format B is a user-driven manual import — skip structural validation
  if (formatId === 'B') return { valid: true, reason: null, score: 1 };

  if (!blocks || !Array.isArray(blocks) || blocks.length === 0) {
    return { valid: false, reason: 'No training blocks found in file.', score: 0 };
  }

  let totalWeeks = 0, totalDays = 0, totalExercises = 0;
  let exercisesWithPrescription = 0;
  let exercisesWithTextName = 0;
  let emptyDays = 0;

  for (const block of blocks) {
    const weeks = block.weeks || [];
    totalWeeks += weeks.length;

    // Sanity: no program has 100+ weeks
    if (weeks.length > 100) {
      return { valid: false, reason: 'Parsed ' + weeks.length + ' weeks — likely a misparse.', score: 0 };
    }

    for (const week of weeks) {
      const days = week.days || [];
      totalDays += days.length;

      // Sanity: no training week has 14+ days
      if (days.length > 14) {
        return { valid: false, reason: 'Week "' + (week.label || '?') + '" has ' + days.length + ' days — likely a misparse.', score: 0 };
      }

      for (const day of days) {
        const exes = day.exercises || [];
        if (exes.length === 0) { emptyDays++; continue; }
        for (const ex of exes) {
          totalExercises++;
          if (ex.prescription && String(ex.prescription).trim()) exercisesWithPrescription++;
          if (ex.name && /[a-zA-Z]{2,}/.test(ex.name)) exercisesWithTextName++;
        }
      }
    }
  }

  // Must have at least 1 week, 1 day, 1 exercise
  if (totalWeeks === 0) return { valid: false, reason: 'No weeks found.', score: 0 };
  if (totalDays === 0) return { valid: false, reason: 'No training days found.', score: 0 };
  if (totalExercises === 0) return { valid: false, reason: 'No exercises found.', score: 0 };

  // Exercise names should contain at least one letter (not all pure numbers/percentages)
  const textNameRatio = exercisesWithTextName / totalExercises;
  if (textNameRatio < 0.5) {
    return { valid: false, reason: 'Most exercise names appear to be numbers or symbols, not real exercises. The file may not match the detected format.', score: textNameRatio };
  }

  // At least 30% of exercises should have prescriptions (sets/reps)
  const prescriptionRatio = exercisesWithPrescription / totalExercises;
  if (prescriptionRatio < 0.3) {
    return { valid: false, reason: 'Less than 30% of exercises have set/rep prescriptions. The parser may have misread the file layout.', score: prescriptionRatio };
  }

  // If more than 60% of days are empty, something went wrong
  const emptyDayRatio = totalDays > 0 ? emptyDays / totalDays : 0;
  if (emptyDayRatio > 0.6 && totalDays > 2) {
    return { valid: false, reason: 'Most training days are empty. The parser may not be reading the correct cells.', score: 1 - emptyDayRatio };
  }

  const score = (textNameRatio * 0.4) + (prescriptionRatio * 0.3) + ((1 - emptyDayRatio) * 0.3);
  return { valid: true, reason: null, score };
}

// ── UNIFIED ENTRY POINT ───────────────────────────────────────────────────────
function parseWorkbook(wb){
  const fmt = detectFormat(wb);
  let blocks;
  if (fmt === 'A') blocks = parseA(wb);
  else if (fmt === 'I') blocks = parseTexasMethod(wb);
  else if (fmt === 'J') blocks = parseHepburn(wb);
  else if (fmt === 'K') blocks = parseBulgarianMethod(wb);
  else if (fmt === 'E') blocks = parseE(wb);
  else if (fmt === 'F') blocks = parseF(wb);
  else if (fmt === 'G') blocks = parseG(wb);
  else if (fmt === 'H') blocks = parseH(wb);
  else if (fmt === 'L') blocks = parseL(wb);
  else if (fmt === 'M') blocks = parseM(wb);
  else if (fmt === 'N') blocks = parseN(wb);
  else if (fmt === 'O') blocks = parseO(wb);
  else if (fmt === 'P') blocks = parseP(wb);
  else if (fmt === 'Q') blocks = parseQ(wb);
  else if (fmt === 'R') blocks = parseR(wb);
  else if (fmt === 'S') blocks = parseS(wb);
  else if (fmt === 'T') blocks = parseT(wb);
  else if (fmt === 'U') blocks = parseU(wb);
  else if (fmt === 'V') blocks = parseV(wb);
  else if (fmt === 'W') blocks = parseW(wb);
  else if (fmt === 'ADAPTIVE') blocks = parseAdaptive(wb);
  else if (fmt === 'C') blocks = parseCAutoFormat(wb);
  else if (fmt === 'D') blocks = parseD(wb);
  else blocks = parseB(wb);

  // ── Post-processing (all fixes apply to ALL parser formats) ──────────────

  if (blocks && Array.isArray(blocks)) {

    // Fix 1: Duplicate Block Name Detection
    if (blocks.length > 1) {
      const nameCount = {};
      for (const b of blocks) nameCount[b.name] = (nameCount[b.name] || 0) + 1;
      const hasDupes = Object.values(nameCount).some(c => c > 1);
      if (hasDupes && blocks.length === wb.SheetNames.length) {
        const sheetNameCount = {};
        for (const sn of wb.SheetNames) sheetNameCount[sn] = (sheetNameCount[sn] || 0) + 1;
        const sheetNamesUnique = Object.values(sheetNameCount).every(c => c === 1);
        if (sheetNamesUnique) {
          for (let i = 0; i < blocks.length; i++) {
            blocks[i].name = wb.SheetNames[i];
          }
        }
      }
      // Fallback: append suffix for any remaining duplicates
      const seen = {};
      for (const b of blocks) {
        seen[b.name] = (seen[b.name] || 0) + 1;
        if (seen[b.name] > 1) {
          b.name = b.name + ' (' + seen[b.name] + ')';
        }
      }
    }

    // Fixes 2–5: Exercise-level post-processing
    for (const block of blocks) {
      if (!block.weeks) continue;
      for (const week of block.weeks) {
        if (!week.days) continue;
        for (const day of week.days) {
          if (!day.exercises) continue;

          // Fix 3: Paren typos on all prescriptions
          for (const ex of day.exercises) {
            if (ex.prescription) ex.prescription = _fixParenTypos(ex.prescription);
          }

          // Fixes 2, 4, 5: Filter and restructure exercises
          const filtered = [];
          let circuitLabel = null;

          for (let i = 0; i < day.exercises.length; i++) {
            const ex = day.exercises[i];
            const name = (ex.name || '').trim();
            const pres = (ex.prescription || '').trim();

            // Fix 5: Section headers — remove
            if (_isSectionHeader(name, pres)) {
              continue;
            }

            // Fix 7: Header term blocklist — remove leaked spreadsheet headers
            if (_isHeaderTerm(name) && (!pres || _isHeaderTerm(pres))) {
              continue;
            }

            // Fix 2: Circuit prefix — remove, capture label
            const prefix = _isCircuitPrefix(name);
            if (prefix && !pres) {
              circuitLabel = prefix;
              continue;
            }

            // Fix 4: Coach instructions — remove, attach to adjacent exercise
            if (_isCoachInstruction(name) && !pres) {
              if (filtered.length > 0) {
                const prev = filtered[filtered.length - 1];
                prev.note = prev.note ? prev.note + '; ' + name : name;
              }
              continue;
            }

            // Fix 2: Extract embedded rep-range from exercise name
            if (!pres) {
              const rangeMatch = name.match(/^(\d+[-\u2013]\d+)\s+(.+)$/);
              if (rangeMatch) {
                ex.prescription = rangeMatch[1];
                ex.name = rangeMatch[2].trim();
              }
            }

            // Fix 2: Apply circuit context (from removed prefix or supersetGroup)
            const effectiveCircuit = circuitLabel || (ex.supersetGroup && !pres ? ex.supersetGroup.label : null);
            if (effectiveCircuit && !pres) {
              if (!ex.note || !ex.note.includes(effectiveCircuit)) {
                ex.note = ex.note ? effectiveCircuit + '; ' + ex.note : effectiveCircuit;
              }
            }

            // Reset circuit label when hitting an exercise that already had a prescription
            if (pres) {
              circuitLabel = null;
            }

            filtered.push(ex);
          }

          day.exercises = filtered;

          // Fix 6: Exercise name normalization (alias dictionary)
          for (const ex of day.exercises) {
            if (ex.name) ex.name = _normalizeExerciseName(ex.name);
          }
        }
      }
    }

    // Normalize null fields to empty strings across all blocks
    for (const block of blocks) {
      if (!block.weeks) continue;
      for (const week of block.weeks) {
        if (!week.days) continue;
        for (const day of week.days) {
          if (!day.exercises) continue;
          for (const ex of day.exercises) {
            if (ex.prescription == null) ex.prescription = '';
            if (ex.note == null) ex.note = '';
            if (ex.loggedWeight == null) ex.loggedWeight = '';
            if (ex.lifterNote == null) ex.lifterNote = '';
            if (!('supersetGroup' in ex)) ex.supersetGroup = null;
          }
        }
      }
    }
  }

  // ── Post-parse validation gate ────────────────────────────────────────────
  const validation = _validateParseOutput(blocks, fmt);
  if (!validation.valid) {
    console.warn('[liftlog] Parse validation failed for format ' + fmt + ': ' + validation.reason);
    return { error: true, reason: validation.reason, format: fmt };
  }

  return blocks;
}

// ── FORMAT E PARSER (flat grid / nSuns-like / week-day text) ─────────────────
function parseE(wb) {
  let result;
  result = parseE_nSuns(wb);
  if (result && result.length > 0) return result;
  result = parseE_tabular(wb);
  if (result && result.length > 0) return result;
  result = parseE_phaseGrid(wb);
  if (result && result.length > 0) return result;
  result = parseE_texasMethod(wb);
  if (result && result.length > 0) return result;
  result = parseE_weekText(wb);
  if (result && result.length > 0) return result;
  result = parseE_cycleWeek(wb);
  if (result && result.length > 0) return result;
  result = parseE_parallel531(wb);
  if (result && result.length > 0) return result;
  return parseB(wb);
}

// Helper: clean row data
function _eCleanRows(ws) {
  return XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
}

// Helper: skip non-training sheets
function _eIsSkipSheet(name) {
  return /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|output)$/i.test(name.trim());
}

// Helper: extract 1RM maxes from header area
function _eExtractMaxes(rows, maxRow) {
  const maxes = {};
  for (let i = 0; i < Math.min(maxRow || 20, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    for (let c = 0; c < row.length - 1; c++) {
      const val = row[c]; if (val == null) continue;
      const label = String(val).trim().toLowerCase();
      const findNum = () => {
        for (let nc = c + 1; nc < Math.min(c + 4, row.length); nc++) {
          if (typeof row[nc] === 'number' && row[nc] > 0) return row[nc];
        }
        return null;
      };
      if (/^squat|back\s*squat/i.test(label) && !maxes.squat) { const v = findNum(); if (v) maxes.squat = v; }
      else if (/^bench/i.test(label) && !maxes.bench) { const v = findNum(); if (v) maxes.bench = v; }
      else if (/^(dead|dl$|sumo\s*dead|conventional)/i.test(label) && !maxes.deadlift) { const v = findNum(); if (v) maxes.deadlift = v; }
      else if (/^(press|ohp|overhead|military|o\.?h\.?p)/i.test(label) && !maxes.press) { const v = findNum(); if (v) maxes.press = v; }
      else if (/max.*\(?lb/i.test(label) || /^max$/i.test(label) || /1\s*rm/i.test(label) || /enter.*1\s*rm/i.test(label) || /current.*1?\s*rm/i.test(label)) {
        const v = findNum();
        if (v && !maxes.squat) maxes.squat = v; // generic max defaults to first lift
      }
    }
  }
  return maxes;
}

// ── E Sub-parser: nSuns / 2-Suns (day names + alternating weight/rep columns) ─
function parseE_nSuns(wb) {
  const allBlocks = [];
  for (const sn of wb.SheetNames) {
    if (_eIsSkipSheet(sn)) continue;
    if (/^(inputs?|maxes?)$/i.test(sn.trim())) continue;
    const rows = _eCleanRows(wb.Sheets[sn]);
    if (!rows || rows.length < 10) continue;

    // Find data column and first day row
    let dataCol = -1, firstDayRow = -1;
    for (let i = 0; i < Math.min(50, rows.length) && firstDayRow < 0; i++) {
      const row = rows[i]; if (!row) continue;
      for (let col = 0; col <= 2; col++) {
        const val = row[col]; if (!val) continue;
        const str = String(val).trim().toLowerCase();
        const isDayName = DAYS.some(d => str === d || str.startsWith(d + ' '));
        const isDayLabel = /^day\s*\d/i.test(String(val).trim());
        if (!isDayName && !isDayLabel) continue;
        // Must be standalone (not multi-day header)
        const otherDays = row.filter((c2, ci) => ci !== col && c2 && DAYS.some(d => String(c2).trim().toLowerCase().startsWith(d)));
        if (otherDays.length > 0) continue;
        // Skip if nearby rows have column headers (Sets/Reps/Load → D/B format, not nSuns)
        let hasColHeaders = false;
        for (let j = i; j < Math.min(i + 3, rows.length); j++) {
          const hr = rows[j]; if (!hr) continue;
          const hstrs = hr.map(_normalizeCell);
          if ((hstrs.includes('sets') || hstrs.includes('set')) && (hstrs.includes('reps') || hstrs.includes('rep goal'))) hasColHeaders = true;
          if (hstrs.includes('movement') || hstrs.includes('exercise movement')) hasColHeaders = true;
        }
        if (hasColHeaders) continue;
        // Verify numeric data below
        let hasNumData = false;
        for (let j = i + 1; j < Math.min(i + 8, rows.length); j++) {
          const dr = rows[j]; if (!dr) continue;
          let numCount = 0;
          for (let c2 = col + 1; c2 < dr.length; c2++) if (typeof dr[c2] === 'number') numCount++;
          if (numCount >= 6) { hasNumData = true; break; }
        }
        if (hasNumData) { dataCol = col; firstDayRow = i; break; }
      }
    }
    if (firstDayRow < 0) continue;

    const maxes = _eExtractMaxes(rows, firstDayRow);

    // Find day section boundaries
    const daySections = [];
    for (let i = firstDayRow; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const val = row[dataCol]; if (!val) continue;
      const str = String(val).trim();
      if (DAYS.some(d => str.toLowerCase() === d || str.toLowerCase().startsWith(d + ' ')) || /^day\s*\d/i.test(str)) {
        daySections.push({ name: str, startRow: i + 1 });
      }
    }
    for (let i = 0; i < daySections.length; i++) {
      daySections[i].endRow = i + 1 < daySections.length ? daySections[i + 1].startRow - 1 : rows.length;
    }
    if (daySections.length === 0) continue;

    // Parse exercises for each day
    const days = [];
    for (const section of daySections) {
      const exercises = [];
      for (let r = section.startRow; r < section.endRow; r++) {
        const row = rows[r]; if (!row) continue;
        const exVal = row[dataCol]; if (exVal == null) continue;
        const exStr = String(exVal).trim();
        if (!exStr) continue;
        if (/^(assist|accessor)/i.test(exStr)) continue;
        if (DAYS.some(d => exStr.toLowerCase() === d || exStr.toLowerCase().startsWith(d + ' ')) || /^day\s*\d/i.test(exStr)) break;
        if (/^\d+\.?\d*$/.test(exStr)) continue;
        if (/^(1RM|TM|lb|kg)/i.test(exStr)) continue;

        // Find first numeric cell after dataCol
        let numStart = -1;
        for (let c = dataCol + 1; c < Math.min(row.length, 40); c++) {
          if (typeof row[c] === 'number') { numStart = c; break; }
        }
        if (numStart < 0) continue;

        // Parse alternating weight/reps pairs
        const pairs = [];
        for (let c = numStart; c < row.length; c += 2) {
          const weightCell = row[c];
          const repsCell = c + 1 < row.length ? row[c + 1] : null;
          if (weightCell == null && repsCell == null) break;
          if (weightCell == null) continue;
          let weight;
          if (typeof weightCell === 'number') weight = weightCell;
          else { const ws = String(weightCell).trim(); if (/^\d+\.?\d*$/.test(ws)) weight = parseFloat(ws); else continue; }
          if (repsCell == null) continue;
          let repStr = String(repsCell).trim().replace(/^[xX]\s*/, '');
          if (!repStr) continue;
          pairs.push({ weight: Math.round(weight * 100) / 100, reps: repStr });
        }
        if (pairs.length === 0) continue;

        // Collapse consecutive same weight+reps into groups
        const groups = [];
        let pi = 0;
        while (pi < pairs.length) {
          let count = 1;
          while (pi + count < pairs.length && pairs[pi].weight === pairs[pi + count].weight && pairs[pi].reps === pairs[pi + count].reps) count++;
          groups.push(pairs[pi].weight > 0 ? `${count}x${pairs[pi].reps}(${pairs[pi].weight})` : `${count}x${pairs[pi].reps}`);
          pi += count;
        }
        exercises.push({ name: exStr.replace(/\s+/g, ' '), prescription: groups.join(', '), note: '', lifterNote: '', loggedWeight: '' });
      }
      if (exercises.length > 0) days.push({ name: section.name, exercises });
    }

    if (days.length > 0) {
      const blockName = (sn === 'Sheet1' || /^sheet\d*$/i.test(sn))
        ? (wb._plFilename || 'nSuns Program').replace(/\.xlsx?$/i, '')
        : sn;
      allBlocks.push({
        id: 'e_' + sn.replace(/\s+/g, '_') + '_' + Date.now(),
        name: blockName, format: 'E', athleteName: '', dateRange: '', maxes,
        weeks: [{ label: 'Week 1', days }]
      });
    }
  }
  return allBlocks.length > 0 ? allBlocks : null;
}

// ── E Sub-parser: Tabular (Prep10-style: Week/Phase/Session/Lift/Sets/Reps columns) ─
function parseE_tabular(wb) {
  for (const sn of wb.SheetNames) {
    if (_eIsSkipSheet(sn)) continue;
    const rows = _eCleanRows(wb.Sheets[sn]);
    if (!rows || rows.length < 5) continue;

    // Find header row with Week/Lift/Sets/Reps columns
    let hdrRow = -1, colMap = {};
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      const wIdx = strs.indexOf('week');
      const lIdx = strs.findIndex(s => s.includes('primary lift') || s === 'exercise');
      const sIdx = strs.indexOf('sets');
      const rIdx = strs.indexOf('reps');
      if (wIdx >= 0 && lIdx >= 0 && (sIdx >= 0 || rIdx >= 0)) {
        hdrRow = i;
        colMap = {
          week: wIdx, phase: strs.indexOf('phase'), session: strs.indexOf('session'),
          lift: lIdx, weight: strs.indexOf('weight'), sets: sIdx, reps: rIdx,
          assist: strs.findIndex(s => s.includes('assistance') || s.includes('suggested')),
          assistPres: strs.findIndex(s => s.includes('sets x reps') || s.includes('sets×reps'))
        };
        break;
      }
    }
    if (hdrRow < 0) continue;

    const maxes = _eExtractMaxes(rows, hdrRow);

    // Parse rows into week→day→exercise structure
    const weekMap = new Map(); // weekNum → Map(session → {name, exercises[]})
    for (let i = hdrRow + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const weekVal = row[colMap.week];
      const liftVal = row[colMap.lift];
      if (weekVal == null && liftVal == null) continue;
      const weekNum = typeof weekVal === 'number' ? weekVal : parseInt(String(weekVal || ''));
      if (isNaN(weekNum)) continue;
      const lift = liftVal ? String(liftVal).trim() : '';
      if (!lift) continue;

      const sessionVal = colMap.session >= 0 ? row[colMap.session] : 1;
      const sessionNum = typeof sessionVal === 'number' ? sessionVal : parseInt(String(sessionVal || '1'));
      const phase = colMap.phase >= 0 && row[colMap.phase] ? String(row[colMap.phase]).trim() : '';

      let setsVal = colMap.sets >= 0 && row[colMap.sets] != null ? String(row[colMap.sets]).trim().replace(/\.0$/, '') : '';
      let repsVal = colMap.reps >= 0 && row[colMap.reps] != null ? String(row[colMap.reps]).trim().replace(/\.0$/, '') : '';
      let weightVal = colMap.weight >= 0 && row[colMap.weight] != null ? row[colMap.weight] : null;

      let pres = '';
      if (setsVal && repsVal) pres = setsVal + 'x' + repsVal;
      else if (setsVal) pres = setsVal;
      else if (repsVal) pres = repsVal;
      if (typeof weightVal === 'number' && weightVal > 0 && weightVal <= 1) {
        pres += '(' + Math.round(weightVal * 100) + '%)';
      } else if (typeof weightVal === 'number' && weightVal > 1) {
        pres += '(' + weightVal + ')';
      }

      if (!weekMap.has(weekNum)) weekMap.set(weekNum, new Map());
      const dayMap = weekMap.get(weekNum);
      const dayKey = sessionNum || 1;
      const dayName = phase ? `Session ${dayKey} (${phase})` : `Session ${dayKey}`;
      if (!dayMap.has(dayKey)) dayMap.set(dayKey, { name: dayName, exercises: [] });
      dayMap.get(dayKey).exercises.push({ name: lift, prescription: pres || null, note: '', lifterNote: '', loggedWeight: '' });

      // Assistance exercise in same row
      if (colMap.assist >= 0 && row[colMap.assist]) {
        const aName = String(row[colMap.assist]).trim();
        const aPres = colMap.assistPres >= 0 && row[colMap.assistPres] ? String(row[colMap.assistPres]).trim() : '';
        if (aName && !/^\d+\.?\d*$/.test(aName)) {
          dayMap.get(dayKey).exercises.push({ name: aName, prescription: aPres || null, note: 'Assistance', lifterNote: '', loggedWeight: '' });
        }
      }
    }

    if (weekMap.size === 0) continue;
    const weeks = [];
    for (const [wk, dayMap] of [...weekMap.entries()].sort((a, b) => a[0] - b[0])) {
      const days = [...dayMap.values()].filter(d => d.exercises.length > 0);
      if (days.length > 0) weeks.push({ label: `Week ${wk}`, days });
    }
    if (weeks.length === 0) continue;

    const blockName = (wb._plFilename || sn).replace(/\.xlsx?$/i, '');
    return [{
      id: 'e_tab_' + Date.now(), name: blockName, format: 'E',
      athleteName: '', dateRange: '', maxes, weeks
    }];
  }
  return null;
}

// ── E Sub-parser: Phase Grid (Juggernaut — phase column groups with Weight/Reps) ─
function parseE_phaseGrid(wb) {
  for (const sn of wb.SheetNames) {
    if (_eIsSkipSheet(sn)) continue;
    const rows = _eCleanRows(wb.Sheets[sn]);
    if (!rows || rows.length < 15) continue;

    // Find the Weight/Reps header row with multiple column groups
    let hdrRow = -1, phaseGroups = []; // [{name, weightCol, repsCol}]
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      const wCols = []; const rCols = [];
      for (let c = 0; c < strs.length; c++) {
        if (strs[c] === 'weight') wCols.push(c);
        if (strs[c] === 'reps') rCols.push(c);
      }
      if (wCols.length >= 2 && rCols.length >= 2) {
        hdrRow = i;
        // Match weight-reps pairs
        for (let wi = 0; wi < wCols.length; wi++) {
          const wc = wCols[wi];
          const rc = rCols.find(r => r > wc && r < (wCols[wi + 1] || 999));
          if (rc != null) {
            // Get phase name from the row above
            let phaseName = `Phase ${wi + 1}`;
            if (i > 0 && rows[i - 1]) {
              for (let pc = wc - 1; pc <= wc + 1; pc++) {
                const pv = rows[i - 1][pc];
                if (pv && /\w{3,}/.test(String(pv))) { phaseName = String(pv).trim(); break; }
              }
            }
            phaseGroups.push({ name: phaseName, weightCol: wc, repsCol: rc });
          }
        }
        break;
      }
    }
    if (hdrRow < 0 || phaseGroups.length < 2) continue;

    const maxes = _eExtractMaxes(rows, hdrRow);

    // Find wave/section boundaries below the header
    const waveSections = []; // [{name, startRow, endRow}]
    for (let i = hdrRow - 2; i >= 0; i--) {
      const row = rows[i]; if (!row) continue;
      const val = row[0]; if (val == null) continue;
      const str = String(val).trim();
      if (/\d+\s*Rep\s*Wave/i.test(str) || /Wave/i.test(str)) {
        waveSections.push({ name: str, startRow: hdrRow + 1 });
        break;
      }
    }

    // Find additional wave sections below
    for (let i = hdrRow + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const val = row[0]; if (val == null) continue;
      const str = String(val).trim();
      if (/\d+\s*Rep\s*Wave/i.test(str)) {
        // End previous wave
        if (waveSections.length > 0) waveSections[waveSections.length - 1].endRow = i;
        // Check for Weight/Reps headers below
        let newHdr = -1;
        for (let j = i + 1; j < Math.min(i + 3, rows.length); j++) {
          const hr = rows[j]; if (!hr) continue;
          const hstrs = hr.map(_normalizeCell);
          if (hstrs.filter(s => s === 'weight').length >= 2) { newHdr = j; break; }
        }
        waveSections.push({ name: str, startRow: (newHdr >= 0 ? newHdr + 1 : i + 2) });
      }
    }
    if (waveSections.length > 0 && !waveSections[waveSections.length - 1].endRow) {
      waveSections[waveSections.length - 1].endRow = rows.length;
    }
    if (waveSections.length === 0) {
      waveSections.push({ name: 'Wave 1', startRow: hdrRow + 1, endRow: rows.length });
    }

    // Parse each wave section into weeks (one per phase column group)
    const allWeeks = [];
    for (const wave of waveSections) {
      for (const pg of phaseGroups) {
        const exercises = [];
        let currentEx = null;
        for (let r = wave.startRow; r < wave.endRow; r++) {
          const row = rows[r]; if (!row) continue;
          // Check for exercise name in col adjacent to weight col
          const nameCol = pg.weightCol - 1;
          const exName = nameCol >= 0 && row[nameCol] ? String(row[nameCol]).trim() : '';
          const weight = row[pg.weightCol];
          const reps = row[pg.repsCol];

          if (exName && /^[A-Z]/i.test(exName) && exName.length >= 3 && !/^(New|Training|Failure|\d)/i.test(exName)) {
            if (currentEx && currentEx.sets.length > 0) {
              exercises.push(currentEx);
            }
            currentEx = { name: exName, sets: [] };
          }
          if (currentEx && typeof weight === 'number' && weight > 0 && reps != null) {
            let repStr = String(reps).trim().replace(/\.0$/, '');
            if (repStr === '-' || !repStr) continue;
            currentEx.sets.push({ weight: Math.round(weight * 100) / 100, reps: repStr });
          }
          // Check for wave boundary (exercise name with no data means new section)
          if (/\d+\s*Rep\s*Wave/i.test(String(row[0] || ''))) break;
        }
        if (currentEx && currentEx.sets.length > 0) exercises.push(currentEx);

        if (exercises.length > 0) {
          const dayExercises = exercises.map(ex => {
            const groups = [];
            let si = 0;
            while (si < ex.sets.length) {
              let count = 1;
              while (si + count < ex.sets.length && ex.sets[si].weight === ex.sets[si + count].weight && ex.sets[si].reps === ex.sets[si + count].reps) count++;
              groups.push(`${count}x${ex.sets[si].reps}(${ex.sets[si].weight})`);
              si += count;
            }
            return { name: ex.name, prescription: groups.join(', '), note: '', lifterNote: '', loggedWeight: '' };
          });
          allWeeks.push({
            label: `${wave.name} - ${pg.name}`.substring(0, 50),
            days: [{ name: 'Training Day', exercises: dayExercises }]
          });
        }
      }
    }

    if (allWeeks.length === 0) continue;
    const blockName = (wb._plFilename || sn).replace(/\.xlsx?$/i, '');
    return [{
      id: 'e_pg_' + Date.now(), name: blockName, format: 'E',
      athleteName: '', dateRange: '', maxes, weeks: allWeeks
    }];
  }
  return null;
}

// ── E Sub-parser: Week-Day Text (Russian Squat, Smolov, MagOrt — Week N + SxR text) ─
function parseE_weekText(wb) {
  const allWeeks = [];
  let maxes = {};
  let programName = '';

  for (const sn of wb.SheetNames) {
    if (_eIsSkipSheet(sn)) continue;
    if (/^(inputs?|maxes?)$/i.test(sn.trim())) continue;
    const rows = _eCleanRows(wb.Sheets[sn]);
    if (!rows || rows.length < 3) continue;

    // Check this sheet has SxR text
    let sheetHasSxR = false;
    for (let i = 0; i < Math.min(40, rows.length) && !sheetHasSxR; i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        if (row[c] != null && /^\d+\s*x\s*\d/i.test(String(row[c]).trim())) { sheetHasSxR = true; break; }
      }
    }
    if (!sheetHasSxR) continue;

    const sheetMaxes = _eExtractMaxes(rows, 20);
    Object.assign(maxes, sheetMaxes);

    // Find Week N section boundaries
    const weekSections = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < Math.min(row.length, 15); c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        const wm = str.match(/^Week\s*(\d+)/i);
        if (wm) weekSections.push({ weekNum: parseInt(wm[1]), startRow: i, col: c });
      }
    }

    // Set end rows for each week section
    for (let i = 0; i < weekSections.length; i++) {
      let endRow = rows.length;
      for (let j = i + 1; j < weekSections.length; j++) {
        if (weekSections[j].startRow > weekSections[i].startRow) { endRow = weekSections[j].startRow; break; }
      }
      weekSections[i].endRow = endRow;
    }

    // Deduplicate (horizontal layout may have multiple Week labels at same row)
    const deduped = [];
    const seenKey = new Set();
    for (const ws of weekSections) {
      const key = `${ws.weekNum}_${ws.col}`;
      if (!seenKey.has(key)) { seenKey.add(key); deduped.push(ws); }
    }

    if (deduped.length === 0) {
      // No Week labels found — treat entire sheet as a single week
      deduped.push({ weekNum: allWeeks.length + 1, startRow: 0, endRow: rows.length, col: 0 });
    }

    for (const ws of deduped) {
      const days = _eParseWeekTextDays(rows, ws.startRow, ws.endRow, sn);
      if (days.length > 0) {
        allWeeks.push({ label: `Week ${ws.weekNum}`, days });
        if (!programName) programName = sn;
      }
    }
  }

  if (allWeeks.length === 0) return null;
  const blockName = programName && !/^sheet\d*$/i.test(programName) && !/^(introductory|base|switching|intense)/i.test(programName)
    ? programName
    : (wb._plFilename || 'Program').replace(/\.xlsx?$/i, '');
  return [{
    id: 'e_wt_' + Date.now(), name: blockName, format: 'E',
    athleteName: '', dateRange: '', maxes, weeks: allWeeks
  }];
}

function _eParseWeekTextDays(rows, startRow, endRow, sheetName) {
  // Check for horizontal Day N labels
  const dayColumns = [];
  for (let i = startRow; i < Math.min(startRow + 4, endRow); i++) {
    const row = rows[i]; if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c]; if (val == null) continue;
      if (/^Day\s*\d/i.test(String(val).trim())) {
        dayColumns.push({ col: c, name: String(val).trim(), headerRow: i });
      }
    }
    if (dayColumns.length >= 2) break;
  }

  if (dayColumns.length >= 2) {
    // Horizontal day layout: Day N headers span the top, exercises listed per row.
    // Two sub-layouts:
    //   RSR-style: col A = exercise name, dc.col = prescription (no weight column)
    //   3-col-style: dc.col = exercise name, dc.col+1 = SxR, dc.col+2 = weight
    const dataStart = dayColumns[0].headerRow + 1;
    const days = [];

    // Detect RSR-style: first Day column is not col 0, and col 0 has text exercise names
    const firstDayCol = dayColumns[0].col;
    let exNamesInColA = false;
    if (firstDayCol > 0) {
      for (let r = dataStart; r < Math.min(dataStart + 6, endRow); r++) {
        const row = rows[r]; if (!row) continue;
        const v = row[0];
        if (v != null) {
          const s = String(v).trim();
          if (s && /[a-zA-Z]{3,}/.test(s) && !/^(Week|Day|Warm|REST|TEST|MAX)/i.test(s)) {
            exNamesInColA = true; break;
          }
        }
      }
    }

    for (const dc of dayColumns) {
      const exercises = [];
      for (let r = dataStart; r < endRow; r++) {
        const row = rows[r]; if (!row) continue;
        // Stop at next Week header
        for (let cc = 0; cc < Math.min(row.length, 3); cc++) {
          if (row[cc] != null && /^Week\s*\d/i.test(String(row[cc]).trim())) return days.length > 0 ? days : [];
        }
        let name, pres;
        if (exNamesInColA) {
          // RSR-style: exercise name in col A, prescription in dc.col
          const colAVal = row[0];
          const presVal = row[dc.col];
          if (colAVal == null && presVal == null) continue;
          name = colAVal ? String(colAVal).trim() : '';
          pres = presVal ? String(presVal).trim() : '';
        } else {
          // 3-col-style: name at dc.col, SxR at dc.col+1, weight at dc.col+2
          const exName = row[dc.col];
          const sxr = row[dc.col + 1];
          const weight = row[dc.col + 2];
          if (exName == null && sxr == null) continue;
          name = exName ? String(exName).trim() : '';
          pres = sxr ? String(sxr).trim() : '';
          if (weight != null && typeof weight === 'number' && weight > 0) {
            pres = pres ? `${pres}(${Math.round(weight * 100) / 100})` : `(${Math.round(weight * 100) / 100})`;
          } else if (weight != null) {
            const ws = String(weight).trim();
            if (ws && !/^\d+\.?\d*$/.test(ws)) pres = pres ? `${pres} ${ws}` : ws;
          }
        }
        if (!name || /^\d+\.?\d*$/.test(name)) continue;
        if (/^(Week|Day|Warm|REST|TEST|MAX|LB|KG|1\s*RM)/i.test(name)) continue;
        exercises.push({ name: name.replace(/\s+/g, ' '), prescription: pres || null, note: '', lifterNote: '', loggedWeight: '' });
      }
      if (exercises.length > 0) days.push({ name: dc.name, exercises });
    }
    return days;
  }

  // Check for horizontal SxR+weight column pairs (Smolov Base Mesocycle style)
  // Pattern: a row has 2+ SxR text values each immediately followed by a number
  let multiDayCols = null;
  let multiDayDataRow = -1;
  for (let r = startRow; r < Math.min(startRow + 25, endRow); r++) {
    const row = rows[r]; if (!row) continue;
    const pairs = [];
    for (let c = 0; c < row.length - 1; c++) {
      if (row[c] != null && /^\d+\s*x\s*\d/i.test(String(row[c]).trim())) {
        if (typeof row[c + 1] === 'number' && row[c + 1] > 0) {
          pairs.push({ sxrCol: c, weightCol: c + 1 });
          c++; // skip weight column
        }
      }
    }
    if (pairs.length >= 2) { multiDayCols = pairs; multiDayDataRow = r; break; }
  }

  if (multiDayCols && multiDayCols.length >= 2) {
    // Determine exercise names — check row above data for per-column names
    const perDayNames = [];
    if (multiDayDataRow > 0) {
      const nameRow = rows[multiDayDataRow - 1];
      if (nameRow) {
        for (const col of multiDayCols) {
          const val = nameRow[col.sxrCol];
          perDayNames.push(val != null && typeof val === 'string' && /[a-zA-Z]{3,}/.test(val.trim()) ? val.trim() : null);
        }
      }
    }

    // Fallback exercise name from header area or sheet name
    let exName = sheetName || 'Exercise';
    for (let i = 0; i < Math.min(startRow + 10, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        if (/^(Back Squat|Front Squat|Squat|Bench|Deadlift|Press)/i.test(str)) { exName = str; break; }
      }
    }
    if (exName === sheetName) {
      for (let i = 0; i < Math.min(startRow + 10, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          if (row[c] == null) continue;
          const s = String(row[c]).trim().toLowerCase();
          if (/\bsquat\b/.test(s)) { exName = 'Squat'; break; }
          if (/\bbench\b/.test(s)) { exName = 'Bench'; break; }
          if (/\bdeadlift\b/.test(s)) { exName = 'Deadlift'; break; }
        }
      }
    }

    const days = multiDayCols.map((_, idx) => ({ name: `Day ${idx + 1}`, exercises: [] }));
    for (let r = startRow; r < endRow; r++) {
      const row = rows[r]; if (!row) continue;
      for (let di = 0; di < multiDayCols.length; di++) {
        const { sxrCol, weightCol } = multiDayCols[di];
        const sxr = row[sxrCol]; if (sxr == null) continue;
        const sxrStr = String(sxr).trim();
        if (!/^\d+\s*x\s*\d/i.test(sxrStr)) continue;
        const wt = row[weightCol];
        if (typeof wt !== 'number' || wt <= 0) continue;
        days[di].exercises.push({
          name: (perDayNames[di] || exName).replace(/\s+/g, ' '),
          prescription: `${sxrStr}(${Math.round(wt * 100) / 100})`,
          note: '', lifterNote: '', loggedWeight: ''
        });
      }
    }
    const result = days.filter(d => d.exercises.length > 0);
    if (result.length >= 2) return result;
  }

  // Vertical layout (MagOrt / Smolov style): SxR text + weight
  const exercises = [];
  let exerciseName = sheetName || 'Exercise';
  // Try to find an explicit exercise name in header area
  for (let i = 0; i < Math.min(startRow + 3, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c]; if (val == null) continue;
      const str = String(val).trim();
      if (/^(Back Squat|Front Squat|Squat|Bench|Deadlift|Press)/i.test(str)) { exerciseName = str; break; }
    }
  }

  for (let r = startRow; r < endRow; r++) {
    const row = rows[r]; if (!row) continue;
    for (let c = 0; c < Math.min(row.length, 12); c++) {
      const val = row[c]; if (val == null) continue;
      const str = String(val).trim();
      if (/^Week\s*\d/i.test(str)) continue;
      if (/^(Warm\s*Up|REST|TEST|MAX|LB|KG)/i.test(str)) continue;
      // If it looks like SxR text
      if (/^\d+\s*x\s*\d/i.test(str) || /^\d+x\d/i.test(str)) {
        let weight = null;
        for (let nc = c + 1; nc < Math.min(c + 3, row.length); nc++) {
          if (typeof row[nc] === 'number' && row[nc] > 0) { weight = row[nc]; break; }
        }
        let pres = str;
        if (weight != null) pres += `(${Math.round(weight * 100) / 100})`;
        exercises.push({ name: exerciseName, prescription: pres, note: '', lifterNote: '', loggedWeight: '' });
        break;
      }
      // Check for exercise name
      if (/^[A-Z][a-z]/.test(str) && str.length >= 4 && !/^(Week|Day|Warm|REST|TEST|MAX|Pound|Kilo|Date)/i.test(str)) {
        exerciseName = str;
      }
    }
  }

  if (exercises.length > 0) return [{ name: 'Day 1', exercises }];
  return [];
}

// ── E Sub-parser: Cycle-Week Grid (531 uZan/wdyK — Cycle headers + Set rows) ─
function parseE_cycleWeek(wb) {
  // Find the main training sheet
  let trainingSheet = null, trainingRows = null;
  for (const sn of wb.SheetNames) {
    if (_eIsSkipSheet(sn)) continue;
    if (/^(inputs?|maxes?)$/i.test(sn.trim())) continue;
    const rows = _eCleanRows(wb.Sheets[sn]);
    if (!rows || rows.length < 10) continue;
    // Must have Cycle or "Week N NxN" headers + exercise names
    let hasCycle = false, hasExNames = false, hasWeekPres = false;
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < Math.min(row.length, 25); c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        if (/^Cycle\s*\d/i.test(str)) hasCycle = true;
        if (/^Week\s*\d.*\d+x\d/i.test(str) || /^Week\s*\d.*\d+\/\d+\/\d/i.test(str)) hasWeekPres = true;
        if (/^(Bench|Squat|Deadlift|Press|Military|Overhead|OHP)/i.test(str)) hasExNames = true;
      }
    }
    if ((hasCycle || hasWeekPres) && hasExNames) { trainingSheet = sn; trainingRows = rows; break; }
  }
  if (!trainingRows) return null;

  // Extract maxes from any inputs/maxes sheet
  let maxes = {};
  for (const sn of wb.SheetNames) {
    if (/^(inputs?|maxes?|start)/i.test(sn.trim())) {
      const rows = _eCleanRows(wb.Sheets[sn]);
      if (rows) Object.assign(maxes, _eExtractMaxes(rows, 30));
    }
  }
  Object.assign(maxes, _eExtractMaxes(trainingRows, 20));

  // Parse: find week column groups, then exercise row groups
  // Pattern: "Week N [NxN]" as column headers, exercises as row groups with Set 1/2/3 sub-rows
  const allWeeks = [];
  const rows = trainingRows;

  // Find cycle boundaries
  const cycleBounds = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    for (let c = 0; c < Math.min(row.length, 20); c++) {
      if (row[c] != null && /^Cycle\s*\d/i.test(String(row[c]).trim())) {
        cycleBounds.push(i);
        break;
      }
    }
  }
  if (cycleBounds.length === 0) cycleBounds.push(0);

  for (let ci = 0; ci < cycleBounds.length; ci++) {
    const cycleStart = cycleBounds[ci];
    const cycleEnd = ci + 1 < cycleBounds.length ? cycleBounds[ci + 1] : rows.length;

    // Find week column positions from "Week N" headers
    const weekCols = [];
    for (let i = cycleStart; i < Math.min(cycleStart + 10, cycleEnd); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        if (row[c] != null && /^Week\s*\d/i.test(String(row[c]).trim())) {
          const wm = String(row[c]).trim().match(/^Week\s*(\d+)/i);
          weekCols.push({ col: c, weekLabel: String(row[c]).trim(), weekNum: parseInt(wm[1]) });
        }
      }
      if (weekCols.length > 0) break;
    }
    if (weekCols.length === 0) continue;

    // Find exercise name rows and Set sub-rows
    const exerciseGroups = []; // [{name, startRow, setRows:[]}]
    let currentEx = null;
    for (let i = cycleStart; i < cycleEnd; i++) {
      const row = rows[i]; if (!row) continue;
      const colA = row[0] != null ? String(row[0]).trim() : '';
      if (/^Cycle\s*\d/i.test(colA)) continue;
      if (/^Week\s*\d/i.test(colA)) continue;
      if (/^Core\s*Lift/i.test(colA)) continue;
      if (!colA) continue;
      // Exercise name: text that's not Set N, not a number
      if (/^Set\s*\d/i.test(colA)) {
        if (currentEx) currentEx.setRows.push(i);
      } else if (/^(PR|New|Training|Boring|BBB|Cycle|1RM|TM)/i.test(colA)) {
        // skip metadata rows
      } else if (/[a-zA-Z]{3,}/.test(colA) && !/^\d+\.?\d*$/.test(colA)) {
        currentEx = { name: colA.replace(/\s+/g, ' '), startRow: i, setRows: [] };
        exerciseGroups.push(currentEx);
        // Check if this row itself has weight data (no separate Set rows)
        let hasWeightData = false;
        for (const wc of weekCols) {
          if (typeof row[wc.col] === 'number') hasWeightData = true;
        }
        if (hasWeightData) currentEx.setRows.push(i);
      }
    }

    // For each exercise that has no explicit Set rows, check rows below for weight data
    for (const eg of exerciseGroups) {
      if (eg.setRows.length === 0) {
        for (let r = eg.startRow + 1; r < Math.min(eg.startRow + 5, cycleEnd); r++) {
          const row = rows[r]; if (!row) continue;
          const colA = String(row[0] || '').trim();
          if (/^(Set\s*\d)/i.test(colA) || !colA) {
            let hasData = false;
            for (const wc of weekCols) {
              if (typeof row[wc.col] === 'number' || (row[wc.col] != null && /^\d/.test(String(row[wc.col]).trim()))) hasData = true;
            }
            if (hasData) eg.setRows.push(r);
          } else break;
        }
      }
    }

    // Build weeks: each week column becomes a week with one "day" per exercise
    // Actually, group exercises into days (many 531 programs have 4 exercises = 4 days)
    for (const wc of weekCols) {
      const weekExercises = [];
      for (const eg of exerciseGroups) {
        const sets = [];
        for (const sr of eg.setRows) {
          const row = rows[sr];
          if (!row) continue;
          const weight = row[wc.col];
          if (weight == null) continue;
          const w = typeof weight === 'number' ? weight : parseFloat(String(weight));
          if (isNaN(w)) continue;
          // Look for reps in adjacent column
          let reps = '';
          if (wc.col + 1 < row.length && row[wc.col + 1] != null) {
            reps = String(row[wc.col + 1]).trim().replace(/^[xX]\s*/, '').replace(/\.0$/, '');
          }
          sets.push({ weight: Math.round(w * 100) / 100, reps: reps || '?' });
        }
        if (sets.length > 0) {
          const groups = [];
          let si = 0;
          while (si < sets.length) {
            let count = 1;
            while (si + count < sets.length && sets[si].weight === sets[si + count].weight && sets[si].reps === sets[si + count].reps) count++;
            groups.push(sets[si].weight > 0 ? `${count}x${sets[si].reps}(${sets[si].weight})` : `${count}x${sets[si].reps}`);
            si += count;
          }
          weekExercises.push({ name: eg.name, prescription: groups.join(', '), note: '', lifterNote: '', loggedWeight: '' });
        }
      }
      if (weekExercises.length > 0) {
        // One day per exercise (each exercise is typically a separate training day in 531)
        const days = weekExercises.map(ex => ({ name: ex.name + ' Day', exercises: [ex] }));
        allWeeks.push({ label: wc.weekLabel, days });
      }
    }
  }

  if (allWeeks.length === 0) return null;
  const blockName = (wb._plFilename || trainingSheet).replace(/\.xlsx?$/i, '');
  return [{
    id: 'e_cw_' + Date.now(), name: blockName, format: 'E',
    athleteName: '', dateRange: '', maxes, weeks: allWeeks
  }];
}

// ── E Sub-parser: Parallel 531 (4 parallel column groups, Week N row separators) ─
function parseE_parallel531(wb) {
  let trainingSheet = null, trainingRows = null;
  for (const sn of wb.SheetNames) {
    if (_eIsSkipSheet(sn)) continue;
    if (/^(inputs?|maxes?|start|how)/i.test(sn.trim())) continue;
    const rows = _eCleanRows(wb.Sheets[sn]);
    if (!rows || rows.length < 10) continue;
    // Check for Week N in first few rows + multiple Reps headers
    for (let i = 0; i < Math.min(3, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      if (/^week\s*\d/i.test(strs[0] || '')) {
        const repsCols = strs.filter(s => s === 'reps').length;
        if (repsCols >= 2) { trainingSheet = sn; trainingRows = rows; break; }
      }
    }
    if (trainingRows) break;
  }
  if (!trainingRows) return null;

  // Extract maxes from a Maxes sheet
  let maxes = {};
  for (const sn of wb.SheetNames) {
    if (/^(inputs?|maxes?)/i.test(sn.trim())) {
      const rows = _eCleanRows(wb.Sheets[sn]);
      if (rows) Object.assign(maxes, _eExtractMaxes(rows, 40));
    }
  }

  const rows = trainingRows;

  // Identify column groups: each group starts with an exercise name in row 2
  // and has weight col, scheme col, reps-done col
  // Week N labels in col A separate week sections

  // Find week boundaries
  const weekBounds = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    const val = row[0]; if (val == null) continue;
    if (/^Week\s*\d/i.test(String(val).trim())) {
      const wm = String(val).trim().match(/Week\s*(\d+)/i);
      weekBounds.push({ row: i, weekNum: parseInt(wm[1]) });
    }
  }
  if (weekBounds.length === 0) return null;

  // Find column groups from the first week section
  // Look at row 1 (or week row + 1) for exercise names
  const firstWeekRow = weekBounds[0].row;
  const exRow = rows[firstWeekRow + 1]; // Exercise names should be in the row after "Week N"
  if (!exRow) return null;

  const colGroups = []; // [{exCol, weightCol, schemeCol, repsCol, name}]
  for (let c = 0; c < exRow.length; c++) {
    const val = exRow[c]; if (val == null) continue;
    const str = String(val).trim();
    if (str.length >= 3 && /[a-zA-Z]/.test(str) && !/^(Reps|Weight|Set|Rep)/i.test(str)) {
      // This looks like an exercise name - find the weight/scheme/reps columns
      // Typically: exName at C, weight at C+1, scheme at C+2, reps at C+3
      colGroups.push({ exCol: c, name: str });
    }
  }

  // For each column group, identify the sub-columns
  // Look for "Reps" header to identify reps column
  const headerRow = rows[firstWeekRow];
  if (headerRow) {
    for (const cg of colGroups) {
      cg.weightCol = cg.exCol + 1;
      cg.schemeCol = cg.exCol + 2;
      cg.repsCol = cg.exCol + 3;
    }
  }

  if (colGroups.length === 0) return null;

  // Parse each week section
  const allWeeks = [];
  for (let wi = 0; wi < weekBounds.length; wi++) {
    const wStart = weekBounds[wi].row;
    const wEnd = wi + 1 < weekBounds.length ? weekBounds[wi + 1].row : rows.length;
    const weekLabel = `Week ${weekBounds[wi].weekNum}`;

    const days = [];
    for (const cg of colGroups) {
      // Read current exercise name (may change per week if accessories rotate)
      let currentExName = cg.name;
      const nameRow = rows[wStart + 1];
      if (nameRow && nameRow[cg.exCol]) {
        const n = String(nameRow[cg.exCol]).trim();
        if (n.length >= 3 && /[a-zA-Z]/.test(n)) currentExName = n;
      }

      const exercises = [];
      let mainSets = [];
      let accessoryExName = null;

      for (let r = wStart + 2; r < wEnd; r++) {
        const row = rows[r]; if (!row) continue;
        // Check if this row has an exercise name in the group's column
        const cellName = row[cg.exCol] != null ? String(row[cg.exCol]).trim() : '';
        const cellWeight = row[cg.weightCol];
        const cellScheme = row[cg.schemeCol] != null ? String(row[cg.schemeCol]).trim() : '';

        if (cellName && /[a-zA-Z]{3,}/.test(cellName) && cellName !== currentExName) {
          // Flush main exercise
          if (mainSets.length > 0) {
            const groups = [];
            let si = 0;
            while (si < mainSets.length) {
              let count = 1;
              while (si + count < mainSets.length && mainSets[si].weight === mainSets[si + count].weight && mainSets[si].reps === mainSets[si + count].reps) count++;
              groups.push(mainSets[si].weight > 0 ? `${count}x${mainSets[si].reps}(${mainSets[si].weight})` : `${count}x${mainSets[si].reps}`);
              si += count;
            }
            exercises.push({ name: currentExName, prescription: groups.join(', '), note: '', lifterNote: '', loggedWeight: '' });
            mainSets = [];
          }
          accessoryExName = cellName;
        }

        if (typeof cellWeight === 'number' && cellWeight > 0) {
          let reps = cellScheme.replace(/^[xX]\s*/, '').replace(/\.0$/, '') || '?';
          if (accessoryExName) {
            // Accessory exercise
            exercises.push({ name: accessoryExName, prescription: `${reps}(${cellWeight})`, note: '', lifterNote: '', loggedWeight: '' });
          } else {
            mainSets.push({ weight: Math.round(cellWeight * 100) / 100, reps });
          }
        } else if (cellScheme && accessoryExName) {
          exercises.push({ name: accessoryExName, prescription: cellScheme, note: '', lifterNote: '', loggedWeight: '' });
        }
      }

      // Flush remaining main sets
      if (mainSets.length > 0) {
        const groups = [];
        let si = 0;
        while (si < mainSets.length) {
          let count = 1;
          while (si + count < mainSets.length && mainSets[si].weight === mainSets[si + count].weight && mainSets[si].reps === mainSets[si + count].reps) count++;
          groups.push(mainSets[si].weight > 0 ? `${count}x${mainSets[si].reps}(${mainSets[si].weight})` : `${count}x${mainSets[si].reps}`);
          si += count;
        }
        exercises.push({ name: currentExName, prescription: groups.join(', '), note: '', lifterNote: '', loggedWeight: '' });
      }

      if (exercises.length > 0) days.push({ name: currentExName + ' Day', exercises });
    }

    if (days.length > 0) allWeeks.push({ label: weekLabel, days });
  }

  if (allWeeks.length === 0) return null;
  const blockName = (wb._plFilename || trainingSheet).replace(/\.xlsx?$/i, '');
  return [{
    id: 'e_p531_' + Date.now(), name: blockName, format: 'E',
    athleteName: '', dateRange: '', maxes, weeks: allWeeks
  }];
}

// ── parseE_texasMethod: Texas Method columnar Mon/Wed/Fri cycle pattern ──────
function parseE_texasMethod(wb) {
  // Find the Texas Method sheet
  let tmSheet = null;
  for (const sn of wb.SheetNames) {
    if (/texas\s*method/i.test(sn.trim())) { tmSheet = wb.Sheets[sn]; break; }
  }
  if (!tmSheet) return null;

  const rows = XLSX.utils.sheet_to_json(tmSheet, { header: 1, defval: null })
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

  // Verify Mon/Wed/Fri pattern exists
  let hasTMHeader = false;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    if (!row.some(c => c && /^texas\s*method$/i.test(String(c).trim()))) continue;
    if (row.filter(c => c && /^(mon|wed|fri)/i.test(String(c).trim())).length >= 3) { hasTMHeader = true; break; }
  }
  if (!hasTMHeader) return null;

  // Extract maxes from Questions/Input sheets
  let maxes = {};
  for (const sn of wb.SheetNames) {
    if (/^(question|input|maxe?s?|start)/i.test(sn.trim())) {
      const mRows = _eCleanRows(wb.Sheets[sn]);
      if (mRows) Object.assign(maxes, _eExtractMaxes(mRows, 30));
    }
  }

  // Find all "Texas Method" header rows — each starts a cycle block
  const blockStarts = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    if (row.some(c => c && /^texas\s*method$/i.test(String(c).trim()))) blockStarts.push(i);
  }
  if (!blockStarts.length) return null;

  const DAY_NAMES = ['Monday', 'Wednesday', 'Friday'];
  const allWeeks = [];

  for (let bi = 0; bi < blockStarts.length; bi++) {
    const bStart = blockStarts[bi];
    const headerRow = rows[bStart]; if (!headerRow) continue;

    // Find first day column
    let firstDayCol = -1;
    for (let c = 0; c < headerRow.length; c++) {
      if (headerRow[c] && /^(mon|wed|fri)/i.test(String(headerRow[c]).trim())) { firstDayCol = c; break; }
    }
    if (firstDayCol < 0) continue;

    // Count day columns (groups of 3 = weeks)
    let dayColCount = 0;
    for (let c = firstDayCol; c < headerRow.length; c++) {
      if (headerRow[c] && /^(mon|wed|fri)/i.test(String(headerRow[c]).trim())) dayColCount++;
      else if (headerRow[c] == null || String(headerRow[c]).trim() === '') continue;
      else break;
    }
    const numWeeks = Math.floor(dayColCount / 3) || 1;

    // Read SxR scheme row (bStart + 2)
    const sxrRow = rows[bStart + 2] || [];

    // Initialize weeks with empty days
    const weekDays = [];
    for (let w = 0; w < numWeeks; w++) {
      const days = [];
      for (let d = 0; d < 3; d++) {
        days.push({
          name: DAY_NAMES[d],
          colIdx: firstDayCol + w * 3 + d,
          sxr: String(sxrRow[firstDayCol + w * 3 + d] || '').trim(),
          exercises: []
        });
      }
      weekDays.push(days);
    }

    // Parse exercise blocks
    const bEnd = (bi + 1 < blockStarts.length) ? blockStarts[bi + 1] : rows.length;
    let curExercise = null;

    for (let i = bStart + 3; i < bEnd; i++) {
      const row = rows[i]; if (!row) continue;
      const col0 = row[0] ? String(row[0]).trim() : '';
      const col1 = row[1] ? String(row[1]).trim() : '';

      if (/^(volume|cycle)/i.test(col0)) break;

      // New exercise name in col 0
      if (col0 && /[a-zA-Z]{2,}/.test(col0) && !/^(VOLUME|Cycle|Reps|Sets)/i.test(col0)) {
        curExercise = col0;
      }
      if (!curExercise) continue;

      // Only capture "Work Sets" rows
      if (col1.toLowerCase() !== 'work sets') continue;

      for (let w = 0; w < numWeeks; w++) {
        for (let d = 0; d < 3; d++) {
          const colIdx = firstDayCol + w * 3 + d;
          const val = row[colIdx];
          if (val == null || val === '') continue;

          const valStr = String(val).trim();
          const daySxr = weekDays[w][d].sxr;
          let prescription;

          if (/^\d+x[\dF]/i.test(valStr)) {
            prescription = valStr;
          } else if (typeof val === 'number' || /^[\d.]+$/.test(valStr)) {
            const wt = typeof val === 'number' ? Math.round(val * 10) / 10 : valStr;
            prescription = /^\d+x\d/.test(daySxr) ? `${daySxr}(${wt})` : String(wt);
          } else {
            prescription = valStr;
          }

          weekDays[w][d].exercises.push({
            name: _hCapitalizeName(curExercise),
            prescription,
            note: '', lifterNote: '', loggedWeight: ''
          });
        }
      }
      curExercise = null;
    }

    for (let w = 0; w < numWeeks; w++) {
      const filteredDays = weekDays[w].filter(d => d.exercises.length > 0).map(d => ({
        name: d.name, exercises: d.exercises
      }));
      if (filteredDays.length > 0) {
        allWeeks.push({ label: `Week ${allWeeks.length + 1}`, days: filteredDays });
      }
    }
  }

  if (!allWeeks.length) return null;
  return [{
    id: 'e_tm_' + Date.now(), name: 'Texas Method', format: 'E',
    athleteName: '', dateRange: '', maxes, weeks: allWeeks
  }];
}

// ── FORMAT C PARSER (split Sets/Reps/Load columns with Day N headers) ────────
function parseCAutoFormat(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    try {
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, defval:null});
      fixDateSerials(wb.Sheets[sn], rawRows);
      const rows = rawRows.map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
      const week = parseCAutoSheet(sn, rows);
      if (week) blocks.push(week);
    } catch(e) { console.error(`[parseCAutoFormat] Error parsing sheet "${sn}":`, e); }
  }
  if (blocks.length === 0) return [];

  const firstRawRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, defval:null});
  fixDateSerials(wb.Sheets[wb.SheetNames[0]], firstRawRows);
  const firstRows = firstRawRows.map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
  const meta = extractCAutoMeta(firstRows);

  const blockName = meta.blockName || wb._plFilename || 'Program';
  const id = 'c_' + blockName.replace(/\s+/g,'_').substring(0,30) + '_' + Date.now();
  return [{
    id, name: blockName, format: 'C', weeks: blocks, athlete: meta.athlete || '',
    maxes: meta.maxes || {}, dateRange: meta.dateRange || '',
    startDate: meta.startDate || null
  }];
}

function extractCAutoMeta(rows) {
  const meta = { maxes: {}, athlete: '', blockName: '', dateRange: '', startDate: null };
  for (let i = 0; i < Math.min(20, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    const a = String(row[0] || '').trim();
    const maxMatch = a.match(/^(squat|bench\s*press?|sumo\s*deadlift|conventional\s*deadlift|deadlift)\s*[:=]\s*(\d+)/i);
    if (maxMatch) {
      const lift = maxMatch[1].toLowerCase();
      const val = parseInt(maxMatch[2]);
      if (lift.includes('squat')) meta.maxes.squat = val;
      else if (lift.includes('bench')) meta.maxes.bench = val;
      else if (lift.includes('deadlift')) meta.maxes.deadlift = val;
    }
    if (i === 0 && a) {
      const dateMatch = a.match(/((?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{1,2}(?:st|nd|rd|th)?,?\s*\d{4})/i);
      if (dateMatch) meta.dateRange = dateMatch[1];
      for (let c = 1; c < (row.length || 0); c++) {
        const v = String(row[c] || '').trim();
        const blockMatch = v.match(/(?:block|phase|cycle)\s*[:\-]?\s*(.+)/i);
        if (blockMatch) meta.blockName = blockMatch[0].substring(0, 50);
      }
    }
  }
  return meta;
}

function parseCAutoSheet(sn, rows) {
  if (!rows || rows.length === 0) return null;

  let setsIdx = -1, repsIdx = -1, loadIdx = -1, exIdx = -1;
  let coachIdx = -1, loggedIdx = -1, rpeIdx = -1, notesIdx = -1;
  let dataStartRow = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    const strs = row.map(_normalizeCell);
    if (strs.includes('sets') && strs.includes('reps')) {
      setsIdx = strs.indexOf('sets');
      repsIdx = strs.indexOf('reps');
      loadIdx = strs.indexOf('load'); if (loadIdx < 0) loadIdx = strs.indexOf('weight');
      coachIdx = strs.indexOf('coach notes'); if (coachIdx < 0) strs.forEach((s, ci) => { if (s.includes('coach') && coachIdx < 0) coachIdx = ci; });
      loggedIdx = strs.findIndex(s => s.includes('weight used'));
      rpeIdx = strs.findIndex(s => s.includes('overall rpe') || s === 'rpe');
      notesIdx = strs.findIndex(s => s.includes('athlete notes') || s.includes('lifter notes'));
      exIdx = strs.findIndex(s => s === 'warm up' || s === 'main exercise');
      if (exIdx < 0) exIdx = setsIdx > 0 ? setsIdx - 1 : 1;
      dataStartRow = (rows[i][0] && /^day\s*\d/i.test(String(rows[i][0]).trim())) ? i :
                     (i > 0 && rows[i-1] && rows[i-1][0] && /^day\s*\d/i.test(String(rows[i-1][0]).trim())) ? i - 1 : i;
      break;
    }
  }

  if (setsIdx < 0 || repsIdx < 0) return null;

  const days = [];
  let curDay = null;
  let curSupersetGroup = null;
  const SUB_SECTIONS = ['warm up','main exercise','accessories','finisher','abs','recovery','cool down','cardio'];
  const _isSupersetLabel = (s) => /^(superset|super\s*set|ss|giant\s*set|circuit)$/i.test(s) || /^\d+\s*rounds?$/i.test(s);

  for (let i = dataStartRow; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    const colA = String(row[0] || '').trim();
    const exVal = String(row[exIdx] || '').trim();
    const strs = row.map(_normalizeCell);

    if (/^day\s*\d/i.test(colA)) {
      let dayName = colA;
      if (i + 1 < rows.length && rows[i+1]) {
        const nextA = String(rows[i+1][0] || '').trim();
        if (/^(mon|tue|wed|thu|fri|sat|sun)/i.test(nextA)) {
          dayName = nextA + ' (' + colA + ')';
        }
      }
      curDay = { name: dayName, exercises: [] };
      curSupersetGroup = null;
      days.push(curDay);
      continue;
    }

    if (/^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$/i.test(colA)) continue;
    if (strs.includes('sets') && strs.includes('reps')) continue;
    if (!curDay) continue;

    // Detect superset/circuit grouping labels in column A or exercise column
    if (_isSupersetLabel(colA) || _isSupersetLabel(exVal)) {
      const matchText = _isSupersetLabel(colA) ? colA : exVal;
      const roundsMatch = matchText.match(/^(\d+)\s*rounds?$/i);
      const rounds = roundsMatch ? parseInt(roundsMatch[1]) : 1;
      const label = rounds > 1 ? rounds + ' rounds' : 'superset';
      curSupersetGroup = { label: label + '_' + curDay.exercises.length, rounds, startIdx: curDay.exercises.length };
      continue;
    }

    // Non-superset sub-section headers clear superset grouping
    if (SUB_SECTIONS.some(s => colA.toLowerCase() === s || exVal.toLowerCase() === s)) {
      curSupersetGroup = null;
      continue;
    }

    if (!exVal) { curSupersetGroup = null; continue; }

    const setsRaw = row[setsIdx] != null ? String(row[setsIdx]).trim() : '';
    const repsRaw = row[repsIdx] != null ? String(row[repsIdx]).trim() : '';
    const loadVal = loadIdx >= 0 && row[loadIdx] != null ? String(row[loadIdx]).trim() : '';

    const isSimpleNum = (v) => /^\d+\.?\d*$/.test(v);
    const isFullPrescription = (v) => !isSimpleNum(v) && v.length > 0 && (
      /\d+x\d+/.test(v) || /,/.test(v) || /[a-zA-Z]{3,}/.test(v) || /bar/i.test(v)
    );

    let pres = '';
    const setsVal = setsRaw.replace(/\.0$/, '');
    const repsVal = repsRaw.replace(/\.0$/, '');

    if (isFullPrescription(setsRaw)) {
      pres = setsRaw;
      if (repsRaw && !isSimpleNum(repsRaw)) pres += ' ' + repsRaw;
      if (loadVal && !isSimpleNum(loadVal) && !/rpe/i.test(loadVal)) pres += ' ' + loadVal;
    } else if (setsVal && repsVal) {
      pres = setsVal + 'x' + repsVal;
    } else if (setsVal) {
      pres = setsVal;
    } else if (repsVal) {
      pres = repsVal;
    }

    if (loadVal && !isFullPrescription(setsRaw)) {
      const cleanLoad = loadVal.replace(/\.0$/, '');
      if (/^\d+/.test(cleanLoad)) {
        pres += '(' + cleanLoad + ')';
      } else if (/rpe/i.test(cleanLoad)) {
        pres += ' ' + cleanLoad;
      } else if (cleanLoad) {
        pres += ' ' + cleanLoad;
      }
    }

    const logged = loggedIdx >= 0 && row[loggedIdx] != null ? String(row[loggedIdx]).trim().replace(/\.0$/, '') || null : null;
    const note = coachIdx >= 0 && row[coachIdx] != null ? String(row[coachIdx]).trim() || null : null;
    const rpeLogged = rpeIdx >= 0 && row[rpeIdx] != null ? String(row[rpeIdx]).trim().replace(/\.0$/, '') || null : null;
    const lifterNote = notesIdx >= 0 && row[notesIdx] != null ? String(row[notesIdx]).trim() || null : null;

    let logStr = logged || null;
    if (rpeLogged && logStr) logStr += ' @' + rpeLogged;
    else if (rpeLogged && !logStr) logStr = '@' + rpeLogged;

    curDay.exercises.push({
      name: exVal.replace(/\s+/g, ' '),
      prescription: pres.trim() || null,
      note,
      lifterNote,
      loggedWeight: logStr,
      supersetGroup: curSupersetGroup || null
    });
  }

  const filtered = days.filter(d => d.exercises.length > 0);
  if (filtered.length === 0) return null;
  return { label: sn, days: filtered };
}

// ── FORMAT F DETECTION ──────────────────────────────────────────────────────
function detectF(wb) {
  const SKIP_SHEET = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|output|pr\s*sheet|notes?|start\s*here|1\.\s*start|predicted)$/i;

  for (let si = 0; si < Math.min(8, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP_SHEET.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows || rows.length < 5) continue;

    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;

      // Count "Week" labels spread horizontally across columns
      const weekCols = [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        if (/^Week\s*\d/i.test(str)) weekCols.push(c);
      }

      if (weekCols.length >= 2) {
        // Verify sub-headers in next 1-2 rows (Weight/Reps, %, or Exercise repeating)
        // Exclude D-style files with Sets/Reps/Intensity/Load column groups
        for (let j = i + 1; j < Math.min(i + 3, rows.length); j++) {
          const sr = rows[j]; if (!sr) continue;
          const strs = sr.map(_normalizeCell);
          // D-style: Sets + Reps + (Intensity or Load) repeating — skip
          const setsRepeat = strs.filter(s => s === 'sets').length;
          const hasIntensity = strs.some(s => s === 'intensity');
          if (setsRepeat >= 2 && hasIntensity) continue;
          // F sub-header patterns — but exclude Sheiko (numbered exercises + % in data rows)
          const subMatch = (strs.filter(s => s === 'weight').length >= 2 && strs.filter(s => s === 'reps').length >= 2)
            || (strs.filter(s => s === '%' || s === 'percentage').length >= 2)
            || (strs.filter(s => s === 'exercise').length >= 2)
            || (strs.filter(s => s === 'percentage').length >= 2 && strs.filter(s => s === 'load').length >= 2);
          if (!subMatch) continue;
          // Sheiko exclusion: data rows below have numbered indices in col 0 + exercise names in col 1 + % in col 2
          let sheikoPat = false;
          for (let k = j + 1; k < Math.min(j + 5, rows.length); k++) {
            const dr = rows[k]; if (!dr) continue;
            const a = dr[0], b = dr[1], cc = dr[2];
            if (typeof a === 'number' && a >= 1 && a <= 20
              && b != null && /[a-zA-Z]{2,}/.test(String(b))
              && typeof cc === 'number' && cc > 0 && cc < 1) {
              sheikoPat = true; break;
            }
          }
          if (sheikoPat) continue;
          return true;
        }
      }

      // F3 Madcow: Day/Exercise header + W1/W2 or sequential numbered columns
      const strs = row.map(_normalizeCell);
      const hasDay = strs.some(s => s === 'day');
      const hasExercise = strs.some(s => s === 'exercise');
      if (hasDay && hasExercise) {
        let wLabels = 0;
        for (const s of strs) { if (/^w\d+/i.test(s)) wLabels++; }
        if (wLabels >= 3) return true;

        let seqCount = 0;
        for (let c = 0; c < row.length; c++) {
          const val = row[c];
          if (typeof val === 'number' && val === Math.floor(val) && val >= 1 && val <= 52) seqCount++;
        }
        if (seqCount >= 4) return true;
      }
    }
  }
  return false;
}

// ── FORMAT F PARSER (horizontal week columns) ────────────────────────────────
function _fCleanRows(ws) {
  return XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
}

function _fIsSkipSheet(name) {
  return /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|output|pr\s*sheet|notes?|start\s*here|1\.\s*start|predicted)$/i.test(name.trim());
}

function parseF(wb) {
  let result;
  result = parseF_stride3(wb);
  if (result && result.length > 0 && result[0].weeks && result[0].weeks.length > 0) return result;
  result = parseF_stride15(wb);
  if (result && result.length > 0 && result[0].weeks && result[0].weeks.length > 0) return result;
  result = parseF_stride6(wb);
  if (result && result.length > 0 && result[0].weeks && result[0].weeks.length > 0) return result;
  result = parseF_pctLoad(wb);
  if (result && result.length > 0 && result[0].weeks && result[0].weeks.length > 0) return result;
  result = parseF_stride1(wb);
  if (result && result.length > 0 && result[0].weeks && result[0].weeks.length > 0) return result;
  result = parseF_coach(wb);
  if (result && result.length > 0 && result[0].weeks && result[0].weeks.length > 0) return result;
  return parseB(wb);
}

// F1: stride-3 — Week headers + Weight/Reps/Notes sub-headers (531 ViolentZen)
function parseF_stride3(wb) {
  const allWeeks = [];
  let blockName = '';
  let globalWeekNum = 0;

  for (const sn of wb.SheetNames) {
    if (_fIsSkipSheet(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = _fCleanRows(ws);
    if (!rows || rows.length < 5) continue;

    // Find week header row: ≥2 "Week N" labels
    let weekHeaderRow = -1;
    let weekPositions = [];
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const wpos = [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        if (/^Week\s*\d/i.test(String(val).trim())) wpos.push({col: c, label: String(val).trim()});
      }
      if (wpos.length >= 2) { weekHeaderRow = i; weekPositions = wpos; break; }
    }
    if (weekHeaderRow === -1) continue;

    // Find sub-header row with Weight/Reps repeating (or Exercise/Sets/Reps/Weight repeating)
    let subHeaderRow = -1;
    let weekColGroups = [];
    for (let i = weekHeaderRow + 1; i < Math.min(weekHeaderRow + 3, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      const wCols = [], rCols = [], sCols = [], eCols = [], rpeCols = [], nCols = [];
      for (let c = 0; c < strs.length; c++) {
        if (strs[c] === 'weight') wCols.push(c);
        if (strs[c] === 'reps') rCols.push(c);
        if (strs[c] === 'sets') sCols.push(c);
        if (strs[c] === 'exercise') eCols.push(c);
        if (strs[c] === 'rpe') rpeCols.push(c);
        if (strs[c] === 'notes' || strs[c] === 'note' || strs[c] === 'coach notes') nCols.push(c);
      }
      if (wCols.length >= 2 && rCols.length >= 2) {
        subHeaderRow = i;
        for (let g = 0; g < wCols.length; g++) {
          weekColGroups.push({
            weightCol: wCols[g],
            repsCol: rCols[g] !== undefined ? rCols[g] : wCols[g] + 1,
            setsCol: sCols[g] !== undefined ? sCols[g] : -1,
            exCol: eCols[g] !== undefined ? eCols[g] : -1,
            rpeCol: rpeCols[g] !== undefined ? rpeCols[g] : -1,
            noteCol: nCols[g] !== undefined ? nCols[g] : -1
          });
        }
        break;
      }
    }
    if (subHeaderRow === -1 || weekColGroups.length === 0) continue;

    if (!blockName) blockName = sn;

    // Find day sections (match day-of-week names OR "Day N" labels)
    const dayStarts = [];
    for (let i = subHeaderRow + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const colA = String(row[0] || '').trim().toLowerCase();
      if (DAYS.some(d => colA === d || colA.startsWith(d + ' ')) || /^day\s*\d/i.test(colA)) {
        dayStarts.push({row: i, name: String(row[0]).trim()});
      }
    }
    if (dayStarts.length === 0) continue;

    // For each day, group rows by exercise name
    const dayData = [];
    for (let d = 0; d < dayStarts.length; d++) {
      const startRow = dayStarts[d].row;
      const endRow = d + 1 < dayStarts.length ? dayStarts[d + 1].row : rows.length;
      const exerciseGroups = [];

      for (let r = startRow; r < endRow; r++) {
        const row = rows[r]; if (!row) continue;
        let exName = '';
        for (let c = 0; c <= 1; c++) {
          const v = row[c]; if (v == null) continue;
          const s = String(v).trim();
          if (DAYS.some(dd => s.toLowerCase() === dd || s.toLowerCase().startsWith(dd + ' '))) continue;
          if (/^day\s*\d/i.test(s)) continue;
          if (/[a-zA-Z]{2,}/.test(s) && !/^(Week|Exercise|Sets?|Reps?|Weight|Notes?|Load|Percentage)$/i.test(s)) { exName = s; break; }
        }
        let hasData = false;
        for (const g of weekColGroups) {
          if (row[g.weightCol] != null || row[g.repsCol] != null) { hasData = true; break; }
        }
        if (!hasData) continue;

        if (exName) {
          exerciseGroups.push({name: exName, rows: [r]});
        } else if (exerciseGroups.length > 0) {
          exerciseGroups[exerciseGroups.length - 1].rows.push(r);
        }
      }

      const dayExercises = [];
      for (const eg of exerciseGroups) {
        const weekSets = [];
        for (let w = 0; w < weekColGroups.length; w++) {
          const g = weekColGroups[w];
          const sets = [];
          const noteItems = [];
          for (const ri of eg.rows) {
            const row = rows[ri];
            const weight = row[g.weightCol];
            const reps = row[g.repsCol];
            const setsNum = g.setsCol >= 0 ? (parseInt(row[g.setsCol]) || 0) : 0;
            // Also check for per-week exercise name (different exercises per week in same row position)
            const weekExName = g.exCol >= 0 && row[g.exCol] ? String(row[g.exCol]).trim() : '';
            // Read RPE from separate column if available
            const rpeVal = g.rpeCol >= 0 && row[g.rpeCol] != null ? String(row[g.rpeCol]).trim() : '';
            // Read notes column if available
            if (g.noteCol >= 0 && row[g.noteCol] != null) {
              const n = String(row[g.noteCol]).trim();
              if (n && !/^(notes?|coach\s*notes?)$/i.test(n)) noteItems.push(n);
            }
            if (weight != null || reps != null) {
              sets.push({weight: typeof weight === 'number' ? weight : null, reps: reps != null ? String(reps) : '', setsNum, weekExName, rpe: rpeVal});
            }
          }
          weekSets.push({sets, note: noteItems.join('; ')});
        }
        dayExercises.push({name: eg.name, weekSets});
      }
      dayData.push({name: dayStarts[d].name, exercises: dayExercises});
    }

    // Build weeks — handle per-week exercise names (exercises that change across weeks in same row)
    for (let w = 0; w < weekColGroups.length; w++) {
      globalWeekNum++;
      const weekDays = [];
      for (const dd of dayData) {
        const exercises = [];
        for (const ex of dd.exercises) {
          if (w >= ex.weekSets.length) continue;
          const weekSetData = ex.weekSets[w];
          const sets = weekSetData.sets;
          const weekNote = weekSetData.note || '';
          if (sets.length === 0) continue;
          // Group sets by resolved weekExName to handle rows with different exercises per week
          const _resolveName = (s) => (s.weekExName && /[a-zA-Z]{2,}/.test(s.weekExName) && !/^(Exercise|Week)/i.test(s.weekExName))
            ? s.weekExName : ex.name;
          const nameGroups = new Map();
          for (const s of sets) {
            const nm = _resolveName(s);
            if (!nameGroups.has(nm)) nameGroups.set(nm, []);
            nameGroups.get(nm).push(s);
          }
          for (const [useName, groupSets] of nameGroups) {
            const presParts = [];
            for (const s of groupSets) {
              const sc = s.setsNum > 0 ? s.setsNum : 1;
              const rpeSuffix = (s.rpe && /^\d+(\.\d+)?$/.test(s.rpe)) ? ' @RPE' + s.rpe : '';
              if (s.weight != null && s.weight > 0 && s.reps) {
                presParts.push(sc + 'x' + s.reps + '(' + (Math.round(s.weight * 10) / 10) + ')' + rpeSuffix);
              } else if (s.reps) {
                presParts.push(sc + 'x' + s.reps + rpeSuffix);
              }
            }
            if (presParts.length === 0) continue;
            exercises.push({name: useName, prescription: presParts.join(', '), note: weekNote, lifterNote: '', loggedWeight: '', supersetGroup: null});
          }
        }
        if (exercises.length > 0) weekDays.push({name: dd.name, exercises});
      }
      if (weekDays.length > 0) {
        allWeeks.push({label: weekPositions[w] ? weekPositions[w].label : 'Week ' + globalWeekNum, days: weekDays});
      }
    }
  }

  if (allWeeks.length === 0) return null;
  return [{id: 'F_' + (blockName || 'block').replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now(), name: blockName, weeks: allWeeks, format: 'F'}];
}

// F2: stride-15 — nested exercises per week (531 Esl3fko)
function parseF_stride15(wb) {
  for (const sn of wb.SheetNames) {
    if (_fIsSkipSheet(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = _fCleanRows(ws);
    if (!rows || rows.length < 5) continue;

    // Find week header row
    let weekHeaderRow = -1;
    let weekPositions = [];
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const wpos = [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        if (/^Week\s*\d/i.test(String(val).trim())) wpos.push({col: c, label: String(val).trim()});
      }
      if (wpos.length >= 2) { weekHeaderRow = i; weekPositions = wpos; break; }
    }
    if (weekHeaderRow === -1) continue;

    // Check next row for exercise names
    const exRow = rows[weekHeaderRow + 1];
    if (!exRow) continue;
    const exerciseCols = [];
    for (let c = 0; c < exRow.length; c++) {
      const val = exRow[c]; if (val == null) continue;
      const str = String(val).trim();
      if (/[a-zA-Z]{2,}/.test(str) && !/^(%|Weight|Sets|Reps|Notes|Day|Week)/i.test(str)) {
        exerciseCols.push({col: c, name: str});
      }
    }
    if (exerciseCols.length < 2) continue;

    // Check for sub-header row with % repeating
    const subRow = rows[weekHeaderRow + 2];
    if (!subRow) continue;
    const strs = subRow.map(_normalizeCell);
    let pctCount = 0;
    for (const s of strs) { if (s === '%' || s === 'percentage') pctCount++; }
    if (pctCount < 2) continue;

    // Determine exercise stride
    const exStride = exerciseCols.length > 1 ? exerciseCols[1].col - exerciseCols[0].col : 5;
    const weekStride = weekPositions.length > 1 ? weekPositions[1].col - weekPositions[0].col : exerciseCols.length * exStride;
    const exPerWeek = Math.round(weekStride / exStride);

    // Build weeks
    const allWeeks = [];
    for (let w = 0; w < weekPositions.length; w++) {
      const weekStart = weekPositions[w].col;
      const dayExercises = [];

      for (let e = 0; e < exPerWeek; e++) {
        const exCol = weekStart + e * exStride;
        const exNameVal = exRow[exCol];
        if (!exNameVal) continue;
        const exName = String(exNameVal).trim();
        if (!/[a-zA-Z]{2,}/.test(exName)) continue;

        // Parse data rows: offsets 0=%, 1=Weight, 2=Sets, 3=Reps from exCol
        const sets = [];
        for (let r = weekHeaderRow + 3; r < rows.length; r++) {
          const row = rows[r]; if (!row) continue;
          const weight = row[exCol + 1];
          const setsVal = row[exCol + 2];
          const repsVal = row[exCol + 3];
          if (weight == null && setsVal == null && repsVal == null) continue;
          const s = parseInt(setsVal) || 1;
          const rr = repsVal != null ? String(repsVal) : '';
          const wt = typeof weight === 'number' && weight > 0 ? weight : null;
          if (rr || wt) {
            if (wt) sets.push(s + 'x' + rr + '(' + (Math.round(wt * 10) / 10) + ')');
            else if (rr) sets.push(s + 'x' + rr);
          }
        }

        if (sets.length > 0) {
          dayExercises.push({name: exName, prescription: sets.join(', '), note: '', lifterNote: '', loggedWeight: '', supersetGroup: null});
        }
      }

      if (dayExercises.length > 0) {
        allWeeks.push({label: weekPositions[w].label, days: [{name: 'Day 1', exercises: dayExercises}]});
      }
    }

    if (allWeeks.length > 0) return [{id: 'F_' + (sn || 'block').replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now(), name: sn, format: 'F', weeks: allWeeks}];
  }
  return null;
}

// F4: stride-6 — "Week N - Day N" column groups (Rip and Tear)
function parseF_stride6(wb) {
  for (const sn of wb.SheetNames) {
    if (_fIsSkipSheet(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = _fCleanRows(ws);
    if (!rows || rows.length < 5) continue;

    // Find row with "Week N - Day N" labels
    let headerRow = -1;
    let weekDayPositions = [];
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const wdPos = [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        const m = String(val).trim().match(/^Week\s*(\d+)\s*[-–]\s*Day\s*(\d+)/i);
        if (m) wdPos.push({col: c, week: parseInt(m[1]), day: parseInt(m[2])});
      }
      if (wdPos.length >= 2) { headerRow = i; weekDayPositions = wdPos; break; }
    }
    if (headerRow === -1) continue;

    // Determine stride
    const stride = weekDayPositions.length > 1 ? weekDayPositions[1].col - weekDayPositions[0].col : 6;

    // Map sub-header offsets from first group
    const sr = rows[headerRow + 1];
    if (!sr) continue;
    const base = weekDayPositions[0].col;
    let exOff = -1, setsOff = -1, repsOff = -1, weightOff = -1, rpeOff = -1, noteOff = -1;
    for (let c = base; c < base + stride && c < sr.length; c++) {
      const s = String(sr[c] || '').toLowerCase().trim();
      if (s === 'exercise') exOff = c - base;
      else if (s === 'sets') setsOff = c - base;
      else if (s === 'reps') repsOff = c - base;
      else if (s === 'weight') weightOff = c - base;
      else if (s === 'rpe') rpeOff = c - base;
      else if (s === 'notes' || s === 'note' || s === 'coach notes') noteOff = c - base;
    }
    if (exOff === -1 && setsOff === -1) continue;

    // Find all day section headers (multiple "Week-Day" header rows stacked vertically)
    const daySections = [{headerRow, positions: weekDayPositions}];
    for (let i = headerRow + 2; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const wdPos = [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        const m = String(val).trim().match(/^Week\s*(\d+)\s*[-–]\s*Day\s*(\d+)/i);
        if (m) wdPos.push({col: c, week: parseInt(m[1]), day: parseInt(m[2])});
      }
      if (wdPos.length >= 2) daySections.push({headerRow: i, positions: wdPos});
    }

    // Parse exercises from each section
    const weekMap = {};
    for (let sec = 0; sec < daySections.length; sec++) {
      const section = daySections[sec];
      const dataStart = section.headerRow + 2; // skip sub-header row
      const dataEnd = sec + 1 < daySections.length ? daySections[sec + 1].headerRow : rows.length;

      for (let r = dataStart; r < dataEnd; r++) {
        const row = rows[r]; if (!row) continue;
        for (const pos of section.positions) {
          const gb = pos.col;
          const exName = exOff >= 0 ? String(row[gb + exOff] || '').trim() : '';
          if (!exName || !/[a-zA-Z]{2,}/.test(exName)) continue;
          const sv = setsOff >= 0 ? parseInt(row[gb + setsOff]) || 0 : 0;
          const rv = repsOff >= 0 ? row[gb + repsOff] : null;
          const wv = weightOff >= 0 ? row[gb + weightOff] : null;
          const rr = rv != null ? String(rv) : '';
          const wt = typeof wv === 'number' && wv > 0 ? Math.round(wv) : null;
          // Read RPE from separate column
          const rpeRaw = rpeOff >= 0 && row[gb + rpeOff] != null ? String(row[gb + rpeOff]).trim() : '';
          const rpeSuffix = (rpeRaw && /^\d+(\.\d+)?$/.test(rpeRaw)) ? ' @RPE' + rpeRaw : '';
          // Read notes column
          let exNote = '';
          if (noteOff >= 0 && row[gb + noteOff] != null) {
            const n = String(row[gb + noteOff]).trim();
            if (n && !/^(notes?|coach\s*notes?)$/i.test(n)) exNote = n;
          }

          let pres = '';
          if (sv > 0 && rr && wt) pres = sv + 'x' + rr + '(' + wt + ')' + rpeSuffix;
          else if (sv > 0 && rr) pres = sv + 'x' + rr + rpeSuffix;

          const wk = pos.week, dy = pos.day;
          if (!weekMap[wk]) weekMap[wk] = {};
          if (!weekMap[wk][dy]) weekMap[wk][dy] = [];
          weekMap[wk][dy].push({name: exName, prescription: pres, note: exNote, lifterNote: '', loggedWeight: '', supersetGroup: null});
        }
      }
    }

    // Build weeks
    const allWeeks = [];
    for (const wk of Object.keys(weekMap).map(Number).sort((a, b) => a - b)) {
      const days = [];
      for (const dy of Object.keys(weekMap[wk]).map(Number).sort((a, b) => a - b)) {
        const exercises = weekMap[wk][dy].filter(e => e.name);
        if (exercises.length > 0) days.push({name: 'Day ' + dy, exercises});
      }
      if (days.length > 0) allWeeks.push({label: 'Week ' + wk, days});
    }

    if (allWeeks.length > 0) return [{id: 'F_' + (sn || 'Training_Program').replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now(), name: sn || 'Training Program', format: 'F', weeks: allWeeks}];
  }
  return null;
}

// F5: stride-2 — Percentage/Load pairs per week (HLM)
function parseF_pctLoad(wb) {
  const allWeeks = [];
  let blockName = '';
  let globalWeekNum = 0;

  for (const sn of wb.SheetNames) {
    if (_fIsSkipSheet(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = _fCleanRows(ws);
    if (!rows || rows.length < 5) continue;

    // Find day sections: day name in col A + Week labels in same row
    const daySections = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const colA = String(row[0] || '').trim().toLowerCase();
      if (!DAYS.some(d => colA === d || colA.startsWith(d + ' '))) continue;

      // Check for Week labels in this row
      const weekPositions = [];
      for (let c = 1; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        if (/^Week\s*\d/i.test(String(val).trim())) weekPositions.push({col: c, label: String(val).trim()});
      }
      if (weekPositions.length < 2) continue;

      // Find sub-header row below with Percentage/Load repeating
      const nextRow = rows[i + 1];
      if (!nextRow) continue;
      const strs = nextRow.map(_normalizeCell);
      const loadCols = [], rpeCols = [];
      let movCol = -1, setCol = -1, repsCol = -1, noteCol = -1;
      for (let c = 0; c < strs.length; c++) {
        if (strs[c] === 'movement' || strs[c] === 'exercise') movCol = c;
        if (strs[c] === 'set' || strs[c] === 'sets') setCol = c;
        if (strs[c] === 'reps') repsCol = c;
        if (strs[c] === 'load') loadCols.push(c);
        if (strs[c] === 'rpe') rpeCols.push(c);
        if (strs[c] === 'notes' || strs[c] === 'note' || strs[c] === 'coach notes') noteCol = c;
      }
      if (loadCols.length < 2 || movCol === -1) continue;

      daySections.push({row: i, name: String(row[0]).trim(), weekPositions, movCol, setCol, repsCol, loadCols, rpeCols, noteCol, subHeaderRow: i + 1});
    }

    if (daySections.length === 0) continue;
    if (!blockName) blockName = sn;
    const numWeeks = daySections[0].loadCols.length;

    // Parse each day section
    const dayData = [];
    for (let d = 0; d < daySections.length; d++) {
      const sec = daySections[d];
      const startRow = sec.subHeaderRow + 1;
      const endRow = d + 1 < daySections.length ? daySections[d + 1].row : rows.length;
      const exercises = [];

      for (let r = startRow; r < endRow; r++) {
        const row = rows[r]; if (!row) continue;
        const exName = String(row[sec.movCol] || '').trim();
        if (!exName || !/[a-zA-Z]{2,}/.test(exName)) continue;
        const sets = sec.setCol >= 0 ? (parseInt(row[sec.setCol]) || 0) : 0;
        const reps = sec.repsCol >= 0 && row[sec.repsCol] != null ? String(row[sec.repsCol]) : '';
        // Read notes column
        let exNote = '';
        if (sec.noteCol >= 0 && row[sec.noteCol] != null) {
          const n = String(row[sec.noteCol]).trim();
          if (n && !/^(notes?|coach\s*notes?)$/i.test(n)) exNote = n;
        }

        const weekData = [];
        for (let w = 0; w < sec.loadCols.length; w++) {
          const load = row[sec.loadCols[w]];
          const wt = typeof load === 'number' && load > 0 ? Math.round(load) : null;
          // Read RPE from separate column if available
          const rpeRaw = (sec.rpeCols && sec.rpeCols[w] != null && row[sec.rpeCols[w]] != null) ? String(row[sec.rpeCols[w]]).trim() : '';
          const rpeSuffix = (rpeRaw && /^\d+(\.\d+)?$/.test(rpeRaw)) ? ' @RPE' + rpeRaw : '';
          let pres = '';
          if (sets > 0 && reps && wt) pres = sets + 'x' + reps + '(' + wt + ')' + rpeSuffix;
          else if (sets > 0 && reps) pres = sets + 'x' + reps + rpeSuffix;
          weekData.push(pres);
        }
        exercises.push({name: exName, weekData, note: exNote});
      }
      dayData.push({name: sec.name, exercises});
    }

    // Build weeks
    for (let w = 0; w < numWeeks; w++) {
      globalWeekNum++;
      const weekDays = [];
      for (const dd of dayData) {
        const dayExercises = [];
        for (const ex of dd.exercises) {
          if (w >= ex.weekData.length || !ex.weekData[w]) continue;
          dayExercises.push({name: ex.name, prescription: ex.weekData[w], note: ex.note || '', lifterNote: '', loggedWeight: '', supersetGroup: null});
        }
        if (dayExercises.length > 0) weekDays.push({name: dd.name, exercises: dayExercises});
      }
      if (weekDays.length > 0) {
        const wpArr = daySections[0].weekPositions;
        allWeeks.push({label: w < wpArr.length ? wpArr[w].label : 'Week ' + globalWeekNum, days: weekDays});
      }
    }
  }

  if (allWeeks.length === 0) return null;
  return [{id: 'F_' + (blockName || 'block').replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now(), name: blockName, format: 'F', weeks: allWeeks}];
}

// F3: stride-1 — single column per week (Madcow 5x5)
function parseF_stride1(wb) {
  for (const sn of wb.SheetNames) {
    if (_fIsSkipSheet(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = _fCleanRows(ws);
    if (!rows || rows.length < 5) continue;

    // Find header row with Day + Exercise + week columns
    let headerRow = -1, dayCol = -1, exCol = -1;
    let weekColumns = [];

    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      const dIdx = strs.indexOf('day');
      const eIdx = strs.indexOf('exercise');
      if (dIdx === -1 || eIdx === -1) continue;

      const wCols = [];
      const repsColumns = [];
      for (let c = 0; c < row.length; c++) {
        if (strs[c] === 'reps') repsColumns.push(c);
        const wMatch = strs[c].match(/^w(\d+)/i);
        if (wMatch) wCols.push({col: c, num: parseInt(wMatch[1]), label: String(row[c]).trim()});
        if (typeof row[c] === 'number' && row[c] === Math.floor(row[c]) && row[c] >= 1 && row[c] <= 52) {
          wCols.push({col: c, num: row[c], label: 'Week ' + row[c]});
        }
      }

      if (wCols.length >= 3) {
        headerRow = i; dayCol = dIdx; exCol = eIdx;
        wCols.sort((a, b) => a.col - b.col);
        // Assign reps column: nearest Reps col to the left of each week col
        for (const wc of wCols) {
          let bestReps = -1;
          for (const rc of repsColumns) { if (rc < wc.col) bestReps = rc; }
          wc.repsCol = bestReps;
        }
        weekColumns = wCols;
        break;
      }
    }
    if (headerRow === -1) continue;

    // Find day markers
    const dayStarts = [];
    for (let i = headerRow + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const val = String(row[dayCol] || '').trim().toLowerCase();
      if (DAYS.some(d => val === d || val.startsWith(d + ' '))) {
        dayStarts.push({row: i, name: String(row[dayCol]).trim()});
      }
    }
    if (dayStarts.length === 0) continue;

    // Build day data
    const dayData = [];
    for (let d = 0; d < dayStarts.length; d++) {
      const startRow = dayStarts[d].row;
      const endRow = d + 1 < dayStarts.length ? dayStarts[d + 1].row : rows.length;
      const exerciseGroups = [];

      for (let r = startRow; r < endRow; r++) {
        const row = rows[r]; if (!row) continue;
        let exName = String(row[exCol] || '').trim();
        let hasData = false;
        for (const wc of weekColumns) {
          if (typeof row[wc.col] === 'number') { hasData = true; break; }
        }
        if (!hasData) continue;

        if (exName && /[a-zA-Z]{2,}/.test(exName)) {
          exerciseGroups.push({name: exName, rows: [r]});
        } else if (exerciseGroups.length > 0) {
          exerciseGroups[exerciseGroups.length - 1].rows.push(r);
        }
      }

      const dayExercises = [];
      for (const eg of exerciseGroups) {
        const weekSets = [];
        for (const wc of weekColumns) {
          const sets = [];
          for (const ri of eg.rows) {
            const row = rows[ri];
            const weight = row[wc.col];
            let reps = '';
            if (wc.repsCol >= 0) {
              const rv = row[wc.repsCol];
              if (rv != null && rv !== '') reps = String(rv);
            }
            if (typeof weight === 'number' && weight > 0) {
              sets.push(reps ? '1x' + reps + '(' + Math.round(weight) + ')' : String(Math.round(weight)));
            }
          }
          weekSets.push(sets);
        }
        dayExercises.push({name: eg.name, weekSets});
      }
      dayData.push({name: dayStarts[d].name, exercises: dayExercises});
    }

    // Build weeks
    const allWeeks = [];
    for (let w = 0; w < weekColumns.length; w++) {
      const weekDays = [];
      for (const dd of dayData) {
        const exercises = [];
        for (const ex of dd.exercises) {
          if (w >= ex.weekSets.length) continue;
          const sets = ex.weekSets[w];
          if (sets.length === 0) continue;
          exercises.push({name: ex.name, prescription: sets.join(', '), note: '', lifterNote: '', loggedWeight: '', supersetGroup: null});
        }
        if (exercises.length > 0) weekDays.push({name: dd.name, exercises});
      }
      if (weekDays.length > 0) allWeeks.push({label: weekColumns[w].label, days: weekDays});
    }

    if (allWeeks.length > 0) return [{id: 'F_' + (sn || 'block').replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now(), name: sn, format: 'F', weeks: allWeeks}];
  }
  return null;
}

// F6: parseF_coach — coaching template with WEEK headers + EXERCISE/SETS/REPS/LOAD repeating columns
// Handles DUP-style programs (Brendan Tietz, DUP Templates) where:
// - Multiple BLOCK sheets each contain 4 weeks
// - WEEK 1/2/3/4 labels in a header row, stride=7
// - Day sections separated by date rows + EXERCISE sub-header rows
// - Multi-row exercises (top set RPE, back-off weight, AMRAP)
function parseF_coach(wb) {
  const allWeeks = [];
  let blockName = '';
  let globalWeekNum = 0;

  const SKIP_SHEET_COACH = /^(homepage|faq|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|output|pr\s*sheet|notes?|start\s*here|1\.\s*start|predicted|movement\s*prep|nutrition|volume)/i;

  for (const sn of wb.SheetNames) {
    if (SKIP_SHEET_COACH.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = _fCleanRows(ws);
    if (!rows || rows.length < 5) continue;

    // Determine column/row offset: sheet_to_json arrays start from the first column in the sheet range
    const sheetRef = ws['!ref'] || 'A1';
    const sheetRange = XLSX.utils.decode_range(sheetRef);
    const colOffset = sheetRange.s.c; // Physical column index of array[0]
    const rowOffset = sheetRange.s.r; // Physical row index of rows[0]

    // Find week header row: ≥2 "WEEK N" labels
    let weekHeaderRow = -1;
    let weekPositions = [];
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const wpos = [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        if (/^Week\s*\d/i.test(String(val).trim())) wpos.push({col: c, label: String(val).trim()});
      }
      if (wpos.length >= 2) { weekHeaderRow = i; weekPositions = wpos; break; }
    }
    if (weekHeaderRow === -1) continue;

    // Determine stride from week positions
    const stride = weekPositions.length > 1 ? weekPositions[1].col - weekPositions[0].col : 7;

    // Find EXERCISE/SETS/REPS/LOAD sub-header rows (can appear multiple times for different day sections)
    // Detect column offsets within a week group
    let exOff = -1, setsOff = -1, repsOff = -1, loadOff = -1, topSetOff = -1, ratingOff = -1;
    let firstSubHeaderRow = -1;

    for (let i = weekHeaderRow + 1; i < Math.min(weekHeaderRow + 4, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const base = weekPositions[0].col;
      for (let c = base; c < base + stride && c < row.length; c++) {
        const s = String(row[c] || '').toLowerCase().trim();
        if (s === 'exercise') exOff = c - base;
        else if (s === 'sets') setsOff = c - base;
        else if (s === 'reps') repsOff = c - base;
        else if (s === 'load' || s === 'load/rpe' || s === 'load / rpe') loadOff = c - base;
        else if (s === 'top set') topSetOff = c - base;
        else if (s === 'rating') ratingOff = c - base;
      }
      if (exOff >= 0 && setsOff >= 0 && repsOff >= 0) { firstSubHeaderRow = i; break; }
    }
    if (firstSubHeaderRow === -1 || exOff < 0 || setsOff < 0 || repsOff < 0) continue;

    if (!blockName) blockName = sn;

    // Find all day sections: a day starts when we see a sub-header row with EXERCISE repeating
    // Each day section: subheader row followed by exercise data rows
    const daySections = [];
    let dayNum = 0;
    for (let i = firstSubHeaderRow; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      // Check if this row is a sub-header (EXERCISE appears at expected offset for ≥2 week groups)
      let exHeaderCount = 0;
      for (const wp of weekPositions) {
        const val = String(row[wp.col + exOff] || '').toLowerCase().trim();
        if (val === 'exercise') exHeaderCount++;
      }
      if (exHeaderCount >= 2) {
        dayNum++;
        daySections.push({subHeaderRow: i, dayNum});
      }
    }
    if (daySections.length === 0) continue;

    // Parse each day section
    const dayData = [];
    for (let d = 0; d < daySections.length; d++) {
      const startRow = daySections[d].subHeaderRow + 1;
      const endRow = d + 1 < daySections.length ? daySections[d + 1].subHeaderRow - 1 : rows.length;

      // Group exercises within the first week's columns (exercise name in exOff col)
      const base = weekPositions[0].col;
      const exerciseGroups = [];
      for (let r = startRow; r < endRow; r++) {
        const row = rows[r]; if (!row) continue;
        const nameVal = row[base + exOff];
        const exName = nameVal ? String(nameVal).trim() : '';
        // Check if this row has any data (sets/reps/load in any week)
        let hasData = false;
        for (const wp of weekPositions) {
          if (row[wp.col + setsOff] != null || row[wp.col + repsOff] != null) { hasData = true; break; }
        }
        if (!hasData) continue;

        if (exName && /[a-zA-Z]{2,}/.test(exName) && !/^(exercise|week|rest\s*day)/i.test(exName)) {
          exerciseGroups.push({name: exName, rows: [r]});
        } else if (exerciseGroups.length > 0 && !exName) {
          exerciseGroups[exerciseGroups.length - 1].rows.push(r);
        }
      }

      // Build per-week exercise data
      const dayExercises = [];
      for (const eg of exerciseGroups) {
        const weekData = [];
        for (let w = 0; w < weekPositions.length; w++) {
          const wBase = weekPositions[w].col;
          const presParts = [];
          const noteParts = [];
          // Check per-week exercise name (may differ from first week)
          let weekExName = '';
          for (const ri of eg.rows) {
            const row = rows[ri];
            const wkName = row[wBase + exOff] ? String(row[wBase + exOff]).trim() : '';
            if (wkName && /[a-zA-Z]{2,}/.test(wkName) && !/^(exercise|week)/i.test(wkName)) {
              if (!weekExName) weekExName = wkName;
            }
            const setsVal = row[wBase + setsOff];
            const repsVal = row[wBase + repsOff];
            const rawLoadVal = row[wBase + loadOff] != null ? row[wBase + loadOff] : null;
            if (setsVal == null && repsVal == null) continue;

            // Get formatted load value from worksheet cell (handles % formatting)
            let ld = '';
            if (rawLoadVal != null) {
              const physRow = ri + rowOffset;
              const physCol = wBase + loadOff + colOffset;
              const loadCellAddr = XLSX.utils.encode_cell({r: physRow, c: physCol});
              const loadCell = ws[loadCellAddr];
              if (loadCell && loadCell.w) {
                ld = loadCell.w.trim();
              } else {
                ld = String(rawLoadVal).trim();
              }
            }

            const sc = parseInt(setsVal) || 0;
            const rp = repsVal != null ? String(repsVal).trim() : '';

            if (sc > 0 && rp) {
              if (ld && /[a-zA-Z%]/.test(ld)) {
                // Load contains text like "72.5lbs", "-5%", "same", "3 clusters"
                presParts.push(sc + 'x' + rp + '(' + ld + ')');
              } else if (ld && /^[0-9]+$/.test(ld) && parseInt(ld) <= 10) {
                // Numeric load that's small (1-10) is likely RPE
                presParts.push(sc + 'x' + rp + ' @' + ld);
              } else if (ld && /^[0-9.,]+$/.test(ld)) {
                // Numeric load with commas = multiple RPE values
                presParts.push(sc + 'x' + rp + ' @' + ld);
              } else {
                presParts.push(sc + 'x' + rp);
              }
            } else if (sc > 0) {
              presParts.push(sc + ' sets');
            }

            // Read TOP SET value for this row
            if (topSetOff >= 0) {
              const tsRaw = row[wBase + topSetOff];
              if (tsRaw != null) {
                const ts = String(tsRaw).trim();
                if (ts && !/^top\s*set$/i.test(ts)) {
                  noteParts.push(_w_formatTopSet(ts));
                }
              }
            }
            // Read RATING value for this row
            if (ratingOff >= 0) {
              const rtRaw = row[wBase + ratingOff];
              if (rtRaw != null) {
                const rt = String(rtRaw).trim();
                if (rt && !/^rating$/i.test(rt)) {
                  noteParts.push('Rating: ' + rt);
                }
              }
            }
          }
          weekData.push({pres: presParts.join(', '), weekExName: weekExName || eg.name, note: noteParts.join('; ')});
        }
        dayExercises.push({name: eg.name, weekData});
      }
      dayData.push({name: 'Day ' + daySections[d].dayNum, exercises: dayExercises});
    }

    // Build weeks
    for (let w = 0; w < weekPositions.length; w++) {
      globalWeekNum++;
      const weekDays = [];
      for (const dd of dayData) {
        const exercises = [];
        for (const ex of dd.exercises) {
          if (w >= ex.weekData.length) continue;
          const wd = ex.weekData[w];
          if (!wd.pres) continue;
          exercises.push({name: wd.weekExName, prescription: wd.pres, note: wd.note || '', lifterNote: '', loggedWeight: '', supersetGroup: null});
        }
        if (exercises.length > 0) weekDays.push({name: dd.name, exercises});
      }
      if (weekDays.length > 0) {
        allWeeks.push({label: 'Week ' + globalWeekNum, days: weekDays});
      }
    }
  }

  if (allWeeks.length === 0) return null;
  return [{id: 'F_' + (blockName || 'block').replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now(), name: blockName, weeks: allWeeks, format: 'F'}];
}

// ── FORMAT G DETECTION (Sheiko numbered exercises) ───────────────────────────
function detectG(wb) {
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|max|output|pr\s*sheet|notes?|start\s*here|1\.\s*start|predicted|volume)/i;

  for (let si = 0; si < Math.min(8, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows || rows.length < 10) continue;

    let hasWeekLabel = false, hasNumberedExercise = false;
    for (let i = 0; i < Math.min(40, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const colA = row[0], colB_val = row[1];
      // Check col A or col B for "Week N"
      if (colA != null && /^week\s*\d/i.test(String(colA).trim())) hasWeekLabel = true;
      if (colB_val != null && /^week\s*\d/i.test(String(colB_val).trim()) && colA == null) hasWeekLabel = true;
      // Numbered exercise: col A = small int, col B = text, col C = decimal 0-1 (percentage)
      if (typeof colA === 'number' && colA >= 1 && colA <= 20 && colA === Math.floor(colA)) {
        const colB = row[1], colC = row[2];
        if (colB != null && /[a-zA-Z]{2,}/.test(String(colB)) && typeof colC === 'number' && colC > 0 && colC < 1) {
          hasNumberedExercise = true;
        }
      }
      if (hasWeekLabel && hasNumberedExercise) return true;
    }
  }
  return false;
}

// ── FORMAT G PARSER (Sheiko numbered exercises) ──────────────────────────────
function parseG(wb) {
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|max|output|pr\s*sheet|notes?|start\s*here|1\.\s*start|predicted|volume)/i;
  const allWeeks = [];
  let blockName = '';
  let globalWeekNum = 0;

  for (const sn of wb.SheetNames) {
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows || rows.length < 10) continue;

    // Find week and day boundaries (check col A, fall back to col B if col A empty)
    const weekStarts = [];
    const dayStarts = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const colA = String(row[0] || '').trim();
      const colBstr = (row[0] == null && row[1] != null) ? String(row[1]).trim() : '';
      const label = colA || colBstr;
      if (/^week\s*\d/i.test(label)) {
        weekStarts.push({row: i, weekNum: parseInt(label.match(/\d+/)[0])});
      }
      // "N day (DayName)" format or "Day N" format or plain day name
      const dayMatch = label.match(/^(\d+)\s*day\s*\((\w+)/i);
      const dayMatch2 = label.match(/^day\s*(\d+)/i);
      if (dayMatch) {
        dayStarts.push({row: i, dayName: dayMatch[2]});
      } else if (dayMatch2) {
        dayStarts.push({row: i, dayName: 'Day ' + dayMatch2[1]});
      } else if (DAYS.some(d => label.toLowerCase() === d)) {
        dayStarts.push({row: i, dayName: label});
      }
    }
    if (weekStarts.length === 0) continue;
    if (!blockName) blockName = sn;

    // Process each week
    for (let w = 0; w < weekStarts.length; w++) {
      const weekStart = weekStarts[w].row;
      const weekEnd = w + 1 < weekStarts.length ? weekStarts[w + 1].row : rows.length;
      const weekDays = dayStarts.filter(d => d.row > weekStart && d.row < weekEnd);
      if (weekDays.length === 0) continue;

      globalWeekNum++;
      const parsedDays = [];

      for (let d = 0; d < weekDays.length; d++) {
        const dayStart = weekDays[d].row;
        const dayEnd = d + 1 < weekDays.length ? weekDays[d + 1].row : weekEnd;
        const exercises = [];
        let currentEx = null;

        for (let r = dayStart + 1; r < dayEnd; r++) {
          const row = rows[r]; if (!row) continue;
          const colA = row[0], colB = row[1], colC = row[2], colD = row[3], colE = row[4], colF = row[5];

          // New exercise: col A is an integer
          if (typeof colA === 'number' && colA >= 1 && colA <= 20 && colA === Math.floor(colA)) {
            if (currentEx && currentEx.sets.length > 0) exercises.push(_gBuildExercise(currentEx));
            const name = colB != null ? String(colB).trim() : '';
            if (!name || !/[a-zA-Z]{2,}/.test(name)) { currentEx = null; continue; }
            currentEx = {name, sets: []};
            if (colC != null || colD != null || colE != null) {
              currentEx.sets.push({
                pct: typeof colC === 'number' ? colC : null,
                reps: colD != null ? String(colD) : '',
                setCount: typeof colE === 'number' ? Math.round(colE) : (parseInt(colE) || 1),
                weight: typeof colF === 'number' && colF > 0 ? colF : null
              });
            }
          } else if (currentEx && (colC != null || colD != null || colE != null)) {
            // Skip volume/tonnage summary rows: colB names a lift but colC has no percentage
            if (colB != null && String(colB).trim() && colC == null) continue;
            // Continuation row for current exercise
            currentEx.sets.push({
              pct: typeof colC === 'number' ? colC : null,
              reps: colD != null ? String(colD) : '',
              setCount: typeof colE === 'number' ? Math.round(colE) : (parseInt(colE) || 1),
              weight: typeof colF === 'number' && colF > 0 ? colF : null
            });
          }
        }
        if (currentEx && currentEx.sets.length > 0) exercises.push(_gBuildExercise(currentEx));

        if (exercises.length > 0) {
          parsedDays.push({name: weekDays[d].dayName, exercises});
        }
      }

      if (parsedDays.length > 0) {
        allWeeks.push({label: 'Week ' + globalWeekNum, days: parsedDays});
      }
    }
  }

  if (allWeeks.length === 0) return parseB(wb);
  const maxes = _extractMaxesFromSheet(wb);
  return [{id: 'G_' + (blockName || 'block').replace(/[^a-zA-Z0-9]/g, '_') + '_' + Date.now(), name: blockName, format: 'G', weeks: allWeeks, maxes}];
}

function _gBuildExercise(ex) {
  const parts = [];
  for (const s of ex.sets) {
    const sc = s.setCount || 1;
    const reps = s.reps || '';
    const wt = s.weight ? Math.round(s.weight) : null;
    if (sc > 0 && reps) {
      parts.push(wt ? sc + 'x' + reps + '(' + wt + ')' : sc + 'x' + reps);
    }
  }
  return {name: ex.name, prescription: parts.join(', '), note: '', lifterNote: '', loggedWeight: '', supersetGroup: null};
}

// ── FORMAT D PARSER (horizontal week layout — Calgary Barbell style) ──────────
function parseD(wb) {
  const skipSheets = new Set();
  for (const sn of wb.SheetNames) {
    const s = sn.toLowerCase();
    if (s.includes('max') || s.includes('instruction') || s.includes('note') || s.includes('info')
        || s.startsWith('copy of') || s.startsWith('sheet')) skipSheets.add(sn);
  }

  let blockMeta = { name: '', athlete: '' };
  const trainingSheets = wb.SheetNames.filter(sn => !skipSheets.has(sn));
  if (trainingSheets.length > 0) {
    const firstSheet = trainingSheets[0];
    if (/week/i.test(firstSheet)) {
      blockMeta.name = 'Training Block';
    } else {
      blockMeta.name = firstSheet;
    }
  }
  if (!blockMeta.name) blockMeta.name = 'Training Block';

  const allWeeks = [];
  let globalWeekNum = 0;

  for (const sn of wb.SheetNames) {
    if (skipSheets.has(sn)) continue;
    const ws = wb.Sheets[sn];
    const rawRows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    fixDateSerials(ws, rawRows);
    const rows = rawRows.map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows.length) continue;

    const sheetResult = parseDSheet(sn, rows);
    if (!sheetResult || !sheetResult.daySections.length) continue;

    const numWeeks = sheetResult.numWeeks;

    for (let w = 0; w < numWeeks; w++) {
      globalWeekNum++;
      const weekDays = [];
      for (const daySection of sheetResult.daySections) {
        const dayExercises = [];
        for (const ex of daySection.exercises) {
          if (!ex.weekData[w]) continue;
          const wd = ex.weekData[w];
          if (!wd.sets && !wd.reps) continue;
          dayExercises.push({
            name: ex.name,
            prescription: wd.prescription,
            note: wd.note || '',
            lifterNote: '',
            loggedWeight: '',
            supersetGroup: null
          });
        }
        if (dayExercises.length > 0) {
          weekDays.push({ name: daySection.dayName, exercises: dayExercises });
        }
      }
      if (weekDays.length > 0) {
        allWeeks.push({ label: 'Week ' + globalWeekNum, days: weekDays });
      }
    }
  }

  if (allWeeks.length === 0) return parseB(wb);

  return [{
    id: (blockMeta.name||'training_block').replace(/[^a-zA-Z0-9]/g,'_') + '_D_' + Date.now(),
    name: blockMeta.name,
    format: 'D',
    athleteName: blockMeta.athlete || '',
    dateRange: '',
    maxes: {},
    weeks: allWeeks
  }];
}

function parseDSheet(sheetName, rows) {
  const daySections = [];
  const dayStartRows = [];

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]; if (!row) continue;
    const colA = String(row[0] || '').trim();
    const wkDayMatch = colA.match(/^WEEK\s*\d+(?:\s*,\s*Day\s*(\d+))?$/i);
    const taperMatch = colA.match(/^(\d+)\s*Days?\s*(from|before)\s*Competition$/i);
    const dayOnlyMatch = colA.match(/^Day\s*(\d+)$/i);
    if (wkDayMatch || taperMatch || dayOnlyMatch) {
      let dayName;
      if (taperMatch) {
        dayName = colA;
      } else if (dayOnlyMatch) {
        dayName = 'Day ' + dayOnlyMatch[1];
      } else if (wkDayMatch[1]) {
        dayName = 'Day ' + wkDayMatch[1];
      } else {
        dayName = 'Day 1';
      }
      dayStartRows.push({ rowIdx: r, dayName });
    }
  }

  if (dayStartRows.length === 0) return null;

  let numWeeks = 0;
  let firstHeaderRow = null;
  for (let r = dayStartRows[0].rowIdx + 1; r < Math.min(dayStartRows[0].rowIdx + 3, rows.length); r++) {
    const row = rows[r]; if (!row) continue;
    const strs = row.map(_normalizeCell);
    const setsCount = strs.filter(s => s === 'sets').length;
    if (setsCount >= 1) {
      numWeeks = setsCount;
      firstHeaderRow = r;
      break;
    }
  }
  if (numWeeks === 0) return null;

  for (let d = 0; d < dayStartRows.length; d++) {
    const startRow = dayStartRows[d].rowIdx;
    const endRow = d + 1 < dayStartRows.length ? dayStartRows[d + 1].rowIdx : rows.length;

    let colHeaderIdx = -1;
    let weekColGroups = [];

    for (let r = startRow + 1; r < Math.min(startRow + 4, endRow); r++) {
      const row = rows[r]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      if (strs.includes('sets')) {
        colHeaderIdx = r;
        for (let c = 0; c < strs.length; c++) {
          if (strs[c] === 'sets') {
            const group = { setsCol: c, repsCol: -1, loadCol: -1, intensityCol: -1, tempoCol: -1, restCol: -1 };
            for (let cc = c + 1; cc < Math.min(c + 8, strs.length); cc++) {
              const h = strs[cc];
              if (h === 'reps' && group.repsCol === -1) group.repsCol = cc;
              else if ((h === 'load' || h === 'weight') && group.loadCol === -1) group.loadCol = cc;
              else if (h === 'intensity' && group.intensityCol === -1) group.intensityCol = cc;
              else if (h === 'tempo' && group.tempoCol === -1) group.tempoCol = cc;
              else if (h.startsWith('rest') && group.restCol === -1) group.restCol = cc;
              else if (h === 'performed' && group.loadCol === -1) { /* skip */ }
              else if (h === 'sets') break;
            }
            weekColGroups.push(group);
          }
        }
        break;
      }
    }

    if (colHeaderIdx === -1 || weekColGroups.length === 0) continue;

    const exercises = [];
    let lastExName = '';
    for (let r = colHeaderIdx + 1; r < endRow; r++) {
      const row = rows[r]; if (!row) continue;
      let exName = String(row[0] || '').trim();
      if (!exName) {
        const hasSets = row[weekColGroups[0].setsCol] != null;
        if (hasSets && lastExName) {
          exName = lastExName;
        } else {
          continue;
        }
      } else {
        lastExName = exName;
      }
      if (/^WEEK\s*\d/i.test(exName) || /^\d+\s*Days?\s*(from|before)/i.test(exName)) break;
      if (/^exercise(\s+movement)?$/i.test(exName)) continue;

      const weekData = [];
      for (let w = 0; w < weekColGroups.length; w++) {
        const g = weekColGroups[w];
        const sets = row[g.setsCol];
        const reps = row[g.repsCol];
        const loadRaw = g.loadCol >= 0 ? row[g.loadCol] : null;
        const intensity = g.intensityCol >= 0 ? row[g.intensityCol] : null;
        const tempo = g.tempoCol >= 0 ? row[g.tempoCol] : null;
        const rest = g.restCol >= 0 ? row[g.restCol] : null;

        let load = loadRaw;
        let effectiveIntensity = intensity;
        if (g.loadCol >= 0) {
          const nextColVal = row[g.loadCol + 1];
          const loadStr = String(loadRaw || '').trim();
          const isRPE = /rpe/i.test(loadStr);
          const isPct = typeof loadRaw === 'number' && loadRaw > 0 && loadRaw <= 1;
          const isOpener = /opener/i.test(loadStr);
          if (isRPE || isPct || isOpener) {
            effectiveIntensity = loadRaw;
            if (typeof nextColVal === 'number' && nextColVal > 1) {
              load = nextColVal;
            } else {
              load = null;
            }
          }
        }

        weekData.push(buildDPrescription(sets, reps, load, effectiveIntensity, tempo, rest));
      }

      exercises.push({ name: exName.replace(/\s+/g,' ').trim(), weekData });
    }

    daySections.push({
      dayName: dayStartRows[d].dayName,
      exercises
    });
  }

  return { daySections, numWeeks };
}

function buildDPrescription(sets, reps, load, intensity, tempo, rest) {
  if (sets == null && reps == null) return { sets: 0, reps: '', prescription: '', note: '' };

  const setsStr = String(sets || '').trim();
  const repsStr = String(reps || '').trim();
  const loadVal = load != null ? load : null;
  const intensityStr = intensity != null ? String(intensity).trim() : '';
  const tempoStr = tempo != null ? String(tempo).trim() : '';

  let totalSets = 0;
  const plusMatch = setsStr.match(/^(\d+)\+(\d+)[FR]$/i);
  if (plusMatch) {
    totalSets = parseInt(plusMatch[1]) + parseInt(plusMatch[2]);
  } else {
    totalSets = parseInt(setsStr) || 0;
  }

  let repsDisplay = repsStr;
  if (repsStr.toLowerCase() === 'x') {
    if (/^\d+s$/i.test(intensityStr)) {
      repsDisplay = intensityStr;
    }
  }

  let pres = '';
  if (totalSets > 0 && repsDisplay) {
    if (typeof loadVal === 'number' && loadVal > 0) {
      pres = totalSets + 'x' + repsDisplay + '(' + loadVal + ')';
    } else {
      pres = totalSets + 'x' + repsDisplay;
    }
  } else if (totalSets > 0) {
    pres = totalSets + ' sets';
  }

  const notes = [];

  if (intensityStr && !/^\d+s$/i.test(intensityStr) && intensityStr.toLowerCase() !== 'x') {
    const numIntensity = parseFloat(intensityStr);
    if (/opener/i.test(intensityStr)) {
      notes.push('Opener');
    } else if (/rpe/i.test(intensityStr)) {
      const rpeNum = parseFloat(intensityStr);
      if (!isNaN(rpeNum)) notes.push('@' + rpeNum);
    } else if (!isNaN(numIntensity) && numIntensity > 0 && numIntensity <= 1) {
      notes.push(Math.round(numIntensity * 100) + '%');
    }
  }

  if (plusMatch) {
    const topSets = plusMatch[1];
    const extraSets = plusMatch[2];
    const fOrR = setsStr.slice(-1).toUpperCase();
    const label = fOrR === 'F' ? 'fatigue' : 'rep';
    notes.push(topSets + ' top + ' + extraSets + ' ' + label);
  }

  if (tempoStr && tempoStr.toLowerCase() !== 'x') {
    notes.push('Tempo: ' + tempoStr);
  }

  return {
    sets: totalSets,
    reps: repsDisplay,
    prescription: pres,
    note: notes.join(' · ')
  };
}

// ── POST-PROCESSING ───────────────────────────────────────────────────────────
function deduplicateExerciseNames(blocks) {
  for (const block of blocks) {
    for (const week of (block.weeks || [])) {
      for (const day of (week.days || [])) {
        const seen = {};
        for (const ex of (day.exercises || [])) {
          const base = ex.name;
          if (seen[base] === undefined) {
            seen[base] = 1;
          } else {
            seen[base]++;
            ex.displayName = base;
            const disambig = ex.prescription
              ? ex.prescription.replace(/\s+/g,' ').trim().substring(0, 30)
              : ('#' + seen[base]);
            ex.name = base + ' [' + disambig + ']';
            const firstEx = day.exercises.find(e => e !== ex && (e.displayName || e.name) === base);
            if (firstEx && !firstEx.displayName) firstEx.displayName = base;
          }
        }
      }
    }
  }
}

// ── SET PARSERS ───────────────────────────────────────────────────────────────
// Session-level memoization caches — cleared via clearParseCache() on block/program change
const _parseSetsCache = new Map();
const _parseSimpleCache = new Map();
const _parseRangeSetCache = new Map();
const _parseBarePresCache = new Map();
const _parseWarmupSetsCache = new Map();

function clearParseCache() {
  _parseSetsCache.clear();
  _parseSimpleCache.clear();
  _parseRangeSetCache.clear();
  _parseBarePresCache.clear();
  _parseWarmupSetsCache.clear();
}

// ── parseSets helper: parses a single prescription segment (no semicolons/commas)
// Handles all @ notation, AMRAP, plus-notation, and (for comma-split context) plain NxM and bare N.
function _parseSingleSetGroup(seg){
  seg = seg.trim();
  if(!seg) return null;
  // NxM(weight) — existing parenthesized format, extended to AMRAP/AMAP/MR/F/? as reps
  {
    const sets=[]; const pat=/(\d+x)?(\d+|AMRAP|AMAP|MR|F|\?)\s*\(([^)]+)\)/gi; let m;
    while((m=pat.exec(seg))!==null){
      const mult=m[1]?parseInt(m[1]):1;
      const rraw=m[2]; const rups=rraw.toUpperCase();
      const reps=/^\d+$/.test(rraw)?parseInt(rraw):(rups==='AMAP'||rups==='MR'||rups==='F')?'AMRAP':rraw;
      const amrapFlag=reps==='AMRAP';
      for(let i=0;i<mult;i++) sets.push({reps,weight:m[3].trim(),isRpe:isNaN(Number(m[3].trim())),...(amrapFlag?{amrap:true}:{})});
    }
    if(sets.length>0) return sets;
  }
  let r;
  // NxAMRAP @RPE — aliases: AMRAP, AMAP, MR, F
  r=seg.match(/^(\d+)\s*x\s*(AMRAP|AMAP|MR|F)\s*@\s*RPE\s*(\d+(?:\.\d+)?)\s*$/i);
  if(r){ const n=parseInt(r[1]),rpe=parseFloat(r[3]); return Array.from({length:n},()=>({reps:'AMRAP',rpe,amrap:true})); }
  // NxAMRAP bare — aliases: AMRAP, AMAP, MR, F
  r=seg.match(/^(\d+)\s*x\s*(AMRAP|AMAP|MR|F)\s*$/i);
  if(r){ const n=parseInt(r[1]); return Array.from({length:n},()=>({reps:'AMRAP',amrap:true})); }
  // NxM @RPE (with keyword)  e.g. "3x5 @RPE8"
  r=seg.match(/^(\d+)\s*x\s*(\d+)\s*@\s*RPE\s*(\d+(?:\.\d+)?)\s*$/i);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),rpe=parseFloat(r[3]); return Array.from({length:n},()=>({reps,rpe})); }
  // NxM @weight lbs/kg  e.g. "3x6 @330lbs"
  r=seg.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(\d+(?:\.\d+)?)\s*(lbs?|kg)\s*$/i);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),weight=parseFloat(r[3]),unit=r[4].toLowerCase(); return Array.from({length:n},()=>({reps,weight,unit})); }
  // NxM @-drop%  e.g. "2x7 @-15%"
  r=seg.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(-\d+(?:\.\d+)?)%\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),dropPercent=parseFloat(r[3])/100; return Array.from({length:n},()=>({reps,dropPercent})); }
  // NxM @%  e.g. "5x8 @75%"
  r=seg.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(\d+(?:\.\d+)?)%\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),percentage=parseFloat(r[3])/100; return Array.from({length:n},()=>({reps,percentage})); }
  // NxM @0.decimal  e.g. "3x5 @0.80"
  r=seg.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(0\.\d+)\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),percentage=parseFloat(r[3]); return Array.from({length:n},()=>({reps,percentage})); }
  // NxM+ (plus-notation, last set AMRAP)  e.g. "3x5+"
  r=seg.match(/^(\d+)\s*x\s*(\d+)\+\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]); const s=Array.from({length:n-1},()=>({reps})); s.push({reps,amrap:true}); return s; }
  // N@RPE bare (single set)  e.g. "5@6"
  r=seg.match(/^(\d+)\s*@\s*(\d+(?:\.\d+)?)\s*$/);
  if(r){ return [{reps:parseInt(r[1]),rpe:parseFloat(r[2])}]; }
  // NxM @bare number (no keyword, treat as RPE)  e.g. "2x5@8"
  r=seg.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(\d+(?:\.\d+)?)\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),rpe=parseFloat(r[3]); return Array.from({length:n},()=>({reps,rpe})); }
  // Nx? (open/unknown reps)  e.g. "3x?" — n≤20 guard
  r=seg.match(/^(\d+)\s*x\s*\?\s*$/);
  if(r){ const n=parseInt(r[1]); if(n>20) return null; return Array.from({length:n},()=>({reps:'?'})); }
  // Range-set with @RPE (for comma-split context)  e.g. "2-3x3@9"
  r=seg.match(/^(\d+)[\-–](\d+)\s*x\s*(\d+)\s*@\s*(?:RPE\s*)?(\d+(?:\.\d+)?)\s*$/i);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[3]),rpe=parseFloat(r[4]); return Array.from({length:n},()=>({reps,rpe})); }
  // +Nunit xR xS (relative weight, from-zero offset)  e.g. "+5kg x6 x2"
  r=seg.match(/^\+(\d+(?:\.\d+)?)\s*(kg|lbs?)\s*x\s*(\d+)\s*x\s*(\d+)\s*$/i);
  if(r){ const weight='+'+r[1],unit=r[2].toLowerCase(),reps=parseInt(r[3]),n=parseInt(r[4]); return Array.from({length:n},()=>({reps,weight,unit})); }
  // NxM plain (for comma-split context only)  e.g. "3x5"  — n≤20 guard prevents warmup "135x8" producing 135 sets
  r=seg.match(/^(\d+)\s*x\s*(\d+)\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]); if(n>20) return null; return Array.from({length:n},()=>({reps})); }
  // Bare number (for comma-split context)  e.g. "5"
  r=seg.match(/^(\d+)\s*$/);
  if(r){ return [{reps:parseInt(r[1])}]; }
  return null;
}

function parseSets(pres){
  if(!pres) return [];
  pres = pres.replace(/×/g, 'x'); /* normalize unicode multiply sign */
  const cached = _parseSetsCache.get(pres);
  if(cached !== undefined) return cached;
  // Bail out if prescription is a set range like "3-5x1 (405)" — let parseRangeSet handle it
  if(/^\d+[\-–]\d+\s*x/i.test(pres.trim())) { _parseSetsCache.set(pres,[]); return []; }

  // Semicolon split: "1x5 @RPE5; 2x7 @-15%" → parse each segment and concat
  if(pres.includes(';')){
    const segs=pres.split(';').map(s=>s.trim()).filter(Boolean);
    if(segs.length>1){
      const sets=[];
      for(const seg of segs) sets.push(...parseSets(seg));
      _parseSetsCache.set(pres,sets); return sets;
    }
  }

  // Existing NxM(weight) global pattern — extended to AMRAP/AMAP/MR/F/? as reps
  {
    const sets=[]; const pat=/(\d+x)?(\d+|AMRAP|AMAP|MR|F|\?)\s*\(([^)]+)\)/gi; let m;
    while((m=pat.exec(pres))!==null){
      const mult=m[1]?parseInt(m[1]):1;
      const rraw=m[2]; const rups=rraw.toUpperCase();
      const reps=/^\d+$/.test(rraw)?parseInt(rraw):(rups==='AMAP'||rups==='MR'||rups==='F')?'AMRAP':rraw;
      const amrapFlag=reps==='AMRAP';
      for(let i=0;i<mult;i++) sets.push({reps,weight:m[3].trim(),isRpe:isNaN(Number(m[3].trim())),...(amrapFlag?{amrap:true}:{})});
    }
    if(sets.length>0){ _parseSetsCache.set(pres,sets); return sets; }
  }

  // Comma split: "5@6, 5@7, 2x5@8" or "3x5, 5x3, 1x1+" — each segment parsed via helper
  if(pres.includes(',')){
    const segs=pres.split(',').map(s=>s.trim()).filter(Boolean);
    const sets=[]; let ok=true;
    for(const seg of segs){
      const parsed=_parseSingleSetGroup(seg);
      if(parsed&&parsed.length>0) sets.push(...parsed);
      else{ ok=false; break; }
    }
    if(ok&&sets.length>0){ _parseSetsCache.set(pres,sets); return sets; }
  }

  // New single-group @ patterns (no plain NxM — parseSets("3x5") still returns [])
  let r;
  // NxAMRAP @RPE — aliases: AMRAP, AMAP, MR, F  e.g. "3xAMRAP @RPE8"
  r=pres.match(/^(\d+)\s*x\s*(AMRAP|AMAP|MR|F)\s*@\s*RPE\s*(\d+(?:\.\d+)?)\s*$/i);
  if(r){ const n=parseInt(r[1]),rpe=parseFloat(r[3]); const sets=Array.from({length:n},()=>({reps:'AMRAP',rpe,amrap:true})); _parseSetsCache.set(pres,sets); return sets; }
  // NxAMRAP bare — aliases: AMRAP, AMAP, MR, F  e.g. "3xAMRAP" or "4xAMAP"
  r=pres.match(/^(\d+)\s*x\s*(AMRAP|AMAP|MR|F)\s*$/i);
  if(r){ const n=parseInt(r[1]); const sets=Array.from({length:n},()=>({reps:'AMRAP',amrap:true})); _parseSetsCache.set(pres,sets); return sets; }
  // NxM @RPE (with keyword)  e.g. "3x5 @RPE8"
  r=pres.match(/^(\d+)\s*x\s*(\d+)\s*@\s*RPE\s*(\d+(?:\.\d+)?)\s*$/i);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),rpe=parseFloat(r[3]); const sets=Array.from({length:n},()=>({reps,rpe})); _parseSetsCache.set(pres,sets); return sets; }
  // NxM @weight lbs/kg  e.g. "3x6 @330lbs"
  r=pres.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(\d+(?:\.\d+)?)\s*(lbs?|kg)\s*$/i);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),weight=parseFloat(r[3]),unit=r[4].toLowerCase(); const sets=Array.from({length:n},()=>({reps,weight,unit})); _parseSetsCache.set(pres,sets); return sets; }
  // NxM @-drop%  e.g. "2x7 @-15%"
  r=pres.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(-\d+(?:\.\d+)?)%\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),dropPercent=parseFloat(r[3])/100; const sets=Array.from({length:n},()=>({reps,dropPercent})); _parseSetsCache.set(pres,sets); return sets; }
  // NxM @%  e.g. "5x8 @75%"
  r=pres.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(\d+(?:\.\d+)?)%\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),percentage=parseFloat(r[3])/100; const sets=Array.from({length:n},()=>({reps,percentage})); _parseSetsCache.set(pres,sets); return sets; }
  // NxM @0.decimal  e.g. "3x5 @0.80"
  r=pres.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(0\.\d+)\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),percentage=parseFloat(r[3]); const sets=Array.from({length:n},()=>({reps,percentage})); _parseSetsCache.set(pres,sets); return sets; }
  // NxM+ (plus-notation)  e.g. "3x5+"
  r=pres.match(/^(\d+)\s*x\s*(\d+)\+\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]); const sets=Array.from({length:n-1},()=>({reps})); sets.push({reps,amrap:true}); _parseSetsCache.set(pres,sets); return sets; }
  // N@RPE bare  e.g. "5@6"
  r=pres.match(/^(\d+)\s*@\s*(\d+(?:\.\d+)?)\s*$/);
  if(r){ const sets=[{reps:parseInt(r[1]),rpe:parseFloat(r[2])}]; _parseSetsCache.set(pres,sets); return sets; }
  // NxM @bare number (no keyword, treat as RPE)  e.g. "2x5@8"
  r=pres.match(/^(\d+)\s*x\s*(\d+)\s*@\s*(\d+(?:\.\d+)?)\s*$/);
  if(r){ const n=parseInt(r[1]),reps=parseInt(r[2]),rpe=parseFloat(r[3]); const sets=Array.from({length:n},()=>({reps,rpe})); _parseSetsCache.set(pres,sets); return sets; }
  // Nx? (open/unknown reps)  e.g. "3x?"
  r=pres.match(/^(\d+)\s*x\s*\?\s*$/);
  if(r){ const n=parseInt(r[1]); const sets=Array.from({length:n},()=>({reps:'?'})); _parseSetsCache.set(pres,sets); return sets; }
  // +Nunit xR xS (relative weight offset)  e.g. "+5kg x6 x2"
  r=pres.match(/^\+(\d+(?:\.\d+)?)\s*(kg|lbs?)\s*x\s*(\d+)\s*x\s*(\d+)\s*$/i);
  if(r){ const weight='+'+r[1],unit=r[2].toLowerCase(),reps=parseInt(r[3]),n=parseInt(r[4]); const sets=Array.from({length:n},()=>({reps,weight,unit})); _parseSetsCache.set(pres,sets); return sets; }

  _parseSetsCache.set(pres,[]);
  return [];
}

function parseSimple(pres){
  if(!pres) return null;
  pres = pres.replace(/×/g, 'x'); /* normalize unicode multiply sign */
  if(_parseSimpleCache.has(pres)) return _parseSimpleCache.get(pres);
  const m=pres.match(/^(\d+)\s*x\s*([\d\-]+(?:\s*(?:sec|min|s|m)\w*)?)/i);
  if(!m){ _parseSimpleCache.set(pres,null); return null; }
  if(pres.match(/^\d+[\-–]\d+x/i)){ _parseSimpleCache.set(pres,null); return null; }
  const sets = parseInt(m[1]);
  if(sets > 20){ _parseSimpleCache.set(pres,null); return null; }
  const result={sets,reps:m[2].trim()};
  _parseSimpleCache.set(pres, result);
  return result;
}

function parseRangeSet(pres){
  if(!pres) return null;
  pres = pres.replace(/×/g, 'x'); /* normalize unicode multiply sign */
  if(_parseRangeSetCache.has(pres)) return _parseRangeSetCache.get(pres);
  // Range with optional weight in parens: "3-5x1 (405)"
  const mw=pres.match(/^(\d+)[\-–](\d+)\s*x\s*([\d\-–]+(?:\s*(?:sec|min|s|m)\w*)?)\s*\(([^)]+)\)/i);
  if(mw){
    const sets=parseInt(mw[1]);
    const maxSets=parseInt(mw[2]);
    const reps=mw[3].trim();
    const weight=mw[4].trim();
    const isRpe=isNaN(Number(weight));
    const result={sets,maxSets,reps,weight,isRpe};
    _parseRangeSetCache.set(pres, result);
    return result;
  }
  // Range with optional weight after @: "3-5x1 @ 435lb" or "3-5x1 @ 435"
  const mAt=pres.match(/^(\d+)[\-–](\d+)\s*x\s*([\d\-–]+(?:\s*(?:sec|min|s|m)\w*)?)\s*@\s*(\d+(?:\.\d+)?)\s*(?:lbs?|kg)?\s*$/i);
  if(mAt){
    const sets=parseInt(mAt[1]);
    const maxSets=parseInt(mAt[2]);
    const reps=mAt[3].trim();
    const weight=mAt[4].trim();
    const result={sets,maxSets,reps,weight,isRpe:false};
    _parseRangeSetCache.set(pres, result);
    return result;
  }
  const m=pres.match(/^(\d+)[\-–]?(\d*)\s*x\s*([\d\-–]+(?:\s*(?:sec|min|s|m)\w*)?)/i);
  if(m){
    const sets=parseInt(m[1]);
    // Sanity: >20 sets in one NxM notation is almost certainly a weight misinterpretation (e.g. "45x45.0 @50%")
    if(sets>20){ _parseRangeSetCache.set(pres, null); return null; }
    const maxSets=m[2]?parseInt(m[2]):null;
    const reps=m[3].trim();
    /* Extract @ weight if present: "3x5 @ 315lb" */
    const atW=pres.match(/@\s*(\d+(?:\.\d+)?)\s*(?:lbs?|kg)?\s*$/i);
    const result=atW?{sets,maxSets,reps,weight:atW[1],isRpe:false}:{sets,maxSets,reps};
    _parseRangeSetCache.set(pres, result);
    return result;
  }
  const s=pres.match(/^(\d+)[\-–](\d+)\s*sets?/i);
  if(s){ const result={sets:parseInt(s[1]),maxSets:parseInt(s[2]),reps:'open'}; _parseRangeSetCache.set(pres,result); return result; }
  _parseRangeSetCache.set(pres, null);
  return null;
}

function parseBarePres(pres){
  if(!pres) return null;
  if(_parseBarePresCache.has(pres)) return _parseBarePresCache.get(pres);
  const p=pres.trim();
  let result = null;
  if(/^\d+$/.test(p)) result={sets:1,reps:p};
  else if(/^\d+:\d{2}$/.test(p)) result={sets:1,reps:p};
  // Rep range like "15-20" or "8-12" (with or without spaces around dash)
  else if(/^\d+\s*[\-–]\s*\d+$/.test(p)) result={sets:1,reps:p.replace(/\s/g,'')};
  // "N reps" or "1 rep (rest note)" — "N" is set count 1, reps = N
  else if(/^\d+\s+reps?(\s.*)?$/i.test(p)) result={sets:1,reps:String(parseInt(p))};
  // "N sets" — N sets of unspecified reps
  else if(/^\d+\s+sets?$/i.test(p)) result={sets:parseInt(p),reps:'open'};
  // "(N)" — parenthesized weight only  e.g. "(312)"
  else if(/^\(\s*\d+(?:\.\d+)?\s*\)$/.test(p)){ const w=p.replace(/[()]/g,'').trim(); result={sets:1,reps:'1',weight:w}; }
  // "NRM" — rep-max notation  e.g. "5RM"
  else if(/^\d+RM$/i.test(p)) result={sets:1,reps:String(parseInt(p))};
  // "Max Attempt" or "Max Out" — 1 set of max weight/reps
  else if(/^max\s+(attempt|out)[!]*/i.test(p)) result={sets:1,reps:'max'};
  _parseBarePresCache.set(pres, result);
  return result;
}

function parseWarmupSets(pres){
  if(!pres) return [];
  if(_parseWarmupSetsCache.has(pres)) return _parseWarmupSetsCache.get(pres);
  const p=pres.trim();
  if(!p.includes(',')){ _parseWarmupSetsCache.set(pres,[]); return []; }
  const parts=p.split(',').map(s=>s.trim()).filter(Boolean);
  const sets=[];
  for(const part of parts){
    const m=part.match(/^(\w+)\s*x\s*(\d+)$/i);
    if(m){
      sets.push({reps:parseInt(m[2]),weight:m[1],isRpe:false});
    } else {
      return [];
    }
  }
  const result = sets.length>=2 ? sets : [];
  _parseWarmupSetsCache.set(pres, result);
  return result;
}

// ── FORMAT H: STRONG / HEVY CSV IMPORT ────────────────────────────────────────

function parseCSVRows(text) {
  // Simple CSV parser handling quoted fields with commas/newlines
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') { inQuotes = false; }
      else { field += ch; }
    } else {
      if (ch === '"') { inQuotes = true; }
      else if (ch === ',') { row.push(field.trim()); field = ''; }
      else if (ch === '\n' || (ch === '\r' && text[i + 1] === '\n')) {
        row.push(field.trim()); field = '';
        if (row.some(c => c !== '')) rows.push(row);
        row = [];
        if (ch === '\r') i++;
      } else { field += ch; }
    }
  }
  // Last row
  row.push(field.trim());
  if (row.some(c => c !== '')) rows.push(row);
  return rows;
}

function detectCSVFormat(headers) {
  const h = headers.map(c => c.toLowerCase().trim());
  if (h.includes('workout name') && h.includes('set order')) return 'strong';
  if (h.includes('exercise_title') && h.includes('set_index')) return 'hevy';
  // FitNotes: has both kg and lbs weight columns + Kind column
  if (h.includes('weight (kg)') && h.includes('weight (lbs)') && h.includes('kind')) return 'fitnotes';
  // Fitbod: has Weight(kg) (no space) + isWarmup boolean + multiplier
  if (h.includes('weight(kg)') && h.includes('iswarmup') && h.includes('multiplier')) return 'fitbod';
  return null;
}

function parseCSVImport(csvText, unitLabel) {
  const rows = parseCSVRows(csvText);
  if (rows.length < 2) return null;

  const headers = rows[0];
  const fmt = detectCSVFormat(headers);
  if (!fmt) return null;

  // Build column index map
  const col = {};
  headers.forEach((h, i) => { col[h.toLowerCase().trim()] = i; });

  // Extract rows into normalized objects
  const entries = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const get = (key) => r[col[key]] || '';

    let date, workoutName, exerciseName, setOrder, weight, reps, rpe, notes, restTime, equipment;

    if (fmt === 'strong') {
      date = get('date');
      workoutName = get('workout name');
      exerciseName = get('exercise name');
      setOrder = parseInt(get('set order')) || 0;
      weight = get('weight');
      reps = get('reps');
      rpe = get('rpe') || '';
      notes = get('notes');
      restTime = get('seconds') || get('rest_seconds') || get('rest time') || '';
      equipment = get('equipment') || '';
    } else if (fmt === 'fitnotes') {
      // FitNotes CSV: Date, Exercise, Category, Weight (kg), Weight (lbs), Reps, Distance, Distance Unit, Time, Notes, Kind
      date = get('date');  // YYYY-MM-DD
      workoutName = date;  // FitNotes has no workout name concept — group by date
      exerciseName = get('exercise');
      setOrder = 0;  // No set order in FitNotes — sets are sequential per exercise
      // Use lbs column by default; fall back to kg * 2.20462
      // Treat non-empty lbs value (including "0") as intentional — don't fall through to kg conversion
      const lbsVal = get('weight (lbs)');
      const kgVal = get('weight (kg)');
      if (lbsVal && lbsVal.trim() !== '') {
        weight = lbsVal;
      } else if (kgVal && parseFloat(kgVal) > 0) {
        weight = String(Math.round(parseFloat(kgVal) * 2.20462 * 10) / 10);
      } else {
        weight = '';
      }
      reps = get('reps');
      rpe = '';
      notes = get('notes');
      // Skip warmup/dropset if Kind indicates it
      const kind = get('kind').toLowerCase();
      if (kind === 'warmup') { notes = (notes ? notes + ' ' : '') + 'Warmup'; }
      if (kind === 'dropset') { notes = (notes ? notes + ' ' : '') + 'Dropset'; }
      restTime = get('time') || get('rest_seconds') || '';
      equipment = get('equipment') || '';
    } else if (fmt === 'fitbod') {
      // Fitbod CSV: Date, Exercise, Reps, Weight(kg), Duration(s), Distance(m), Incline, Resistance, isWarmup, Note, multiplier
      const rawDate = get('date');
      // Fitbod dates include time: "YYYY-MM-DD HH:MM:SS" — extract date portion
      date = rawDate.split(' ')[0] || rawDate;
      workoutName = date;
      exerciseName = get('exercise');
      setOrder = 0;
      // Weight is always in kg — convert to lbs
      const kgWeight = get('weight(kg)');
      if (kgWeight && parseFloat(kgWeight) > 0) {
        weight = String(Math.round(parseFloat(kgWeight) * 2.20462 * 10) / 10);
      } else {
        weight = '';
      }
      reps = get('reps');
      rpe = '';
      notes = get('note');
      // Flag warmup sets
      const isWarmup = (get('iswarmup') || '').toLowerCase().trim();
      if (isWarmup === 'true' || isWarmup === '1' || isWarmup === 'yes') {
        notes = (notes ? notes + ' ' : '') + 'Warmup';
      }
      restTime = get('duration(s)') || get('rest_seconds') || '';
      const _resistance = get('resistance'), _incline = get('incline');
      const _eqParts = [];
      if (_resistance && _resistance.trim() && _resistance !== '0') _eqParts.push(_resistance.trim());
      if (_incline && _incline.trim() && _incline !== '0') _eqParts.push('Incline ' + _incline.trim());
      equipment = _eqParts.join(', ') || get('equipment') || '';
    } else {
      // Hevy — parse date from start_time
      const startRaw = get('start_time');
      // Hevy dates can be "28 Mar 2025, 17:29" or ISO "2025-03-28T17:29:00Z"
      const d = new Date(startRaw);
      date = !isNaN(d) ? `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}` : startRaw;
      workoutName = get('title');
      exerciseName = get('exercise_title');
      setOrder = parseInt(get('set_index')) || 0;
      weight = get('weight_lbs');
      reps = get('reps');
      rpe = get('rpe');
      notes = get('exercise_notes');
      restTime = get('duration_seconds') || get('rest_seconds') || '';
      equipment = get('equipment') || '';
    }

    // Skip rows with no exercise name
    if (!exerciseName) continue;
    // Skip cardio: no weight AND no reps
    const w = parseFloat(weight);
    const rep = parseInt(reps);
    if ((isNaN(w) || w === 0) && (isNaN(rep) || rep === 0)) continue;

    entries.push({ date, workoutName, exerciseName, setOrder, weight: isNaN(w) ? '' : String(w), reps: isNaN(rep) ? '' : String(rep), rpe, notes, restTime: restTime || '', equipment: equipment || '' });
  }

  if (!entries.length) return null;

  // Group entries by date + workout name → "sessions"
  const sessionMap = new Map();
  for (const e of entries) {
    const key = e.date + '|||' + e.workoutName;
    if (!sessionMap.has(key)) sessionMap.set(key, { date: e.date, name: e.workoutName, entries: [] });
    sessionMap.get(key).entries.push(e);
  }

  // Sort sessions chronologically
  const sessions = [...sessionMap.values()].sort((a, b) => {
    const da = new Date(a.date), db = new Date(b.date);
    return da - db;
  });

  // Group sessions into weeks by ISO week boundary (Monday-based)
  function isoWeekKey(dateStr) {
    const d = new Date(dateStr);
    if (isNaN(d)) return 'unknown';
    // Adjust to Monday-based week
    const day = d.getDay() || 7; // Sunday = 7
    const monday = new Date(d);
    monday.setDate(d.getDate() - day + 1);
    return monday.toISOString().slice(0, 10);
  }

  const weekMap = new Map();
  for (const sess of sessions) {
    const wk = isoWeekKey(sess.date);
    if (!weekMap.has(wk)) weekMap.set(wk, []);
    weekMap.get(wk).push(sess);
  }

  // Build block structure
  const weeks = [];
  let weekNum = 0;
  for (const [, weekSessions] of weekMap) {
    weekNum++;
    const days = [];
    for (const sess of weekSessions) {
      // Group entries by exercise name (preserve order)
      const exMap = new Map();
      for (const e of sess.entries) {
        if (!exMap.has(e.exerciseName)) exMap.set(e.exerciseName, []);
        exMap.get(e.exerciseName).push(e);
      }

      const exercises = [];
      for (const [exName, sets] of exMap) {
        // Build prescription from actual data: collapse identical sets
        const setParts = [];
        let runWeight = null, runReps = null, runCount = 0;
        for (const s of sets) {
          if (s.weight === runWeight && s.reps === runReps) {
            runCount++;
          } else {
            if (runCount > 0) {
              const hasWeight = runWeight && runWeight !== '0';
              setParts.push(hasWeight ? `${runCount}x${runReps}(${runWeight})` : `${runCount}x${runReps}`);
            }
            runWeight = s.weight;
            runReps = s.reps;
            runCount = 1;
          }
        }
        if (runCount > 0) {
          const hasWeight = runWeight && runWeight !== '0';
          setParts.push(hasWeight ? `${runCount}x${runReps}(${runWeight})` : `${runCount}x${runReps}`);
        }

        const prescription = setParts.join(', ');
        // Build enriched note with supplementary data (rest time, equipment)
        const _noteParts = [];
        if (sets[0].notes) _noteParts.push(sets[0].notes);
        const _restVal = sets.find(s => s.restTime && s.restTime.trim() && s.restTime !== '0')?.restTime;
        if (_restVal) {
          const _sec = parseFloat(_restVal);
          if (!isNaN(_sec) && _sec > 0) {
            _noteParts.push(_sec >= 120 ? 'Rest: ' + Math.round(_sec / 60) + 'm' : 'Rest: ' + Math.round(_sec) + 's');
          }
        }
        if (sets[0].equipment) _noteParts.push('Equipment: ' + sets[0].equipment);
        const note = _noteParts.length > 0 ? _noteParts.join(' | ') : null;

        exercises.push({
          name: exName,
          prescription: prescription || null,
          note: note,
          // Store raw set data for log population
          _csvSets: sets.map(s => ({ weight: s.weight, reps: s.reps, rpe: s.rpe })),
          _csvDate: sess.date
        });
      }

      days.push({ name: sess.name || sess.date, exercises });
    }
    weeks.push({ label: 'Week ' + weekNum, days });
  }

  if (!weeks.length) return null;

  // Determine date range
  const firstDate = sessions[0].date;
  const lastDate = sessions[sessions.length - 1].date;

  const ts = Date.now();
  const source = fmt === 'strong' ? 'Strong' : fmt === 'fitnotes' ? 'FitNotes' : fmt === 'fitbod' ? 'Fitbod' : 'Hevy';
  const blockName = source + ' Import';

  return [{
    id: 'csv_' + ts,
    name: blockName,
    format: 'CSV',
    athleteName: '',
    dateRange: firstDate + ' – ' + lastDate,
    startDate: firstDate,
    maxes: {},
    historical: true,
    _csvUnit: unitLabel || 'lbs',
    weeks
  }];
}

// ── FORMAT I: TEXAS METHOD ────────────────────────────────────────────────────

function detectTexasMethodFmt(wb) {
  if (wb.SheetNames.length > 2) return false;
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    // Must be a standalone "Texas Method" cell (not embedded in a longer description)
    // AND the same row must have Mon/Wed/Fri day columns
    const hasTM = row.some(c => c && /^texas\s*method$/i.test(String(c).trim()));
    if (!hasTM) continue;
    const dayHits = row.filter(c => c && /^(mon|wed|fri)/i.test(String(c).trim()));
    if (dayHits.length >= 3) return true;
  }
  return false;
}

function parseTexasMethod(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

  // Extract 1RM maxes from top area (col H=7 has lift names, col I=8 has weights)
  const maxes = {};
  for (let i = 0; i < 10; i++) {
    const row = rows[i]; if (!row) continue;
    const label = row[7], weight = row[8];
    if (label && typeof weight === 'number' && weight > 0) {
      const n = String(label).trim().toLowerCase();
      if (n === 'squat') maxes.squat = weight;
      else if (n === 'bench') maxes.bench = weight;
      else if (/^(press|ohp)$/i.test(n)) maxes.press = weight;
      else if (n === 'deadlift') maxes.deadlift = weight;
      else if (n === 'clean') maxes.clean = weight;
    }
  }

  // Find all "Texas Method" header rows — each starts a 3-week block
  const blockStarts = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    if (row.some(c => c && /^texas\s*method$/i.test(String(c).trim()))) blockStarts.push(i);
  }
  if (!blockStarts.length) return null;

  const DAY_NAMES = ['Monday', 'Wednesday', 'Friday'];
  const allWeeks = [];

  for (let bi = 0; bi < blockStarts.length; bi++) {
    const bStart = blockStarts[bi];
    const headerRow = rows[bStart]; if (!headerRow) continue;

    // Find first day column (Mon)
    let firstDayCol = -1;
    for (let c = 0; c < headerRow.length; c++) {
      if (headerRow[c] && /^(mon|tue|wed|thu|fri|sat|sun)/i.test(String(headerRow[c]).trim())) {
        firstDayCol = c; break;
      }
    }
    if (firstDayCol < 0) continue;

    // Count how many day columns (groups of 3 = weeks)
    let dayColCount = 0;
    for (let c = firstDayCol; c < headerRow.length; c++) {
      if (headerRow[c] && /^(mon|tue|wed|thu|fri|sat|sun)/i.test(String(headerRow[c]).trim())) dayColCount++;
      else if (headerRow[c] == null || String(headerRow[c]).trim() === '') continue;
      else break;
    }
    const numWeeks = Math.floor(dayColCount / 3) || 1;

    // Read SxR scheme row (bStart + 2): e.g. "5x5", "2x5", "1x5"
    const sxrRow = rows[bStart + 2] || [];

    // Initialize weeks with empty days
    const weekDays = [];
    for (let w = 0; w < numWeeks; w++) {
      const days = [];
      for (let d = 0; d < 3; d++) {
        days.push({
          name: DAY_NAMES[d],
          colIdx: firstDayCol + w * 3 + d,
          sxr: String(sxrRow[firstDayCol + w * 3 + d] || '').trim(),
          exercises: []
        });
      }
      weekDays.push(days);
    }

    // Parse exercise blocks from bStart+3 until VOLUME or next block
    const bEnd = (bi + 1 < blockStarts.length) ? blockStarts[bi + 1] : rows.length;
    let curExercise = null;

    for (let i = bStart + 3; i < bEnd; i++) {
      const row = rows[i]; if (!row) continue;
      const col2 = row[2] ? String(row[2]).trim() : '';
      const col3 = row[3] ? String(row[3]).trim() : '';

      // VOLUME row = end of exercises
      if (/^volume$/i.test(col2)) break;

      // New exercise name in col 2
      if (col2 && /[a-zA-Z]{2,}/.test(col2) && col2 !== curExercise) {
        curExercise = col2;
      }
      if (!curExercise) continue;

      // Only capture "Work Sets" rows
      if (col3.toLowerCase() !== 'work sets') continue;

      for (let w = 0; w < numWeeks; w++) {
        for (let d = 0; d < 3; d++) {
          const colIdx = firstDayCol + w * 3 + d;
          const val = row[colIdx];
          if (val == null || val === '') continue;

          const valStr = String(val).trim();
          const daySxr = weekDays[w][d].sxr;
          let prescription;

          // If value is a SxR pattern itself (e.g. "5x10", "3xF"), use as-is
          if (/^\d+x[\dF]/i.test(valStr)) {
            prescription = valStr;
          } else if (typeof val === 'number' || /^[\d.]+$/.test(valStr)) {
            // Numeric weight — combine with day's SxR scheme
            const wt = typeof val === 'number' ? Math.round(val * 10) / 10 : valStr;
            if (/^\d+x\d/.test(daySxr)) {
              prescription = `${daySxr}(${wt})`;
            } else {
              prescription = String(wt);
            }
          } else {
            prescription = valStr;
          }

          weekDays[w][d].exercises.push({
            name: _hCapitalizeName(curExercise),
            prescription,
            note: '',
            lifterNote: '',
            loggedWeight: '',
            supersetGroup: null
          });
        }
      }
      curExercise = null; // Done with this exercise's work sets
    }

    // Add to allWeeks
    for (let w = 0; w < numWeeks; w++) {
      const filteredDays = weekDays[w].filter(d => d.exercises.length > 0).map(d => ({
        name: d.name,
        exercises: d.exercises
      }));
      if (filteredDays.length > 0) {
        allWeeks.push({ label: `Week ${allWeeks.length + 1}`, days: filteredDays });
      }
    }
  }

  if (!allWeeks.length) return null;
  return [{
    id: 'texas_' + Date.now(),
    name: 'Texas Method',
    format: 'I',
    maxes,
    weeks: allWeeks
  }];
}


// ── FORMAT J: HEPBURN ────────────────────────────────────────────────────────

function detectHepburnFmt(wb) {
  const names = wb.SheetNames.map(s => s.toLowerCase());
  return names.some(n => /1rm\s*input/i.test(n)) &&
         names.some(n => /power\s*phase/i.test(n) || /pump\s*phase/i.test(n));
}

function parseHepburn(wb) {
  const SKIP = /^(notes|1rm\s*input|readme|info|instructions?)$/i;

  // Extract maxes from "1RM Input" sheet
  const maxes = {};
  const maxSheet = wb.SheetNames.find(s => /1rm\s*input/i.test(s));
  if (maxSheet) {
    const mRows = XLSX.utils.sheet_to_json(wb.Sheets[maxSheet], {header:1, defval:null});
    for (let i = 0; i < mRows.length; i++) {
      const row = mRows[i]; if (!row) continue;
      const label = row[0] ? String(row[0]).trim().toLowerCase() : '';
      const val = row[1];
      if (typeof val === 'number' && val > 0) {
        if (/^squat/.test(label)) maxes.squat = maxes.squat || val;
        else if (/^bench/.test(label)) maxes.bench = maxes.bench || val;
        else if (/^dead/.test(label)) maxes.deadlift = maxes.deadlift || val;
        else if (/^ohp|^press/.test(label)) maxes.press = maxes.press || val;
      }
    }
  }

  const blocks = [];

  for (const sn of wb.SheetNames) {
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

    const weeks = [];
    let curWeekNum = 0, curWorkoutNum = 0;
    let curDayExercises = null;
    let curWeek = null;

    const flushDay = () => {
      if (curDayExercises && curDayExercises.length > 0) {
        if (!curWeek) curWeek = { label: `Week ${curWeekNum || 1}`, days: [] };
        curWeek.days.push({ name: `Workout ${curWorkoutNum}`, exercises: curDayExercises });
      }
      curDayExercises = null;
    };

    const flushWeek = () => {
      flushDay();
      if (curWeek && curWeek.days.length > 0) {
        weeks.push(curWeek);
      }
      curWeek = null;
    };

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const col0 = row[0] ? String(row[0]).trim() : '';
      const col1 = row[1];

      // Week marker
      if (/^week$/i.test(col0) && typeof col1 === 'number') {
        const newWeekNum = col1;
        if (newWeekNum !== curWeekNum) {
          flushWeek();
          curWeekNum = newWeekNum;
        }
        continue;
      }

      // Workout marker
      if (/^workout$/i.test(col0) && typeof col1 === 'number') {
        flushDay();
        curWorkoutNum = col1;
        curDayExercises = [];
        continue;
      }

      // Header row (Sets/Reps/Weight) — skip
      if (row[3] && /^sets$/i.test(String(row[3]).trim()) &&
          row[4] && /^reps$/i.test(String(row[4]).trim())) continue;

      // Exercise row: col 2 = exercise name, col 3 = sets, col 4 = reps, col 5 = weight
      if (curDayExercises !== null && row[2]) {
        const exName = String(row[2]).trim();
        if (!exName || /^optional/i.test(exName)) continue;
        if (!/[a-zA-Z]{2,}/.test(exName)) continue;

        const sets = typeof row[3] === 'number' ? row[3] : null;
        const reps = typeof row[4] === 'number' ? row[4] : null;
        const weight = typeof row[5] === 'number' ? row[5] : null;

        if (sets && reps) {
          const corrected = _hCapitalizeName(exName);
          // Check if last exercise in current day has the same name — merge sets
          const last = curDayExercises.length > 0 ? curDayExercises[curDayExercises.length - 1] : null;
          if (last && last.name === corrected) {
            // Append to prescription
            const wt = weight ? `(${Math.round(weight * 10) / 10})` : '';
            last.prescription += `, ${sets}x${reps}${wt}`;
          } else {
            const wt = weight ? `(${Math.round(weight * 10) / 10})` : '';
            curDayExercises.push({
              name: corrected,
              prescription: `${sets}x${reps}${wt}`,
              note: ''
            });
          }
        }
      }
    }

    // Flush remaining
    flushWeek();

    if (weeks.length > 0) {
      blocks.push({
        id: 'hepburn_' + Date.now() + '_' + blocks.length,
        name: sn,
        format: 'J',
        maxes,
        weeks
      });
    }
  }

  return blocks.length > 0 ? blocks : null;
}


// ── FORMAT K: BULGARIAN METHOD ───────────────────────────────────────────────

function detectBulgarianFmt(wb) {
  const names = wb.SheetNames.map(s => s.toLowerCase());
  if (!names.some(n => n.includes('start here'))) return false;
  // Check for (A) markers in any training sheet
  for (const sn of wb.SheetNames) {
    if (/start\s*here|bare\s*minimum/i.test(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      if (row[0] && /^\(\s*[A-E]\s*\)$/i.test(String(row[0]).trim())) return true;
    }
  }
  return false;
}

function parseBulgarianMethod(wb) {
  const SKIP = /^(start\s*here|the\s*bare\s*minimum|readme|notes|info)$/i;

  // Extract 1RM from training sheets (rows 1-2 typically have 1RM values)
  const maxes = {};
  const blocks = [];

  for (const sn of wb.SheetNames) {
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

    if (rows.length < 5) continue;

    // Extract 1RM from this sheet (look for "1RM" row with Squat/Bench/Deadlift values)
    let sheetMaxes = {};
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const col1 = row[1] ? String(row[1]).trim().toLowerCase() : '';
      if (/1rm|1\s*rep\s*max/i.test(col1)) {
        // Next cells should be lift maxes — look at header row above for lift names
        const headerRow = rows[i - 1];
        if (headerRow) {
          for (let c = 2; c < Math.min(8, row.length); c++) {
            const liftName = headerRow[c] ? String(headerRow[c]).trim().toLowerCase() : '';
            const val = row[c];
            if (typeof val === 'number' && val > 0) {
              if (/squat/.test(liftName)) sheetMaxes.squat = val;
              else if (/bench/.test(liftName)) sheetMaxes.bench = val;
              else if (/dead/.test(liftName)) sheetMaxes.deadlift = val;
            }
          }
        }
      }
    }
    Object.assign(maxes, sheetMaxes);

    // Find week sections
    const weekStarts = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < Math.min(3, row.length); c++) {
        if (row[c] && /^week\s*[\d_]/i.test(String(row[c]).trim())) {
          weekStarts.push(i);
          break;
        }
      }
    }
    if (weekStarts.length === 0) continue;

    const weeks = [];

    for (let wi = 0; wi < weekStarts.length; wi++) {
      const wStart = weekStarts[wi];
      const wEnd = (wi + 1 < weekStarts.length) ? weekStarts[wi + 1] : rows.length;
      const weekRow = rows[wStart];

      // Detect day columns from week header row
      // Look for "Day N" cells to find column positions
      const dayColumns = []; // {col, dayNum, label}

      // Check for split layout (Frequency First): two groups side by side
      let isSplit = false;
      const weekLabels = [];
      for (let c = 0; c < (weekRow ? weekRow.length : 0); c++) {
        const v = weekRow[c] ? String(weekRow[c]).trim() : '';
        if (/^week/i.test(v)) weekLabels.push(c);
        if (/^day\s*\d/i.test(v)) {
          const dayNum = parseInt(v.match(/\d+/)[0]);
          dayColumns.push({ col: c, dayNum, label: `Day ${dayNum}` });
        }
      }
      if (weekLabels.length >= 2) isSplit = true;

      if (dayColumns.length === 0) continue;

      // Find reps column(s) — the column with "Reps" header in the row below week header
      const repsRow = rows[wStart + 1];
      const repsColumns = []; // col indices where "Reps" appears
      if (repsRow) {
        for (let c = 0; c < repsRow.length; c++) {
          if (repsRow[c] && /^reps$/i.test(String(repsRow[c]).trim())) repsColumns.push(c);
        }
      }

      // For split layout, there are two reps columns. Map each day to its reps column.
      // For simple layout, one reps column for all days.
      const dayToRepsCol = {};
      if (repsColumns.length === 1) {
        dayColumns.forEach(dc => { dayToRepsCol[dc.dayNum] = repsColumns[0]; });
      } else if (repsColumns.length >= 2 && isSplit) {
        // Left group days use first reps col, right group days use second
        // Sort reps columns by position
        const sorted = [...repsColumns].sort((a, b) => a - b);
        for (const dc of dayColumns) {
          // Assign to nearest reps column that is to the left
          let best = sorted[0];
          for (const rc of sorted) {
            if (rc < dc.col) best = rc;
          }
          dayToRepsCol[dc.dayNum] = best;
        }
      }

      // Parse exercise blocks within this week
      // Structure: (A) marker in col 0 or col 7 (split), exercise name next col, then rows of sets
      const dayExercises = {}; // dayNum -> [{name, sets: [{reps, weight}]}]
      dayColumns.forEach(dc => { dayExercises[dc.dayNum] = []; });
      const dayContext = {}; // dayNum -> { bw, form, sessionNotes }

      let curExName = null;
      let curExSets = {}; // dayNum -> [{reps, weight}]

      const flushExercise = () => {
        if (!curExName) return;
        for (const dc of dayColumns) {
          const sets = curExSets[dc.dayNum] || [];
          if (sets.length === 0) continue;

          // Collapse consecutive identical sets
          const collapsed = [];
          let runReps = null, runWeight = null, runCount = 0;
          for (const s of sets) {
            if (s.reps === runReps && s.weight === runWeight) {
              runCount++;
            } else {
              if (runCount > 0) collapsed.push({ count: runCount, reps: runReps, weight: runWeight });
              runReps = s.reps; runWeight = s.weight; runCount = 1;
            }
          }
          if (runCount > 0) collapsed.push({ count: runCount, reps: runReps, weight: runWeight });

          const parts = collapsed.map(c => {
            const wt = (c.weight != null && c.weight !== '' && c.weight !== 0) ? `(${c.weight})` : '';
            return `${c.count}x${c.reps}${wt}`;
          });

          // Build note from daily context (bodyweight, form quality, session notes)
          const _ctx = dayContext[dc.dayNum];
          const _ctxParts = [];
          if (_ctx) {
            if (_ctx.bw) _ctxParts.push('BW: ' + _ctx.bw);
            if (_ctx.form) _ctxParts.push('Form: ' + _ctx.form);
            if (_ctx.sessionNotes) _ctxParts.push(_ctx.sessionNotes);
          }
          dayExercises[dc.dayNum].push({
            name: _hCapitalizeName(curExName),
            prescription: parts.join(', '),
            note: _ctxParts.length > 0 ? _ctxParts.join(' | ') : ''
          });
        }
        curExName = null;
        curExSets = {};
      };

      for (let i = wStart + 2; i < wEnd; i++) {
        const row = rows[i]; if (!row) continue;

        // Check for (A)/(B)/(C) group markers in col 0 or col 7 (split right group)
        const col0 = row[0] ? String(row[0]).trim() : '';
        let groupMarkerCol = -1;
        if (/^\(\s*[A-Z]\s*\)$/i.test(col0)) groupMarkerCol = 0;
        // Split layout: right group marker might be in another column
        for (let c = 5; c < Math.min(10, row.length); c++) {
          if (row[c] && /^\(\s*[A-Z]\s*\)$/i.test(String(row[c]).trim())) {
            // This starts a new exercise in the right group
            // Handle: flush left group exercise, parse right group
            groupMarkerCol = c;
          }
        }

        // "Your typical warmup" = warmup separator, skip
        const col1 = row[1] ? String(row[1]).trim() : '';
        if (/warmup/i.test(col1)) { continue; }

        // Detect daily context labels (bodyweight, form quality, session notes)
        const _col1Lower = col1.toLowerCase();
        if (/^(body\s*weight|bw|current\s*weight)$/.test(_col1Lower)) {
          for (const dc of dayColumns) {
            const val = row[dc.col];
            if (typeof val === 'number' && val > 0) {
              if (!dayContext[dc.dayNum]) dayContext[dc.dayNum] = {};
              dayContext[dc.dayNum].bw = val;
            }
          }
          continue;
        }
        if (/^(form|quality|rating)$/.test(_col1Lower)) {
          for (const dc of dayColumns) {
            const val = row[dc.col];
            if (val != null && String(val).trim()) {
              if (!dayContext[dc.dayNum]) dayContext[dc.dayNum] = {};
              dayContext[dc.dayNum].form = String(val).trim();
            }
          }
          continue;
        }
        if (/^(notes?|session\s*notes?|comments?)$/.test(_col1Lower)) {
          for (const dc of dayColumns) {
            const val = row[dc.col];
            if (val != null && String(val).trim()) {
              if (!dayContext[dc.dayNum]) dayContext[dc.dayNum] = {};
              dayContext[dc.dayNum].sessionNotes = String(val).trim();
            }
          }
          continue;
        }

        // "(Optional)" or "Daily Min" / "Daily Max" labels
        if (/^\(optional\)|daily\s*(min|max)/i.test(col1)) {
          // These are placeholder rows — include if they have a rep count
          // (they indicate the program structure even without predetermined weights)
        }

        if (groupMarkerCol >= 0) {
          // New exercise — flush previous
          flushExercise();
          const nameCol = groupMarkerCol + 1;
          curExName = row[nameCol] ? String(row[nameCol]).trim() : null;
          if (!curExName || !/[a-zA-Z]{2,}/.test(curExName)) { curExName = null; continue; }
          curExSets = {};

          // Read first set from this same row
          for (const dc of dayColumns) {
            const repsCol = dayToRepsCol[dc.dayNum];
            // Find which reps column is associated with this group
            let actualRepsCol = repsCol;
            if (isSplit) {
              // For split: if day col is in left group, use left reps col; right group, use right reps col
              if (repsColumns.length >= 2) {
                const sorted = [...repsColumns].sort((a, b) => a - b);
                actualRepsCol = dc.col > sorted[sorted.length - 1] ? sorted[sorted.length - 1] :
                               dc.col > sorted[0] ? sorted[sorted.length - 1] : sorted[0];
                // Simpler: use the reps column closest to (and before) this day column
                for (const rc of sorted) {
                  if (rc <= dc.col) actualRepsCol = rc;
                }
              }
            }
            const reps = row[actualRepsCol];
            const weight = row[dc.col];
            if (reps == null) continue;
            const repsStr = String(reps).trim();
            const weightStr = weight != null ? String(weight).trim() : '';

            // Skip rest days
            if (/^x$/i.test(weightStr) || /^x\s*=/i.test(weightStr)) continue;
            // Skip if reps is not numeric or range
            if (!/^\d/.test(repsStr)) continue;

            if (!curExSets[dc.dayNum]) curExSets[dc.dayNum] = [];

            // Handle percentage text like "70-85% of 1RM"
            if (/\d+.*%/.test(weightStr)) {
              curExSets[dc.dayNum].push({ reps: repsStr, weight: weightStr });
            } else if (typeof weight === 'number') {
              curExSets[dc.dayNum].push({ reps: repsStr, weight: Math.round(weight * 10) / 10 });
            } else if (weightStr && /^\d/.test(weightStr)) {
              curExSets[dc.dayNum].push({ reps: repsStr, weight: weightStr });
            } else if (weightStr === '') {
              // No weight (placeholder for daily min/max)
              curExSets[dc.dayNum].push({ reps: repsStr, weight: '' });
            }
          }
          continue;
        }

        // Continuation rows for current exercise (no group marker)
        if (curExName) {
          for (const dc of dayColumns) {
            let actualRepsCol = dayToRepsCol[dc.dayNum] || repsColumns[0];
            if (isSplit && repsColumns.length >= 2) {
              const sorted = [...repsColumns].sort((a, b) => a - b);
              for (const rc of sorted) {
                if (rc <= dc.col) actualRepsCol = rc;
              }
            }
            const reps = row[actualRepsCol];
            const weight = row[dc.col];
            if (reps == null) continue;
            const repsStr = String(reps).trim();
            const weightStr = weight != null ? String(weight).trim() : '';
            if (/^x$/i.test(weightStr) || /^x\s*=/i.test(weightStr)) continue;
            if (!/^\d/.test(repsStr)) continue;

            if (!curExSets[dc.dayNum]) curExSets[dc.dayNum] = [];
            if (/\d+.*%/.test(weightStr)) {
              curExSets[dc.dayNum].push({ reps: repsStr, weight: weightStr });
            } else if (typeof weight === 'number') {
              curExSets[dc.dayNum].push({ reps: repsStr, weight: Math.round(weight * 10) / 10 });
            } else if (weightStr && /^\d/.test(weightStr)) {
              curExSets[dc.dayNum].push({ reps: repsStr, weight: weightStr });
            } else if (weightStr === '') {
              curExSets[dc.dayNum].push({ reps: repsStr, weight: '' });
            }
          }
        }
      }
      flushExercise();

      // Build week with sorted days
      const dayNums = Object.keys(dayExercises).map(Number).sort((a, b) => a - b);
      const weekDays = [];
      for (const dn of dayNums) {
        if (dayExercises[dn].length > 0) {
          weekDays.push({ name: `Day ${dn}`, exercises: dayExercises[dn] });
        }
      }
      if (weekDays.length > 0) {
        weeks.push({ label: `Week ${weeks.length + 1}`, days: weekDays });
      }
    }

    if (weeks.length > 0) {
      blocks.push({
        id: 'bulgarian_' + Date.now() + '_' + blocks.length,
        name: sn,
        format: 'K',
        maxes: Object.keys(sheetMaxes).length > 0 ? sheetMaxes : maxes,
        weeks
      });
    }
  }

  return blocks.length > 0 ? blocks : null;
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPER FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

function getCellValue(ws, cellAddr) {
  try {
    if (!ws[cellAddr]) return null;
    const cell = ws[cellAddr];
    if (cell.v instanceof Date) return cell.w || String(cell.v);
    return cell.v !== undefined ? cell.v : null;
  } catch (e) {
    return null;
  }
}

function parseSetsReps(text) {
  if (!text) return { sets: null, reps: null };
  const str = String(text).toLowerCase().trim();
  const match = str.match(/(\d+)\s*[*x×]\s*(\d+)/);
  if (match) {
    return { sets: parseInt(match[1]), reps: parseInt(match[2]) };
  }
  return { sets: null, reps: null };
}

function buildPrescription(sets, reps, weight) {
  if (sets === null || reps === null || weight === null) {
    return null;
  }
  return `${sets}x${reps}(${Math.round(weight)})`;
}

function buildNote(percentage, additionalText) {
  const parts = [];
  if (percentage !== null && percentage !== undefined) {
    parts.push(`${percentage}%`);
  }
  if (additionalText && additionalText.trim()) {
    parts.push(additionalText.trim());
  }
  return parts.length > 0 ? parts.join(' · ') : '';
}

function _hCapitalizeName(name) {
  return name
    .split(/\s+/)
    .map(word => {
      if (word.length <= 2) return word;
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

// ─────────────────────────────────────────────────────────────────────────────
// SUB-PARSER 1: H-COAN (Coan-Phillipi Deadlift)
// ─────────────────────────────────────────────────────────────────────────────

function parseH_coan(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      const block = parseH_coanSheet(sn, ws);
      if (block) blocks.push(block);
    } catch (e) {
      console.error(`[parseH_coan] Error parsing sheet "${sn}":`, e);
    }
  }
  return blocks;
}

function parseH_coanSheet(sn, ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 20) return null;

  // Detect Coan structure: "Exercise", "Week N" headers + "X" patterns
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const row = rows[i];
    if (!row) continue;
    const b = String(row[1] || '').toLowerCase().trim();
    const c = String(row[2] || '').toLowerCase().trim();
    if (b === 'exercise' && /week\s*\d/i.test(c)) {
      headerRowIdx = i;
      break;
    }
  }

  if (headerRowIdx === -1) return null;

  // Extract maxes from E4, E5 (row 3, 4 in 0-indexed)
  let maxes = {};
  const e4 = getCellValue(ws, 'E4');
  const e5 = getCellValue(ws, 'E5');
  if (e4 !== null) maxes.deadlift = e4;
  if (e5 !== null) maxes.current_1rm = e5;

  // Parse weeks in pairs
  const weeks = [];
  let weekNum = 1;

  for (let i = headerRowIdx; i < rows.length; i++) {
    const hdrRow = rows[i];
    if (!hdrRow) continue;

    const b = String(hdrRow[1] || '').toLowerCase().trim();
    const c = String(hdrRow[2] || '').toLowerCase().trim();

    if (b === 'exercise' && /week\s*(\d+)/i.test(c)) {
      // Check if right side also has week header
      const g = String(hdrRow[6] || '').toLowerCase().trim();
      const h = String(hdrRow[7] || '').toLowerCase().trim();

      if (g === 'exercise' && /week\s*(\d+)/i.test(h)) {
        // Parse left side (cols B-D)
        const dataRow1 = rows[i + 1];
        const dataRow2 = rows[i + 2];

        if (dataRow1) {
          const leftEx = String(dataRow1[1] || '').trim();
          const leftWt = getCellValue(ws, `C${i + 2}`);
          const leftPres = String(dataRow1[3] || '').trim();

          if (leftEx && leftWt !== null) {
            const { sets, reps } = parseCoagPrescription(leftPres);
            const { percentage, restTime } = extractCoagPercent(leftPres);

            weeks.push({
              label: `Week ${weekNum}`,
              days: [{
                name: 'Day 1',
                exercises: [{
                  name: _hCapitalizeName(leftEx),
                  prescription: buildPrescription(sets || 1, reps, leftWt),
                  note: buildNote(percentage, restTime),
                  lifterNote: '',
                  loggedWeight: ''
                }]
              }]
            });
            weekNum++;
          }

          // Backoff set in row i+2
          if (dataRow2) {
            const backoffWt = getCellValue(ws, `C${i + 3}`);
            const backoffPres = String(dataRow2[3] || '').trim();

            if (backoffWt !== null && weeks.length > 0) {
              const { percentage: bp, restTime: br } = extractCoagPercent(backoffPres);
              const { sets: bSets, reps: bReps } = parseCoagPrescription(backoffPres);

              weeks[weeks.length - 1].days[0].exercises.push({
                name: 'Deadlift',
                prescription: buildPrescription(bSets, bReps, backoffWt),
                note: buildNote(bp, br),
                lifterNote: '',
                loggedWeight: '',
                supersetGroup: null
              });
            }
          }
        }

        // Parse right side (cols G-I)
        if (dataRow1) {
          const rightEx = String(dataRow1[6] || '').trim();
          const rightWt = getCellValue(ws, `H${i + 2}`);
          const rightPres = String(dataRow1[8] || '').trim();

          if (rightEx && rightWt !== null) {
            const { sets, reps } = parseCoagPrescription(rightPres);
            const { percentage, restTime } = extractCoagPercent(rightPres);

            weeks.push({
              label: `Week ${weekNum}`,
              days: [{
                name: 'Day 1',
                exercises: [{
                  name: _hCapitalizeName(rightEx),
                  prescription: buildPrescription(sets || 1, reps, rightWt),
                  note: buildNote(percentage, restTime),
                  lifterNote: '',
                  loggedWeight: ''
                }]
              }]
            });
            weekNum++;
          }

          // Backoff set
          if (dataRow2) {
            const backoffWt = getCellValue(ws, `H${i + 3}`);
            const backoffPres = String(dataRow2[8] || '').trim();

            if (backoffWt !== null && weeks.length > 0) {
              const { percentage: bp, restTime: br } = extractCoagPercent(backoffPres);
              const { sets: bSets, reps: bReps } = parseCoagPrescription(backoffPres);

              weeks[weeks.length - 1].days[0].exercises.push({
                name: 'Deadlift',
                prescription: buildPrescription(bSets, bReps, backoffWt),
                note: buildNote(bp, br),
                lifterNote: '',
                loggedWeight: '',
                supersetGroup: null
              });
            }
          }
        }
      }
    }
  }

  if (weeks.length === 0) return null;

  return {
    id: `${sn}_${Date.now()}`,
    name: sn,
    format: 'H',
    athleteName: '',
    dateRange: '',
    maxes,
    weeks
  };
}

function parseCoagPrescription(text) {
  if (!text) return { sets: null, reps: null };
  const str = String(text).toLowerCase();

  // "X2 Reps" or just "X2" → sets=1, reps=2
  const m1 = str.match(/^x(\d+)(?:\s*reps?)?$/);
  if (m1) return { sets: 1, reps: parseInt(m1[1]) };

  // "X3 sets of 3 reps" → sets=3, reps=3
  const m2 = str.match(/x(\d+)\s+sets?\s+of\s+(\d+)/);
  if (m2) return { sets: parseInt(m2[1]), reps: parseInt(m2[2]) };

  // "8 sets of 3" → sets=8, reps=3
  const m3 = str.match(/(\d+)\s+sets?\s+of\s+(\d+)/);
  if (m3) return { sets: parseInt(m3[1]), reps: parseInt(m3[2]) };

  return { sets: null, reps: null };
}

function extractCoagPercent(text) {
  if (!text) return { percentage: null, restTime: null };
  const str = String(text).trim();

  let percentage = null;
  const pm = str.match(/(\d+(?:\.\d+)?)\s*%/);
  if (pm) percentage = parseFloat(pm[1]);

  let restTime = null;
  const rm = str.match(/(\d+(?:-\d+)?)\s*(?:sec|second)s?\.?\s*(?:rest)?/i);
  if (rm) restTime = rm[1] + ' sec rest';

  return { percentage, restTime };
}

// ─────────────────────────────────────────────────────────────────────────────
// SUB-PARSER 2: H-HATCH (Hatch Squat)
// ─────────────────────────────────────────────────────────────────────────────

function parseH_hatch(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      const block = parseH_hatchSheet(sn, ws);
      if (block) blocks.push(block);
    } catch (e) {
      console.error(`[parseH_hatch] Error parsing sheet "${sn}":`, e);
    }
  }
  return blocks;
}

function parseH_hatchSheet(sn, ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 50) return null;

  // Detect Hatch: "back sqt" or "front sqt" in first 10 rows
  let hasBackSquat = false;
  let hasFrontSquat = false;

  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const row = rows[i];
    if (!row) continue;
    for (let j = 0; j < Math.min(8, row.length); j++) {
      const val = String(row[j] || '').toLowerCase().trim();
      if (val.includes('back') && val.includes('sqt')) hasBackSquat = true;
      if (val.includes('front') && val.includes('sqt')) hasFrontSquat = true;
    }
  }

  if (!hasBackSquat && !hasFrontSquat) return null;

  // Extract maxes from B4, C4
  let maxes = {};
  const b4 = getCellValue(ws, 'B4');
  const c4 = getCellValue(ws, 'C4');
  if (b4 !== null) maxes.squat = b4;
  if (c4 !== null) maxes.front_squat = c4;

  // Parse weeks
  const weeks = [];
  let currentWeek = null;
  let currentDay = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const a = String(row[0] || '').trim();

    if (/^Week\s+\d/i.test(a)) {
      if (currentWeek && currentWeek.days.length > 0) {
        weeks.push(currentWeek);
      }
      currentWeek = { label: a, days: [] };
      continue;
    }

    if (/^Day\s+\d/i.test(a)) {
      if (currentDay && currentDay.exercises.length > 0 && currentWeek) {
        currentWeek.days.push(currentDay);
      }
      currentDay = { name: a, exercises: [] };
      continue;
    }

    if (currentDay && currentWeek) {
      const bSetsReps = String(row[1] || '').trim();
      const bPercent = row[2];
      const bWt = getCellValue(ws, `D${i + 1}`);

      const eSetsReps = String(row[4] || '').trim();
      const ePercent = row[5];
      const eWt = getCellValue(ws, `G${i + 1}`);

      if (bSetsReps && bWt !== null) {
        const { sets, reps } = parseSetsReps(bSetsReps);
        if (sets && reps) {
          currentDay.exercises.push({
            name: 'Back Squat',
            prescription: buildPrescription(sets, reps, bWt),
            note: bPercent ? `${Math.round(bPercent * 100)}%` : '',
            lifterNote: '',
            loggedWeight: ''
          });
        }
      }

      if (eSetsReps && eWt !== null) {
        const { sets, reps } = parseSetsReps(eSetsReps);
        if (sets && reps) {
          currentDay.exercises.push({
            name: 'Front Squat',
            prescription: buildPrescription(sets, reps, eWt),
            note: ePercent ? `${Math.round(ePercent * 100)}%` : '',
            lifterNote: '',
            loggedWeight: ''
          });
        }
      }
    }
  }

  if (currentDay && currentDay.exercises.length > 0 && currentWeek) {
    currentWeek.days.push(currentDay);
  }
  if (currentWeek && currentWeek.days.length > 0) {
    weeks.push(currentWeek);
  }

  if (weeks.length === 0) return null;

  return {
    id: `${sn}_${Date.now()}`,
    name: sn,
    format: 'H',
    athleteName: '',
    dateRange: '',
    maxes,
    weeks
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// SUB-PARSER 3: H-EDCOAN (Ed Coan Peaking)
// ─────────────────────────────────────────────────────────────────────────────

function parseH_edcoan(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      const block = parseH_edcoanSheet(sn, ws);
      if (block) blocks.push(block);
    } catch (e) {
      console.error(`[parseH_edcoan] Error parsing sheet "${sn}":`, e);
    }
  }
  return blocks;
}

function parseH_edcoanSheet(sn, ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 25) return null;

  // Detect Ed Coan: "Enter your current 1RM" label
  let hasPattern = false;
  for (let i = 0; i < Math.min(12, rows.length); i++) {
    const row = rows[i];
    if (!row) continue;
    const a = String(row[0] || '').toLowerCase();
    if (a.includes('enter') && a.includes('1rm')) {
      hasPattern = true;
      break;
    }
  }

  if (!hasPattern) return null;

  // Extract 1RM from D7 (row 6), D9 (row 8)
  let maxes = {};
  const d7 = getCellValue(ws, 'D7');
  const d9 = getCellValue(ws, 'D9');
  if (d7 !== null) maxes.current_1rm = d7;
  if (d9 !== null) maxes.projected_1rm = d9;

  // Data rows are at fixed indices: 13, 16, 19, 22 (0-indexed)
  // Each row contains 3 weeks of data (Week 1-3, 4-6, 7-9, 10-12)
  const dataRowIndices = [13, 16, 19, 22];
  const weeks = [];
  let weekNum = 1;

  for (const dataRowIdx of dataRowIndices) {
    const dataRow = rows[dataRowIdx];
    if (!dataRow) continue;

    // Each data row has 3 weeks: columns A-C (Week 1), E-G (Week 2), I-K (Week 3)
    const weekPositions = [
      { col: 0, weight: 1, reps: 2 }, // Week N: col A = "2 @", col B = weight, col C = "x reps"
      { col: 4, weight: 5, reps: 6 }, // Week N+1
      { col: 8, weight: 9, reps: 10 } // Week N+2
    ];

    for (const pos of weekPositions) {
      const setsText = String(dataRow[pos.col] || '').trim();
      const weight = getCellValue(ws, XLSX.utils.encode_cell({ r: dataRowIdx, c: pos.weight }));
      const repsText = String(dataRow[pos.reps] || '').trim();

      if (setsText && weight !== null && repsText) {
        const sm = setsText.match(/(\d+)\s*@/i);
        const rm = repsText.match(/x\s*(\d+)/i);

        if (sm && rm) {
          const sets = parseInt(sm[1]);
          const reps = parseInt(rm[1]);
          const prescription = buildPrescription(sets, reps, weight);

          weeks.push({
            label: `Week ${weekNum}`,
            days: [{
              name: 'Day 1',
              exercises: [{
                name: 'Main Lift',
                prescription: prescription,
                note: '',
                lifterNote: '',
                loggedWeight: ''
              }]
            }]
          });

          weekNum++;
        }
      }
    }
  }

  if (weeks.length === 0) return null;

  return {
    id: `${sn}_${Date.now()}`,
    name: sn,
    format: 'H',
    athleteName: '',
    dateRange: '',
    maxes,
    weeks
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// SUB-PARSER 4: H-HATFIELD (Fred Hatfield 12-Week)
// ─────────────────────────────────────────────────────────────────────────────

function parseH_hatfield(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      const block = parseH_hatfieldSheet(sn, ws);
      if (block) blocks.push(block);
    } catch (e) {
      console.error(`[parseH_hatfield] Error parsing sheet "${sn}":`, e);
    }
  }
  return blocks;
}

function parseH_hatfieldSheet(sn, ws) {
  if (!/hatfield/i.test(sn)) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 50) return null;

  // Extract maxes from C4, C5, C6 (row 3, 4, 5 in 0-indexed)
  let maxes = {};
  const c4 = getCellValue(ws, 'C4');
  const c5 = getCellValue(ws, 'C5');
  const c6 = getCellValue(ws, 'C6');
  if (c4 !== null) maxes.bench = c4;
  if (c5 !== null) maxes.squat = c5;
  if (c6 !== null) maxes.deadlift = c6;

  // Hatfield dual-column layout — scan for "lbs." anchor rows:
  // LEFT group:  col C (idx 2) = "lbs.", weights in D-H (idx 3-7)
  // RIGHT group: col J (idx 9) = "lbs.", weights in K-O (idx 10-14)
  // Look backwards from lbs. row for exercise name and day number
  // Look forwards for reps. row

  const COL_GROUPS = [
    { labelCol: 2, weightStart: 3 },   // LEFT: cols C-H
    { labelCol: 9, weightStart: 10 },   // RIGHT: cols J-O
  ];

  const entries = [];

  for (let i = 11; i < Math.min(205, rows.length); i++) {
    const row = rows[i];
    if (!row) continue;
    // Anchor on LEFT "lbs." — both groups share the same row
    if (String(row[2] || '').toLowerCase().trim() !== 'lbs.') continue;

    for (const cg of COL_GROUPS) {
      if (String(row[cg.labelCol] || '').toLowerCase().trim() !== 'lbs.') continue;

      const weights = [];
      for (let s = 0; s < 5; s++) {
        const w = row[cg.weightStart + s];
        if (typeof w === 'number') weights.push(w);
      }
      if (weights.length === 0) continue;

      // Find reps from next row
      let reps = null;
      for (let j = i + 1; j < Math.min(i + 3, rows.length); j++) {
        const nr = rows[j];
        if (!nr) continue;
        if (String(nr[cg.labelCol] || '').toLowerCase().trim() === 'reps.') {
          reps = nr[cg.weightStart];
          break;
        }
      }

      // Look backwards for exercise name and day number
      let exName = null, dayNum = null;
      for (let j = i - 1; j >= Math.max(i - 6, 10); j--) {
        const lr = rows[j];
        if (!lr) continue;
        const potEx = String(lr[cg.labelCol] || '').trim();
        if (!exName && /^(bench|squat|deadlift)/i.test(potEx)) exName = potEx;
        if (dayNum === null && typeof lr[1] === 'number') {
          const d = parseInt(lr[1]);
          if (d >= 1 && d <= 83) dayNum = d;
        }
        if (exName && dayNum !== null) break;
      }

      if (!exName) continue;

      entries.push({
        dayNum: dayNum || 0,
        exName,
        sets: weights.length,
        reps: reps !== null ? Math.round(reps) : 3,
        avgWeight: Math.round(weights.reduce((a, b) => a + b) / weights.length)
      });
    }
  }

  if (entries.length === 0) return null;

  // Group entries into weeks by day number
  const weeks = [];
  let currentWeek = { label: 'Week 1', days: [] };
  let currentDay = null;

  for (const e of entries) {
    const weekNum = e.dayNum > 0 ? Math.ceil(e.dayNum / 7) : (weeks.length + 1);

    if (weekNum > parseInt(currentWeek.label.match(/\d+/)[0])) {
      if (currentDay) { currentWeek.days.push(currentDay); currentDay = null; }
      if (currentWeek.days.length > 0) weeks.push(currentWeek);
      currentWeek = { label: `Week ${weekNum}`, days: [] };
    }

    const dayLabel = e.dayNum > 0 ? `Day ${e.dayNum}` : 'Day';
    if (!currentDay || currentDay.name !== dayLabel) {
      if (currentDay) currentWeek.days.push(currentDay);
      currentDay = { name: dayLabel, exercises: [] };
    }

    currentDay.exercises.push({
      name: _hCapitalizeName(e.exName),
      prescription: `${e.sets}x${e.reps}(${e.avgWeight})`,
      note: '',
      lifterNote: '',
      loggedWeight: ''
    });
  }

  if (currentDay) currentWeek.days.push(currentDay);
  if (currentWeek.days.length > 0) weeks.push(currentWeek);
  if (weeks.length === 0) return null;

  return {
    id: `${sn}_${Date.now()}`,
    name: sn,
    format: 'H',
    athleteName: '',
    dateRange: '',
    maxes,
    weeks
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// SUB-PARSER 5: H-DEATHBENCH (Disbrow's 10x3)
// ─────────────────────────────────────────────────────────────────────────────

function parseH_deathbench(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    try {
      const ws = wb.Sheets[sn];
      const block = parseH_deathbenchSheet(sn, ws);
      if (block) blocks.push(block);
    } catch (e) {
      console.error(`[parseH_deathbench] Error parsing sheet "${sn}":`, e);
    }
  }
  return blocks;
}

function _parseDeathbenchSide(rows, ws, cfg) {
  const weeks = [];
  for (let i = 3; i < Math.min(100, rows.length); i++) {
    const row = rows[i];
    if (!row) continue;
    const label = String(row[cfg.labelCol] || '').trim();

    if (/^Week\s+\d/i.test(label)) {
      weeks.push({ label, days: [] });
      continue;
    }

    if (/^Day\s+\d/i.test(label) && weeks.length > 0) {
      const day = { name: label, exercises: [] };
      const benchName = String(row[cfg.nameCol] || '').trim() || 'Bench';
      const benchSets = row[cfg.setsCol];
      const benchReps = row[cfg.repsCol];
      const benchWt = getCellValue(ws, `${cfg.weightCol}${i + 1}`);

      const benchParts = [];
      if (typeof benchSets === 'number' && typeof benchReps === 'number') {
        benchParts.push(benchWt !== null ? buildPrescription(benchSets, benchReps, benchWt) : `${benchSets}x${benchReps}`);
      }

      // Scan forward for bench continuation sets and accessories
      for (let j = i + 1; j < Math.min(i + 15, rows.length); j++) {
        const fwdRow = rows[j];
        if (!fwdRow) continue;
        const fwdLabel = String(fwdRow[cfg.labelCol] || '').trim();
        if (/^(Day|Week)\s+\d/i.test(fwdLabel)) break;

        const fwdName = String(fwdRow[cfg.nameCol] || '').trim();
        const fwdSets = fwdRow[cfg.setsCol];
        const fwdReps = fwdRow[cfg.repsCol];

        if (!fwdName && typeof fwdSets === 'number' && typeof fwdReps === 'number') {
          // Bench continuation set (no name, has numeric sets/reps)
          const fwdWt = getCellValue(ws, `${cfg.weightCol}${j + 1}`);
          benchParts.push(fwdWt !== null ? buildPrescription(fwdSets, fwdReps, fwdWt) : `${fwdSets}x${fwdReps}`);
        } else if (fwdName && !/^(Set|Reps|Weight)/i.test(fwdName) && typeof fwdSets === 'number') {
          // Accessory exercise
          const fwdWt = getCellValue(ws, `${cfg.weightCol}${j + 1}`);
          day.exercises.push({
            name: _hCapitalizeName(fwdName),
            prescription: fwdWt !== null ? buildPrescription(fwdSets, fwdReps || 0, fwdWt) : `${fwdSets}x${fwdReps || 0}`,
            note: '', lifterNote: '', loggedWeight: ''
          });
        }
      }

      if (benchParts.length > 0) {
        day.exercises.unshift({
          name: _hCapitalizeName(benchName),
          prescription: benchParts.join(', '),
          note: '', lifterNote: '', loggedWeight: ''
        });
      }

      if (day.exercises.length > 0) {
        weeks[weeks.length - 1].days.push(day);
      }
    }
  }
  return weeks;
}

function parseH_deathbenchSheet(sn, ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 50) return null;

  // Detect Deathbench
  const sheetMatch = /disbrow|10x3/i.test(sn);
  let hasOneRMLabel = false;
  if (rows[1] && String(rows[1][0] || '').toLowerCase().includes('1rm')) {
    hasOneRMLabel = true;
  }

  if (!sheetMatch && !hasOneRMLabel) return null;

  // Extract 1RM from B2
  let maxes = {};
  const b2 = getCellValue(ws, 'B2');
  if (b2 !== null) maxes.bench = b2;

  // Parse left side (weeks 1-5) and right side (weeks 6-10)
  const leftWeeks = _parseDeathbenchSide(rows, ws, {
    labelCol: 0, nameCol: 1, setsCol: 2, repsCol: 3, weightCol: 'E'
  });
  const rightWeeks = _parseDeathbenchSide(rows, ws, {
    labelCol: 6, nameCol: 7, setsCol: 8, repsCol: 9, weightCol: 'K'
  });

  const weeks = [...leftWeeks, ...rightWeeks].sort((a, b) => {
    const na = parseInt(a.label.match(/\d+/)[0]);
    const nb = parseInt(b.label.match(/\d+/)[0]);
    return na - nb;
  });

  if (weeks.length === 0) return null;

  return {
    id: `${sn}_${Date.now()}`,
    name: sn,
    format: 'H',
    athleteName: '',
    dateRange: '',
    maxes,
    weeks
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// DETECTION & MAIN PARSER
// ─────────────────────────────────────────────────────────────────────────────

function detectH(wb) {
  // H-deathbench: Sheet named "Disbrows" or containing "10x3"
  for (const sn of wb.SheetNames) {
    if (/disbrow/i.test(sn) || /\b10x3\b/i.test(sn)) return true;
  }

  // H-hatfield: Sheet named "Hatfield" AND has "One Rep Max" labels in first 6 rows
  for (const sn of wb.SheetNames) {
    if (/hatfield/i.test(sn)) {
      const ws = wb.Sheets[sn];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      for (let i = 0; i < Math.min(6, rows.length); i++) {
        const row = rows[i];
        if (!row) continue;
        if (row.some(c => c && /one\s*rep\s*max/i.test(String(c)))) return true;
      }
    }
  }

  // H-edcoan: "Ed Coan" in title area (rows 0-5) OR "Projected new 1RM" label
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i];
      if (!row) continue;
      const joined = row.map(c => String(c || '')).join(' ').toLowerCase();
      if (joined.includes('ed coan') && joined.includes('peaking')) return true;
      if (joined.includes('projected new 1rm')) return true;
    }
  }

  // H-hatch: "back sqt" AND "front sqt" BOTH appearing in the SAME row (header row)
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i];
      if (!row) continue;
      const vals = row.map(c => String(c || '').toLowerCase());
      const hasBackSqt = vals.some(v => v.includes('back') && v.includes('sqt'));
      const hasFrontSqt = vals.some(v => v.includes('front') && v.includes('sqt'));
      if (hasBackSqt && hasFrontSqt) return true;
    }
  }

  // H-coan: "Desired max" AND "Current Max" labels within first 10 rows of any sheet
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    let hasDesired = false, hasCurrent = false;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i];
      if (!row) continue;
      const joined = row.map(c => String(c || '')).join(' ').toLowerCase();
      if (joined.includes('desired max')) hasDesired = true;
      if (joined.includes('current max')) hasCurrent = true;
      if (hasDesired && hasCurrent) return true;
    }
  }

  return false;
}

function parseH(wb) {
  let result = parseH_edcoan(wb);
  if (result && result.length > 0) return result;

  result = parseH_hatch(wb);
  if (result && result.length > 0) return result;

  result = parseH_hatfield(wb);
  if (result && result.length > 0) return result;

  result = parseH_deathbench(wb);
  if (result && result.length > 0) return result;

  result = parseH_coan(wb);
  if (result && result.length > 0) return result;

  return [];
}

// ── FORMAT L PARSER (Candito Family) ─────────────────────────────────────────
// Handles: candito_6week, candito_6week_bench_hybrid, candito_advanced_bench,
//          candito_advanced_deadlift, candito_advanced_squat, candito_linear

function detectL(wb) {
  // Check for "Candito" in first 5 rows of any sheet
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const row = rows[i];
      if (!row) continue;
      const joined = row.map(c => String(c || '')).join(' ');
      if (/candito/i.test(joined)) return true;
    }
  }
  // Also detect linear by its distinctive sheet names + structure
  const snLower = wb.SheetNames.map(s => s.toLowerCase());
  if (snLower.some(s => s.includes('strength hypertrophy')) ||
      snLower.some(s => s.includes('strength control')) ||
      snLower.some(s => s.includes('strength power'))) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i];
      if (!row) continue;
      const a = String(row[0] || '').toLowerCase();
      if (/^(monday|tuesday|thursday|friday)$/.test(a)) return true;
    }
  }
  return false;
}

function parseL(wb) {
  const subType = _classifyCandito(wb);
  switch (subType) {
    case 'advanced_bench': return _parseL_advancedBench(wb);
    case 'advanced_deadlift': return _parseL_advancedDeadlift(wb);
    case 'advanced_squat': return _parseL_advancedSquat(wb);
    case 'bench_hybrid': return _parseL_benchHybrid(wb);
    case 'linear': return _parseL_linear(wb);
    case 'standard': return _parseL_standard(wb);
    default: return [];
  }
}

function _classifyCandito(wb) {
  const snLower = wb.SheetNames.map(s => s.toLowerCase());
  if (snLower.some(s => s === 'bench program')) {
    const ws = wb.Sheets[wb.SheetNames.find(s => s.toLowerCase() === 'bench program')];
    const cell = ws['A12'];
    if (cell && typeof cell.v === 'number') {
      const b16 = ws['B16'];
      if (b16 && b16.f && /SWITCH/i.test(b16.f)) return 'advanced_bench';
    }
  }
  if (snLower.some(s => /^phase\s*\d/.test(s))) return 'advanced_deadlift';
  if (snLower.includes('inputs') && snLower.some(s => /^week\s*\d/.test(s))) {
    const weekSn = wb.SheetNames.find(s => /^week\s*1$/i.test(s));
    if (weekSn) {
      const ws = wb.Sheets[weekSn];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const row = rows[i];
        if (!row) continue;
        const a = String(row[0] || '').trim();
        if (/^day\s*\d+\s*\(/i.test(a)) return 'advanced_squat';
      }
    }
  }
  if (!snLower.some(s => s === 'inputs') &&
      snLower.some(s => s.includes('strength'))) return 'linear';
  if (snLower.includes('inputs')) {
    for (const sn of wb.SheetNames) {
      if (!/^week\s*\d/i.test(sn)) continue;
      const ws = wb.Sheets[sn];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        const row = rows[i];
        if (!row) continue;
        if (row.some(c => String(c || '').toLowerCase() === 'rpe')) return 'bench_hybrid';
      }
      break;
    }
  }
  return 'standard';
}

// L helpers
function _lGetCellVal(ws, addr) {
  const c = ws[addr];
  if (!c) return null;
  return (c.v instanceof Date) ? (c.w || String(c.v)) : c.v;
}

function _lIsDateSerial(v) {
  return (v instanceof Date) || (typeof v === 'number' && v > EXCEL_DATE_SERIAL_MIN && v < 50000);
}

function _lParseReps(text) {
  if (!text) return null;
  const s = String(text).trim();
  const m = s.match(/^x(\d+(?:-\d+)?|MR)$/i);
  if (m) return m[1].toUpperCase() === 'MR' ? 'MR' : m[1];
  if (/^\d+$/.test(s)) return s;
  return s;
}

// SUB-PARSER: Standard (candito_6week)
function _parseL_standard(wb) {
  const inputsSn = wb.SheetNames.find(s => /^inputs$/i.test(s));
  if (!inputsSn) return [];
  const inputsWs = wb.Sheets[inputsSn];
  const maxes = {};
  let programName = '';
  const inputRows = XLSX.utils.sheet_to_json(inputsWs, { header: 1, defval: null });
  for (let i = 0; i < Math.min(25, inputRows.length); i++) {
    const row = inputRows[i];
    if (!row) continue;
    const a = String(row[0] || '').toLowerCase().trim();
    const b = row[1];
    if (/candito.*6\s*week\s*strength/i.test(String(row[0] || '') + ' ' + String(row[1] || '') + ' ' + String(row[2] || '')))
      programName = 'Candito 6 Week Strength';
    if (/candito.*9\s*week\s*squat/i.test(String(row[0] || '') + ' ' + String(row[1] || '') + ' ' + String(row[2] || '')))
      programName = 'Candito 9 Week Squat';
    if (a === 'bench press' && typeof b === 'number') maxes.bench = b;
    if (a === 'squat' && typeof b === 'number') maxes.squat = b;
    if (a === 'deadlift' && typeof b === 'number') maxes.deadlift = b;
    if (a === 'current squat max' && typeof b === 'number') maxes.squat = b;
    if (a === 'goal max' && typeof b === 'number') maxes['squat_goal'] = b;
  }
  if (!programName) programName = 'Candito Program';

  const weekSheets = wb.SheetNames.filter(s => /^week\s*\d+/i.test(s));
  weekSheets.sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));

  const weeks = [];
  for (const sn of weekSheets) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const weekNum = parseInt(sn.match(/\d+/)[0]);
    const week = { label: `Week ${weekNum}`, days: [] };
    let currentDay = null;
    let dayCounter = 0;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      if (_lIsDateSerial(row[0])) {
        if (currentDay) week.days.push(currentDay);
        dayCounter++;
        currentDay = { name: `Day ${dayCounter}`, exercises: [] };
        continue;
      }
      if (!currentDay) continue;
      const a = String(row[0] || '').trim();
      if (!a || /^set\s*\d/i.test(a)) continue;
      const exName = a;
      if (/^(week|complete|choose|what|do you|date|weights|also|to save)/i.test(exName)) continue;

      const setDefs = [];
      const repCols = [3, 5, 7, 9];
      const wtCols = [2, 4, 6, 8];
      for (let s = 0; s < repCols.length; s++) {
        const reps = row[repCols[s]];
        const wt = row[wtCols[s]];
        if (reps !== null && reps !== undefined && String(reps).trim() !== '') {
          setDefs.push({ weight: typeof wt === 'number' ? Math.round(wt) : null, reps: _lParseReps(reps) });
        }
      }
      if (setDefs.length === 0) continue;

      let prescription;
      const allSameWeight = setDefs.every(s => s.weight === setDefs[0].weight);
      const allSameReps = setDefs.every(s => s.reps === setDefs[0].reps);
      if (allSameWeight && allSameReps) {
        const wPart = setDefs[0].weight ? `(${setDefs[0].weight})` : '';
        prescription = `${setDefs.length}x${setDefs[0].reps}${wPart}`;
      } else {
        prescription = setDefs.map(s => {
          const wPart = s.weight ? `(${s.weight})` : '';
          return `1x${s.reps}${wPart}`;
        }).join(', ');
      }

      currentDay.exercises.push({
        name: exName, prescription: prescription,
        note: row[1] === 'Warm Up' ? 'Warm Up' : '', lifterNote: '', loggedWeight: '', supersetGroup: null
      });
    }
    if (currentDay) week.days.push(currentDay);
    if (week.days.length > 0) weeks.push(week);
  }
  if (weeks.length === 0) return [];
  return [{ id: `candito_standard_${Date.now()}`, name: programName, format: 'L',
    athleteName: '', dateRange: '', maxes: maxes, weeks: weeks }];
}

// SUB-PARSER: Bench Hybrid
function _parseL_benchHybrid(wb) {
  const inputsSn = wb.SheetNames.find(s => /^inputs$/i.test(s));
  if (!inputsSn) return [];
  const inputsWs = wb.Sheets[inputsSn];
  const inputRows = XLSX.utils.sheet_to_json(inputsWs, { header: 1, defval: null });
  const maxes = {};
  for (let i = 0; i < Math.min(30, inputRows.length); i++) {
    const row = inputRows[i];
    if (!row) continue;
    const a = String(row[0] || '').toLowerCase().trim();
    const b = row[1];
    if (a === 'squat' && typeof b === 'number') maxes.squat = b;
    if (a === 'deadlift' && typeof b === 'number') maxes.deadlift = b;
    if (a === 'bench' && typeof b === 'number') maxes.bench = b;
    if (a === 'desired bench max' && typeof b === 'number') maxes['bench_goal'] = b;
  }
  const weekSheets = wb.SheetNames.filter(s => /^week\s*\d+/i.test(s));
  const specialSheets = wb.SheetNames.filter(s => /^(projected|deload|taper)/i.test(s));
  const allSheets = [...weekSheets, ...specialSheets];
  allSheets.sort((a, b) => {
    const order = s => {
      const wm = s.match(/week\s*(\d+)/i);
      if (wm) return parseInt(wm[1]);
      if (/projected/i.test(s)) return 100;
      if (/deload/i.test(s)) return 101;
      if (/taper/i.test(s)) return 102;
      return 200;
    };
    return order(a) - order(b);
  });
  const weeks = [];
  for (const sn of allSheets) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const wm = sn.match(/week\s*(\d+)/i);
    const weekLabel = wm ? `Week ${wm[1]}` : sn;
    const week = { label: weekLabel, days: [] };
    let currentDay = null;
    let dayCounter = 0;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      if (_lIsDateSerial(row[0])) {
        if (currentDay) week.days.push(currentDay);
        dayCounter++;
        currentDay = { name: `Day ${dayCounter}`, exercises: [] };
        continue;
      }
      if (!currentDay) continue;
      const b = String(row[1] || '').trim();
      if (!b || b.toLowerCase() === 'lift') continue;
      const exName = b;
      if (/^(week|complete|note)/i.test(exName)) continue;
      const weight = typeof row[2] === 'number' ? Math.round(row[2]) : null;
      const rpe = String(row[3] || '').trim();
      const sets = typeof row[4] === 'number' ? Math.round(row[4]) : null;
      let reps = row[5];
      if (reps instanceof Date) reps = null;
      if (typeof reps === 'number' && reps > 1000) reps = null;
      if (typeof reps === 'number') reps = Math.round(reps);
      if (!sets && !reps) continue;
      const wPart = weight ? `(${weight})` : '';
      const rPart = reps || '?';
      const prescription = sets ? `${sets}x${rPart}${wPart}` : `1x${rPart}${wPart}`;
      const note = (rpe && rpe !== '-' && rpe !== '') ? rpe : '';
      currentDay.exercises.push({ name: exName, prescription: prescription, note: note, lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    if (currentDay) week.days.push(currentDay);
    if (week.days.length > 0) weeks.push(week);
  }
  if (weeks.length === 0) return [];
  return [{ id: `candito_bench_hybrid_${Date.now()}`, name: 'Candito Bench Hybrid Program', format: 'L',
    athleteName: '', dateRange: '', maxes: maxes, weeks: weeks }];
}

// SUB-PARSER: Advanced Bench (SWITCH formula evaluator)
function _parseL_advancedBench(wb) {
  const sn = wb.SheetNames.find(s => s.toLowerCase() === 'bench program');
  if (!sn) return [];
  const ws = wb.Sheets[sn];
  const unit = String(_lGetCellVal(ws, 'B5') || 'LB').toUpperCase();
  const currentRM = _lGetCellVal(ws, 'B6') || 0;
  const desiredRM = _lGetCellVal(ws, 'B7') || 0;
  const acc1 = String(_lGetCellVal(ws, 'B9') || 'Accessory #1');
  const acc2 = String(_lGetCellVal(ws, 'B10') || 'Accessory #2');
  const roundTo = unit === 'KG' ? 2.5 : 5;
  const maxes = { bench: currentRM, bench_goal: desiredRM };
  function mround(val, mult) { return Math.round(val / mult) * mult; }

  const weeks = [];
  for (let weekNum = 1; weekNum <= 6; weekNum++) {
    const week = { label: `Week ${weekNum}`, days: [] };

    // Day 1
    const day1 = { name: 'Day 1', exercises: [] };
    let r16name = weekNum === 6 ? 'High Pin Press' : 'Barbell Flat Bench';
    let r16weight;
    switch (weekNum) {
      case 1: r16weight = mround(0.7 * currentRM, roundTo); break;
      case 2: r16weight = mround(0.875 * currentRM, roundTo); break;
      case 3: r16weight = mround(0.9 * currentRM, roundTo); break;
      case 4: r16weight = `${mround(0.8 * desiredRM, roundTo)} - ${mround(0.825 * desiredRM, roundTo)}`; break;
      case 5: r16weight = `${mround(0.875 * desiredRM, roundTo)} - ${mround(0.9 * desiredRM, roundTo)}`; break;
      case 6: r16weight = mround(1.075 * desiredRM, roundTo); break;
    }
    const r16sets = weekNum === 4 ? 10 : 3;
    const r16reps = (weekNum === 2 || weekNum === 3) ? 1 : 3;
    const w16 = typeof r16weight === 'number' ? `(${r16weight})` : '';
    const n16 = typeof r16weight === 'string' ? r16weight : '';
    day1.exercises.push({ name: r16name, prescription: `${r16sets}x${r16reps}${w16}`, note: n16, lifterNote: '', loggedWeight: '', supersetGroup: null });

    // Day 1 secondary
    let r17name, r17weight = null, r17sets = null, r17reps;
    if (weekNum === 1) { r17name = acc1; r17sets = 3; r17reps = 8; }
    else if (weekNum === 2) { r17name = 'Barbell Flat Bench'; r17weight = mround(0.75 * currentRM, roundTo); r17sets = 3; r17reps = 3; }
    else if (weekNum === 3) { r17name = 'Barbell Flat Bench'; r17weight = mround(0.75 * currentRM, roundTo); r17sets = 3; r17reps = 3; }
    else if (weekNum === 4) { r17name = acc1; r17reps = '15 - 20'; }
    else { r17name = null; }
    if (r17name) {
      const w17 = typeof r17weight === 'number' ? `(${r17weight})` : '';
      const sPart = r17sets ? `${r17sets}x` : '';
      day1.exercises.push({ name: r17name, prescription: `${sPart}${r17reps}${w17}`,
        note: weekNum === 4 ? 'PR Peak Set' : '', lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    // Day 1 isolation
    if (weekNum !== 4) {
      const isoSets = weekNum < 4 ? 3 : 1;
      day1.exercises.push({ name: 'Isolation Accessory', prescription: `${isoSets}x20`, note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day1.exercises.push({ name: 'Isolation Accessory', prescription: `${isoSets}x20`, note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    if (day1.exercises.length > 0) week.days.push(day1);

    // Day 2
    const day2 = { name: 'Day 2', exercises: [] };
    let r22name, r22weight = null, r22sets = null, r22reps = null;
    if (weekNum === 6) { r22name = 'MAX OUT!!!!!'; }
    else if (weekNum === 5) { r22name = 'Barbell Flat Bench'; r22reps = '5RM'; }
    else if (weekNum === 4) {
      r22name = 'Barbell Flat Bench';
      r22weight = `${mround(0.8 * desiredRM + roundTo, roundTo)} - ${mround(0.825 * desiredRM + 3 * roundTo, roundTo)}`;
    } else {
      r22name = 'Barbell Flat Bench';
      switch (weekNum) {
        case 1: r22weight = mround(0.7 * currentRM, roundTo); break;
        case 2: r22weight = mround(0.75 * currentRM, roundTo); break;
        case 3: r22weight = `${mround(0.75 * currentRM, roundTo)} - ${mround(0.8 * currentRM, roundTo)}`; break;
      }
      r22sets = 3; r22reps = 3;
    }
    if (r22name) {
      const w22 = typeof r22weight === 'number' ? `(${r22weight})` : '';
      const n22 = typeof r22weight === 'string' ? r22weight : '';
      const sPart = r22sets ? `${r22sets}x` : '';
      const rPart = r22reps || '';
      day2.exercises.push({ name: r22name, prescription: `${sPart}${rPart}${w22}`.trim() || 'Max Attempt',
        note: n22 || (weekNum === 6 ? `Better be at least ${desiredRM}!` : ''), lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    // Day 2 secondary
    let r23name, r23reps;
    if (weekNum === 1) { r23name = acc2; r23reps = 6; }
    else if (weekNum === 2) { r23name = 'Low Pin Press'; r23reps = 6; }
    else if (weekNum === 3) { r23name = acc1; r23reps = 12; }
    else if (weekNum === 4) { r23name = acc2; r23reps = '15 - 20'; }
    else if (weekNum === 5) {
      r23name = 'Barbell Flat Bench';
      const w23 = mround(desiredRM * 0.85, roundTo);
      day2.exercises.push({ name: r23name, prescription: `5x3(${w23})`, note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      r23name = null;
    } else { r23name = null; }
    if (r23name) {
      const sPart23 = weekNum === 4 ? '' : '3x';
      day2.exercises.push({ name: r23name, prescription: `${sPart23}${r23reps}`,
        note: weekNum === 4 ? 'PR Peak Set' : '', lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    // Day 2 isolation
    if (weekNum < 4) {
      day2.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day2.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
    } else if (weekNum === 5) {
      day2.exercises.push({ name: 'Isolation Accessory', prescription: '1x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day2.exercises.push({ name: 'Isolation Accessory', prescription: '1x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    if (day2.exercises.length > 0) week.days.push(day2);

    // Days 3-5 (weeks 1-3 only)
    if (weekNum <= 3) {
      let baseWeight;
      switch (weekNum) {
        case 1: baseWeight = mround(0.7 * currentRM, roundTo); break;
        case 2: baseWeight = mround(0.75 * currentRM, roundTo); break;
        case 3: baseWeight = `${mround(0.75 * currentRM, roundTo)} - ${mround(0.8 * currentRM, roundTo)}`; break;
      }
      const bw = typeof baseWeight === 'number' ? `(${baseWeight})` : '';
      const bn = typeof baseWeight === 'string' ? baseWeight : '';

      const day3 = { name: 'Day 3', exercises: [] };
      day3.exercises.push({ name: 'Barbell Flat Bench', prescription: `3x3${bw}`, note: bn, lifterNote: '', loggedWeight: '', supersetGroup: null });
      let r29name, r29reps;
      if (weekNum === 1) { r29name = acc1; r29reps = 6; }
      else if (weekNum === 2) { r29name = 'High Pin Press'; r29reps = 6; }
      else { r29name = acc2; r29reps = 12; }
      day3.exercises.push({ name: r29name, prescription: `3x${r29reps}`, note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day3.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day3.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      week.days.push(day3);

      const day4 = { name: 'Day 4', exercises: [] };
      day4.exercises.push({ name: 'Barbell Flat Bench', prescription: `3x3${bw}`, note: bn, lifterNote: '', loggedWeight: '', supersetGroup: null });
      let r35name, r35reps;
      if (weekNum === 1) { r35name = acc2; r35reps = 3; }
      else if (weekNum === 2) { r35name = 'Low Pin Press'; r35reps = 3; }
      else { r35name = acc1; r35reps = 10; }
      day4.exercises.push({ name: r35name, prescription: `3x${r35reps}`, note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day4.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day4.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      week.days.push(day4);

      const day5 = { name: 'Day 5', exercises: [] };
      day5.exercises.push({ name: 'Barbell Flat Bench', prescription: `3x3${bw}`, note: bn, lifterNote: '', loggedWeight: '', supersetGroup: null });
      let r41name, r41reps;
      if (weekNum === 1) { r41name = acc1; r41reps = 3; }
      else if (weekNum === 2) { r41name = 'High Pin Press'; r41reps = 3; }
      else { r41name = acc2; r41reps = 10; }
      day5.exercises.push({ name: r41name, prescription: `3x${r41reps}`, note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day5.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      day5.exercises.push({ name: 'Isolation Accessory', prescription: '3x20', note: '', lifterNote: '', loggedWeight: '', supersetGroup: null });
      week.days.push(day5);
    }
    if (week.days.length > 0) weeks.push(week);
  }
  return [{ id: `candito_adv_bench_${Date.now()}`, name: 'Candito Advanced Bench Program', format: 'L',
    athleteName: '', dateRange: '', maxes: maxes, weeks: weeks }];
}

// SUB-PARSER: Advanced Deadlift
function _parseL_advancedDeadlift(wb) {
  const inputsSn = wb.SheetNames.find(s => /^inputs$/i.test(s));
  const maxes = {};
  if (inputsSn) {
    const inputsWs = wb.Sheets[inputsSn];
    const inputRows = XLSX.utils.sheet_to_json(inputsWs, { header: 1, defval: null });
    for (let i = 0; i < Math.min(25, inputRows.length); i++) {
      const row = inputRows[i];
      if (!row) continue;
      const b = String(row[1] || '').toLowerCase();
      const c = row[2];
      if (b.includes('close variation') && typeof c === 'number') maxes['close_variation'] = c;
      if (b.includes('5 rm') && typeof c === 'number') maxes['distant_5rm'] = c;
      if (b.includes('deadlift goal') && typeof c === 'number') maxes.deadlift = c;
    }
  }
  const phaseSheets = wb.SheetNames.filter(s => /^phase\s*\d/i.test(s));
  phaseSheets.sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
  const weeks = [];
  for (const sn of phaseSheets) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const phaseNum = parseInt(sn.match(/\d+/)[0]);
    let currentWeekLabel = null;
    let currentDay = null;
    let dayCounter = 0;
    const phaseDays = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const a = String(row[0] || '').trim();
      if (/^week\s*\d+/i.test(a)) { currentWeekLabel = a; dayCounter = 0; continue; }
      if (_lIsDateSerial(row[1])) {
        if (currentDay) phaseDays.push(currentDay);
        dayCounter++;
        currentDay = { name: `Day ${dayCounter}`, weekLabel: currentWeekLabel, exercises: [] };
      }
      if (!currentDay) continue;
      const bVal = String(row[1] || '').trim();
      if (!bVal || _lIsDateSerial(row[1])) {
        const c = row[2];
        if (c !== null && c !== undefined) {
          if (typeof c === 'number' && c > 100) {
            const e = String(row[4] || '').trim();
            const lastEx = currentDay.exercises[currentDay.exercises.length - 1];
            if (lastEx) {
              const setsReps = e.match(/(\d+)\s*sets?\s*x\s*(\d+)\s*reps?/i);
              if (setsReps) lastEx.prescription = `${setsReps[1]}x${setsReps[2]}(${Math.round(c)})`;
              else { lastEx.prescription = `1x1(${Math.round(c)})`; lastEx.note = e || lastEx.note; }
            }
          }
        }
        continue;
      }
      const exName = bVal;
      if (/^(phase|week|the program|to make|note)/i.test(exName)) continue;
      const c = row[2];
      const e = String(row[4] || '').trim();
      const g = String(row[6] || '').trim();
      const h = String(row[7] || '').trim();
      let prescription = '';
      let note = '';
      if (typeof c === 'number' && c > 50) {
        const setsReps = e.match(/(\d+)\s*sets?\s*x\s*(\d+)\s*reps?/i);
        if (setsReps) prescription = `${setsReps[1]}x${setsReps[2]}(${Math.round(c)})`;
        else {
          const simpleReps = e.match(/(\d+)\s*x\s*(\d+)/);
          if (simpleReps) prescription = `${simpleReps[1]}x${simpleReps[2]}(${Math.round(c)})`;
          else { prescription = `(${Math.round(c)})`; note = e; }
        }
      } else if (typeof c === 'string') {
        note = c;
        const hMatch = h.match(/(\d+)\s*sets?\s*x\s*(\d+)\s*reps?/i);
        if (hMatch) prescription = `${hMatch[1]}x${hMatch[2]}`;
        else prescription = e || c;
      } else { prescription = e || 'See instructions'; }
      if (!prescription) prescription = 'See program notes';
      currentDay.exercises.push({ name: exName, prescription: prescription, note: note, lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    if (currentDay) phaseDays.push(currentDay);
    let weekGroup = null;
    for (const day of phaseDays) {
      const wl = day.weekLabel || `Phase ${phaseNum}`;
      if (!weekGroup || weekGroup.label !== wl) {
        if (weekGroup && weekGroup.days.length > 0) weeks.push(weekGroup);
        weekGroup = { label: wl, days: [] };
      }
      weekGroup.days.push({ name: day.name, exercises: day.exercises });
    }
    if (weekGroup && weekGroup.days.length > 0) weeks.push(weekGroup);
  }
  if (weeks.length === 0) return [];
  return [{ id: `candito_adv_dl_${Date.now()}`, name: 'Candito Advanced Deadlift Program', format: 'L',
    athleteName: '', dateRange: '', maxes: maxes, weeks: weeks }];
}

// SUB-PARSER: Advanced Squat
function _parseL_advancedSquat(wb) {
  const inputsSn = wb.SheetNames.find(s => /^inputs$/i.test(s));
  if (!inputsSn) return [];
  const inputsWs = wb.Sheets[inputsSn];
  const inputRows = XLSX.utils.sheet_to_json(inputsWs, { header: 1, defval: null });
  const maxes = {};
  for (let i = 0; i < Math.min(20, inputRows.length); i++) {
    const row = inputRows[i];
    if (!row) continue;
    const a = String(row[0] || '').toLowerCase().trim();
    const b = row[1];
    if (a === 'current squat max' && typeof b === 'number') maxes.squat = b;
    if (a === 'goal max' && typeof b === 'number') maxes['squat_goal'] = b;
  }
  const weekSheets = wb.SheetNames.filter(s => /^week\s*\d+$/i.test(s));
  weekSheets.sort((a, b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
  const weeks = [];
  for (const sn of weekSheets) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const weekNum = parseInt(sn.match(/\d+/)[0]);
    const week = { label: `Week ${weekNum}`, days: [] };
    let currentDay = null;
    let setsCol = 3, repsCol = 4;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const a = String(row[0] || '').trim();
      const dayMatch = a.match(/^day\s*(\d+)/i);
      if (dayMatch) {
        if (currentDay) week.days.push(currentDay);
        currentDay = { name: a, exercises: [] };
        continue;
      }
      for (let c = 0; c < row.length; c++) {
        const v = String(row[c] || '').toLowerCase().trim();
        if (v === 'sets') setsCol = c;
        if (v === 'reps') repsCol = c;
      }
      if (!currentDay) continue;
      if (!a || /^(difficulty|weight|sets|reps|checkpoint|enter|percentage|is it|input|look in|leave gym)/i.test(a)) continue;
      const exName = a;
      let weight = row[2];
      let weightNote = '';
      if (typeof weight === 'string') { weightNote = weight; weight = null; }
      else if (typeof weight === 'number') { weight = Math.round(weight); }
      else { weight = null; }
      const dVal = String(row[3] || '');
      if (/^or$/i.test(dVal) && typeof row[4] === 'number') {
        weightNote = `${weight} or ${Math.round(row[4])}`;
      }
      const sets = typeof row[setsCol] === 'number' ? Math.round(row[setsCol]) : null;
      let reps = row[repsCol];
      if (typeof reps === 'number') reps = Math.round(reps);
      else if (typeof reps === 'string') reps = reps.trim();
      else reps = null;
      if (!sets && !reps) continue;
      const wPart = weight ? `(${weight})` : '';
      const sPart = sets || 1;
      const rPart = reps || '?';
      const prescription = `${sPart}x${rPart}${wPart}`;
      const diff = String(row[1] || '').trim();
      let note = weightNote;
      if (diff === 'H') note = (note ? note + ' | ' : '') + 'High Difficulty';
      else if (diff === 'M') note = (note ? note + ' | ' : '') + 'Moderate';
      else if (diff === 'L') note = (note ? note + ' | ' : '') + 'Light';
      currentDay.exercises.push({ name: exName, prescription: prescription, note: note, lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    if (currentDay) week.days.push(currentDay);
    if (week.days.length > 0) weeks.push(week);
  }
  if (weeks.length === 0) return [];
  return [{ id: `candito_adv_squat_${Date.now()}`, name: 'Candito 9 Week Squat Program', format: 'L',
    athleteName: '', dateRange: '', maxes: maxes, weeks: weeks }];
}

// SUB-PARSER: Linear
function _parseL_linear(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    if (rows.length < 10) continue;
    const hasDays = rows.some(r => r && /^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$/i.test(String(r[0] || '')));
    if (!hasDays) continue;
    const week = { label: 'Week 1', days: [] };
    let currentDay = null;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      const a = String(row[0] || '').trim();
      if (/^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$/i.test(a)) {
        if (currentDay) week.days.push(currentDay);
        currentDay = { name: a, exercises: [] };
        continue;
      }
      if (/^(heavy|hypertrophy|strength)/i.test(a) && row.some(c => /^set\s*\d/i.test(String(c || '')))) continue;
      if (!currentDay) continue;
      if (/^set\s*\d/i.test(a) || !a) continue;
      if (/^(to save|also|complete)/i.test(a)) continue;
      const exName = a;
      const repCols = [3, 5, 7, 9, 11];
      const repsFound = [];
      for (const rc of repCols) {
        const v = row[rc];
        if (v !== null && v !== undefined && String(v).trim() !== '') repsFound.push(_lParseReps(v));
      }
      if (repsFound.length === 0) continue;
      const allSame = repsFound.every(r => r === repsFound[0]);
      let prescription;
      if (allSame) prescription = `${repsFound.length}x${repsFound[0]}`;
      else prescription = repsFound.map(r => `1x${r}`).join(', ');
      currentDay.exercises.push({ name: exName, prescription: prescription,
        note: row[1] === 'Warm Up' ? 'Warm Up' : '', lifterNote: '', loggedWeight: '', supersetGroup: null });
    }
    if (currentDay) week.days.push(currentDay);
    if (week.days.length > 0) {
      blocks.push({ id: `candito_linear_${sn}_${Date.now()}`, name: `Candito Linear - ${sn}`, format: 'L',
        athleteName: '', dateRange: '', maxes: {}, weeks: [week] });
    }
  }
  return blocks;
}

// ── FORMAT M PARSER (Nuckols 28 Free Programs) ──────────────────────────────
// Handles: greg_nuckols_28_programs.xlsx (29 training sheets + 1 Maxes sheet)

function detectM(wb) {
  const hasMaxes = wb.SheetNames.some(s => /^maxes$/i.test(s.trim()));
  if (!hasMaxes) return false;
  const hasTraining = wb.SheetNames.some(s => /^(bench|squat|dl)\s+\dx/i.test(s.trim()));
  if (!hasTraining) return false;
  for (const sn of wb.SheetNames) {
    if (/^maxes$/i.test(sn.trim())) continue;
    if (!/^(bench|squat|dl)\s+\dx/i.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i];
      if (!row) continue;
      if (/^week\s*1$/i.test(String(row[0] || '').trim())) return true;
    }
    break;
  }
  return false;
}

function parseM(wb) {
  const maxesSn = wb.SheetNames.find(s => /^maxes$/i.test(s.trim()));
  const maxes = {};
  let rounding = 5;
  if (maxesSn) {
    const ws = wb.Sheets[maxesSn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i];
      if (!row) continue;
      const label = String(row[4] || '').toLowerCase().trim();
      const val = row[5];
      if (label === 'bench' && typeof val === 'number') maxes.bench = val;
      if (label === 'squat' && typeof val === 'number') maxes.squat = val;
      if (label === 'deadlift' && typeof val === 'number') maxes.deadlift = val;
      if (label === 'rounding' && typeof val === 'number') rounding = val;
    }
  }
  const blocks = [];
  for (const sn of wb.SheetNames) {
    if (/^maxes$/i.test(sn.trim())) continue;
    if (!/^(bench|squat|dl)\s+\dx/i.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    const headerRow = rows[0] || [];
    const dayColumns = [];
    dayColumns.push({ label: 'Day 1', exCol: 1, wtCol: 2, setsCol: 3, repsCol: 4, noteCol: -1 });
    if (headerRow[6] && /day\s*2/i.test(String(headerRow[6])))
      dayColumns.push({ label: 'Day 2', exCol: 6, wtCol: 7, setsCol: 8, repsCol: 9, noteCol: 10 });
    if (headerRow[11] && /day\s*3/i.test(String(headerRow[11])))
      dayColumns.push({ label: 'Day 3', exCol: 11, wtCol: 12, setsCol: 13, repsCol: 14, noteCol: -1 });
    const liftType = /^bench/i.test(sn) ? 'bench' : /^squat/i.test(sn) ? 'squat' : 'deadlift';
    const weeks = [];
    const weekStartRows = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row) continue;
      if (/^week\s*\d+$/i.test(String(row[0] || '').trim()))
        weekStartRows.push({ row: i, label: String(row[0]).trim() });
    }
    for (let wi = 0; wi < weekStartRows.length; wi++) {
      const weekStart = weekStartRows[wi].row;
      const weekEnd = wi + 1 < weekStartRows.length ? weekStartRows[wi + 1].row : rows.length;
      const week = { label: weekStartRows[wi].label, days: [] };
      for (const dc of dayColumns) {
        const day = { name: dc.label, exercises: [] };
        let currentExName = null;
        let setLines = [];
        const headerExRow = rows[weekStart];
        if (headerExRow) {
          const exName = String(headerExRow[dc.exCol] || '').trim();
          if (exName && !/^day\s*\d/i.test(exName)) currentExName = exName;
        }
        for (let ri = weekStart + 1; ri < weekEnd; ri++) {
          const row = rows[ri];
          if (!row) continue;
          const exVal = row[dc.exCol];
          const wtVal = row[dc.wtCol];
          const setsVal = row[dc.setsCol];
          const repsVal = row[dc.repsCol];
          const noteVal = dc.noteCol >= 0 ? row[dc.noteCol] : null;
          const exStr = String(exVal || '').trim();
          const isPercentage = typeof exVal === 'number' && exVal > 0 && exVal <= 1;
          const isNewExercise = exStr && !isPercentage && exStr !== 'Warm Up' &&
            !/^week\s*\d/i.test(exStr) && !/^\d+\.?\d*$/.test(exStr);
          if (isNewExercise) {
            if (currentExName && setLines.length > 0) {
              day.exercises.push(_mBuildExercise(currentExName, setLines));
              setLines = [];
            }
            currentExName = exStr;
            if (typeof setsVal === 'number' || repsVal != null) {
              setLines.push({
                pct: null, weight: typeof wtVal === 'number' ? wtVal : null,
                sets: typeof setsVal === 'number' ? Math.round(setsVal) : null,
                reps: _mFmtReps(repsVal),
                note: noteVal ? String(noteVal).trim() : ''
              });
            }
            continue;
          }
          if (isPercentage || typeof setsVal === 'number' || repsVal != null) {
            let weight = null;
            let wtNote = '';
            if (typeof wtVal === 'number' && wtVal > 1) { weight = Math.round(wtVal); }
            else if (typeof wtVal === 'string') {
              const wStr = wtVal.trim();
              if (/warm\s*up/i.test(wStr)) { if (setsVal == null && repsVal == null) continue; }
              else { wtNote = wStr; }
            }
            setLines.push({
              pct: isPercentage ? exVal : null, weight: weight,
              sets: typeof setsVal === 'number' ? Math.round(setsVal) : null,
              reps: _mFmtReps(repsVal),
              note: (noteVal ? String(noteVal).trim() : '') + (wtNote ? (noteVal ? ' ' : '') + wtNote : '')
            });
          }
        }
        if (currentExName && setLines.length > 0) {
          day.exercises.push(_mBuildExercise(currentExName, setLines));
        }
        if (day.exercises.length > 0) week.days.push(day);
      }
      if (week.days.length > 0) weeks.push(week);
    }
    if (weeks.length === 0) continue;
    blocks.push({
      id: `nuckols_${sn.replace(/\s+/g, '_')}_${Date.now()}`,
      name: `Nuckols ${sn}`, format: 'M', athleteName: '', dateRange: '',
      maxes: { [liftType]: maxes[liftType] || 0 }, weeks: weeks
    });
  }
  return blocks;
}

function _mFmtReps(val) {
  if (val == null) return null;
  if (typeof val === 'number') return String(Math.round(val));
  const s = String(val).trim();
  return s || null;
}

function _mBuildExercise(name, setLines) {
  const parts = [];
  const notes = [];
  for (const sl of setLines) {
    if (sl.note) notes.push(sl.note);
    if (!sl.sets && !sl.reps) continue;
    const sets = sl.sets || 1;
    const reps = sl.reps || '?';
    const wPart = sl.weight ? `(${sl.weight})` : '';
    parts.push(`${sets}x${reps}${wPart}`);
  }
  let prescription = '';
  if (parts.length === 0) prescription = '';
  else if (parts.length === 1) prescription = parts[0];
  else if (parts.every(p => p === parts[0])) {
    const m = parts[0].match(/^(\d+)x(.+)$/);
    if (m) prescription = `${parseInt(m[1]) * parts.length}x${m[2]}`;
    else prescription = parts.join(', ');
  } else prescription = parts.join(', ');
  return {
    name: name, prescription: prescription || 'See program notes',
    note: notes.filter(n => n).join(' | '), lifterNote: '', loggedWeight: '', supersetGroup: null
  };
}

// ── FORMAT N PARSER (Beginner LP: Greyskull, Ivysaur, Starting Strength, StrongLifts) ──
function detectN(wb) {
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());
  // Greyskull LP: has "Lift" + "Setup" sheets
  if (names.some(s => s === 'lift') && names.some(s => s === 'setup')) {
    const ws = wb.Sheets[wb.SheetNames[names.indexOf('lift')]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    if (rows[0] && String(rows[0][4]||'').match(/week\s*1/i)) return true;
  }
  // Ivysaur: has "Template" sheet with "4-4-8" or "ivysaur" in first few rows
  if (names.some(s => s === 'template')) {
    const ws = wb.Sheets[wb.SheetNames[names.indexOf('template')]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const txt = (rows[i]||[]).map(c => String(c||'')).join(' ').toLowerCase();
      if (txt.includes('4-4-8') || txt.includes('ivysaur') || txt.includes('4.4.8')) return true;
    }
  }
  // Starting Strength: sheet names contain "novice program" or "onus wunsler"
  if (names.some(s => s.includes('novice program') || s.includes('onus wunsler') || s.includes('wichita falls'))) {
    return true;
  }
  // StrongLifts 5x5: must have "stronglifts" AND "beginner" or "experienced" in sheet names
  // (Madcow also has "stronglifts" in sheet name but uses "advanced" instead)
  if (names.some(s => s.includes('stronglifts')) &&
      names.some(s => s.includes('beginner') || s.includes('experienced'))) {
    return true;
  }
  return false;
}

/* ---------- classify which sub-format -------------------------- */
function _nClassify(wb) {
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());
  if (names.some(s => s === 'lift') && names.some(s => s === 'setup')) return 'GS';
  if (names.some(s => s === 'template')) {
    const ws = wb.Sheets[wb.SheetNames[names.indexOf('template')]];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    for (let i = 0; i < 5; i++) {
      const txt = (rows[i]||[]).map(c => String(c||'')).join(' ').toLowerCase();
      if (txt.includes('4-4-8') || txt.includes('ivysaur') || txt.includes('4.4.8')) return 'IV';
    }
  }
  if (names.some(s => s.includes('novice program') || s.includes('onus wunsler') || s.includes('wichita falls'))) return 'SS';
  if (names.some(s => s.includes('stronglifts')) &&
      names.some(s => s.includes('beginner') || s.includes('experienced'))) return 'SL';
  return null;
}

/* ---------- router --------------------------------------------- */
function parseN(wb) {
  const sub = _nClassify(wb);
  switch(sub) {
    case 'GS': return _parseN_greyskull(wb);
    case 'IV': return _parseN_ivysaur(wb);
    case 'SS': return _parseN_startingStrength(wb);
    case 'SL': return _parseN_stronglifts(wb);
    default: return [];
  }
}

/* ---------- helpers -------------------------------------------- */
function _nRound(v) {
  return typeof v === 'number' ? Math.round(v * 100) / 100 : v;
}

function _nFmtReps(setsReps, weight) {
  // setsReps like "5x5", "3x5", "1x5", "5+/5+", "3xF", "3 sets to failure"
  const w = typeof weight === 'number' ? `(${_nRound(weight)})` : '';
  if (!setsReps) return w ? `1x1${w}` : null;
  const sr = String(setsReps).trim();
  // Already NxN format
  if (/^\d+x\d+/.test(sr)) return `${sr}${w}`;
  // "5+/5+" means AMRAP last set — we'll represent the set scheme
  if (/^\d+\+\/\d+\+/.test(sr)) return `${sr}${w}`;
  // "3 sets to failure"
  if (/sets?\s*to\s*fail/i.test(sr)) return `3xAMRAP${w}`;
  // "3xF"
  if (/^\d+xF$/i.test(sr)) return `${sr}${w}`;
  return `${sr}${w}`;
}

/* ================================================================
   N-GS: Greyskull LP
   ================================================================
   Lift sheet layout:
   R0: Week headers at cols 4,7,10,...  (3 cols per week)
   R1: Day 1/2/3 repeating
   R2: Program name
   Exercise groups start at rows 4, 14, 20, 30, 36
   Each group: name row, warmup rows, working weight row, rep rows
   ================================================================ */
function _parseN_greyskull(wb) {
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());
  const liftSheet = wb.Sheets[wb.SheetNames[names.indexOf('lift')]];
  const rows = XLSX.utils.sheet_to_json(liftSheet, {header:1, defval:null});

  // Determine weeks from R0
  const weekHeaders = [];
  const r0 = rows[0] || [];
  for (let c = 4; c < r0.length; c++) {
    const v = String(r0[c] || '').trim();
    if (/^week\s*\d+/i.test(v)) weekHeaders.push({col: c, label: v});
  }

  // Each week has 3 days (columns: weekCol, weekCol+1, weekCol+2)
  // Find exercise groups: scan for rows where col 0 has a group name like "Bench/OHP"
  const groups = [];
  for (let r = 3; r < Math.min(50, rows.length); r++) {
    const c0 = String(rows[r]?.[0] || '').trim();
    if (c0 && c0 !== 'Warm-Up' && !c0.startsWith('Working') && c0.length > 2) {
      groups.push(r);
    }
  }

  // For each group, identify:
  // - Name row (has alternating exercise names across day columns)
  // - Working weight row (explicit "Working Weight" label OR row with no col-2 value but weight-sized exercise col values)
  // - Rep scheme rows (col 2 has value, exercise cols have small numbers or are empty)
  // - Warmup rows (col 2 has value AND exercise cols have large weights) — skipped
  //
  // Key insight: warmup rows AND rep rows both have col 2 values.
  // Distinction: warmup rows have large values (weights) in exercise cols (>20),
  // rep rows have small values (logged reps <=20) or are empty.
  // Working weight row has NO col 2 value but HAS large exercise col values.
  const parsedGroups = [];
  for (let gi = 0; gi < groups.length; gi++) {
    const gRow = groups[gi];
    const nextGRow = gi + 1 < groups.length ? groups[gi + 1] : Math.min(gRow + 12, rows.length);

    const groupLabel = String(rows[gRow]?.[0] || '').trim();
    const nameRow = gRow;

    let workingRow = -1;
    let repRows = [];
    for (let r = gRow + 1; r < nextGRow; r++) {
      const c0 = String(rows[r]?.[0] || '').trim();
      const c2raw = rows[r]?.[2];
      const c2str = String(c2raw ?? '').trim();
      const hasC2 = c2raw !== null && c2raw !== undefined && c2str !== '';

      // Explicit "Working Weight" label — always trust it
      if (c0.includes('Working Weight')) {
        workingRow = r;
        continue;
      }

      // Check if any exercise column (4+) has a weight-sized number (>20)
      let maxExVal = 0;
      for (let c = 4; c < (rows[r]?.length || 0); c++) {
        const v = rows[r][c];
        if (typeof v === 'number' && v > maxExVal) maxExVal = v;
      }

      if (!hasC2 && maxExVal > 20) {
        // No col 2 value + large exercise col values → working weight row
        workingRow = r;
      } else if (hasC2 && maxExVal > 20) {
        // Has col 2 value + large exercise col values → warmup row (skip)
      } else if (hasC2 && workingRow >= 0) {
        // Has col 2 value + small/no exercise col values, after working weight → rep row
        repRows.push(r);
      }
    }

    parsedGroups.push({ groupLabel, nameRow, workingRow, repRows, nextGRow });
  }

  // Build weeks/days
  const weeks = [];
  for (const wh of weekHeaders) {
    const wLabel = wh.label;
    const days = [];
    for (let d = 0; d < 3; d++) {
      const dayCol = wh.col + d;
      const exercises = [];

      for (const pg of parsedGroups) {
        const exName = String(rows[pg.nameRow]?.[dayCol] || '').trim();
        if (!exName) continue;

        const weight = pg.workingRow >= 0 ? rows[pg.workingRow]?.[dayCol] : null;

        // Build prescription from rep rows
        // Col 2 values can be: "5/5", "5+/5+", plain number 5, or "5+"
        // For "N/N" format, left of slash = this exercise's reps
        // For plain numbers, the number IS the reps
        let prescParts = [];
        if (pg.repRows.length > 0) {
          const reps = pg.repRows.map(rr => {
            const c2raw = rows[rr]?.[2];
            if (typeof c2raw === 'number') return String(c2raw);
            const c2 = String(c2raw || '').trim();
            // "5/5" or "5+/5+" → take left of slash
            if (c2.includes('/')) return c2.split('/')[0].trim();
            return c2;
          });
          // Group identical rep values: e.g. [5, 5, 5+] → 2x5, 1x5+AMRAP
          const w = typeof weight === 'number' ? _nRound(weight) : null;
          const wStr = w !== null ? `(${w})` : '';
          let i = 0;
          while (i < reps.length) {
            let count = 1;
            while (i + count < reps.length && reps[i + count] === reps[i]) count++;
            const rep = reps[i];
            if (rep.endsWith('+')) {
              prescParts.push(`${count}x${rep.replace('+', '')}+AMRAP${wStr}`);
            } else {
              prescParts.push(`${count}x${rep}${wStr}`);
            }
            i += count;
          }
        } else if (typeof weight === 'number') {
          // No rep rows — accessories might just have weight
          prescParts.push(`(${_nRound(weight)})`);
        }

        const prescription = prescParts.join(', ') || null;
        exercises.push({ name: exName, prescription, note: null, lifterNote: null, loggedWeight: null, supersetGroup: null });
      }

      if (exercises.length > 0) {
        days.push({ name: `Day ${d + 1}`, exercises });
      }
    }
    weeks.push({ label: wLabel, days });
  }

  return [{
    id: 'greyskull_lp',
    name: 'Greyskull LP',
    format: 'N',
    athleteName: null,
    dateRange: null,
    maxes: {},
    weeks
  }];
}

/* ================================================================
   N-IV: Ivysaur 4-4-8
   ================================================================
   Template sheet layout:
   R17: Week headers (W1-W30)
   R18-R22: Starting weights per lift per week
   R24: "Week A - Weight estimates"
   R25: Odd week headers
   R26-29: Day 1 exercises, R32-36: Day 2, R38-41: Day 3
   R43: "Week B - Weight estimates"
   R44: Even week headers
   R45-48: Day 1, R51-54: Day 2, R57-60: Day 3

   Exercise names include set/rep: "Bench 4x4", "Squat 4x8"
   Weeks alternate: A(odd), B(even)
   ================================================================ */
function _parseN_ivysaur(wb) {
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());
  const ws = wb.Sheets[wb.SheetNames[names.indexOf('template')]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

  // Find Week A and Week B section starts
  let weekARow = -1, weekBRow = -1;
  for (let r = 20; r < Math.min(70, rows.length); r++) {
    const c1 = String(rows[r]?.[1] || '').trim().toLowerCase();
    if (c1.includes('week a')) weekARow = r;
    if (c1.includes('week b')) weekBRow = r;
  }
  if (weekARow === -1 || weekBRow === -1) return [];

  // Parse day blocks from a section
  // Each section: header row (week labels), then 3 day blocks separated by blank rows
  function parseSectionDays(startRow, endRow) {
    // First row after section header has week column labels
    const weekLabelRow = startRow + 1;
    const weekCols = [];
    const wlr = rows[weekLabelRow] || [];
    for (let c = 1; c < wlr.length; c++) {
      if (wlr[c] && /^W\d+$/i.test(String(wlr[c]).trim())) {
        weekCols.push({ col: c, label: String(wlr[c]).trim() });
      }
    }

    // Scan for exercise rows (non-empty col 1 with exercise name)
    const dayBlocks = []; // array of arrays of exercise rows
    let currentDay = [];
    for (let r = weekLabelRow + 1; r < endRow; r++) {
      const c1 = String(rows[r]?.[1] || '').trim();
      if (!c1) {
        if (currentDay.length > 0) {
          dayBlocks.push(currentDay);
          currentDay = [];
        }
        continue;
      }
      currentDay.push(r);
    }
    if (currentDay.length > 0) dayBlocks.push(currentDay);

    return { weekCols, dayBlocks };
  }

  const weekA = parseSectionDays(weekARow, weekBRow);
  const weekB = parseSectionDays(weekBRow, Math.min(weekBRow + 25, rows.length));

  // Build output: interleave Week A (odd) and Week B (even) weeks
  const totalWeeks = Math.max(
    weekA.weekCols.length > 0 ? parseInt(weekA.weekCols[weekA.weekCols.length-1].label.replace('W','')) : 0,
    weekB.weekCols.length > 0 ? parseInt(weekB.weekCols[weekB.weekCols.length-1].label.replace('W','')) : 0
  );

  const weeks = [];
  for (let w = 1; w <= totalWeeks; w++) {
    const isA = w % 2 === 1;
    const section = isA ? weekA : weekB;
    const wc = section.weekCols.find(wc => wc.label === `W${w}`);
    if (!wc) continue;

    const days = [];
    for (let di = 0; di < section.dayBlocks.length; di++) {
      const block = section.dayBlocks[di];
      const exercises = [];
      for (const exRow of block) {
        const rawName = String(rows[exRow]?.[1] || '').trim();
        if (!rawName) continue;

        // Parse exercise name + set/rep from name like "Bench 4x4" or "Chinups 4x8"
        const match = rawName.match(/^(.+?)\s+(\d+x\d+\+?\d*)/);
        let exName, setRep;
        if (match) {
          exName = match[1].trim();
          setRep = match[2];
        } else {
          exName = rawName;
          setRep = null;
        }

        const weight = rows[exRow]?.[wc.col];
        let prescription = null;
        if (setRep && typeof weight === 'number') {
          prescription = `${setRep}(${_nRound(weight)})`;
        } else if (setRep) {
          prescription = setRep; // e.g., Chinups with no weight
        } else if (typeof weight === 'number') {
          prescription = `(${_nRound(weight)})`;
        }

        exercises.push({ name: exName, prescription, note: null, lifterNote: null, loggedWeight: null, supersetGroup: null });
      }
      if (exercises.length > 0) {
        days.push({ name: `Day ${di + 1}`, exercises });
      }
    }

    weeks.push({ label: `Week ${w}`, days });
  }

  return [{
    id: 'ivysaur_448',
    name: 'Ivysaur 4-4-8',
    format: 'N',
    athleteName: null,
    dateRange: null,
    maxes: {},
    weeks
  }];
}

/* ================================================================
   N-SS: Starting Strength
   ================================================================
   Multiple sheets, each is a program variant.
   Each sheet:
   - "Workout A" / "Workout B" sections (or Monday/Wednesday/Friday)
   - "Session #N" columns
   - Exercises with warmup + working set rows
   - Col 1: exercise name, Col 2: warmup/working label, Col 3: SetsxReps
   - Cols 4+: weight per session

   We treat each session pair (A+B or Mon/Wed/Fri) as a "week" with 2-3 days
   ================================================================ */
function _parseN_startingStrength(wb) {
  const SKIP = /^(pws|read|app|progress|chart|image)/i;
  const blocks = [];

  for (const sn of wb.SheetNames) {
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    if (!rows || rows.length < 20) continue;

    // Find workout sections: "Workout A"/"Workout B" in col 0 OR "Monday"/"Wednesday"/"Friday"
    const sections = [];
    for (let r = 0; r < rows.length; r++) {
      const c0 = String(rows[r]?.[0] || '').trim();
      if (/^workout\s+[a-c]$/i.test(c0) || /^(monday|wednesday|friday)$/i.test(c0)) {
        // Find session columns from this row
        const sessionCols = [];
        const row = rows[r];
        for (let c = 4; c < (row ? row.length : 0); c++) {
          const v = String(row[c] || '').trim();
          if (/^session\s*#?\d+/i.test(v)) {
            const num = parseInt(v.match(/\d+/)[0]);
            sessionCols.push({ col: c, num });
          }
        }
        sections.push({ row: r, label: c0, sessionCols });
      }
    }
    if (sections.length === 0) continue;

    // For each section, parse exercises
    function parseSection(secIdx) {
      const sec = sections[secIdx];
      const endRow = secIdx + 1 < sections.length ? sections[secIdx + 1].row : rows.length;

      const exercises = []; // {name, warmups: [{setsReps, row}], working: {setsReps, row}}
      let currentEx = null;

      for (let r = sec.row + 1; r < endRow; r++) {
        const c1 = String(rows[r]?.[1] || '').trim();
        const c2 = String(rows[r]?.[2] || '').trim();
        const c3 = String(rows[r]?.[3] || '').trim();

        // New exercise name
        if (c1 && c1.length > 1 && !/^\(/.test(c1) && !/^warmup$/i.test(c1) && !/^working/i.test(c1)) {
          if (currentEx) exercises.push(currentEx);
          currentEx = { name: c1, warmups: [], working: null, row: r };
        }

        if (!currentEx) continue;

        // Working set
        if (/^working\s*set/i.test(c2)) {
          currentEx.working = { setsReps: c3, row: r };
        } else if (/^warmup$/i.test(c2)) {
          currentEx.warmups.push({ setsReps: c3, row: r });
        }

        // Special cases: "3 sets to failure", "3-5x10"
        if (/sets?\s*to\s*fail/i.test(c3) || /^\d+-?\d*x\d+/i.test(c3)) {
          if (/^working/i.test(c2)) {
            currentEx.working = { setsReps: c3, row: r };
          }
        }
      }
      if (currentEx) exercises.push(currentEx);

      return { ...sec, exercises };
    }

    const parsedSections = sections.map((_, i) => parseSection(i));

    // Determine session pairing: sessions come in pairs (A sessions = odd, B = even)
    // OR tripled (Mon=1,4,7, Wed=2,5,8, Fri=3,6,9)
    const allSessionCols = parsedSections.flatMap(s => s.sessionCols);
    const maxSession = Math.max(...allSessionCols.map(s => s.num), 0);
    const numSections = parsedSections.length;
    const sessionsPerWeek = numSections; // 2 for A/B, 3 for Mon/Wed/Fri
    const numWeeks = Math.ceil(maxSession / sessionsPerWeek);

    const weeks = [];
    for (let w = 0; w < numWeeks; w++) {
      const days = [];
      for (let d = 0; d < sessionsPerWeek; d++) {
        const sessionNum = w * sessionsPerWeek + d + 1;
        const sec = parsedSections[d];
        const sc = sec.sessionCols.find(s => s.num === sessionNum);
        if (!sc) continue;

        const exercises = [];
        for (const ex of sec.exercises) {
          // Only extract working sets (skip warmups)
          if (ex.working) {
            const weight = rows[ex.working.row]?.[sc.col];
            const prescription = _nFmtReps(ex.working.setsReps, weight);
            exercises.push({ name: ex.name, prescription, note: null, lifterNote: null, loggedWeight: null, supersetGroup: null });
          } else if (ex.name) {
            // Some exercises like "Back Extensions" have fixed prescription, no weight columns
            let fixedPrescription = null;
            if (ex.warmups.length === 0 && ex.working === null) {
              // Look for set/rep info at the exercise row
              const c2 = String(rows[ex.row]?.[2] || '').trim();
              const c3 = String(rows[ex.row]?.[3] || '').trim();
              if (c3) fixedPrescription = c3;
              else if (c2 && /working/i.test(c2)) {
                // Check next cell
              }
            }
            if (fixedPrescription) {
              exercises.push({ name: ex.name, prescription: fixedPrescription, note: null, lifterNote: null, loggedWeight: null, supersetGroup: null });
            }
          }
        }

        if (exercises.length > 0) {
          days.push({ name: sec.label, exercises });
        }
      }
      if (days.length > 0) {
        weeks.push({ label: `Week ${w + 1}`, days });
      }
    }

    blocks.push({
      id: sn.trim().toLowerCase().replace(/\s+/g, '_'),
      name: `Starting Strength - ${sn.trim()}`,
      format: 'N',
      athleteName: null,
      dateRange: null,
      maxes: {},
      weeks
    });
  }

  return blocks;
}

/* ================================================================
   N-SL: StrongLifts 5x5
   ================================================================
   Beginner sheet layout:
   R9:  Week headers at cols 0,3,6,...
   R10: Day labels (mo/we/fr)
   R13: Exercise 1 name (all "Squat")
   R14: Rep scheme (all "5x5")
   R16: Weight row
   R20: Exercise 2 names (alternating Bench/OHP)
   R21: Rep scheme
   R23: Weight row
   R27: Exercise 3 names (alternating Row/DL)
   R28: Rep scheme (5x5/1x5)
   R30: Weight row

   Experienced sheet is similar but starts at different rows:
   R12: Week headers
   R13: Day labels
   R16: Ex1 names, R17: rep, R19: weight
   R23: Ex2 names, R24: rep, R26: weight
   R30: Ex3 names, R31: rep, R33: weight
   ================================================================ */
function _parseN_stronglifts(wb) {
  const blocks = [];

  for (const sn of wb.SheetNames) {
    if (!/stronglifts/i.test(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    if (!rows || rows.length < 30) continue;

    const isBeginner = /beginner/i.test(sn);
    const variantName = isBeginner ? 'StrongLifts 5x5 (Beginner)' : 'StrongLifts 5x5 (Experienced)';

    // Find the week header row (has "Week1" in col 0)
    let weekRow = -1;
    for (let r = 5; r < 20; r++) {
      if (/^week\s*1$/i.test(String(rows[r]?.[0] || '').trim())) {
        weekRow = r;
        break;
      }
    }
    if (weekRow === -1) continue;

    // Week columns: Week1 at col 0, Week2 at col 3, Week3 at col 6, etc.
    const weekCols = [];
    for (let c = 0; c < (rows[weekRow] || []).length; c++) {
      const v = String(rows[weekRow][c] || '').trim();
      if (/^week\s*\d+$/i.test(v)) {
        weekCols.push({ col: c, label: v });
      }
    }

    // Day row is weekRow+1 (mo/we/fr labels)
    const dayRow = weekRow + 1;

    // Find exercise blocks: scan for rows where all columns have exercise names
    // Exercise name rows have text in multiple columns (Squat, Bench Press, etc.)
    const exBlocks = []; // {nameRow, repRow, weightRow}
    for (let r = weekRow + 2; r < Math.min(weekRow + 30, rows.length); r++) {
      const c0 = String(rows[r]?.[0] || '').trim();
      // Exercise name row: has known exercise names
      if (/^(squat|bench|oh\s*press|barbell\s*row|deadlift)/i.test(c0)) {
        const nameRow = r;
        // Next row should be rep scheme
        const repRow = r + 1;
        // Weight row: skip empty rows, find first row with numbers
        let weightRow = -1;
        for (let wr = repRow + 1; wr < Math.min(repRow + 4, rows.length); wr++) {
          if (typeof rows[wr]?.[0] === 'number') { weightRow = wr; break; }
        }
        if (weightRow >= 0) {
          exBlocks.push({ nameRow, repRow, weightRow });
        }
      }
    }

    // Build weeks: 3 days per week (cols offset 0,1,2 from week start)
    const weeks = [];
    for (const wc of weekCols) {
      const days = [];
      for (let d = 0; d < 3; d++) {
        const col = wc.col + d;
        const dayLabel = String(rows[dayRow]?.[col] || '').trim();
        const dayName = dayLabel === 'mo' ? 'Monday' : dayLabel === 'we' ? 'Wednesday' : dayLabel === 'fr' ? 'Friday' : `Day ${d+1}`;

        const exercises = [];
        for (const eb of exBlocks) {
          const exName = String(rows[eb.nameRow]?.[col] || '').trim();
          if (!exName) continue;
          const repScheme = String(rows[eb.repRow]?.[col] || '').trim();
          const weight = rows[eb.weightRow]?.[col];

          const prescription = _nFmtReps(repScheme, weight);
          exercises.push({ name: exName, prescription, note: null, lifterNote: null, loggedWeight: null, supersetGroup: null });
        }

        if (exercises.length > 0) {
          days.push({ name: dayName, exercises });
        }
      }
      weeks.push({ label: wc.label.replace(/(\d+)/, ' $1').trim(), days });
    }

    blocks.push({
      id: sn.trim().toLowerCase().replace(/\s+/g, '_'),
      name: variantName,
      format: 'N',
      athleteName: null,
      dateRange: null,
      maxes: {},
      weeks
    });
  }

  return blocks;
}

/* ---------- detection ------------------------------------------ */
function detectO(wb) {
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());
  if (!names.includes('training')) return false;

  const ws = wb.Sheets[wb.SheetNames[names.indexOf('training')]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
  if (!rows || rows.length < 20) return false;

  // Check R8 for header pattern: "sets", "reps", "intensity", "load" in cols 4-7
  const r8 = rows[8] || [];
  const h4 = String(r8[4] || '').trim().toLowerCase();
  const h5 = String(r8[5] || '').trim().toLowerCase();
  const h6 = String(r8[6] || '').trim().toLowerCase();
  const h7 = String(r8[7] || '').trim().toLowerCase();
  if (h4 !== 'sets' || h5 !== 'reps' || h6 !== 'intensity' || h7 !== 'load') return false;

  // Check for "Day 1" in col 3 of R8
  const d1 = String(r8[3] || '').trim().toLowerCase();
  if (!d1.startsWith('day')) return false;

  // Check for SQ/BN/DL exercise labels in col 0
  let hasSQ = false, hasBN = false;
  for (let r = 9; r < Math.min(40, rows.length); r++) {
    const c0 = String(rows[r]?.[0] || '').trim().toLowerCase();
    if (c0.startsWith('sq ')) hasSQ = true;
    if (c0.startsWith('bn ')) hasBN = true;
  }
  return hasSQ && hasBN;
}

/* ---------- parser --------------------------------------------- */
function parseO(wb) {
  const names = wb.SheetNames.map(s => s.trim().toLowerCase());
  const ws = wb.Sheets[wb.SheetNames[names.indexOf('training')]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

  // Determine number of weeks by scanning R2 for week numbers
  const weekStarts = []; // {col, weekNum, blockName}
  for (let c = 0; c < (rows[2]?.length || 0); c += 14) {
    const wNum = rows[2]?.[c + 3];
    if (typeof wNum === 'number' && wNum > 0) {
      const blockName = String(rows[1]?.[c + 3] || '').trim();
      weekStarts.push({ col: c, weekNum: wNum, blockName });
    }
  }
  if (weekStarts.length === 0) return [];

  // Find day boundaries by scanning col 3 (offset +3 from week 1 col 0)
  // Day markers appear at fixed rows for all weeks
  const dayRows = []; // {row, dayNum}
  for (let r = 8; r < rows.length; r++) {
    const c3 = String(rows[r]?.[3] || '').trim();
    const dm = c3.match(/^Day\s+(\d+)$/i);
    if (dm) dayRows.push({ row: r, dayNum: parseInt(dm[1]) });
  }
  if (dayRows.length === 0) return [];

  // Find exercise rows: rows where col 0 has content (exercise name/category)
  // Also track secondary rows (col 0 empty, col 3 = "-" with sets/reps)
  // Structure: between each pair of day markers

  // Extract maxes from Personal Info sheet if available
  const maxes = {};
  if (names.includes('personal info')) {
    const piWs = wb.Sheets[wb.SheetNames[names.indexOf('personal info')]];
    const piRows = XLSX.utils.sheet_to_json(piWs, {header:1, defval:null});
    for (let r = 0; r < Math.min(20, piRows.length); r++) {
      const label = String(piRows[r]?.[0] || '').trim().toLowerCase();
      const val = piRows[r]?.[1];
      if (typeof val === 'number') {
        if (label.includes('squat')) maxes.squat = val;
        else if (label.includes('bench')) maxes.bench = val;
        else if (label.includes('dead')) maxes.deadlift = val;
      }
    }
  }

  // Build weeks
  const weeks = [];
  for (const ws of weekStarts) {
    const wLabel = `Week ${ws.weekNum}`;
    const days = [];

    for (let di = 0; di < dayRows.length; di++) {
      const dayStart = dayRows[di].row;
      const dayEnd = di + 1 < dayRows.length ? dayRows[di + 1].row : rows.length;
      const dayName = `Day ${dayRows[di].dayNum}`;

      const exercises = [];
      let currentEx = null;

      for (let r = dayStart + 1; r < dayEnd; r++) {
        const base = ws.col;
        const exNameRaw = String(rows[r]?.[base + 0] || '').trim();
        const movement = String(rows[r]?.[base + 3] || '').trim();
        const setsRaw = rows[r]?.[base + 4];
        const repsRaw = rows[r]?.[base + 5];
        const intensityRaw = rows[r]?.[base + 6];
        const loadRaw = rows[r]?.[base + 7];

        // Normalize: treat empty strings as null for sets/reps
        const hasSets = setsRaw != null && setsRaw !== '' && setsRaw !== null;
        const hasReps = repsRaw != null && repsRaw !== '' && repsRaw !== null;

        // Skip empty rows
        if (!exNameRaw && !movement && !hasSets && !hasReps) continue;
        // Skip note rows
        if (movement.toLowerCase().startsWith('note:')) continue;

        // New exercise (has name in col +0)
        if (exNameRaw && exNameRaw.length > 1) {
          // Save previous (only if it has prescription data)
          if (currentEx && currentEx.lines.length > 0) {
            exercises.push(_oFinalize(currentEx));
          }
          currentEx = {
            name: movement && movement !== '-' && movement !== 'SELECT' ? movement : exNameRaw,
            category: exNameRaw,
            lines: []
          };
          // Add this row's data if it has sets/reps
          if (hasSets && hasReps) {
            currentEx.lines.push({ sets: setsRaw, reps: repsRaw, intensity: intensityRaw, load: loadRaw });
          }
        } else if (currentEx && movement === '-' && hasSets && hasReps) {
          // Secondary row for same exercise
          currentEx.lines.push({ sets: setsRaw, reps: repsRaw, intensity: intensityRaw, load: loadRaw });
        }
      }
      // Save last exercise (only if it has actual prescription data)
      if (currentEx && currentEx.lines.length > 0) exercises.push(_oFinalize(currentEx));

      if (exercises.length > 0) {
        days.push({ name: dayName, exercises });
      }
    }

    weeks.push({ label: wLabel, days });
  }

  return [{
    id: 'tsa_9week_intermediate',
    name: 'TSA 9-Week Intermediate',
    format: 'O',
    athleteName: null,
    dateRange: null,
    maxes,
    weeks
  }];
}

/* ---------- helpers -------------------------------------------- */
function _oFinalize(ex) {
  // Build prescription from lines
  const parts = [];
  const notes = [];

  for (const line of ex.lines) {
    const sets = line.sets;
    const reps = line.reps;
    const intensity = line.intensity;
    const load = line.load;

    // Format sets
    let setsStr;
    if (sets === 'x' || sets === 'X') {
      setsStr = '1'; // "x" means total reps in one go
    } else {
      setsStr = String(sets);
    }

    // Format reps (can be number or "N-N" range string)
    const repsStr = String(reps);

    // Format load and intensity
    let prescLine;
    if (typeof load === 'number' && load > 0) {
      const w = Math.round(load * 100) / 100;
      prescLine = `${setsStr}x${repsStr}(${w})`;
    } else {
      prescLine = `${setsStr}x${repsStr}`;
    }
    parts.push(prescLine);

    // Track intensity as note
    if (intensity != null) {
      const iStr = String(intensity).trim();
      if (iStr.startsWith('@')) {
        notes.push(`${prescLine} ${iStr}`);
      } else if (typeof intensity === 'number') {
        notes.push(`${prescLine} @${intensity}%`);
      }
    }
  }

  const prescription = parts.join(', ') || null;
  // Build note from intensity info
  const note = notes.length > 0 ? notes.join('; ') : null;

  return {
    name: ex.name,
    prescription,
    note,
    lifterNote: null,
    loggedWeight: null,
    supersetGroup: null
  };
}


// ── FORMAT P PARSER (RTS: Reactive Training Systems) ──
// Helper: Get cell value safely, handling errors and empty cells
function _pGetCell(sheet, row, col) {
  if (!sheet || row < 0 || col < 0) return null;
  const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
  if (!cell) return null;
  const val = (cell.v instanceof Date) ? (cell.w || String(cell.v)) : cell.v;
  if (val === undefined || val === null || val === '#N/A') return null;
  return val;
}

// Helper: Convert cell range to 2D array
function _pSheetToArray(sheet, maxRows = 1000) {
  const result = [];
  let emptyStreak = 0;
  for (let r = 0; r < maxRows; r++) {
    const row = [];
    for (let c = 0; c < 35; c++) {
      row.push(_pGetCell(sheet, r, c));
    }
    result.push(row);
    // Stop after 10+ consecutive empty rows (end of data)
    if (row.every(v => v === null)) { emptyStreak++; }
    else { emptyStreak = 0; }
    if (emptyStreak >= 10) break;
  }
  return result;
}

// Helper: Find sheet by name (case-insensitive)
function _pGetSheet(workbook, name) {
  const lower = name.toLowerCase();
  for (const sheetName of workbook.SheetNames) {
    if (sheetName.toLowerCase() === lower) {
      return workbook.Sheets[sheetName];
    }
  }
  return null;
}

// Helper: Check if a row is a week marker (col 1 contains "Week N")
function _pIsWeekMarker(row) {
  if (!row || !row[1]) return null;
  const exerciseCell = row[1].toString().trim();
  const match = exerciseCell.match(/^Week\s+(\d+)$/i);
  return match ? match[1] : null;
}

// Helper: Check if exercise row is empty/blank (col 1 is null/empty)
function _pIsBlankRow(row) {
  return !row || !row[1] || (typeof row[1] === 'string' && row[1].trim() === '');
}

// Helper: Parse comma-separated weights
function _pParseWeights(weightsStr) {
  if (!weightsStr) return [];
  const str = weightsStr.toString().trim();
  if (!str || str === ',,') return [];
  return str.split(',').map(w => {
    const trimmed = w.trim();
    return trimmed ? parseFloat(trimmed) : null;
  }).filter(w => w !== null && !isNaN(w));
}

// Helper: Build exercise name with modifiers
function _pBuildExerciseName(exercise, modifiers) {
  if (!exercise) return '';
  let name = exercise.toString().trim();
  if (modifiers && modifiers.toString().trim()) {
    name += ' ' + modifiers.toString().trim();
  }
  return name;
}

// Helper: Build prescription from reps and weights
function _pBuildPrescription(reps, estWorkingWeights, estTopWeight) {
  const weights = _pParseWeights(estWorkingWeights);

  if (weights.length === 0) {
    // Fall back to top weight if available
    if (estTopWeight && !isNaN(parseFloat(estTopWeight)) && parseFloat(estTopWeight) > 0) {
      return `1x${reps}(${Math.round(parseFloat(estTopWeight))})`;
    }
    // No weight data — return reps-only prescription
    return `${reps} reps`;
  }

  // Check if all weights are the same
  const allSame = weights.length > 0 && weights.every(w => w === weights[0]);

  if (allSame) {
    return `${weights.length}x${reps}(${Math.round(weights[0])})`;
  } else {
    // Build individual sets
    return weights.map(w => `1x${reps}(${Math.round(w)})`).join(', ');
  }
}

// Helper: Build note from RPE and fatigue info
function _pBuildNote(rpe, fatigueMethod, fatiguePercent) {
  let note = '';

  // Always include RPE
  if (rpe !== null && rpe !== undefined) {
    note = `@${rpe} RPE`;
  }

  // Add fatigue method if present
  if (fatigueMethod && fatigueMethod.toString().trim()) {
    const method = fatigueMethod.toString().trim();
    if (method === 'Load Drop' && fatiguePercent && !isNaN(parseFloat(fatiguePercent))) {
      const pct = Math.round(parseFloat(fatiguePercent) * 100);
      note += `, Load Drop -${pct}%`;
    } else if (method === 'Repeats') {
      note += ', Repeats';
    } else if (method === 'Rep Drop' && fatiguePercent && !isNaN(parseFloat(fatiguePercent))) {
      const pct = Math.round(parseFloat(fatiguePercent) * 100);
      note += `, Rep Drop -${pct}%`;
    } else if (method && method !== '') {
      note += `, ${method}`;
    }
  }

  return note || null;
}

// Helper: Extract maxes from Inputs sheet
function _pExtractMaxes(inputsSheet) {
  const maxes = {};
  if (!inputsSheet) return maxes;

  const data = _pSheetToArray(inputsSheet, 60);

  // Rows 14-49 have lift names in col 0, values in col 1
  for (let r = 14; r < Math.min(50, data.length); r++) {
    const liftName = data[r][0];
    const value = data[r][1];

    if (!liftName || !value) continue;

    const name = liftName.toString().toLowerCase().trim();
    const val = parseFloat(value);
    if (isNaN(val)) continue;

    // Match main lifts (avoid variants)
    if (name === 'squat' && !maxes.squat) {
      maxes.squat = Math.round(val);
    } else if ((name === 'bench competition' || name === 'bench') && !maxes.bench) {
      maxes.bench = Math.round(val);
    } else if (name === 'deadlift' && !maxes.deadlift) {
      maxes.deadlift = Math.round(val);
    }
  }

  return maxes;
}

/**
 * DETECT: Check if workbook is Format P (RTS)
 * Requirements:
 * 1. Must have sheets "Main" and "Inputs" (case-insensitive)
 * 2. Must find headers with "Date" and "Exercise" in cols 0-1
 * 3. Must have "Week 1" in col 1 somewhere in first ~10 rows
 * 4. Must have "Fatigue" in headers
 */
function detectP(workbook) {
  if (!workbook || !workbook.SheetNames) return false;

  // Check for Main and Inputs sheets
  const mainSheet = _pGetSheet(workbook, 'Main');
  const inputsSheet = _pGetSheet(workbook, 'Inputs');
  if (!mainSheet || !inputsSheet) return false;

  const mainData = _pSheetToArray(mainSheet, 20);

  // Find header row (look for "Date" and "Exercise" in cols 0-1)
  let headerRow = -1;
  for (let r = 0; r < Math.min(5, mainData.length); r++) {
    const row = mainData[r];
    if (!row) continue;
    const h0 = row[0];
    const h1 = row[1];
    if (h0 === 'Date' && h1 === 'Exercise') {
      headerRow = r;
      break;
    }
  }

  if (headerRow === -1) return false;

  // Check for "Fatigue" in headers (cols 5 or 6)
  const headerData = mainData[headerRow];
  const h5 = headerData[5] ? headerData[5].toString() : '';
  const h6 = headerData[6] ? headerData[6].toString() : '';
  if (!h5.includes('Fatigue') && !h6.includes('Fatigue')) return false;

  // Check for "Week 1" in col 1 within first ~10 rows
  let foundWeek1 = false;
  for (let r = 0; r < Math.min(15, mainData.length); r++) {
    if (mainData[r] && mainData[r][1]) {
      const cell = mainData[r][1].toString();
      if (cell.match(/^Week\s+1$/i)) {
        foundWeek1 = true;
        break;
      }
    }
  }

  return foundWeek1;
}

/**
 * PARSE: Extract all training blocks from Format P
 */
function parseP(workbook) {
  if (!workbook || !workbook.SheetNames) return [];

  const mainSheet = _pGetSheet(workbook, 'Main');
  const inputsSheet = _pGetSheet(workbook, 'Inputs');
  if (!mainSheet) return [];

  // Extract maxes
  const maxes = _pExtractMaxes(inputsSheet);

  // Detect program type from workbook properties or content
  let programName = 'RTS Program';
  let programId = 'rts_program';

  // Try to detect which file this is from workbook properties or sheet content
  let detectedName = '';
  if (workbook.Props && workbook.Props.Title) {
    detectedName = workbook.Props.Title.toLowerCase();
  }

  // Also check the first row of Main sheet for title hints
  let mainData = _pSheetToArray(mainSheet, 5);
  if (mainData[0] && mainData[0][2]) {
    const titleHint = mainData[0][2].toString().toLowerCase();
    if (titleHint) detectedName += ' ' + titleHint;
  }

  if (detectedName.includes('generalized')) {
    programName = 'RTS Generalized Intermediate';
    programId = 'rts_generalized_intermediate';
  } else if (detectedName.includes('community')) {
    programName = 'RTS Community v2';
    programId = 'rts_community_v2';
  }

  mainData = _pSheetToArray(mainSheet, 1000);

  // Find header row (look for "Date" and "Exercise" in cols 0-1)
  let headerRow = -1;
  for (let r = 0; r < Math.min(5, mainData.length); r++) {
    const row = mainData[r];
    if (!row) continue;
    const h0 = row[0];
    const h1 = row[1];
    if (h0 === 'Date' && h1 === 'Exercise') {
      headerRow = r;
      break;
    }
  }

  if (headerRow === -1) return [];

  // Data starts after header row + 1 (sub-header row)
  const dataStartRow = headerRow + 2;

  // Parse structure:
  // headerRow = Headers
  // headerRow+1 = Sub-headers (skip)
  // dataStartRow+ = Data, with:
  //   - Week markers (col[1] = "Week N")
  //   - Blank rows (col[1] = empty, separate days)
  //   - Exercise rows (col[1] = exercise name)

  const weeks = [];
  let currentWeek = null;
  let currentDay = null;
  let dayCount = 0;

  for (let r = dataStartRow; r < mainData.length; r++) {
    const row = mainData[r];
    if (!row) continue;

    // Check for week marker
    const weekNum = _pIsWeekMarker(row);
    if (weekNum) {
      currentWeek = {
        label: `Week ${weekNum}`,
        days: []
      };
      weeks.push(currentWeek);
      currentDay = null;
      dayCount = 0;
      continue;
    }

    // Skip blank rows (they separate days within a week)
    if (_pIsBlankRow(row)) {
      currentDay = null; // Next exercise starts a new day
      continue;
    }

    // This is an exercise row
    if (!currentWeek) {
      // Skip exercises before first week marker
      continue;
    }

    // If no current day, start a new one
    if (!currentDay) {
      dayCount++;
      currentDay = {
        name: `Day ${dayCount}`,
        exercises: []
      };
      currentWeek.days.push(currentDay);
    }

    // Extract exercise data
    const exercise = row[1] ? row[1].toString().trim() : null;
    if (!exercise) continue;

    const modifiers = row[2] ? row[2].toString().trim() : null;
    const reps = row[3] !== null && row[3] !== undefined ? parseInt(row[3]) : null;
    const rpe = row[4] !== null && row[4] !== undefined ? parseFloat(row[4]) : null;
    const fatiguePercent = row[5] !== null && row[5] !== undefined ? parseFloat(row[5]) : null;
    const fatigueMethod = row[6] ? row[6].toString().trim() : null;
    const estTopWeight = row[7] !== null && row[7] !== undefined ? parseFloat(row[7]) : null;
    const estWorkingWeights = row[8] ? row[8].toString().trim() : null;

    const name = _pBuildExerciseName(exercise, modifiers);
    const prescription = reps ? _pBuildPrescription(reps, estWorkingWeights, estTopWeight) : null;
    const note = _pBuildNote(rpe, fatigueMethod, fatiguePercent);

    currentDay.exercises.push({
      name,
      prescription,
      note,
      lifterNote: null,
      loggedWeight: null,
      supersetGroup: null
    });
  }

  // Return as array with single block
  return [{
    id: programId,
    name: programName,
    format: 'P',
    athleteName: null,
    dateRange: null,
    maxes,
    weeks
  }];
}


// ── FORMAT Q PARSER (Barbell Medicine Bridge v1/v2/v3) ──
/**
 * Format Q: Barbell Medicine Bridge 1.0 Parser
 * Supports 3 variants:
 * - v1: Single sheet with prescription strings (barbell_medicine_bridge_v1.xlsx)
 * - v2: Multi-sheet with vertical set layout (barbell_medicine_bridge_v2.xlsx)
 * - v3: Per-week sheets with set rows (barbell_medicine_bridge_v3.xlsx)
 */

// ============================================================================
// DETECTION
// ============================================================================

function detectQ(wb) {
  if (!wb || !wb.SheetNames) return false;

  // Try v1 detection
  if (_qDetectV1(wb)) return true;

  // Try v2 detection
  if (_qDetectV2(wb)) return true;

  // Try v3 detection
  if (_qDetectV3(wb)) return true;

  return false;
}

function _qDetectV1(wb) {
  // v1: Single sheet containing "bridge" with specific headers
  // v1 has headers at R2 (row index 1), col B/C (cols 1/2)
  for (const sheetName of wb.SheetNames) {
    if (!sheetName.toLowerCase().includes('bridge')) continue;

    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    // Check headers in R2 (row index 1): Exercise in col B (1), Prescription in col C (2)
    const cell1 = _qGetCell(ws, 1, 1);
    const cell2 = _qGetCell(ws, 1, 2);

    if (cell1 === 'Exercise' && cell2 === 'Prescription') {
      return true;
    }
  }
  return false;
}

function _qDetectV2(wb) {
  // v2: Sheet named "The Bridge 1.0" with Week marker pattern and Weight/Reps headers
  const sheetName = 'The Bridge 1.0';
  if (!wb.Sheets[sheetName]) return false;

  const ws = wb.Sheets[sheetName];

  // Look for "Week 1" marker around row 7
  let foundWeek = false;
  for (let r = 5; r <= 10; r++) {
    const cell = _qGetCell(ws, r, 0);
    if (cell && cell.toString().includes('Week')) {
      foundWeek = true;
      break;
    }
  }

  if (!foundWeek) return false;

  // Look for Weight/Reps pattern in col 2 (around row 8-15)
  let foundPattern = false;
  for (let r = 7; r <= 20; r++) {
    const cell = _qGetCell(ws, r, 2);
    if (cell && (cell.toString().includes('Weight') || cell.toString().includes('Reps'))) {
      foundPattern = true;
      break;
    }
  }

  return foundPattern;
}

function _qDetectV3(wb) {
  // v3: Multiple sheets matching "Week N - *Stress" pattern
  const weekPattern = /^Week\s+\d+\s*-\s*\w+\s*Stress/i;
  let weekCount = 0;

  for (const sheetName of wb.SheetNames) {
    if (weekPattern.test(sheetName)) {
      weekCount++;
    }
  }

  // Need at least 3 week sheets to be confident
  return weekCount >= 3;
}

// ============================================================================
// CLASSIFICATION
// ============================================================================

function _qClassify(wb) {
  // Order matters: check specific patterns before falling back
  if (_qDetectV1(wb)) return 'v1';
  if (_qDetectV2(wb)) return 'v2';
  if (_qDetectV3(wb)) return 'v3';
  return null;
}

// ============================================================================
// MAIN PARSER DISPATCHER
// ============================================================================

function parseQ(wb) {
  const variant = _qClassify(wb);

  if (variant === 'v1') return _parseQ_v1(wb);
  if (variant === 'v2') return _parseQ_v2(wb);
  if (variant === 'v3') return _parseQ_v3(wb);

  return [];
}

// ============================================================================
// VARIANT 1 PARSER: Single sheet with prescription strings
// ============================================================================

function _parseQ_v1(wb) {
  let sheetName = null;
  for (const name of wb.SheetNames) {
    if (name.toLowerCase().includes('bridge')) {
      sheetName = name;
      break;
    }
  }

  if (!sheetName) return [];

  const ws = wb.Sheets[sheetName];
  const weeks = [];
  let currentWeek = null;
  let currentDay = null;

  // v1 structure: Week/Day/Exercise markers in col B (col index 1)
  // Exercise data: exercise name in col B (1), prescription in col C (2)
  const maxRows = 300;
  for (let r = 2; r < maxRows; r++) {
    const colB = _qGetCell(ws, r, 1);
    const colC = _qGetCell(ws, r, 2);

    if (!colB) continue;

    const colBText = colB.toString().trim();

    // Skip empty rows
    if (!colBText) continue;

    // Skip special rows BEFORE processing week/day markers
    if (colBText.includes('Daily Total') || colBText.includes('Weekly Total')) {
      continue;
    }

    // Detect week marker (but not "Weekly Total")
    if (colBText.includes('Week') && !colBText.includes('Total')) {
      if (currentDay && currentWeek) {
        currentWeek.days.push(currentDay);
      }
      if (currentWeek) {
        weeks.push(currentWeek);
      }
      currentWeek = {
        label: colBText,
        days: []
      };
      currentDay = null;
      continue;
    }

    // Detect day marker (includes "Day")
    if (colBText.includes('Day')) {
      if (currentDay && currentWeek) {
        currentWeek.days.push(currentDay);
      }
      // Skip GPP days (they have no structured exercise data)
      if (colBText.includes('GPP')) {
        currentDay = null;
        continue;
      }
      currentDay = {
        name: colBText,
        exercises: []
      };
      continue;
    }

    // Skip GPP text rows (no prescription data)
    if (colBText.toLowerCase().includes('pull-ups') ||
        colBText.toLowerCase().includes('cardio') ||
        colBText.toLowerCase().includes('accumulate')) {
      continue;
    }

    // Exercise row: exercise name in colB, prescription in colC
    if (colC && colC.toString().trim()) {
      const exerciseName = colBText;
      const prescriptionStr = colC.toString().trim();

      // Skip if looks like a header or label
      if (exerciseName.toLowerCase() === 'exercise') continue;

      if (currentDay) {
        currentDay.exercises.push({
          name: exerciseName,
          prescription: prescriptionStr,
          note: null,
          lifterNote: null,
          loggedWeight: null,
          supersetGroup: null
        });
      }
    }
  }

  // Flush last day/week
  if (currentDay && currentWeek) {
    currentWeek.days.push(currentDay);
  }
  if (currentWeek) {
    weeks.push(currentWeek);
  }

  return [{
    id: 'barbell_medicine_bridge_v1',
    name: 'Barbell Medicine - The Bridge 1.0',
    format: 'Q',
    athleteName: null,
    dateRange: null,
    maxes: {},
    weeks: weeks
  }];
}

// ============================================================================
// VARIANT 2 PARSER: Multi-sheet with vertical set layout
// ============================================================================

function _parseQ_v2(wb) {
  const ws = wb.Sheets['The Bridge 1.0'];
  if (!ws) return [];

  const weeks = [];
  let currentWeek = null;
  let currentDay = null;
  let currentExercise = null;
  let exerciseReps = [];
  let exerciseRPE = [];

  const maxRows = 1500;
  for (let r = 0; r < maxRows; r++) {
    const col0 = _qGetCell(ws, r, 0);
    const col1 = _qGetCell(ws, r, 1);
    const col2 = _qGetCell(ws, r, 2);

    // Detect week marker (col 0 = "Week N", col 1 = stress level)
    if (col0 && col0.toString().trim().match(/^Week\s+\d+$/i)) {
      if (currentDay && currentWeek) {
        currentWeek.days.push(currentDay);
      }
      if (currentWeek) {
        weeks.push(currentWeek);
      }
      const stressLevel = col1 ? col1.toString().trim() : '';
      const weekLabel = `Week ${col0.toString().match(/\d+/)[0]}${stressLevel ? ' - ' + stressLevel : ''}`;
      currentWeek = {
        label: weekLabel,
        days: []
      };
      currentDay = null;
      continue;
    }

    // Detect day marker (col 0 = "Day N" or "REST DAY")
    if (col0) {
      const dayText = col0.toString().trim();
      if (dayText.match(/^Day\s+\d+/i)) {
        // Flush current exercise if any
        if (currentExercise && currentDay) {
          _qFlushExerciseV2(currentDay, currentExercise, exerciseReps, exerciseRPE);
        }
        currentExercise = null;
        exerciseReps = [];
        exerciseRPE = [];

        // Skip GPP days
        if (dayText.includes('GPP')) {
          currentDay = null;
          continue;
        }

        // Skip REST DAY
        if (dayText.includes('REST')) {
          currentDay = null;
          continue;
        }

        if (currentDay && currentWeek) {
          currentWeek.days.push(currentDay);
        }
        currentDay = {
          name: dayText,
          exercises: []
        };
        continue;
      }
    }

    // Exercise row: col 1 has exercise name, col 2 = "Weight"
    if (col1 && col2 && col2.toString().trim().toLowerCase() === 'weight') {
      // Flush previous exercise
      if (currentExercise && currentDay) {
        _qFlushExerciseV2(currentDay, currentExercise, exerciseReps, exerciseRPE);
      }
      currentExercise = col1.toString().trim();
      exerciseReps = [];
      exerciseRPE = [];
      continue;
    }

    // Reps row with data (col 2 = "Reps", cols 3+ have numbers)
    if (col2 && col2.toString().trim().toLowerCase() === 'reps' && currentExercise) {
      exerciseReps = [];
      for (let c = 3; c <= 8; c++) {
        const val = _qGetCell(ws, r, c);
        if (val && val !== '') {
          const numVal = parseInt(val, 10);
          if (!isNaN(numVal)) {
            exerciseReps.push(numVal);
          }
        }
      }
    }

    // RPE row with data (col 2 = "RPE", cols 3+ have numbers)
    if (col2 && col2.toString().trim().toLowerCase() === 'rpe' && currentExercise) {
      exerciseRPE = [];
      for (let c = 3; c <= 8; c++) {
        const val = _qGetCell(ws, r, c);
        if (val && val !== '') {
          const numVal = parseInt(val, 10);
          if (!isNaN(numVal)) {
            exerciseRPE.push(numVal);
          }
        }
      }
    }
  }

  // Flush last exercise/day/week
  if (currentExercise && currentDay) {
    _qFlushExerciseV2(currentDay, currentExercise, exerciseReps, exerciseRPE);
  }
  if (currentDay && currentWeek) {
    currentWeek.days.push(currentDay);
  }
  if (currentWeek) {
    weeks.push(currentWeek);
  }

  return [{
    id: 'barbell_medicine_bridge_v2',
    name: 'Barbell Medicine - The Bridge 1.0',
    format: 'Q',
    athleteName: null,
    dateRange: null,
    maxes: {},
    weeks: weeks
  }];
}

function _qFlushExerciseV2(day, exerciseName, reps, rpe) {
  if (!exerciseName || reps.length === 0 || rpe.length === 0) {
    return;
  }

  // Reconstruct prescription: group consecutive same reps/rpe
  const prescription = _qBuildPrescription(reps, rpe);

  day.exercises.push({
    name: exerciseName,
    prescription: prescription,
    note: null,
    lifterNote: null,
    loggedWeight: null,
    supersetGroup: null
  });
}

// ============================================================================
// VARIANT 3 PARSER: Per-week sheets with set rows
// ============================================================================

function _parseQ_v3(wb) {
  const weekPattern = /^Week\s+(\d+)\s*-\s*(.+Stress)$/i;
  const weekSheets = [];

  for (const sheetName of wb.SheetNames) {
    const match = sheetName.match(weekPattern);
    if (match) {
      weekSheets.push({
        sheetName: sheetName,
        weekNum: parseInt(match[1], 10),
        label: sheetName
      });
    }
  }

  // Sort by week number
  weekSheets.sort((a, b) => a.weekNum - b.weekNum);

  const weeks = [];
  for (const ws_info of weekSheets) {
    const ws = wb.Sheets[ws_info.sheetName];
    if (!ws) continue;

    const week = {
      label: ws_info.label,
      days: []
    };

    let currentDay = null;
    let currentExerciseName = null;
    let currentSets = [];
    let inGPP = false;

    const maxRows = 1500;
    for (let r = 0; r < maxRows; r++) {
      const col0 = _qGetCell(ws, r, 0);
      const col1 = _qGetCell(ws, r, 1);
      const col2 = _qGetCell(ws, r, 2);
      const col3 = _qGetCell(ws, r, 3);
      const col4 = _qGetCell(ws, r, 4);

      // Detect GPP marker and skip the GPP section entirely
      if (col0) {
        const col0Text = col0.toString().trim().toUpperCase();
        if (col0Text === 'GPP') {
          // Flush current day
          if (currentExerciseName && currentSets.length > 0 && currentDay) {
            _qFlushExerciseV3(currentDay, currentExerciseName, currentSets);
          }
          if (currentDay) {
            week.days.push(currentDay);
          }
          currentDay = null;
          currentExerciseName = null;
          currentSets = [];
          inGPP = true;
          continue;
        }
      }

      // Skip any remaining rows in GPP section
      if (inGPP) {
        // GPP section ends when we reach end of sheet or next real week marker
        if (col0 && col0.toString().trim().toUpperCase() === 'STRENGTH TRAINING') {
          inGPP = false;
        }
        continue;
      }

      // Detect day marker (col 0 = "Day N")
      if (col0) {
        const dayText = col0.toString().trim();
        if (dayText.match(/^Day\s+\d+/i) && !dayText.includes('Strength Training')) {
          // Flush current exercise
          if (currentExerciseName && currentSets.length > 0 && currentDay) {
            _qFlushExerciseV3(currentDay, currentExerciseName, currentSets);
          }
          currentExerciseName = null;
          currentSets = [];

          if (currentDay) {
            week.days.push(currentDay);
          }
          currentDay = {
            name: dayText,
            exercises: []
          };
          continue;
        }
      }

      // Headers row (Exercise, Set, Reps, Assigned RPE, etc.) - skip it
      if (col0 && col0.toString().trim().toLowerCase() === 'exercise' &&
          col1 && col1.toString().trim().toLowerCase().includes('set')) {
        continue;
      }

      // Exercise data row: col1 has exercise name OR col2 has set number
      // If col1 is non-empty text (not numeric) and we're in a day, it's a new exercise
      if (col1 && col1.toString().trim() && !_qIsNumeric(col1)) {
        const exerciseText = col1.toString().trim();

        // Skip labels like "Date", "Bodyweight", headers
        if (exerciseText.toLowerCase() === 'date' ||
            exerciseText.toLowerCase() === 'bodyweight' ||
            exerciseText.toLowerCase() === 'exercise' ||
            exerciseText.toLowerCase().includes('strength training')) {
          continue;
        }

        // Flush previous exercise
        if (currentExerciseName && currentSets.length > 0 && currentDay) {
          _qFlushExerciseV3(currentDay, currentExerciseName, currentSets);
        }

        currentExerciseName = exerciseText;
        currentSets = [];

        // Parse first set from this row
        const setNum = col2 ? parseInt(col2, 10) : null;
        const reps = col3 ? parseInt(col3, 10) : null;
        const rpe = col4 ? parseInt(col4, 10) : null;

        if (setNum && reps && rpe) {
          currentSets.push({ set: setNum, reps: reps, rpe: rpe });
        }
      } else if (col2 && _qIsNumeric(col2) && currentExerciseName && currentDay) {
        // Subsequent set row for current exercise: col2 = set number
        const setNum = parseInt(col2, 10);
        const reps = col3 ? parseInt(col3, 10) : null;
        const rpe = col4 ? parseInt(col4, 10) : null;

        if (reps && rpe) {
          currentSets.push({ set: setNum, reps: reps, rpe: rpe });
        }
      }
    }

    // Flush last exercise/day
    if (currentExerciseName && currentSets.length > 0 && currentDay) {
      _qFlushExerciseV3(currentDay, currentExerciseName, currentSets);
    }
    if (currentDay) {
      week.days.push(currentDay);
    }

    weeks.push(week);
  }

  return [{
    id: 'barbell_medicine_bridge_v3',
    name: 'Barbell Medicine - The Bridge 1.0',
    format: 'Q',
    athleteName: null,
    dateRange: null,
    maxes: {},
    weeks: weeks
  }];
}

function _qFlushExerciseV3(day, exerciseName, sets) {
  if (!exerciseName || sets.length === 0) {
    return;
  }

  // Extract reps and RPE from sets
  const reps = sets.map(s => s.reps);
  const rpe = sets.map(s => s.rpe);

  const prescription = _qBuildPrescription(reps, rpe);

  day.exercises.push({
    name: exerciseName,
    prescription: prescription,
    note: null,
    lifterNote: null,
    loggedWeight: null,
    supersetGroup: null
  });
}

// ============================================================================
// PRESCRIPTION BUILDER
// ============================================================================

function _qBuildPrescription(reps, rpe) {
  // reps: [5, 5, 5]
  // rpe: [6, 7, 8]
  // Returns: "5@6, 5@7, 5@8"

  // reps: [8, 8, 8, 8]
  // rpe: [6, 7, 8, 8]
  // Returns: "8@6, 8@7, 2x8@8"

  if (!reps || !rpe || reps.length === 0 || rpe.length === 0) {
    return null;
  }

  // Ensure lengths match
  const len = Math.min(reps.length, rpe.length);
  const pairs = [];
  for (let i = 0; i < len; i++) {
    pairs.push({ reps: reps[i], rpe: rpe[i] });
  }

  // Group consecutive identical reps/rpe
  const groups = [];
  let currentGroup = null;
  for (const pair of pairs) {
    if (!currentGroup) {
      currentGroup = { reps: pair.reps, rpe: pair.rpe, count: 1 };
    } else if (currentGroup.reps === pair.reps && currentGroup.rpe === pair.rpe) {
      currentGroup.count++;
    } else {
      groups.push(currentGroup);
      currentGroup = { reps: pair.reps, rpe: pair.rpe, count: 1 };
    }
  }
  if (currentGroup) {
    groups.push(currentGroup);
  }

  // Build prescription string
  const parts = [];
  for (const group of groups) {
    if (group.count > 1) {
      parts.push(`${group.count}x${group.reps}@${group.rpe}`);
    } else {
      parts.push(`${group.reps}@${group.rpe}`);
    }
  }

  return parts.join(', ');
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function _qGetCell(ws, row, col) {
  if (!ws) return null;
  const cellRef = XLSX.utils.encode_col(col) + (row + 1);
  const cell = ws[cellRef];
  if (!cell) return null;

  const val = (cell.v instanceof Date) ? (cell.w || String(cell.v)) : cell.v;

  // Handle error/NA values
  if (val === '#N/A' || val === null || val === undefined) {
    return null;
  }

  // Handle various empty/invalid strings
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (trimmed === '' || trimmed === 'N/A' || trimmed === '---' || trimmed === '--') {
      return null;
    }
  }

  return val;
}

function _qIsNumeric(val) {
  if (val === null || val === undefined || val === '') return false;
  const num = Number(val);
  return !isNaN(num) && num !== '';
}

// ============================================================================
// EXPORTS
// ============================================================================


// ── FORMAT R: PHUL (Power Hypertrophy Upper Lower) ────────────────────────
/**
 * Format R Parser - PHUL (Power Hypertrophy Upper Lower) Powerlifting Program
 *
 * Parses PHUL spreadsheets with day-of-week sheets (Monday-Saturday)
 * Each day contains exercise blocks with Target/Set(s)/Rep(s)/Percentage data
 * Optional TUT and RPS rows provide additional exercise notes
 */

/**
 * Detect if a workbook is in Format R (PHUL-style)
 * @param {object} workbook - XLSX workbook object
 * @returns {boolean}
 */
function detectR(workbook) {
  const sheetNames = workbook.SheetNames || [];

  // Check for day-of-week sheet names (need at least 3)
  const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const daySheets = sheetNames.filter(name => dayNames.includes(name));
  if (daySheets.length < 3) return false;

  // Check first day sheet
  const firstDaySheet = daySheets[0];
  const sheet = workbook.Sheets[firstDaySheet];
  if (!sheet) return false;

  // Row 0 should have "Week" in col A
  const A0 = _rGetCell(sheet, 0, 0);
  if (A0 !== 'Week') return false;

  // Check for Target, Set(s), Rep(s) labels within first 10 rows
  let hasTarget = false, hasSetLabel = false, hasRepLabel = false;
  for (let r = 0; r < 10; r++) {
    const cellA = _rGetCell(sheet, r, 0);
    if (cellA === 'Target') hasTarget = true;
    if (cellA === 'Set(s)') hasSetLabel = true;
    if (cellA === 'Rep(s)') hasRepLabel = true;
  }

  if (!hasTarget || !hasSetLabel || !hasRepLabel) return false;

  // Should NOT have "Exercise" header in row 0-3 col A
  for (let r = 0; r < 4; r++) {
    const cellA = _rGetCell(sheet, r, 0);
    if (cellA === 'Exercise') return false;
  }

  return true;
}

/**
 * Parse Format R workbook into PowerliftingLog format
 * @param {object} workbook - XLSX workbook object
 * @returns {array} Array of parsed block objects
 */
function parseR(workbook) {
  const blocks = [];
  const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const daySheets = workbook.SheetNames.filter(name => dayNames.includes(name));

  if (daySheets.length === 0) return blocks;

  // Get week count and dates from first day sheet
  const firstSheet = workbook.Sheets[daySheets[0]];
  const weekCount = _rGetWeekCount(firstSheet);

  // Create a single block containing all days
  const block = {
    id: _rGenerateId(),
    name: 'PHUL',
    format: 'R',
    athleteName: null,
    dateRange: null,
    maxes: {},
    weeks: []
  };

  // Build weeks with days
  for (let w = 0; w < weekCount; w++) {
    const weekLabel = `Week ${w + 1}`;
    const weekDays = [];

    for (const daySheetName of daySheets) {
      const sheet = workbook.Sheets[daySheetName];
      const dayName = daySheetName;
      const exercises = _rParseDay(sheet, w);

      if (exercises.length > 0) {
        weekDays.push({
          name: dayName,
          exercises: exercises
        });
      }
    }

    if (weekDays.length > 0) {
      block.weeks.push({
        label: weekLabel,
        days: weekDays
      });
    }
  }

  blocks.push(block);
  return blocks;
}

/**
 * Parse a single day sheet for a specific week
 * @param {object} sheet - XLSX sheet object
 * @param {number} weekIndex - 0-based week index
 * @returns {array} Array of exercise objects
 */
function _rParseDay(sheet, weekIndex) {
  const exercises = [];
  const columnIndex = weekIndex + 1; // Week columns start at B (col 1)

  // Skip header rows (0-2: Week, Date, blank)
  let rowIndex = 3;

  while (rowIndex < 100) { // Reasonable upper limit
    const cellA = _rGetCell(sheet, rowIndex, 0);

    // Stop at RPE/reaction rows or empty section
    if (!cellA || _rIsRPERow(cellA)) break;

    // Exercise name found (no cell content in other columns at this row)
    if (cellA && !['Target', 'Set(s)', 'Rep(s)', 'Percentage', 'Approx. Percentage', 'TUT', 'RPS'].includes(cellA)) {
      const exerciseName = cellA;
      const exerciseData = _rParseExerciseBlock(sheet, rowIndex, columnIndex);

      if (exerciseData && exerciseData.rows > 0) {
        // Only include exercise if it's not skipped for this week
        if (exerciseData.prescription) {
          exercises.push({
            name: exerciseName,
            prescription: exerciseData.prescription,
            note: exerciseData.note,
            lifterNote: null,
            loggedWeight: null,
            supersetGroup: null
          });
        }
        rowIndex += exerciseData.rows;
      } else {
        rowIndex++;
      }
    } else {
      rowIndex++;
    }
  }

  return exercises;
}

/**
 * Parse an exercise block starting from the exercise name row
 * @param {object} sheet - XLSX sheet object
 * @param {number} startRow - Row of exercise name
 * @param {number} columnIndex - Column index for the week (1=week1, 2=week2, etc.)
 * @returns {object|null} { prescription, note, rows } or null if skipped
 */
function _rParseExerciseBlock(sheet, startRow, columnIndex) {
  let rowIndex = startRow + 1;
  let targetValue = null;
  let setsValue = null;
  let repsValue = null;
  let percentage = null;
  let tut = null;
  let rps = null;
  let rowsConsumed = 1;

  // Parse exercise block rows
  while (rowIndex < startRow + 20) { // Reasonable limit for exercise block
    const cellA = _rGetCell(sheet, rowIndex, 0);
    const cellValue = _rGetCell(sheet, rowIndex, columnIndex);

    if (!cellA) {
      // Blank row signals end of exercise block
      rowsConsumed = rowIndex - startRow + 1;
      break;
    }

    if (cellA === 'Target') {
      targetValue = cellValue;
      rowIndex++;
    } else if (cellA === 'Set(s)') {
      setsValue = cellValue;
      rowIndex++;
    } else if (cellA === 'Rep(s)') {
      repsValue = cellValue;
      rowIndex++;
    } else if (cellA === 'Percentage' || cellA === 'Approx. Percentage') {
      percentage = cellValue;
      rowIndex++;
    } else if (cellA === 'TUT') {
      tut = cellValue;
      rowIndex++;
    } else if (cellA === 'RPS') {
      rps = cellValue;
      rowIndex++;
    } else {
      // Unknown row or next exercise
      rowsConsumed = rowIndex - startRow;
      break;
    }
  }

  // Check if exercise is skipped for this week
  if (!targetValue || !setsValue || !repsValue) {
    return { prescription: null, note: null, rows: rowsConsumed };
  }

  const targetStr = String(targetValue).trim();
  const setsStr = String(setsValue).trim();
  const repsStr = String(repsValue).trim();

  if (targetStr === '-' || setsStr === '-' || repsStr === '-' || !targetStr || !setsStr || !repsStr) {
    return { prescription: null, note: null, rows: rowsConsumed };
  }

  // Build prescription
  const prescription = _rBuildPrescription(setsStr, repsStr, targetStr);

  // Build note
  const note = _rBuildNote(percentage, tut, rps);

  return { prescription, note, rows: rowsConsumed };
}

/**
 * Build prescription string from sets, reps, and target
 * @param {string} sets - Sets value (already trimmed; can be "3", "AMRAP", "3 + AMRAP")
 * @param {string} reps - Reps value (already trimmed)
 * @param {string} target - Target weight (already trimmed)
 * @returns {string}
 */
function _rBuildPrescription(sets, reps, target) {
  if (!sets || !reps) return '';

  // Handle AMRAP variations
  if (sets.includes('+')) {
    // "3 + AMRAP" format
    const parts = sets.split('+').map(p => p.trim());
    const mainSets = parts[0];
    const prescription = `${mainSets}x${reps}(${target})`;

    if (parts[1] && parts[1].toUpperCase() === 'AMRAP') {
      return `${prescription}, AMRAP(${target})`;
    }
    return prescription;
  }

  if (sets.toUpperCase() === 'AMRAP') {
    return `AMRAP(${target})`;
  }

  // Normal format: sets x reps (target)
  return `${sets}x${reps}(${target})`;
}

/**
 * Build note string from percentage, TUT, and RPS
 * @param {string|number} percentage - Percentage value (e.g., 0.725)
 * @param {string} tut - Time under tension (e.g., "4111")
 * @param {string} rps - RPS value
 * @returns {string|null}
 */
function _rBuildNote(percentage, tut, rps) {
  const notes = [];

  if (percentage) {
    const pctStr = String(percentage).trim();
    if (pctStr && pctStr !== '-') {
      const pctNum = parseFloat(pctStr);
      if (!isNaN(pctNum)) {
        const pctDisplay = Math.round(pctNum * 1000) / 10; // Convert to percentage
        notes.push(`${pctDisplay}% 1RM`);
      }
    }
  }

  if (tut) {
    const tutStr = String(tut).trim();
    if (tutStr && tutStr !== '-') {
      notes.push(`TUT: ${tutStr}`);
    }
  }

  if (rps) {
    const rpsStr = String(rps).trim();
    if (rpsStr && rpsStr !== '-') {
      notes.push(`RPS: ${rpsStr}`);
    }
  }

  return notes.length > 0 ? notes.join(', ') : null;
}

/**
 * Check if a row label is an RPE/reaction row
 * @param {string} label - Cell value from col A
 * @returns {boolean}
 */
function _rIsRPERow(label) {
  if (!label) return false;
  const upper = String(label).toUpperCase();
  return upper.includes('EASY') ||
         upper.includes('TOO') ||
         upper.includes('RPE') ||
         upper.includes('REACTION');
}

/**
 * Get week count from sheet (number of week columns)
 * @param {object} sheet - XLSX sheet object
 * @returns {number}
 */
function _rGetWeekCount(sheet) {
  // Row 0 has "Week" in col A and week numbers in B onwards
  let count = 0;
  for (let col = 1; col <= 20; col++) {
    const val = _rGetCell(sheet, 0, col);
    // Check if this is a numeric week column
    if (val !== undefined && val !== '' && !isNaN(parseInt(val))) {
      count++; // Count actual week columns found
    } else if (val === undefined || val === '') {
      // Stop at the first empty column
      break;
    }
  }
  // Return the actual count of weeks detected
  return count > 0 ? count : 13; // Default to 13 if detection fails
}

// _rGetDates removed — defined but never called (audit W6)

/**
 * Get cell value from sheet by row and column indices
 * @param {object} sheet - XLSX sheet object
 * @param {number} row - 0-based row index
 * @param {number} col - 0-based column index (0=A, 1=B, etc.)
 * @returns {*}
 */
function _rGetCell(sheet, row, col) {
  if (!sheet) return undefined;

  const colLetter = _rColIndexToLetter(col);
  const cellRef = `${colLetter}${row + 1}`; // XLSX uses 1-based row numbers
  const cell = sheet[cellRef];

  if (!cell) return undefined;

  // Return the value
  return (cell.v instanceof Date) ? (cell.w || String(cell.v)) : cell.v;
}

/**
 * Convert 0-based column index to Excel column letter
 * @param {number} index - 0-based column index
 * @returns {string}
 */
function _rColIndexToLetter(index) {
  let letter = '';
  let num = index;
  while (num >= 0) {
    letter = String.fromCharCode(65 + (num % 26)) + letter;
    num = Math.floor(num / 26) - 1;
  }
  return letter;
}

/**
 * Generate a unique ID for a block
 * @returns {string}
 */
function _rGenerateId() {
  return 'phul_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}


// ── FORMAT S: PHAT (Power Hypertrophy Adaptive Training) ──────────────────
/**
 * Format S Parser for PHAT (Power Hypertrophy Adaptive Training)
 * Parses the PHAT powerlifting program spreadsheet structure
 */

/**
 * Detects if a workbook is a PHAT-style Format S file
 * @param {Object} workbook - XLSX workbook object
 * @returns {boolean}
 */
function detectS(workbook) {
  if (!workbook || !workbook.SheetNames) {
    return false;
  }

  const sheetNames = workbook.SheetNames;

  // Must have "Inputs" sheet
  if (!sheetNames.includes('Inputs')) {
    return false;
  }

  // Must have at least one "Week N" sheet
  const weekSheets = sheetNames.filter(name => /^Week\s+\d+/.test(name));
  if (weekSheets.length === 0) {
    return false;
  }

  // Check first week sheet for "Set Scheme" header and section pattern
  const firstWeekSheet = weekSheets[0];
  const ws = workbook.Sheets[firstWeekSheet];

  if (!ws) {
    return false;
  }

  // Look for "Set Scheme" in first 3 rows
  let hasSetScheme = false;
  for (let row = 1; row <= 3; row++) {
    const cell = ws[`C${row}`];
    if (cell && cell.v && typeof cell.v === 'string' && cell.v.includes('Set Scheme')) {
      hasSetScheme = true;
      break;
    }
  }

  if (!hasSetScheme) {
    return false;
  }

  // Check for section header pattern (e.g., "1: Upper Body Power", "2: Lower Body Power", etc.)
  const range = ws['!ref'];
  if (!range) {
    return false;
  }

  // Scan for section headers matching pattern: digit: ... (Power|HT|Hypertrophy)
  const sectionPattern = /^\d+:\s+.*(Power|HT|Hypertrophy)/;
  for (const cellRef in ws) {
    if (cellRef.startsWith('A') && cellRef !== 'A!') {
      const cell = ws[cellRef];
      if (cell && cell.v && typeof cell.v === 'string') {
        if (sectionPattern.test(cell.v)) {
          return true;
        }
      }
    }
  }

  return false;
}

/**
 * Parses a PHAT Format S workbook
 * @param {Object} workbook - XLSX workbook object
 * @returns {Array} Array of parsed program blocks
 */
function parseS(workbook) {
  const sheetNames = workbook.SheetNames;
  const weekSheets = sheetNames.filter(name => /^Week\s+\d+/.test(name)).sort();

  if (weekSheets.length === 0) {
    return [];
  }

  const weeks = weekSheets.map(sheetName => ({
    sheetName,
    weekNumber: _sExtractWeekNumber(sheetName)
  }));

  const blocks = [];

  for (const week of weeks) {
    const ws = workbook.Sheets[week.sheetName];
    if (!ws) continue;

    const weekData = _sParseWeekSheet(ws, week.weekNumber);

    if (blocks.length === 0) {
      // First block - create with first week
      blocks.push({
        id: 'phat-1',
        name: 'PHAT',
        format: 'S',
        athleteName: null,
        dateRange: null,
        maxes: {},
        weeks: [weekData]
      });
    } else {
      // Subsequent weeks - add to first block
      blocks[0].weeks.push(weekData);
    }
  }

  return blocks;
}

/**
 * Extract week number from sheet name like "Week 1"
 * @param {string} sheetName
 * @returns {number}
 */
function _sExtractWeekNumber(sheetName) {
  const match = sheetName.match(/Week\s+(\d+)/i);
  return match ? parseInt(match[1], 10) : 0;
}

/**
 * Parse a single week sheet
 * @param {Object} ws - XLSX worksheet object
 * @param {number} weekNumber
 * @returns {Object} Week data structure
 */
function _sParseWeekSheet(ws, weekNumber) {
  const days = [];
  const sections = _sExtractSections(ws);

  // Day mapping: section header prefix to day number
  const sectionToDayName = {
    '1': 'Day 1 - Upper Body Power',
    '2': 'Day 2 - Lower Body Power',
    '4': 'Day 4 - Back & Shoulders HT',
    '5': 'Day 5 - Lower Body HT',
    '6': 'Day 6 - Chest & Arms HT'
  };

  // Process each section and build days
  for (const section of sections) {
    const dayNum = section.dayNumber;
    const dayLabel = sectionToDayName[dayNum] || `Day ${dayNum}`;

    // Extract display name from section title
    const sectionName = _sExtractSectionName(section.title);

    const exercises = section.exercises.map(ex => ({
      name: ex.name,
      prescription: _sFormatPrescription(ex.setScheme, ex.weight),
      note: _sFormatNote(ex.purpose, ex.percentage),
      lifterNote: null,
      loggedWeight: null,
      supersetGroup: null
    }));

    days.push({
      name: sectionName,
      exercises
    });
  }

  return {
    label: `Week ${weekNumber}`,
    days
  };
}

/**
 * Extract sections from worksheet
 * @param {Object} ws
 * @returns {Array} Array of section objects
 */
function _sExtractSections(ws) {
  const sections = [];
  const rows = _sGetWorksheetRows(ws);
  let currentSection = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    // Check if this is a section header row
    const sectionTitle = row[0]; // Column A

    if (_sIsSectionHeader(sectionTitle)) {
      // Save previous section if exists
      if (currentSection && currentSection.exercises.length > 0) {
        sections.push(currentSection);
      }

      const dayNum = _sExtractDayNumber(sectionTitle);

      currentSection = {
        title: sectionTitle,
        dayNumber: dayNum,
        exercises: []
      };
      continue;
    }

    // Skip blank rows
    if (!_sIsNonBlankRow(row)) {
      continue;
    }

    // If we have a current section and this is an exercise row
    if (currentSection && !_sIsSectionHeader(row[0])) {
      const exercise = {
        name: row[0],         // Column A
        purpose: row[1],      // Column B (Purpose)
        setScheme: row[2],    // Column C (Set Scheme)
        weight: _sParseNumber(row[3]),  // Column D (Weight)
        percentage: _sParseDecimal(row[4])  // Column E (Percentage)
      };

      currentSection.exercises.push(exercise);
    }
  }

  // Don't forget the last section
  if (currentSection && currentSection.exercises.length > 0) {
    sections.push(currentSection);
  }

  return sections;
}

/**
 * Get all rows from worksheet as arrays
 * @param {Object} ws
 * @returns {Array<Array>}
 */
function _sGetWorksheetRows(ws) {
  const rows = [];

  if (!ws['!ref']) {
    return rows;
  }

  // Parse the range to get max row
  const range = ws['!ref'];
  const [, endRef] = range.split(':');
  const endMatch = endRef.match(/^[A-Z]+(\d+)$/);
  const maxRow = endMatch ? parseInt(endMatch[1], 10) : 0;

  for (let rowNum = 1; rowNum <= maxRow; rowNum++) {
    const row = [];

    // Extract columns A through F
    for (let col = 0; col < 6; col++) {
      const colLetter = String.fromCharCode(65 + col); // A, B, C, D, E, F
      const cellRef = `${colLetter}${rowNum}`;
      const cell = ws[cellRef];

      if (cell && cell.v !== undefined) {
        if (cell.v instanceof Date) { row.push(cell.w || String(cell.v)); }
        else { row.push(cell.v); }
      } else {
        row.push('');
      }
    }

    rows.push(row);
  }

  return rows;
}

/**
 * Check if a row value is a section header
 * @param {*} value
 * @returns {boolean}
 */
function _sIsSectionHeader(value) {
  if (!value || typeof value !== 'string') {
    return false;
  }

  // Pattern: "1: Upper Body Power", "2: Lower Body Power", etc.
  return /^\d+:\s+.*(Power|HT|Hypertrophy)/.test(value);
}

/**
 * Check if a row has non-empty values
 * @param {Array} row
 * @returns {boolean}
 */
function _sIsNonBlankRow(row) {
  return row.some(cell => cell !== '' && cell !== undefined && cell !== null);
}

/**
 * Extract day number from section title
 * @param {string} title
 * @returns {string}
 */
function _sExtractDayNumber(title) {
  const match = title.match(/^(\d+):/);
  return match ? match[1] : '1';
}

/**
 * Extract section name from title (remove day number prefix)
 * @param {string} title
 * @returns {string}
 */
function _sExtractSectionName(title) {
  // Remove the "1: " prefix
  return title.replace(/^\d+:\s+/, '');
}

/**
 * Parse a number value, return 0 if empty or non-numeric
 * @param {*} value
 * @returns {number}
 */
function _sParseNumber(value) {
  if (value === '' || value === undefined || value === null) {
    return 0;
  }

  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

/**
 * Parse a decimal value
 * @param {*} value
 * @returns {number|null}
 */
function _sParseDecimal(value) {
  if (value === '' || value === undefined || value === null) {
    return null;
  }

  const num = Number(value);
  return isNaN(num) ? null : num;
}

/**
 * Format prescription from set scheme and weight
 * e.g., "3 sets of 3-5" + 70 -> "3x3-5(70)"
 * @param {string} setScheme
 * @param {number} weight
 * @returns {string}
 */
function _sFormatPrescription(setScheme, weight) {
  if (!setScheme || typeof setScheme !== 'string') {
    return '';
  }

  // Parse "3 sets of 3-5" -> "3x3-5"
  const match = setScheme.match(/(\d+)\s+sets?\s+of\s+([\d\-]+)/i);

  if (!match) {
    return setScheme.trim();
  }

  const sets = match[1];
  const reps = match[2];
  let prescription = `${sets}x${reps}`;

  // Append weight if > 0
  if (weight > 0) {
    prescription += `(${weight})`;
  }

  return prescription;
}

/**
 * Format note from purpose and percentage
 * @param {string} purpose
 * @param {number|null} percentage
 * @returns {string|null}
 */
function _sFormatNote(purpose, percentage) {
  const parts = [];

  if (purpose && purpose !== '') {
    parts.push(purpose);
  }

  if (percentage !== null && percentage !== undefined) {
    const percentStr = (percentage * 100).toFixed(0);
    parts.push(`${percentStr}% 1RM`);
  }

  return parts.length > 0 ? parts.join(', ') : null;
}


// ── FORMAT T: Mike Israetel Hypertrophy ───────────────────────────────────
/**
 * Format T Parser - Mike Israetel Hypertrophy Program
 * Parses the standardized XLSX format with Week 1-5 sheets
 */

/**
 * Detect if workbook is Format T
 */
function detectT(workbook) {
  const sheets = workbook.SheetNames;

  const hasWeeks = sheets.filter(s => /^Week \d+$/.test(s)).length >= 3;
  const hasMetaSheet = sheets.includes('General Overview') || sheets.includes('Exercise Table');

  if (!hasWeeks || !hasMetaSheet) return false;

  // Check first week sheet for expected headers
  const firstWeekSheet = sheets.find(s => /^Week \d+$/.test(s));
  if (!firstWeekSheet) return false;

  const ws = workbook.Sheets[firstWeekSheet];

  // Headers are in row 4 (index 4) for Week 1 or row 0 (index 0) for Week 2+
  let headerRow = null;

  // Check row 4 (Week 1 style)
  const row4A = ws['A4'];
  if (row4A && row4A.v === 'Day') headerRow = 4;

  // Check row 1 (Week 2+ style — XLSX cells are 1-based)
  const row1A = ws['A1'];
  if (row1A && row1A.v === 'Day') headerRow = 1;

  if (headerRow === null) return false;

  // Verify expected headers
  const expectedHeaders = ['Day', 'Muscle Group', 'Exercise', 'Sets', 'Reps'];
  const row = {};
  for (const col of ['A', 'B', 'C', 'D', 'E']) {
    const cell = ws[`${col}${headerRow}`];
    if (cell) row[col] = (cell.v instanceof Date) ? (cell.w || String(cell.v)) : cell.v;
  }

  return (
    row['A'] === 'Day' &&
    row['B'] === 'Muscle Group' &&
    row['C'] === 'Exercise' &&
    row['D'] === 'Sets' &&
    row['E'] === 'Reps'
  );
}

/**
 * Parse Format T workbook
 */
function parseT(workbook) {
  const sheets = workbook.SheetNames;
  const volumeLandmarks = _tExtractVolumeLandmarks(workbook);
  const weekSheets = sheets.filter(s => /^Week \d+$/.test(s)).sort((a, b) => {
    const numA = parseInt(a.match(/\d+/)[0]);
    const numB = parseInt(b.match(/\d+/)[0]);
    return numA - numB;
  });

  const weeks = [];

  for (const weekSheet of weekSheets) {
    const ws = workbook.Sheets[weekSheet];
    const weekNum = parseInt(weekSheet.match(/\d+/)[0]);

    // Determine header row and data start row
    // Week 1 has extra calculator rows, headers at row index 3 (0-indexed), data at 4
    // Week 2+ have headers at row index 0, data at 1
    // XLSX uses 0-based row addressing in cell refs like A0, A1, etc.
    let dataStartRow;
    const firstA = ws['A1'];
    if (firstA && firstA.v === 'Day') {
      // Week 2-5 style: headers at row 1 (XLSX cells are 1-based)
      dataStartRow = 2;
    } else {
      // Week 1 style: headers at row 3-4 area
      dataStartRow = 5;
    }

    const days = _tParseDays(ws, dataStartRow, volumeLandmarks);

    weeks.push({
      label: `Week ${weekNum}`,
      days: days
    });
  }

  return [{
    id: 'mike_israetel_hypertrophy',
    name: 'Mike Israetel Hypertrophy',
    format: 'T',
    athleteName: null,
    dateRange: null,
    maxes: {},
    weeks: weeks
  }];
}

/**
 * Parse days from worksheet starting at given row
 */
function _tParseDays(ws, startRow, volumeLandmarks) {
  const days = [];
  let currentDayNum = null;
  let currentExercises = [];
  let row = startRow;
  let blankRowCount = 0;

  // Find the maximum row with data to avoid infinite loops
  let maxRow = startRow;
  for (let i = startRow; i < startRow + 1000; i++) {
    const dayCell = ws[`A${i}`];
    const exerciseCell = ws[`C${i}`];
    if ((dayCell && dayCell.v) || (exerciseCell && exerciseCell.v)) {
      maxRow = i;
    }
  }

  while (row <= maxRow + 10) {
    const dayCell = ws[`A${row}`];
    const exerciseCell = ws[`C${row}`];

    const dayNum = dayCell ? dayCell.v : null;
    const hasExercise = exerciseCell && exerciseCell.v;

    // Check if row is completely empty
    if (!dayNum && !hasExercise) {
      blankRowCount++;
      row++;
      // Skip blank rows but continue looking for more days
      if (blankRowCount > 5) {
        // Too many blank rows, likely end of data
        break;
      }
      continue;
    }

    blankRowCount = 0;

    // If we have a day number, this is a new day
    if (dayNum !== null && dayNum !== undefined && dayNum !== '') {
      // Save previous day
      if (currentDayNum !== null && currentExercises.length > 0) {
        days.push({
          name: `Day ${currentDayNum}`,
          exercises: currentExercises
        });
      }

      currentDayNum = dayNum;
      currentExercises = [];

      // Day 3 is rest - skip it
      if (dayNum === 3) {
        row++;
        continue;
      }
    }

    // If we're on day 3 (rest), skip this row
    if (currentDayNum === 3) {
      row++;
      continue;
    }

    // Parse exercise from this row
    const exercise = _tParseExercise(ws, row, volumeLandmarks);
    if (exercise) {
      currentExercises.push(exercise);
    }

    row++;
  }

  // Save last day
  if (currentDayNum !== null && currentExercises.length > 0) {
    days.push({
      name: `Day ${currentDayNum}`,
      exercises: currentExercises
    });
  }

  return days;
}

/**
 * Parse a single exercise row
 */
function _tParseExercise(ws, row, volumeLandmarks) {
  const muscleCell = ws[`B${row}`];
  const exerciseCell = ws[`C${row}`];
  const setsCell = ws[`D${row}`];
  const repsCell = ws[`E${row}`];
  const weightCell = ws[`F${row}`];
  const noteCell = ws[`H${row}`];

  // Must have exercise name
  if (!exerciseCell || !exerciseCell.v) return null;

  const exerciseName = String(exerciseCell.v).trim();

  // Skip rest days
  if (exerciseName.includes('Rest') && exerciseName.includes('Conditioning')) {
    return null;
  }

  const muscle = muscleCell ? String(muscleCell.v).trim() : '';
  const sets = setsCell ? setsCell.v : null;
  const reps = repsCell ? repsCell.v : null;
  const weight = weightCell ? weightCell.v : null;
  const notes = noteCell ? String(noteCell.v).trim() : '';

  // Build prescription
  const prescription = _tBuildPrescription(sets, reps, weight);

  if (!prescription) return null;

  // Build note
  const note = _tBuildNote(muscle, notes, exerciseName, volumeLandmarks);

  return {
    name: exerciseName,
    prescription: prescription,
    note: note,
    lifterNote: null,
    loggedWeight: null
  };
}

/**
 * Build prescription string from sets, reps, weight
 */
function _tBuildPrescription(sets, reps, weight) {
  if (sets === null || reps === null) return null;

  const setsStr = String(sets).toUpperCase().trim();
  const repsStr = String(reps).toUpperCase().trim();
  const weightVal = weight !== null ? weight : 0;

  let weightStr;
  if (String(weightVal).toLowerCase() === 'body' || weightVal === 0 || weightVal === '0') {
    weightStr = 'BW';
  } else {
    weightStr = String(weightVal);
  }

  // AFAP x{reps}({weight})
  if (setsStr === 'AFAP') {
    return `AFAP x${repsStr}(${weightStr})`;
  }

  // {sets}xAMAP({weight})
  if (repsStr === 'AMAP') {
    return `${setsStr}xAMAP(${weightStr})`;
  }

  // {sets}x{reps}({weight})
  return `${setsStr}x${repsStr}(${weightStr})`;
}

/**
 * Extract MRV/MEV volume landmarks from General Overview or Exercise Table sheet.
 * Returns map: { "chest": { mrv: 20, mev: 10 }, ... } (keys lowercased)
 */
function _tExtractVolumeLandmarks(workbook) {
  const volumeMap = {};
  const metaSheet = workbook.SheetNames.find(s => s === 'General Overview' || s === 'Exercise Table');
  if (!metaSheet) return volumeMap;

  const ws = workbook.Sheets[metaSheet];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Find header row with MRV/MEV columns
  let headerRow = -1;
  let muscleCol = -1, mrvCol = -1, mevCol = -1;

  for (let i = 0; i < Math.min(30, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = row[c] ? String(row[c]).trim().toLowerCase() : '';
      if (val === 'mrv' || val.includes('maximum recoverable')) { mrvCol = c; headerRow = i; }
      if (val === 'mev' || val.includes('minimum effective')) { mevCol = c; headerRow = i; }
      if (/^(muscle\s*group|muscle|body\s*part)$/.test(val)) muscleCol = c;
    }
    if (mrvCol >= 0 || mevCol >= 0) break;
  }

  if (headerRow < 0 || muscleCol < 0) return volumeMap;

  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    const muscle = row[muscleCol] ? String(row[muscleCol]).trim() : '';
    if (!muscle || muscle.length < 2) continue;
    // Stop at blank muscle rows
    if (!muscle) break;

    const mrv = mrvCol >= 0 && row[mrvCol] != null ? row[mrvCol] : null;
    const mev = mevCol >= 0 && row[mevCol] != null ? row[mevCol] : null;

    if (mrv != null || mev != null) {
      volumeMap[muscle.toLowerCase()] = {
        mrv: typeof mrv === 'number' ? mrv : null,
        mev: typeof mev === 'number' ? mev : null
      };
    }
  }

  return volumeMap;
}

/**
 * Build note string from muscle group and other notes
 */
function _tBuildNote(muscle, notes, exerciseName, volumeLandmarks) {
  const parts = [];

  if (muscle) {
    parts.push(muscle);
  }

  // Check if exercise name contains SS (superset indicator)
  if (exerciseName && exerciseName.includes('SS')) {
    parts.push('Superset');
  }

  // Extract special notes
  if (notes) {
    if (notes.includes('Superset') || notes.includes('SS')) {
      if (!parts.includes('Superset')) {
        parts.push('Superset');
      }
    }
    if (notes.includes('RPE') || notes.includes('RIR')) {
      parts.push('RPE/RIR: check notes');
    }
  }

  // Add MRV/MEV volume landmarks if available for this muscle group
  if (muscle && volumeLandmarks) {
    const vl = volumeLandmarks[muscle.toLowerCase()];
    if (vl) {
      const vlParts = [];
      if (vl.mrv != null) vlParts.push('MRV: ' + vl.mrv + ' sets');
      if (vl.mev != null) vlParts.push('MEV: ' + vl.mev + ' sets');
      if (vlParts.length > 0) parts.push(vlParts.join(' | '));
    }
  }

  return parts.length > 0 ? parts.join(', ') : null;
}

// ── SHARED: extract 1RM maxes from a calculator/maxes sheet ──────────────────
// Used by Format U (GZCL) and Format V (PH3) — both have dedicated maxes sheets
function _extractMaxesFromSheet(wb) {
  const maxes = {};
  for (const sn of wb.SheetNames) {
    if (/^(max(es)?|inputs?|1rm|calculator|calc|max\s*weight)/i.test(sn.trim())) {
      const ws = wb.Sheets[sn];
      if (!ws) continue;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
        .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
      for (let i = 0; i < Math.min(30, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        for (let c = 0; c < row.length - 1; c++) {
          const label = row[c]; if (label == null) continue;
          const n = String(label).trim().toLowerCase();
          const findNum = () => {
            for (let nc = c + 1; nc < Math.min(c + 4, row.length); nc++) {
              if (typeof row[nc] === 'number' && row[nc] > 0) return row[nc];
            }
            return null;
          };
          if (/squat/i.test(n) && !maxes['Squat']) { const v = findNum(); if (v) maxes['Squat'] = v; }
          else if (/bench/i.test(n) && !maxes['Bench Press']) { const v = findNum(); if (v) maxes['Bench Press'] = v; }
          else if (/dead/i.test(n) && !maxes['Deadlift']) { const v = findNum(); if (v) maxes['Deadlift'] = v; }
          else if (/ohp|overhead|press/i.test(n) && !/bench/i.test(n) && !maxes['OHP']) { const v = findNum(); if (v) maxes['OHP'] = v; }
        }
      }
      break;
    }
  }
  return maxes;
}

// ── FORMAT U DETECTION (GZCL T1/T2/T3 tier-based) ────────────────────────────
// Detects GZCL family programs: GZCLP, Jacked & Tan 2.0, The Rippler, UHF, General Gainz
// Key marker: T1/T2/T3 tier labels in col 0 or col 1, combined with Sets/Reps/Weight headers
// Won't false-positive on Format T (requires Day/Muscle Group/Exercise/Sets/Reps exact headers)
// Won't false-positive on Format E (no SxR text patterns or alternating weight/rep pairs needed)

function detectU(wb) {
  if (!wb || !wb.SheetNames) return false;
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|maxes?|inputs?|output|calculator|1rm)$/i;

  for (let si = 0; si < Math.min(8, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows || rows.length < 3) continue;

    let hasTier = false;
    let hasSetsRepsWeight = false;
    let hasTierHeader = false;

    for (let i = 0; i < Math.min(25, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c <= Math.min(1, row.length - 1); c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        // T1/T2/T3 as standalone cell values (not embedded in longer text)
        if (/^T[123]$/i.test(str)) { hasTier = true; }
      }
      // Check for "Tier" column header
      for (let c = 0; c < Math.min(6, row.length); c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim().toLowerCase();
        if (str === 'tier') hasTierHeader = true;
      }
      // Check for Sets/Reps/Weight headers in same row
      const strs = row.map(_normalizeCell);
      const hasSets = strs.some(s => s === 'sets' || s === 'set');
      const hasReps = strs.some(s => s === 'reps' || s === 'rep');
      const hasWeight = strs.some(s => s === 'weight' || s === 'load' || s === 'lbs' || s === 'kg');
      if (hasSets && hasReps && hasWeight) hasSetsRepsWeight = true;
    }

    // Need T1/T2/T3 markers (or "Tier" header) AND Sets/Reps/Weight column structure
    if ((hasTier || hasTierHeader) && hasSetsRepsWeight) return true;
  }

  // Variant B (single-sheet horizontal): look for "T1:" / "T2:" / "T3:" prefixes on exercise names
  for (let si = 0; si < Math.min(4, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows || rows.length < 5) continue;

    let tierPrefixCount = 0;
    let hasWeekCols = false;
    for (let i = 0; i < Math.min(40, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const col0 = row[0]; if (col0 == null) continue;
      const str = String(col0).trim();
      if (/^T[123]\s*[:\-]/i.test(str)) tierPrefixCount++;
      // Check for Week N column headers
      for (let c = 1; c < Math.min(row.length, 30); c++) {
        const hv = row[c]; if (hv == null) continue;
        if (/^Week\s*\d/i.test(String(hv).trim())) hasWeekCols = true;
      }
    }
    if (tierPrefixCount >= 3 && hasWeekCols) return true;
  }

  return false;
}

// ── FORMAT U PARSER (GZCL family) ─────────────────────────────────────────────

function parseU(wb) {
  // Try Variant A (multi-sheet, one per week) first, then Variant B (single-sheet horizontal)
  let result = _parseU_multiSheet(wb);
  if (result && result.length > 0 && result[0].weeks.length > 0) return result;

  result = _parseU_singleSheet(wb);
  if (result && result.length > 0 && result[0].weeks.length > 0) return result;

  return [];
}

// Variant A: Multi-sheet GZCLP — one sheet per week with workout blocks
function _parseU_multiSheet(wb) {
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|maxes?|inputs?|output|calculator|1rm)$/i;

  // Find week sheets
  const weekSheets = [];
  for (const sn of wb.SheetNames) {
    if (/^Week\s*\d+/i.test(sn.trim())) {
      weekSheets.push(sn);
    }
  }

  // If no explicit week sheets, try non-skip sheets that have T1/T2/T3 content
  if (weekSheets.length === 0) {
    for (const sn of wb.SheetNames) {
      if (SKIP.test(sn.trim())) continue;
      const ws = wb.Sheets[sn];
      if (!ws) continue;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      let hasTier = false;
      for (let i = 0; i < Math.min(30, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        for (let c = 0; c <= 1; c++) {
          if (row[c] && /^T[123]$/i.test(String(row[c]).trim())) { hasTier = true; break; }
        }
        if (hasTier) break;
      }
      if (hasTier) weekSheets.push(sn);
    }
  }

  if (weekSheets.length === 0) return [];

  const maxes = _extractMaxesFromSheet(wb);

  const weeks = [];
  for (let wi = 0; wi < weekSheets.length; wi++) {
    const sn = weekSheets[wi];
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

    // Find header row (contains Sets/Reps/Weight or Tier)
    let tierCol = -1, exCol = -1, setsCol = -1, repsCol = -1, weightCol = -1;
    let headerRow = -1;

    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      const si2 = strs.indexOf('sets');
      const ri = strs.indexOf('reps');
      if (si2 >= 0 && ri >= 0) {
        setsCol = si2;
        repsCol = ri;
        headerRow = i;
        // Find weight column
        for (let c = 0; c < strs.length; c++) {
          if (/^(weight|load|lbs|kg)$/.test(strs[c])) { weightCol = c; break; }
        }
        // Find tier column
        for (let c = 0; c < strs.length; c++) {
          if (strs[c] === 'tier') { tierCol = c; break; }
        }
        // Find exercise column
        for (let c = 0; c < strs.length; c++) {
          if (/^(exercise|movement|lift)$/.test(strs[c])) { exCol = c; break; }
        }
        break;
      }
    }

    // If no explicit header, infer from T1/T2/T3 positions
    if (headerRow === -1) {
      for (let i = 0; i < Math.min(15, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        for (let c = 0; c <= 1; c++) {
          if (row[c] && /^T[123]$/i.test(String(row[c]).trim())) {
            tierCol = c;
            exCol = c + 1;
            // Look for numeric columns after exercise
            for (let nc = exCol + 1; nc < Math.min(row.length, exCol + 5); nc++) {
              if (typeof row[nc] === 'number') {
                if (setsCol === -1) setsCol = nc;
                else if (repsCol === -1) repsCol = nc;
                else if (weightCol === -1) { weightCol = nc; break; }
              }
            }
            headerRow = i;
            break;
          }
        }
        if (headerRow >= 0) break;
      }
    }

    // Parse workout blocks — separated by blank rows or workout labels
    const days = [];
    let currentDay = null;
    let dayNum = 0;

    for (let i = (headerRow >= 0 ? headerRow : 0); i < rows.length; i++) {
      const row = rows[i]; if (!row) {
        // Blank row may separate workout blocks — don't close current day yet
        continue;
      }

      // Check for workout label (A1/A2/B1/B2 or Mon/Wed/Fri etc.)
      const col0 = row[0] != null ? String(row[0]).trim() : '';
      const col1 = row[1] != null ? String(row[1]).trim() : '';
      const isWorkoutLabel = /^(Workout\s*)?[A-D][1-4]$/i.test(col0) ||
                              /^(Workout\s*)?[A-D][1-4]$/i.test(col1) ||
                              DAYS.some(d => col0.toLowerCase().startsWith(d)) ||
                              /^Day\s*\d/i.test(col0);

      if (isWorkoutLabel) {
        if (currentDay && currentDay.exercises.length > 0) {
          days.push(currentDay);
        }
        dayNum++;
        const label = /^(Workout\s*)?[A-D][1-4]$/i.test(col0) ? col0.replace(/^Workout\s*/i, '') :
                      /^(Workout\s*)?[A-D][1-4]$/i.test(col1) ? col1.replace(/^Workout\s*/i, '') :
                      col0;
        currentDay = { name: label || ('Day ' + dayNum), exercises: [] };
        continue;
      }

      // Check for tier + exercise data
      let tier = null, exName = null, sets = null, reps = null, weight = null;

      if (tierCol >= 0 && row[tierCol] != null) {
        const tv = String(row[tierCol]).trim();
        if (/^T[123]$/i.test(tv)) tier = tv.toUpperCase();
      }

      if (exCol >= 0 && row[exCol] != null) {
        exName = String(row[exCol]).trim();
      } else if (tierCol >= 0 && row[tierCol + 1] != null) {
        exName = String(row[tierCol + 1]).trim();
      }

      if (!exName || exName.length < 2) continue;
      // Skip header rows
      if (/^(exercise|movement|lift|tier)$/i.test(exName)) continue;
      // Skip non-exercise content
      if (/^(sets|reps|weight|load|logged|notes)$/i.test(exName)) continue;

      if (setsCol >= 0 && row[setsCol] != null) sets = row[setsCol];
      if (repsCol >= 0 && row[repsCol] != null) reps = row[repsCol];
      if (weightCol >= 0 && row[weightCol] != null) weight = row[weightCol];

      // Build prescription
      const prescription = _uBuildPrescription(sets, reps, weight);
      const note = tier ? tier : '';

      // Create a default day if none exists yet
      if (!currentDay) {
        dayNum++;
        currentDay = { name: 'Day ' + dayNum, exercises: [] };
      }

      currentDay.exercises.push({
        name: _hCapitalizeName(exName),
        prescription: prescription,
        note: note,
        lifterNote: '',
        loggedWeight: '',
        supersetGroup: null
      });
    }

    // Push last day
    if (currentDay && currentDay.exercises.length > 0) {
      days.push(currentDay);
    }

    if (days.length > 0) {
      weeks.push({ label: sn.trim() || ('Week ' + (wi + 1)), days });
    }
  }

  if (weeks.length === 0) return [];

  // Determine program name from first sheet or workbook title
  let progName = 'GZCL Program';
  for (const sn of wb.SheetNames) {
    const lo = sn.toLowerCase();
    if (lo.includes('gzclp')) { progName = 'GZCLP'; break; }
    if (lo.includes('jacked') || lo.includes('j&t') || lo.includes('jt2')) { progName = 'Jacked & Tan 2.0'; break; }
    if (lo.includes('rippler')) { progName = 'The Rippler'; break; }
    if (lo.includes('uhf')) { progName = 'GZCL UHF'; break; }
    if (lo.includes('general gainz') || lo.includes('gg')) { progName = 'General Gainz'; break; }
  }

  const blocks = [{
    id: 'gzcl_' + Date.now(),
    name: progName,
    format: 'U',
    athleteName: '',
    dateRange: '',
    maxes: maxes,
    weeks: weeks
  }];

  deduplicateExerciseNames(blocks);
  return blocks;
}

// Variant B: Single-sheet horizontal — exercises as rows, weeks as column groups
function _parseU_singleSheet(wb) {
  const SKIP = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|maxes?|inputs?|output|calculator|1rm)$/i;

  for (let si = 0; si < Math.min(4, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows || rows.length < 5) continue;

    // Find week column positions from header area
    const weekCols = []; // [{col, label}]
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 1; c < row.length; c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        if (/^Week\s*\d+/i.test(str)) {
          weekCols.push({ col: c, label: str });
        }
      }
    }

    if (weekCols.length === 0) continue;

    // Parse exercise rows with day separators
    const weekData = weekCols.map(wc => ({ label: wc.label, days: [] }));
    let currentDayName = 'Day 1';
    let dayExercises = weekCols.map(() => []);
    let dayNum = 0;

    for (let i = 2; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const col0 = row[0] != null ? String(row[0]).trim() : '';

      // Day separator (e.g., "Day 1", "Push A", "Lower", workout label)
      if (col0 && !(/^T[123]\s*[:\-]/i.test(col0)) &&
          (/^Day\s*\d/i.test(col0) || /^(Push|Pull|Legs|Upper|Lower|Full)/i.test(col0) ||
           DAYS.some(d => col0.toLowerCase().startsWith(d)))) {
        // Save previous day's exercises
        if (dayNum > 0) {
          for (let w = 0; w < weekCols.length; w++) {
            if (dayExercises[w].length > 0) {
              weekData[w].days.push({ name: currentDayName, exercises: dayExercises[w] });
            }
          }
        }
        dayNum++;
        currentDayName = col0;
        dayExercises = weekCols.map(() => []);
        continue;
      }

      // Exercise row with tier prefix (e.g., "T1: Squat")
      if (!col0 || col0.length < 2) continue;

      let tier = '';
      let exName = col0;
      const tierMatch = col0.match(/^(T[123])\s*[:\-]\s*(.+)/i);
      if (tierMatch) {
        tier = tierMatch[1].toUpperCase();
        exName = tierMatch[2].trim();
      }

      if (!exName || exName.length < 2) continue;
      // Skip header-like rows
      if (/^(exercise|movement|lift)$/i.test(exName)) continue;

      // Extract per-week values
      for (let w = 0; w < weekCols.length; w++) {
        const wCol = weekCols[w].col;
        const weight = row[wCol];
        // Look for sets/reps in adjacent cells or use SxR text
        let prescription = null;
        if (weight != null) {
          const wStr = String(weight).trim();
          // If cell contains SxR text directly
          if (/^\d+\s*x\s*\d/i.test(wStr)) {
            prescription = wStr;
          } else if (typeof weight === 'number' && weight > 0) {
            prescription = String(Math.round(weight));
          } else if (wStr) {
            prescription = wStr;
          }
        }

        if (!dayExercises[w]) dayExercises[w] = [];
        dayExercises[w].push({
          name: _hCapitalizeName(exName),
          prescription: prescription,
          note: tier,
          lifterNote: '',
          loggedWeight: '',
          supersetGroup: null
        });
      }
    }

    // Push last day
    if (dayNum >= 0) {
      for (let w = 0; w < weekCols.length; w++) {
        if (dayExercises[w].length > 0) {
          weekData[w].days.push({ name: currentDayName, exercises: dayExercises[w] });
        }
      }
    }

    const weeks = weekData.filter(w => w.days.length > 0);
    if (weeks.length === 0) continue;

    let progName = 'GZCL Program';
    for (const s of wb.SheetNames) {
      const lo = s.toLowerCase();
      if (lo.includes('jacked') || lo.includes('j&t') || lo.includes('jt2')) { progName = 'Jacked & Tan 2.0'; break; }
      if (lo.includes('rippler')) { progName = 'The Rippler'; break; }
      if (lo.includes('uhf')) { progName = 'GZCL UHF'; break; }
      if (lo.includes('gzclp')) { progName = 'GZCLP'; break; }
    }

    const blocks = [{
      id: 'gzcl_h_' + Date.now(),
      name: progName,
      format: 'U',
      athleteName: '',
      dateRange: '',
      maxes: {},
      weeks: weeks
    }];

    deduplicateExerciseNames(blocks);
    return blocks;
  }

  return [];
}

// Helper: build GZCL prescription string
function _uBuildPrescription(sets, reps, weight) {
  const s = sets != null ? String(sets).trim() : '';
  const r = reps != null ? String(reps).trim() : '';
  const w = weight != null ? weight : null;

  if (!s && !r) return null;

  let setsStr = s || '1';
  let repsStr = r || '?';

  // Handle percentage weights
  if (w != null) {
    const wStr = String(w).trim();
    if (!wStr || wStr === '0') {
      return setsStr + 'x' + repsStr;
    }
    if (wStr.includes('%') || (typeof w === 'number' && w > 0 && w < 1)) {
      const pct = typeof w === 'number' && w < 1 ? Math.round(w * 100) : wStr.replace('%', '');
      return setsStr + 'x' + repsStr + ' @' + pct + '%';
    }
    if (typeof w === 'number' && w > 0) {
      return setsStr + 'x' + repsStr + '(' + Math.round(w) + ')';
    }
    // RPE or text weight
    if (/rpe/i.test(wStr)) {
      return setsStr + 'x' + repsStr + ' ' + wStr;
    }
    return setsStr + 'x' + repsStr + '(' + wStr + ')';
  }

  return setsStr + 'x' + repsStr;
}

// ── FORMAT V DETECTION (Layne Norton PH3 — phase + RIR column structure) ─────
// Detects PH3 and similar programs with:
// - A calculator/1RM sheet
// - Multiple week-numbered sheets
// - Sets/Reps/Load columns with RIR/RPE column
// Won't false-positive on Format H (single-sheet percentage, no week sheets)
// Won't false-positive on Format Q (BBM uses specific "Bridge"/"Week N - Stress" naming)
// Won't false-positive on Format T (requires Day/Muscle Group/Exercise exact headers)

function detectV(wb) {
  if (!wb || !wb.SheetNames) return false;

  // Need a calculator/1RM sheet
  let hasCalcSheet = false;
  for (const sn of wb.SheetNames) {
    const lo = sn.toLowerCase().trim();
    if (lo.includes('calculator') || lo.includes('1rm') || lo === 'calc' ||
        lo.includes('max weight') || lo.includes('maxes')) {
      // Verify it actually has 1RM-related content
      const ws = wb.Sheets[sn];
      if (!ws) continue;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      for (let i = 0; i < Math.min(15, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        const joined = row.map(c => String(c || '')).join(' ').toLowerCase();
        if (joined.includes('1rm') || joined.includes('max') || joined.includes('squat') ||
            joined.includes('bench') || joined.includes('deadlift')) {
          hasCalcSheet = true;
          break;
        }
      }
      if (hasCalcSheet) break;
    }
  }

  if (!hasCalcSheet) return false;

  // Need multiple week-numbered sheets
  let weekCount = 0;
  for (const sn of wb.SheetNames) {
    if (/^Week\s*\d/i.test(sn.trim())) weekCount++;
  }

  if (weekCount < 3) return false;

  // Verify at least one week sheet has RIR/RPE column OR percentage load values
  for (const sn of wb.SheetNames) {
    if (!/^Week\s*\d/i.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

    let hasRIR = false;
    let hasPercentLoad = false;
    let hasDayLabel = false;
    let hasExerciseHeader = false;
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      if (strs.some(s => s === 'rir' || s === 'rpe' || s === 'rir/rpe')) hasRIR = true;
      if (strs.some(s => s.includes('%') && /\d/.test(s))) hasPercentLoad = true;
      if (row[0] && /^Day\s*\d/i.test(String(row[0]).trim())) hasDayLabel = true;
      // Require an exercise/movement/lift header to distinguish from Format H/Q
      if (strs.some(s => s === 'exercise' || s === 'movement' || s === 'lift')) hasExerciseHeader = true;
    }

    if ((hasRIR || hasPercentLoad) && hasDayLabel && hasExerciseHeader) return true;
  }

  return false;
}

// ── FORMAT V PARSER (Layne Norton PH3) ───────────────────────────────────────

function parseV(wb) {
  const maxes = _extractMaxesFromSheet(wb);

  // Parse week sheets
  const weekSheets = wb.SheetNames.filter(sn => /^Week\s*\d/i.test(sn.trim()));
  weekSheets.sort((a, b) => {
    const na = parseInt(a.match(/\d+/)[0]);
    const nb = parseInt(b.match(/\d+/)[0]);
    return na - nb;
  });

  const weeks = [];
  for (const sn of weekSheets) {
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);

    // Find column headers: Exercise/Sets/Reps/Load/RIR
    let exCol = -1, setsCol = -1, repsCol = -1, loadCol = -1, rirCol = -1;
    let headerRow = -1;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(_normalizeCell);
      const hasEx = strs.some(s => s === 'exercise' || s === 'movement' || s === 'lift');
      const hasSets = strs.some(s => s === 'sets' || s === 'set');
      const hasReps = strs.some(s => s === 'reps' || s === 'rep');
      if (hasEx && hasSets && hasReps) {
        headerRow = i;
        for (let c = 0; c < strs.length; c++) {
          if (/^(exercise|movement|lift)$/.test(strs[c])) exCol = c;
          else if (/^(sets?|set)$/.test(strs[c])) setsCol = c;
          else if (/^(reps?|rep)$/.test(strs[c])) repsCol = c;
          else if (/^(load|weight|intensity|%)$/.test(strs[c])) loadCol = c;
          else if (/^(rir|rpe|rir\/rpe)$/.test(strs[c])) rirCol = c;
        }
        break;
      }
    }

    // If no explicit header row, try to infer from Day N labels
    // Validate inferred columns contain exercise-like data before using them
    if (headerRow === -1) {
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const row = rows[i]; if (!row) continue;
        if (row[0] && /^Day\s*\d/i.test(String(row[0]).trim())) {
          // Look at next non-blank row to verify it has exercise data
          let verified = false;
          for (let j = i + 1; j < Math.min(i + 5, rows.length); j++) {
            const dataRow = rows[j]; if (!dataRow) continue;
            // Try col 0 first, then col 1 for exercise name
            for (let ec = 0; ec <= 1; ec++) {
              if (dataRow[ec] != null) {
                const val = String(dataRow[ec]).trim();
                if (val.length >= 3 && !/^(Day|Week|Phase|Set|Rep|Load|Note|Total)/i.test(val) && !/^\d+$/.test(val)) {
                  exCol = ec;
                  setsCol = ec + 1;
                  repsCol = ec + 2;
                  loadCol = ec + 3;
                  rirCol = ec + 4;
                  headerRow = i - 1;
                  verified = true;
                  break;
                }
              }
            }
            if (verified) break;
          }
          break;
        }
      }
    }

    // Extract phase label from sheet content
    let phaseLabel = '';
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const joined = row.map(c => String(c || '')).join(' ');
      if (/accumulation/i.test(joined)) phaseLabel = 'Accumulation';
      else if (/transition/i.test(joined)) phaseLabel = 'Transition';
      else if (/intensification/i.test(joined)) phaseLabel = 'Intensification';
      else if (/peaking/i.test(joined)) phaseLabel = 'Peaking';
      if (phaseLabel) break;
    }

    // Parse days within the sheet
    const days = [];
    let currentDay = null;
    let curSupersetGroup = null;
    const _isSupersetLabel = (s) => /^(superset|super\s*set|ss|giant\s*set|circuit)$/i.test(s) || /^\d+\s*rounds?$/i.test(s);

    const startRow = headerRow >= 0 ? headerRow + 1 : 0;
    for (let i = startRow; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;

      const col0 = row[0] != null ? String(row[0]).trim() : '';

      // Day separator
      if (/^Day\s*\d/i.test(col0)) {
        if (currentDay && currentDay.exercises.length > 0) {
          days.push(currentDay);
        }
        currentDay = { name: col0, exercises: [] };
        curSupersetGroup = null;
        continue;
      }

      // Exercise row
      const exName = exCol >= 0 && row[exCol] != null ? String(row[exCol]).trim() : col0;
      if (!exName || exName.length < 2) continue;
      // Skip header-like or summary rows
      if (/^(exercise|movement|lift|sets|reps|load|rir|rpe|total|daily|weekly|phase|notes?)$/i.test(exName)) continue;

      // Detect superset/circuit grouping labels in column 0 or exercise column
      if (_isSupersetLabel(col0) || _isSupersetLabel(exName)) {
        const matchText = _isSupersetLabel(col0) ? col0 : exName;
        const roundsMatch = matchText.match(/^(\d+)\s*rounds?$/i);
        const rounds = roundsMatch ? parseInt(roundsMatch[1]) : 1;
        const label = rounds > 1 ? rounds + ' rounds' : 'superset';
        curSupersetGroup = { label: label + '_' + (currentDay ? currentDay.exercises.length : 0), rounds, startIdx: currentDay ? currentDay.exercises.length : 0 };
        continue;
      }

      const sets = setsCol >= 0 && row[setsCol] != null ? row[setsCol] : null;
      const reps = repsCol >= 0 && row[repsCol] != null ? row[repsCol] : null;
      const load = loadCol >= 0 && row[loadCol] != null ? row[loadCol] : null;
      const rir = rirCol >= 0 && row[rirCol] != null ? row[rirCol] : null;

      if (sets == null && reps == null && load == null) {
        curSupersetGroup = null;
        continue;
      }

      // Build prescription
      const prescription = _vBuildPrescription(sets, reps, load, rir);

      // Determine note (warmup, AMRAP, etc.)
      let note = '';
      const exLo = exName.toLowerCase();
      if (exLo.includes('warmup') || exLo.includes('warm-up') || exLo.includes('warm up')) note = 'Warmup';

      // Clean exercise name (remove warmup/work annotations)
      let cleanName = exName.replace(/\s*\(?(warm\s*-?\s*up|work|amrap)\)?/gi, '').trim();
      if (!cleanName) cleanName = exName;

      if (!currentDay) {
        currentDay = { name: 'Day 1', exercises: [] };
      }

      currentDay.exercises.push({
        name: _hCapitalizeName(cleanName),
        prescription: prescription,
        note: note,
        lifterNote: '',
        loggedWeight: '',
        supersetGroup: curSupersetGroup || null
      });
    }

    // Push last day
    if (currentDay && currentDay.exercises.length > 0) {
      days.push(currentDay);
    }

    if (days.length > 0) {
      const weekLabel = phaseLabel ? sn.trim() + ' — ' + phaseLabel : sn.trim();
      weeks.push({ label: weekLabel, days });
    }
  }

  if (weeks.length === 0) return [];

  const blocks = [{
    id: 'ph3_' + Date.now(),
    name: 'Layne Norton PH3',
    format: 'V',
    athleteName: '',
    dateRange: '',
    maxes: maxes,
    weeks: weeks
  }];

  deduplicateExerciseNames(blocks);
  return blocks;
}

// Helper: build PH3 prescription string
function _vBuildPrescription(sets, reps, load, rir) {
  const s = sets != null ? String(sets).trim() : '';
  const r = reps != null ? String(reps).trim() : '';
  const l = load != null ? load : null;
  const rirVal = rir != null ? String(rir).trim() : '';

  if (!s && !r) return null;

  const setsStr = s || '1';
  let repsStr = r || '?';

  // Handle AMRAP notation (5+, MR, etc.)
  if (/\+/.test(repsStr) || /amrap|mr/i.test(repsStr)) {
    repsStr = repsStr.replace(/\s+/g, '');
  }

  let parts = setsStr + 'x' + repsStr;

  // Handle load/weight
  if (l != null) {
    const lStr = String(l).trim();
    if (lStr) {
      // Percentage-based load
      if (lStr.includes('%') || (typeof l === 'number' && l > 0 && l < 1)) {
        const pct = typeof l === 'number' && l < 1 ? Math.round(l * 100) : lStr.replace('%', '');
        parts += ' @' + pct + '%';
      } else if (/^rpe\s*\d/i.test(lStr)) {
        // RPE-based accessories
        parts += ' ' + lStr;
      } else if (typeof l === 'number' && l > 0) {
        parts += '(' + Math.round(l) + ')';
      } else if (lStr && lStr !== '0') {
        parts += ' ' + lStr;
      }
    }
  }

  // Append RIR/RPE
  if (rirVal && rirVal !== '0' && !/rpe|rir/i.test(parts)) {
    if (/^\d+$/.test(rirVal)) {
      parts += ' RIR ' + rirVal;
    } else {
      parts += ' ' + rirVal;
    }
  }

  return parts;
}

// ── FORMAT W PARSER (Simple Specific Scientific coaching template) ────────────
// Layout: HOMEPAGE_FAQ for athlete/maxes; PROGRAM has 4 horizontal week groups
// (stride=7 columns), 7 days per section, multi-row exercises combined.
function parseW(wb) {
  // ── Step 1: Extract from HOMEPAGE_FAQ ──────────────────────────────────────
  let athleteName = '';
  let maxes = { squat: null, bench: null, deadlift: null };
  let programName = 'Simple Specific Scientific';

  for (const sn of wb.SheetNames) {
    if (!/homepage|faq/i.test(sn)) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
    // Row 1 (0-indexed): title (e.g., "SIMPLE SPECIFIC SCIENTIFIC")
    if (rows[1] && rows[1][0]) {
      const raw = String(rows[1][0]).trim();
      programName = raw.replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
    }
    // Row 3 (0-indexed): Athlete name at col 1
    if (rows[3] && rows[3][1] != null) athleteName = String(rows[3][1]).trim();
    // Rows 8-10 (0-indexed): SQUAT/DEADLIFT/BENCH 1RM-LB at col 11 (M column)
    if (rows[8]  && rows[8][11]  != null) maxes.squat    = Number(rows[8][11]);
    if (rows[9]  && rows[9][11]  != null) maxes.deadlift = Number(rows[9][11]);
    if (rows[10] && rows[10][11] != null) maxes.bench    = Number(rows[10][11]);
    break;
  }

  // ── Step 2: Parse PROGRAM sheet ────────────────────────────────────────────
  const programSn = wb.SheetNames.find(sn => /^program$/i.test(sn.trim()));
  if (!programSn) return null;

  const ws = wb.Sheets[programSn];
  const rows = _w_cleanRows(ws);
  const sheetRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const colOffset = sheetRange.s.c;
  const rowOffset = sheetRange.s.r;
  if (!rows || rows.length < 5) return null;

  // Find week header row: ≥2 "WEEK N" labels in a single row
  let weekHeaderRow = -1;
  const weekPositions = []; // [{col, label}]
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    const wpos = [];
    for (let c = 0; c < row.length; c++) {
      if (row[c] && /^WEEK\s*\d/i.test(String(row[c]).trim()))
        wpos.push({ col: c, label: String(row[c]).trim() });
    }
    if (wpos.length >= 2) { weekHeaderRow = i; weekPositions.push(...wpos); break; }
  }
  if (weekHeaderRow === -1 || weekPositions.length < 2) return null;

  const base0 = weekPositions[0].col;
  const stride = weekPositions.length > 1 ? weekPositions[1].col - weekPositions[0].col : 7;

  // Detect column offsets within a week group from first sub-header row
  let exOff = 0, setsOff = 1, repsOff = 2, loadOff = 3, topSetOff = -1, ratingOff = -1;
  for (let i = weekHeaderRow + 1; i < Math.min(weekHeaderRow + 5, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    let found = 0;
    for (let c = base0; c < base0 + stride && c < row.length; c++) {
      const s = String(row[c] || '').toLowerCase().trim();
      if (s === 'exercise')                            { exOff   = c - base0; found++; }
      else if (s === 'sets')                           { setsOff = c - base0; found++; }
      else if (s === 'reps')                           { repsOff = c - base0; found++; }
      else if (s === 'load' || s === 'load/rpe' || s === 'load / rpe') { loadOff = c - base0; found++; }
      else if (s === 'top set')                        { topSetOff = c - base0; }
      else if (s === 'rating')                         { ratingOff = c - base0; }
    }
    if (found >= 3) break;
  }

  // Find day sections: rows where "exercise" appears at ≥2 week column positions
  const daySections = []; // [{subHeaderRow, dayNum}]
  let dayNum = 0;
  for (let i = weekHeaderRow + 1; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    let exCount = 0;
    for (const wp of weekPositions) {
      if (String(row[wp.col + exOff] || '').toLowerCase().trim() === 'exercise') exCount++;
    }
    if (exCount >= 2) { dayNum++; daySections.push({ subHeaderRow: i, dayNum }); }
  }
  if (daySections.length === 0) return null;

  // ── Parse exercise groups per day section ──────────────────────────────────
  const dayData = []; // [{dayNum, weekExerciseSets: Array[nWeeks][]}]

  for (let d = 0; d < daySections.length; d++) {
    const startRow = daySections[d].subHeaderRow + 1;
    const endRow   = d + 1 < daySections.length ? daySections[d + 1].subHeaderRow - 1 : rows.length;

    // Build exercise groups from first week column (primary structure reference)
    const exerciseGroups = []; // [{name, rows: [arrRowIdx]}]
    for (let r = startRow; r < endRow; r++) {
      const row = rows[r]; if (!row) continue;

      // Skip rows with no sets/reps data in any week
      let hasData = false;
      for (const wp of weekPositions) {
        if (row[wp.col + setsOff] != null || row[wp.col + repsOff] != null) { hasData = true; break; }
      }
      if (!hasData) continue;

      const nameRaw = row[base0 + exOff];
      const exName  = nameRaw ? String(nameRaw).trim() : '';

      if (/^rest\s*day$/i.test(exName)) continue; // skip REST DAY rows

      const isNamed = exName && /[a-zA-Z]{2,}/.test(exName) && !/^(exercise|week)/i.test(exName);
      if (isNamed) {
        const prev = exerciseGroups.length > 0 ? exerciseGroups[exerciseGroups.length - 1] : null;
        // Same name as previous = backdown row with name repeated → combine
        if (prev && prev.name === exName) {
          prev.rows.push(r);
        } else {
          exerciseGroups.push({ name: exName, rows: [r] });
        }
      } else if (exerciseGroups.length > 0 && !exName) {
        // Anonymous continuation row → belongs to previous exercise
        exerciseGroups[exerciseGroups.length - 1].rows.push(r);
      }
    }

    // Build per-week exercise lists from exercise groups
    const weekExerciseSets = weekPositions.map(() => []);
    for (const eg of exerciseGroups) {
      for (let w = 0; w < weekPositions.length; w++) {
        const wBase = weekPositions[w].col;

        // Prefer this week's exercise name if it differs from first week's
        let weekExName = eg.name;
        for (const ri of eg.rows) {
          const r2 = rows[ri];
          const wkRaw = r2[wBase + exOff];
          const wkName = wkRaw ? String(wkRaw).trim() : '';
          if (wkName && /[a-zA-Z]{2,}/.test(wkName) && !/^(exercise|week)/i.test(wkName)) {
            weekExName = wkName; break;
          }
        }

        // Build prescription parts and collect TOP SET / RATING notes
        const presParts = [];
        const noteParts = [];
        for (const ri of eg.rows) {
          const r2 = rows[ri];
          const setsVal = r2[wBase + setsOff];
          const repsVal = r2[wBase + repsOff];
          if (setsVal == null && repsVal == null) continue;

          const sc = parseInt(setsVal) || 0;
          if (sc <= 0) continue;
          const rp = repsVal != null ? String(repsVal).trim() : '';

          // Read formatted load value from worksheet cell (handles formula strings like "330lbs")
          let ld = '';
          const rawLoad = r2[wBase + loadOff];
          if (rawLoad != null) {
            const physRow = ri + rowOffset;
            const physCol = wBase + loadOff + colOffset;
            const cellAddr = XLSX.utils.encode_cell({ r: physRow, c: physCol });
            const cell = ws[cellAddr];
            ld = (cell && cell.w) ? String(cell.w).trim() : String(rawLoad).trim();
          }

          const pres = _w_buildPrescription(sc, rp, ld);
          if (pres) presParts.push(pres);

          // Read TOP SET value for this row
          if (topSetOff >= 0) {
            const tsRaw = r2[wBase + topSetOff];
            if (tsRaw != null) {
              const ts = String(tsRaw).trim();
              if (ts && !/^top\s*set$/i.test(ts)) {
                noteParts.push(_w_formatTopSet(ts));
              }
            }
          }

          // Read RATING value for this row
          if (ratingOff >= 0) {
            const rtRaw = r2[wBase + ratingOff];
            if (rtRaw != null) {
              const rt = String(rtRaw).trim();
              if (rt && !/^rating$/i.test(rt)) {
                noteParts.push('Rating: ' + rt);
              }
            }
          }
        }

        if (presParts.length > 0) {
          weekExerciseSets[w].push({
            name: weekExName,
            prescription: presParts.join('; '),
            note: noteParts.join('; '), lifterNote: '', loggedWeight: '', supersetGroup: null
          });
        }
      }
    }

    dayData.push({ dayNum: daySections[d].dayNum, weekExerciseSets });
  }

  // ── Build week objects ──────────────────────────────────────────────────────
  const allWeeks = [];
  for (let w = 0; w < weekPositions.length; w++) {
    const weekDays = [];
    let seqDay = 0;
    for (const dd of dayData) {
      const exercises = dd.weekExerciseSets[w];
      if (!exercises || exercises.length === 0) continue;
      seqDay++;
      weekDays.push({ name: 'Day ' + seqDay, exercises });
    }
    if (weekDays.length > 0) {
      const rawLabel = weekPositions[w].label;
      const label = rawLabel.replace(/^WEEK\s*(\d+)$/i, (_, n) => 'Week ' + n) || rawLabel;
      allWeeks.push({ label, days: weekDays });
    }
  }
  if (allWeeks.length === 0) return null;

  const dateRange = _w_computeDateRange(rows, weekHeaderRow, weekPositions);
  const blockName = programName + (athleteName ? ' \u2014 ' + athleteName : '');

  return [{
    id: 'W_' + Date.now(),
    name: blockName,
    format: 'W',
    athleteName,
    dateRange,
    maxes,
    weeks: allWeeks
  }];
}

// ── Helper: strip leading apostrophes from sheet cells ───────────────────────
function _w_cleanRows(ws) {
  return XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
}

// ── Helper: build Format W prescription string ───────────────────────────────
// Handles RPE (integer 1-10), absolute weight ("330lbs"), percent drop ("-15%"), AMRAP
function _w_buildPrescription(sets, reps, loadStr) {
  if (!sets || sets <= 0) return null;
  const repsStr = /^amrap$/i.test(reps) ? 'AMRAP' : (reps || '');
  const ld = loadStr || '';
  let loadPart = '';
  if (ld) {
    if (/^-?\d+%$/.test(ld)) {
      // Percentage drop like "-15%" → @-15%
      loadPart = ' @' + ld;
    } else if (/lbs$/i.test(ld) || /kg$/i.test(ld)) {
      // Absolute weight like "330lbs" → @330lbs
      loadPart = ' @' + ld;
    } else if (/^\d+$/.test(ld) && parseInt(ld) >= 1 && parseInt(ld) <= 10) {
      // Integer 1-10 → RPE
      loadPart = ' @RPE' + ld;
    } else if (ld) {
      loadPart = ' @' + ld;
    }
  }
  return sets + 'x' + repsStr + loadPart;
}

// ── Helper: format TOP SET cell value into a note string ─────────────────────
// "396-407" → "Top set: 396-407 lbs"
// "374lbs CAP" → "Cap: 374 lbs"
function _w_formatTopSet(val) {
  const capMatch = val.match(/^(\d+)\s*lbs?\s+CAP\s*$/i);
  if (capMatch) return 'Cap: ' + capMatch[1] + ' lbs';
  // Weight range like "396-407" or "496-507, 523-535"
  if (/^\d[\d,\s\-–]+$/.test(val)) return 'Top set: ' + val + ' lbs';
  // Fallback: return as-is
  return 'Top set: ' + val;
}

// ── Helper: compute date range string from PROGRAM sheet rows ─────────────────
function _w_computeDateRange(rows, weekHeaderRow, weekPositions) {
  const base0 = weekPositions[0].col;
  let firstSerial = null;
  let lastSerial  = null;

  // First date: row immediately after week header with a date serial at base column
  for (let i = weekHeaderRow + 1; i < Math.min(weekHeaderRow + 4, rows.length); i++) {
    const row = rows[i]; if (!row) continue;
    const val = row[base0];
    if (typeof val === 'number' && val > 40000 && val < 60000) { firstSerial = val; break; }
  }

  // Last date: scan backwards for a date serial at the last week's column
  const lastWp = weekPositions[weekPositions.length - 1];
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i]; if (!row) continue;
    const val = row[lastWp.col];
    if (typeof val === 'number' && val > 40000 && val < 60000) { lastSerial = val; break; }
  }

  if (!firstSerial) return '';
  const toDate = s => new Date((s - 25569) * 86400 * 1000);
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const fmt = d => MONTHS[d.getUTCMonth()] + ' ' + d.getUTCDate();
  const startDate = toDate(firstSerial);
  const endDate   = lastSerial ? toDate(lastSerial) : null;
  if (!endDate) return fmt(startDate) + ', ' + startDate.getUTCFullYear();
  const yr = endDate.getUTCFullYear();
  const sameYr = startDate.getUTCFullYear() === yr;
  return sameYr
    ? fmt(startDate) + ' \u2013 ' + fmt(endDate) + ', ' + yr
    : fmt(startDate) + ', ' + startDate.getUTCFullYear() + ' \u2013 ' + fmt(endDate) + ', ' + yr;
}


// ── ADAPTIVE PARSER — SESSION 1: OCCUPANCY MATRIX + REGION DETECTION ─────────
// Pure utility functions — no side effects on existing parsers or pipeline.

/**
 * _resolveMergedCells(ws)
 * Build a lookup map from every non-top-left cell address inside a merge range
 * to the top-left cell address of that range.
 * Must run before _buildOccupancyMatrix so merged cells get the correct value/type.
 *
 * @param {object} ws — SheetJS worksheet
 * @returns {object} { cellAddr: topLeftAddr } for all non-top-left merged cells
 */
function _resolveMergedCells(ws) {
  const map = {};
  if (!ws || !ws['!merges']) return map;
  for (const merge of ws['!merges']) {
    const tlAddr = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    for (let r = merge.s.r; r <= merge.e.r; r++) {
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        if (r === merge.s.r && c === merge.s.c) continue; // top-left is the source cell
        map[XLSX.utils.encode_cell({ r, c })] = tlAddr;
      }
    }
  }
  return map;
}

/**
 * _buildOccupancyMatrix(ws)
 * Convert a SheetJS worksheet into a 2D grid of typed cells.
 * Cell types: 'empty' | 'string' | 'number' | 'formula' | 'date'
 * Merged cells resolve to the top-left cell's value and type.
 * Formula cells carry their computed value; numericValue is set when result is numeric.
 *
 * @param {object} ws — SheetJS worksheet
 * @returns {{ cells: Array<Array<{type, value, numericValue?}>>,
 *             rowCount, colCount, startRow, startCol, range }}
 */
function _buildOccupancyMatrix(ws) {
  if (!ws || !ws['!ref']) {
    return { cells: [], rowCount: 0, colCount: 0, startRow: 0, startCol: 0, range: null };
  }

  const mergeMap = _resolveMergedCells(ws);
  const range = XLSX.utils.decode_range(ws['!ref']);
  const rowCount = range.e.r - range.s.r + 1;
  const colCount = range.e.c - range.s.c + 1;

  const cells = [];
  for (let r = 0; r < rowCount; r++) {
    cells[r] = [];
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
      const resolvedAddr = mergeMap[addr] || addr; // follow merge to top-left
      const cell = ws[resolvedAddr];

      if (!cell || cell.t === 'z' || cell.v == null || cell.v === '') {
        cells[r][c] = { type: 'empty', value: null };
      } else if (cell.t === 'd' || cell.v instanceof Date) {
        cells[r][c] = { type: 'date', value: cell.v };
      } else if (cell.f) {
        // Formula: tag 'formula', expose numericValue when computed result is numeric
        const v = cell.v;
        cells[r][c] = { type: 'formula', value: v, numericValue: typeof v === 'number' ? v : null };
      } else if (cell.t === 'n') {
        cells[r][c] = { type: 'number', value: cell.v };
      } else if (cell.t === 's') {
        const v = String(cell.v).trim();
        cells[r][c] = v === '' ? { type: 'empty', value: null } : { type: 'string', value: v };
      } else if (cell.t === 'b') {
        // Boolean — treat as string for text-pattern matching
        cells[r][c] = { type: 'string', value: String(cell.v) };
      } else {
        cells[r][c] = { type: 'empty', value: null };
      }
    }
  }

  return { cells, rowCount, colCount, startRow: range.s.r, startCol: range.s.c, range };
}

/**
 * _scoreRegionWorkoutLikeness(cells, minR, maxR, minC, maxC)
 * Score a rectangular sub-region of the occupancy matrix for how likely it is
 * to contain workout training data (exercises + numeric prescriptions).
 *
 * Signals (spec section 1b):
 *   +0.50 — column with >=3 exercise-name-like strings (L176: /[a-zA-Z]{2,}/, not pure numbers)
 *   +0.30 — column of small integers (sets 1-10 / reps 1-30; >=40% of non-empty column cells)
 *   +0.20 — region has both string AND numeric cells (mixed-type region)
 *   +0.10 — weight indicator (numbers >20 or strings matching /\d+\s*(lbs?|kg)/)
 *   -> 0.0  — pure text block (>80% string, zero numeric cells) — rejected immediately
 *
 * @returns {number} score clamped to 0.0-1.0
 */
function _scoreRegionWorkoutLikeness(cells, minR, maxR, minC, maxC) {
  let totalCells = 0;
  let stringCells = 0;
  let numericCells = 0;

  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      if (!cells[r] || !cells[r][c]) continue;
      const t = cells[r][c].type;
      totalCells++;
      if (t === 'string') stringCells++;
      if (t === 'number' || (t === 'formula' && cells[r][c].numericValue != null)) numericCells++;
    }
  }

  if (totalCells === 0) return 0;

  // Reject pure text blocks: >80% string AND zero numeric cells
  if (stringCells / totalCells > 0.80 && numericCells === 0) return 0;

  let hasExerciseColumn = false;
  let hasSmallIntColumn = false;
  let hasWeightColumn = false;

  for (let c = minC; c <= maxC; c++) {
    let exerciseNameCount = 0;
    let smallIntCount = 0;
    let weightCount = 0;
    let totalInCol = 0;

    for (let r = minR; r <= maxR; r++) {
      if (!cells[r] || !cells[r][c]) continue;
      const cell = cells[r][c];
      if (cell.type === 'empty') continue;
      totalInCol++;

      if (cell.type === 'string' && typeof cell.value === 'string') {
        const v = cell.value;
        // L176: two consecutive alpha chars minimum; skip pure-number strings
        if (/[a-zA-Z]{2,}/.test(v) && !/^\d+\.?\d*$/.test(v.trim())) {
          exerciseNameCount++;
        }
        if (/\d+\s*(lbs?|kg)\b/i.test(v)) weightCount++;
      }

      const numVal = cell.type === 'number' ? cell.value :
                     (cell.type === 'formula' ? cell.numericValue : null);
      if (numVal != null) {
        if (Number.isInteger(numVal) && numVal >= 1 && numVal <= 30) smallIntCount++;
        if (numVal > 20) weightCount++;
      }
    }

    if (exerciseNameCount >= 3) hasExerciseColumn = true;
    if (totalInCol > 0 && smallIntCount >= 2 && smallIntCount / totalInCol >= 0.40) hasSmallIntColumn = true;
    if (weightCount >= 2) hasWeightColumn = true;
  }

  let score = 0;
  if (hasExerciseColumn) score += 0.50;
  if (hasSmallIntColumn) score += 0.30;
  if (numericCells > 0 && stringCells > 0) score += 0.20;
  if (hasWeightColumn) score += 0.10;

  return Math.min(score, 1.0);
}

/**
 * _detectRegions(matrix, ws)
 * Flood-fill (BFS) over the occupancy matrix to find connected components of
 * non-empty cells. Each component's bounding box becomes a candidate region.
 * Regions smaller than 3 rows x 2 columns are discarded.
 * Survivors are scored via _scoreRegionWorkoutLikeness and sorted highest-first.
 *
 * Coordinate conventions in returned regions:
 *   startRow/endRow/startCol/endCol -- absolute 0-indexed sheet position
 *   relStartRow/relEndRow/relStartCol/relEndCol -- relative to matrix cells[][]
 *
 * @param {{ cells, rowCount, colCount, startRow, startCol }} matrix
 * @param {object} ws -- original worksheet (available to future passes)
 * @returns {Array<{startRow,endRow,startCol,endCol,
 *                  relStartRow,relEndRow,relStartCol,relEndCol,
 *                  height,width,cellCount,workoutScore}>}
 */
function _detectRegions(matrix, ws) {
  const { cells, rowCount, colCount, startRow, startCol } = matrix;
  if (!rowCount || !colCount || !cells.length) return [];

  const visited = Array.from({ length: rowCount }, () => new Uint8Array(colCount));
  const regions = [];

  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      if (visited[r][c]) continue;
      if (!cells[r] || !cells[r][c] || cells[r][c].type === 'empty') {
        visited[r][c] = 1;
        continue;
      }

      // BFS to collect full connected component
      const queue = [[r, c]];
      visited[r][c] = 1;
      let head = 0;
      let minR = r, maxR = r, minC = c, maxC = c;
      let cellCount = 0;

      while (head < queue.length) {
        const [cr, cc] = queue[head++];
        cellCount++;
        if (cr < minR) minR = cr;
        if (cr > maxR) maxR = cr;
        if (cc < minC) minC = cc;
        if (cc > maxC) maxC = cc;

        // 4-directional neighbors only (no diagonals)
        const next = [[cr - 1, cc], [cr + 1, cc], [cr, cc - 1], [cr, cc + 1]];
        for (const [nr, nc] of next) {
          if (nr < 0 || nr >= rowCount || nc < 0 || nc >= colCount) continue;
          if (visited[nr][nc]) continue;
          visited[nr][nc] = 1;
          if (cells[nr] && cells[nr][nc] && cells[nr][nc].type !== 'empty') {
            queue.push([nr, nc]);
          }
        }
      }

      const height = maxR - minR + 1;
      const width  = maxC - minC + 1;

      // Minimum size: 3 rows x 2 columns (spec requirement)
      if (height < 3 || width < 2) continue;

      const workoutScore = _scoreRegionWorkoutLikeness(cells, minR, maxR, minC, maxC);

      regions.push({
        startRow:    minR + startRow,   // absolute sheet row (0-indexed)
        endRow:      maxR + startRow,
        startCol:    minC + startCol,   // absolute sheet col (0-indexed)
        endCol:      maxC + startCol,
        relStartRow: minR,              // relative to matrix cells[][]
        relEndRow:   maxR,
        relStartCol: minC,
        relEndCol:   maxC,
        height,
        width,
        cellCount,
        workoutScore
      });
    }
  }

  // Sort: highest workout score first; ties broken by cell count (denser = more data)
  regions.sort((a, b) => b.workoutScore - a.workoutScore || b.cellCount - a.cellCount);
  return regions;
}

// ── SESSION 2: Layout Classification + Header/Boundary Detection ──────────────

/**
 * _classifySheetLayout(ws)
 * Internal helper — classify a single worksheet's structural pattern.
 * Checks for horizontal-weeks first (week labels in row spanning columns),
 * then block-based (2+ header rows at distinct vertical positions),
 * then defaults to vertical-columnar.
 *
 * @param {object} ws — SheetJS worksheet
 * @returns {'horizontal-weeks'|'block-based'|'vertical'}
 */
function _classifySheetLayout(ws) {
  if (!ws || !ws['!ref']) return 'vertical';
  const range = XLSX.utils.decode_range(ws['!ref']);
  const maxRow = Math.min(range.e.r, range.s.r + 60); // scan first 60 rows
  const maxCol = Math.min(range.e.c, range.s.c + 30);

  const HEADER_KW = /^(exercise|movement|lift|name|sets?|reps?|weight|wt|load|lbs?|kg|rpe|rir|effort|intensity|%|percent)$/i;
  const WEEK_LABEL_RE = /\bweek\s*\d+|\bwk\s*\d+/i;

  const headerRowPositions = [];
  let weekLabelRowCount = 0;
  const weekLabelCols = new Set();

  for (let r = range.s.r; r <= maxRow; r++) {
    let headerKwCount = 0;
    let weekLabelsInRow = 0;

    for (let c = range.s.c; c <= maxCol; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const v = String(cell.v).trim();
      if (HEADER_KW.test(v)) headerKwCount++;
      if (WEEK_LABEL_RE.test(v)) {
        weekLabelsInRow++;
        weekLabelCols.add(c);
      }
    }

    if (headerKwCount >= 2) headerRowPositions.push(r);
    if (weekLabelsInRow >= 1) weekLabelRowCount++;
  }

  // Horizontal-weeks: 2+ week labels spread across different columns in same/adjacent rows
  if (weekLabelRowCount >= 1 && weekLabelCols.size >= 2) return 'horizontal-weeks';

  // Block-based: 2+ header rows at non-adjacent positions (gap > 3 rows)
  const distinctBlocks = [];
  for (const pos of headerRowPositions) {
    if (!distinctBlocks.length || pos - distinctBlocks[distinctBlocks.length - 1] > 3) {
      distinctBlocks.push(pos);
    }
  }
  if (distinctBlocks.length >= 2) return 'block-based';

  return 'vertical';
}

/**
 * _classifyLayout(wb, regions)
 * Classify the workbook's structural pattern into one of four options:
 *   'week-per-tab'     — majority of tab names match /week|wk|w\d+|phase|block|cycle/i
 *   'block-based'      — multiple header rows at different positions separated by blank rows
 *   'horizontal-weeks' — week labels span column groups within a sheet
 *   'vertical'         — single header row, exercises down (default fallback)
 *
 * Checked in priority order: week-per-tab → horizontal-weeks → block-based → vertical.
 *
 * @param {object} wb      — SheetJS workbook
 * @param {Array}  regions — workout-like regions from _detectRegions (may be empty)
 * @returns {{ pattern: string, sheetLayouts: {[sheetName]: string}, weekTabNames: string[] }}
 */
function _classifyLayout(wb, regions) {
  if (!wb || !wb.SheetNames || !wb.SheetNames.length) {
    return { pattern: 'vertical', sheetLayouts: {}, weekTabNames: [] };
  }

  const WEEK_TAB_RE = /week|wk|w\d+|phase|block|cycle/i;
  const sheetLayouts = {};

  // ── Priority 1: week-per-tab ─────────────────────────────────────────────
  // Require ≥2 matching tabs AND they make up ≥40% of all tabs
  const weekTabs = wb.SheetNames.filter(s => WEEK_TAB_RE.test(s.trim()));
  if (weekTabs.length >= 2 && weekTabs.length / wb.SheetNames.length >= 0.40) {
    for (const sn of wb.SheetNames) {
      sheetLayouts[sn] = WEEK_TAB_RE.test(sn.trim()) ? 'week-per-tab' : 'meta';
    }
    return { pattern: 'week-per-tab', sheetLayouts, weekTabNames: weekTabs };
  }

  // ── Per-sheet classification ─────────────────────────────────────────────
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws || !ws['!ref']) { sheetLayouts[sn] = 'vertical'; continue; }
    sheetLayouts[sn] = _classifySheetLayout(ws);
  }

  // ── Determine dominant pattern ───────────────────────────────────────────
  const counts = { 'block-based': 0, 'horizontal-weeks': 0, 'vertical': 0 };
  for (const p of Object.values(sheetLayouts)) {
    if (p in counts) counts[p]++;
  }

  // Priority: horizontal-weeks > block-based > vertical
  let pattern = 'vertical';
  if (counts['horizontal-weeks'] > 0) pattern = 'horizontal-weeks';
  else if (counts['block-based'] >= counts['vertical'] && counts['block-based'] > 0) pattern = 'block-based';

  return { pattern, sheetLayouts, weekTabNames: weekTabs };
}

/**
 * _detectHeaders(ws, region)
 * Score candidate rows in the first 20 rows of a region to find the header row.
 *
 * Scoring per row:
 *   +40 per keyword — domain keywords: exercise, sets, reps, weight, load, rpe, intensity, %
 *   +30             — row is all-string AND row below has >50% numeric cells (type transition)
 *   +10             — all cells non-empty AND all values unique
 *
 * @param {object} ws     — SheetJS worksheet
 * @param {object} region — from _detectRegions: { startRow, endRow, startCol, endCol, ... }
 * @returns {{ row: number, score: number, keywords: string[] }}
 *          row = -1 if no header found (pure data region or tiny region)
 */
function _detectHeaders(ws, region) {
  const { startRow, endRow, startCol, endCol } = region;
  const scanEnd = Math.min(endRow, startRow + 19); // first 20 rows only
  const width = endCol - startCol + 1;

  // Domain keywords that signal a header row (checked case-insensitively)
  const DOMAIN_KW = /^(exercise|exercises|movement|lift|name|sets?|reps?|weight|wt|load|lbs?|kg|rpe|rir|effort|intensity|%|percent|notes?|comments?|cues?)$/i;

  let bestRow = -1;
  let bestScore = 0;
  let bestKeywords = [];

  for (let r = startRow; r <= scanEnd; r++) {
    let score = 0;
    const keywords = [];
    let allString = true;
    let nonEmpty = 0;
    const values = [];

    for (let c = startCol; c <= endCol; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      const rawVal = cell ? cell.v : null;
      const v = rawVal != null ? String(rawVal).trim() : '';

      if (!v) {
        allString = false; // empty cell means row isn't fully string
        continue;
      }
      nonEmpty++;
      values.push(v);

      // Track if all non-empty cells are string type
      if (!cell || cell.t !== 's') allString = false;

      // Check for domain keywords — each match scores +40
      if (DOMAIN_KW.test(v)) {
        score += 40;
        keywords.push(v);
      }
    }

    // Type-transition: this row all-string, next row >50% numeric
    if (allString && nonEmpty > 0 && r + 1 <= endRow) {
      let numericNext = 0, totalNext = 0;
      for (let c = startCol; c <= endCol; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: r + 1, c })];
        if (!cell || cell.v == null) continue;
        totalNext++;
        if (cell.t === 'n' || (cell.f && typeof cell.v === 'number')) numericNext++;
      }
      if (totalNext > 0 && numericNext / totalNext > 0.50) score += 30;
    }

    // All non-empty + unique values across the full width
    if (nonEmpty === width && new Set(values).size === values.length) score += 10;

    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
      bestKeywords = keywords;
    }
  }

  return { row: bestRow, score: bestScore, keywords: bestKeywords };
}

/**
 * _detectDayBoundaries(ws, region, layout)
 * Find where training days start and end within a detected region.
 *
 * Signals (checked in priority order):
 *   1. Blank rows (all cells empty within column span) — most common separator
 *   2. Repeated header rows (EXERCISE/SETS/REPS keywords re-appearing)
 *   3. Day label text in first column ("Day 1", "Monday", "Upper", etc.)
 *   4. Date serial in first column of new block
 *
 * @param {object} ws     — SheetJS worksheet
 * @param {object} region — from _detectRegions: { startRow, endRow, startCol, endCol }
 * @param {object} layout — from _classifyLayout (unused currently, reserved for pattern-specific logic)
 * @returns {Array<{ startRow: number, endRow: number, label: string|null, dateSerial: number|null }>}
 */
function _detectDayBoundaries(ws, region, layout) {
  const { startRow, endRow, startCol, endCol } = region;

  const DAY_LABEL_RE = /^(day\s*\d+|monday|tuesday|wednesday|thursday|friday|saturday|sunday|upper|lower|push|pull|legs?|squat\s*day|bench\s*day|deadlift\s*day|[a-z]+\s+day\b)/i;
  const HEADER_KW_RE = /^(exercise|movement|lift|sets?|reps?|weight|load|rpe)$/i;

  // ── Pass 1: classify every row within the region ─────────────────────────
  const rowInfo = [];
  for (let r = startRow; r <= endRow; r++) {
    let isEmpty = true;
    let dayLabel = null;
    let headerKwCount = 0;
    let dateSerial = null;

    for (let c = startCol; c <= endCol; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null || String(cell.v).trim() === '') continue;
      isEmpty = false;
      const v = String(cell.v).trim();

      // First column only: look for day labels and dates
      if (c === startCol) {
        if (DAY_LABEL_RE.test(v)) dayLabel = v;
        // SheetJS date: cell.t === 'd', OR numeric Excel serial (40000–50000 ≈ 2009–2036)
        if (cell.t === 'd' || (cell.t === 'n' && cell.v > 40000 && cell.v < 55000)) {
          dateSerial = cell.v;
        }
      }
      if (HEADER_KW_RE.test(v.toLowerCase())) headerKwCount++;
    }

    rowInfo.push({ r, isEmpty, dayLabel, headerKwCount, dateSerial });
  }

  // ── Pass 2: group rows into day blocks ───────────────────────────────────
  const days = [];
  let blockStart = -1;
  let blockLabel = null;
  let blockDate = null;
  let pendingLabel = null; // day label row seen just before data starts
  let pendingDate = null;
  let blankRunLength = 0;

  const closeBlock = (endR) => {
    if (blockStart >= 0) {
      days.push({ startRow: blockStart, endRow: endR, label: blockLabel, dateSerial: blockDate });
      blockStart = -1;
      blockLabel = null;
      blockDate = null;
    }
  };

  for (let i = 0; i < rowInfo.length; i++) {
    const info = rowInfo[i];

    if (info.isEmpty) {
      blankRunLength++;
      // Close block on first blank row (or start of a blank run)
      if (blankRunLength === 1 && blockStart >= 0) {
        closeBlock(info.r - 1);
      }
      continue;
    }

    blankRunLength = 0;

    // Standalone day-label row (no data keywords) → store as pending label for next block
    const isDayLabelOnlyRow = info.dayLabel && info.headerKwCount === 0 && blockStart < 0;
    if (isDayLabelOnlyRow) {
      pendingLabel = info.dayLabel;
      pendingDate = info.dateSerial;
      continue;
    }

    // Repeated header row mid-region → close current block, skip this row
    if (info.headerKwCount >= 2 && blockStart >= 0) {
      closeBlock(info.r - 1);
      // This header row belongs to the next block; don't open the block yet
      pendingLabel = pendingLabel; // carry over any pending label
      continue;
    }

    // Start of a new data block
    if (blockStart < 0) {
      blockStart = info.r;
      blockLabel = pendingLabel || info.dayLabel;
      blockDate = pendingDate || info.dateSerial;
      pendingLabel = null;
      pendingDate = null;
    }
  }

  // Close final open block
  closeBlock(endRow);

  // If no boundaries found, treat the whole region as a single day
  if (days.length === 0) {
    days.push({ startRow, endRow, label: null, dateSerial: null });
  }

  return days;
}

/**
 * _detectWeekBoundaries(wb, regions, layout)
 * Find week divisions in the workbook based on the detected layout pattern.
 *
 * Returns shape varies by pattern:
 *   week-per-tab:      [{ label, sheetName }]
 *   horizontal-weeks:  [{ label, sheetName, columns: [startCol, endCol] }]
 *   block-based/vert:  [{ label, sheetName, startRow, endRow }]
 *
 * @param {object} wb      — SheetJS workbook
 * @param {Array}  regions — workout-like regions from _detectRegions (primary sheet)
 * @param {object} layout  — from _classifyLayout
 * @returns {Array<object>}
 */
function _detectWeekBoundaries(wb, regions, layout) {
  if (!wb || !wb.SheetNames) return [];
  const { pattern, sheetLayouts, weekTabNames } = layout;

  // ── week-per-tab: each matching tab is one week ──────────────────────────
  if (pattern === 'week-per-tab') {
    return (weekTabNames || []).map(sn => ({ label: sn, sheetName: sn }));
  }

  const WEEK_LABEL_RE = /\b(week|wk|w)\s*(\d+)\b/i;
  const weeks = [];

  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws || !ws['!ref']) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const sheetPat = sheetLayouts ? (sheetLayouts[sn] || pattern) : pattern;

    if (sheetPat === 'horizontal-weeks') {
      // ── Horizontal: scan header rows for "WEEK N" spanning column groups ──
      const maxScanRow = Math.min(range.e.r, range.s.r + 10);
      for (let r = range.s.r; r <= maxScanRow; r++) {
        const weekLabelsInRow = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (!cell || cell.v == null) continue;
          const v = String(cell.v).trim();
          const m = v.match(WEEK_LABEL_RE);
          if (m) weekLabelsInRow.push({ label: v, col: c, weekNum: parseInt(m[2], 10) });
        }
        if (weekLabelsInRow.length >= 2) {
          // Build column-group ranges from the label positions
          for (let i = 0; i < weekLabelsInRow.length; i++) {
            const startC = weekLabelsInRow[i].col;
            const endC = (i + 1 < weekLabelsInRow.length)
              ? weekLabelsInRow[i + 1].col - 1
              : range.e.c;
            weeks.push({ label: weekLabelsInRow[i].label, sheetName: sn, columns: [startC, endC] });
          }
          break; // found week-label row for this sheet — don't scan further
        }
      }
    } else {
      // ── Vertical / block-based: scan for "WEEK N" label rows ─────────────
      let lastWeekStart = -1;
      let lastWeekLabel = null;

      for (let r = range.s.r; r <= range.e.r; r++) {
        // Check first 4 columns for a week label
        for (let c = range.s.c; c <= Math.min(range.e.c, range.s.c + 3); c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (!cell || cell.v == null) continue;
          const v = String(cell.v).trim();
          const m = v.match(WEEK_LABEL_RE);
          if (m) {
            // Close previous week
            if (lastWeekStart >= 0) {
              weeks.push({ label: lastWeekLabel, sheetName: sn, startRow: lastWeekStart, endRow: r - 1 });
            }
            lastWeekStart = r;
            lastWeekLabel = v;
            break;
          }
        }
      }
      // Close final week for this sheet
      if (lastWeekStart >= 0) {
        weeks.push({ label: lastWeekLabel, sheetName: sn, startRow: lastWeekStart, endRow: range.e.r });
      }
    }
  }

  return weeks;
}

// ── SESSION 3: COLUMN INFERENCE + DATA EXTRACTION ─────────────────────────────
// Pure functions — no side effects on existing parsers or pipeline.

/**
 * _scoreColumnForType(type, headerVal, values, colIdx, totalCols, prevAssigned)
 * Internal helper for _inferColumns.
 * Scores a single column for how well it fits a given data type.
 *
 * Scoring:
 *   Header keyword match → +50 pts
 *   Value pattern match  → up to +30 pts (proportional to % of cells matching)
 *   Position signal      → +20 pts ideal position, +8 pts reasonable position
 *
 * @param {string} type         — 'exercise'|'sets'|'reps'|'weight'|'rpe'|'percentage'|'notes'
 * @param {string} headerVal    — raw header cell text for this column
 * @param {Array}  values       — cell values from data rows in this column
 * @param {number} colIdx       — 0-based index of this column in [startCol, endCol] range
 * @param {number} totalCols    — total columns in the scan range
 * @param {object} prevAssigned — column assignments so far: { exercise: absColIdx, sets: absColIdx, ... }
 * @returns {number} score 0-100
 */
function _scoreColumnForType(type, headerVal, values, colIdx, totalCols, prevAssigned) {
  let score = 0;
  const hv = String(headerVal || '').toLowerCase().trim();
  const nonNull = values.filter(v => v != null && String(v).trim() !== '');

  if (type === 'exercise') {
    if (/^(exercise|exercises?|movement|lift|name)$/i.test(hv)) score += 50;
    // Value pattern: strings with 2+ consecutive alpha chars, not pure numbers (L176)
    const exCount = nonNull.filter(v => {
      const s = String(v).trim();
      return /[a-zA-Z]{2,}/.test(s) && !/^\d+\.?\d*$/.test(s);
    }).length;
    if (nonNull.length > 0) score += Math.round((exCount / nonNull.length) * 30);
    // Position: leftmost column is ideal for exercise names
    if (colIdx === 0) score += 20;
    else if (colIdx === 1 && prevAssigned.exercise == null) score += 10;

  } else if (type === 'sets') {
    if (/^(sets?|s)$/i.test(hv)) score += 50;
    // Value pattern: small integers 1-10
    const intCount = nonNull.filter(v => {
      const n = Number(v);
      return Number.isInteger(n) && n >= 1 && n <= 10;
    }).length;
    if (nonNull.length > 0) score += Math.round((intCount / nonNull.length) * 30);
    // Position: right after exercise column
    const exCol = prevAssigned.exercise != null ? prevAssigned.exercise : -1;
    if (exCol >= 0 && colIdx === exCol + 1) score += 20;
    else if (colIdx >= 1 && colIdx <= 3) score += 8;

  } else if (type === 'reps') {
    if (/^(reps?|r)$/i.test(hv)) score += 50;
    // Value pattern: 1-30, AMRAP variants, "N+" notation
    const repCount = nonNull.filter(v => {
      const s = String(v).trim();
      const n = Number(s);
      return /^(amrap|amap|mr|f)$/i.test(s) ||
             /^\d+\+$/.test(s) ||
             (Number.isFinite(n) && n >= 1 && n <= 30);
    }).length;
    if (nonNull.length > 0) score += Math.round((repCount / nonNull.length) * 30);
    // Position: right after sets column
    const setsCol = prevAssigned.sets != null ? prevAssigned.sets : -1;
    if (setsCol >= 0 && colIdx === setsCol + 1) score += 20;
    else if (colIdx >= 2 && colIdx <= 4) score += 8;

  } else if (type === 'weight') {
    if (/^(weight|wt|load|lbs?|kg|load\/rpe|load \/ rpe)$/i.test(hv)) score += 50;
    // Value pattern: numbers >20, strings with "lbs"/"kg", negative percentages
    const wCount = nonNull.filter(v => {
      const s = String(v).trim();
      // Strip trailing unit for numeric check
      const numPart = s.replace(/[a-zA-Z%\s]+$/, '').replace(/^-/, '').trim();
      const n = Number(numPart);
      return /\d+\s*(lbs?|kg)\b/i.test(s) ||
             /^-?\d+(\.\d+)?%$/.test(s) ||  // percentage drop like -15%
             (Number.isFinite(n) && n > 20 && n < 2000);
    }).length;
    if (nonNull.length > 0) score += Math.round((wCount / nonNull.length) * 30);
    // Position: after reps column
    const repsCol = prevAssigned.reps != null ? prevAssigned.reps : -1;
    if (repsCol >= 0 && colIdx === repsCol + 1) score += 20;
    else if (colIdx >= 3) score += 5;

  } else if (type === 'rpe') {
    if (/^(rpe|rir|effort)$/i.test(hv)) score += 50;
    // Value pattern: numbers 5-10 in half-step increments (5, 5.5, 6, ..., 10)
    const rpeCount = nonNull.filter(v => {
      const n = Number(v);
      return Number.isFinite(n) && n >= 5 && n <= 10 &&
             Math.round(n * 2) === n * 2;
    }).length;
    if (nonNull.length > 0) score += Math.round((rpeCount / nonNull.length) * 30);
    // Position: after weight column
    const wtCol = prevAssigned.weight != null ? prevAssigned.weight : -1;
    if (wtCol >= 0 && colIdx === wtCol + 1) score += 20;

  } else if (type === 'percentage') {
    if (/^(%|percent|intensity|pct)$/i.test(hv)) score += 50;
    // Value pattern: 50-105 (whole percent) or 0.50-1.05 (decimal fraction)
    const pctCount = nonNull.filter(v => {
      const n = Number(v);
      return Number.isFinite(n) &&
             ((n >= 50 && n <= 105) || (n > 0.49 && n < 1.06));
    }).length;
    if (nonNull.length > 0) score += Math.round((pctCount / nonNull.length) * 30);
    // Position: near weight column
    const wtCol2 = prevAssigned.weight != null ? prevAssigned.weight : -1;
    if (wtCol2 >= 0 && Math.abs(colIdx - wtCol2) <= 2) score += 10;

  } else if (type === 'notes') {
    if (/^(notes?|comments?|cues?|tips?)$/i.test(hv)) score += 50;
    // Value pattern: longer strings (>10 chars)
    const noteCount = nonNull.filter(v => String(v).trim().length > 10).length;
    if (nonNull.length > 0) score += Math.round((noteCount / nonNull.length) * 30);
    // Position: last column in region
    if (colIdx === totalCols - 1) score += 20;
  }

  return score;
}

/**
 * _inferColumns(ws, region, headerRow, colRange)
 * Score each column in a region against known column types and return
 * the best-matching absolute column index for each type.
 *
 * Handles dual-purpose LOAD/RPE columns per spec §2a disambiguation rule:
 * if a column has BOTH weight strings ("330lbs") and RPE integers (5-10),
 * mark isDualLoadRpe=true and let extraction parse each cell individually.
 *
 * @param {object}  ws         — SheetJS worksheet
 * @param {object}  region     — { startRow, endRow, startCol, endCol }
 * @param {object}  headerRow  — { row: N, score, keywords } from _detectHeaders (row=-1 = none)
 * @param {object}  [colRange] — optional { startCol, endCol } override for column scan range
 * @returns {{ exercise, sets, reps, weight, rpe, percentage, notes: number (-1 = not found),
 *             isDualLoadRpe: boolean, colScores: object }}
 */
function _inferColumns(ws, region, headerRow, colRange) {
  if (!ws || !region) {
    return { exercise: -1, sets: -1, reps: -1, weight: -1, rpe: -1,
             percentage: -1, notes: -1, isDualLoadRpe: false, colScores: {} };
  }

  const scanStartCol = (colRange && colRange.startCol != null) ? colRange.startCol : region.startCol;
  const scanEndCol   = (colRange && colRange.endCol   != null) ? colRange.endCol   : region.endCol;
  const startRow = region.startRow;
  const endRow   = region.endRow;

  const headerRowNum = (headerRow && headerRow.row >= startRow && headerRow.row <= endRow)
    ? headerRow.row : -1;
  const dataStartRow = headerRowNum >= 0 ? headerRowNum + 1 : startRow;
  const totalCols    = scanEndCol - scanStartCol + 1;

  // ── Build per-column stats ────────────────────────────────────────────────
  const colData = [];
  for (let c = scanStartCol; c <= scanEndCol; c++) {
    const colIdx = c - scanStartCol;

    let headerVal = '';
    if (headerRowNum >= 0) {
      const hcell = ws[XLSX.utils.encode_cell({ r: headerRowNum, c })];
      if (hcell && hcell.v != null) headerVal = String(hcell.v).trim();
    }

    // Prefer formatted string (.w) for values like "330lbs", "AMRAP"
    const values = [];
    for (let r = dataStartRow; r <= endRow; r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v != null && String(cell.v).trim() !== '') {
        values.push(cell.w !== undefined ? cell.w : cell.v);
      }
    }

    colData.push({ c, colIdx, headerVal, values });
  }

  // ── Greedy assignment: process types in priority order ────────────────────
  const TYPES = ['exercise', 'sets', 'reps', 'weight', 'rpe', 'percentage', 'notes'];
  const result = { isDualLoadRpe: false, colScores: {} };
  const assignedCols = new Set();

  for (const type of TYPES) {
    let bestC = -1, bestScore = 0;
    for (const col of colData) {
      if (assignedCols.has(col.c)) continue;
      const score = _scoreColumnForType(
        type, col.headerVal, col.values, col.colIdx, totalCols, result
      );
      if (score > bestScore) {
        bestScore = score;
        bestC = col.c;
      }
    }
    result[type] = bestScore > 0 ? bestC : -1;
    if (bestC >= 0) assignedCols.add(bestC);
  }

  // ── Dual-purpose LOAD/RPE detection ──────────────────────────────────────
  if (result.weight >= 0) {
    const wd = colData.find(cd => cd.c === result.weight);
    if (wd) {
      let hasAbsWeight = false, hasRpeInt = false;
      for (const v of wd.values) {
        const s = String(v).trim();
        const rawNum = Number(s);
        if (/\d+\s*(lbs?|kg)/i.test(s) || (Number.isFinite(rawNum) && rawNum > 20)) hasAbsWeight = true;
        if (Number.isInteger(rawNum) && rawNum >= 5 && rawNum <= 10) hasRpeInt = true;
      }
      if (hasAbsWeight && hasRpeInt) result.isDualLoadRpe = true;
    }
  }

  // ── Store per-column scores for debugging / confidence ────────────────────
  for (const col of colData) {
    result.colScores[col.c] = {};
    for (const type of TYPES) {
      result.colScores[col.c][type] = _scoreColumnForType(
        type, col.headerVal, col.values, col.colIdx, totalCols, {}
      );
    }
  }

  return result;
}

/**
 * _buildPrescription(sets, reps, loadStr, rpeVal, pctVal)
 * Assemble a prescription string compatible with parseSets() / parseSimple().
 *
 * Priority for load annotation: explicit RPE column > explicit % column > parsed loadStr.
 * loadStr parsing: "330lbs" → @330lbs, "-15%" → @-15%, "-0.15" → @-15%,
 *                  "8" (RPE int 5-10) → @RPE8, "75" (50-105) → @75%, ">20" → @Nlbs.
 *
 * @param {number|string} sets    — number of sets (must parse to integer 1-50)
 * @param {number|string} reps    — reps value (number, "AMRAP", "5+", etc.)
 * @param {string}        loadStr — formatted load value from worksheet (.w or .v)
 * @param {number|null}   rpeVal  — explicit RPE value from a dedicated RPE column
 * @param {number|null}   pctVal  — explicit percentage from a dedicated % column
 * @returns {string|null}
 */
function _buildPrescription(sets, reps, loadStr, rpeVal, pctVal) {
  const sc = parseInt(sets);
  if (!sc || sc <= 0 || sc > 50) return null;

  const repsStr = reps != null ? String(reps).trim() : '';
  if (!repsStr) return null;

  const repsOut = /^(amrap|amap|mr|f)$/i.test(repsStr) ? 'AMRAP' : repsStr;
  const base = sc + 'x' + repsOut;

  // Build load annotation
  let loadPart = '';

  if (rpeVal != null) {
    const r = Number(rpeVal);
    if (Number.isFinite(r) && r >= 5 && r <= 10) loadPart = ' @RPE' + r;
  }

  if (!loadPart && pctVal != null) {
    const p = Number(pctVal);
    if (Number.isFinite(p)) {
      if (p > 0.49 && p < 1.06) loadPart = ' @' + Math.round(p * 100) + '%';
      else if (p >= 50 && p <= 105) loadPart = ' @' + Math.round(p) + '%';
    }
  }

  if (!loadPart) {
    const ld = String(loadStr || '').trim();
    if (ld) {
      if (/^-?\d+(\.\d+)?%$/.test(ld)) {
        loadPart = ' @' + ld;                                   // "-15%" → "@-15%"
      } else if (/\d+\s*(lbs?|kg)\b/i.test(ld)) {
        loadPart = ' @' + ld.replace(/\s+/g, '');               // "330 lbs" → "@330lbs"
      } else if (/^-?\d+(\.\d+)?$/.test(ld)) {
        const n = Number(ld);
        if (n < 0) {
          // Negative decimal → percentage drop (e.g. -0.15 → @-15%)
          loadPart = ' @' + Math.round(n * 100) + '%';
        } else if (n >= 5 && n <= 10) {
          loadPart = ' @RPE' + n;                               // integer 5-10 → RPE
        } else if (n >= 50 && n <= 105) {
          loadPart = ' @' + n + '%';                            // 50-105 → percentage
        } else if (n > 20) {
          loadPart = ' @' + n + 'lbs';                         // bare number > 20 → lbs
        }
      } else if (ld) {
        loadPart = ' @' + ld;                                   // passthrough
      }
    }
  }

  return base + loadPart;
}

/**
 * _handleMultiRowExercise(ws, rowIndices, columns)
 * Process row indices from a day block, detect multi-row exercise patterns,
 * and return combined exercise objects.
 *
 * Multi-row patterns (spec §2d):
 *   1. Empty exercise cell + has prescription data → continuation of previous
 *   2. Same exercise name as previous row → additional set prescription
 *   3. Negative number in LOAD column → relative percentage drop (e.g. -0.15 = @-15%)
 *
 * @param {object}  ws         — SheetJS worksheet
 * @param {Array}   rowIndices — absolute row indices to process
 * @param {object}  columns    — { exercise, sets, reps, weight, rpe, percentage, isDualLoadRpe }
 * @returns {Array<{name, prescription, note, lifterNote, loggedWeight, supersetGroup}>}
 */
function _handleMultiRowExercise(ws, rowIndices, columns) {
  const groups = []; // { name: string, parts: string[] }

  for (const r of rowIndices) {
    // ── Read exercise name ─────────────────────────────────────────────────
    let exName = '';
    if (columns.exercise >= 0) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: columns.exercise })];
      if (cell && cell.v != null) exName = String(cell.v).trim();
    }

    // Skip rows that are themselves header rows or day-label rows
    if (/^(exercise|exercises?|movement|lift)$/i.test(exName)) continue;
    if (/^(week|day)\s*\d/i.test(exName)) continue;
    if (/^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/i.test(exName)) continue;

    // ── Read sets ──────────────────────────────────────────────────────────
    let setsVal = null;
    if (columns.sets >= 0) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: columns.sets })];
      if (cell && cell.v != null) setsVal = cell.v;
    }

    // ── Read reps (prefer .w for AMRAP strings) ───────────────────────────
    let repsVal = null;
    if (columns.reps >= 0) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: columns.reps })];
      if (cell && cell.v != null) {
        repsVal = cell.w !== undefined ? cell.w : String(cell.v).trim();
      }
    }

    // ── Read load/weight (prefer formatted string .w) ─────────────────────
    let loadStr = '';
    if (columns.weight >= 0) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: columns.weight })];
      if (cell && cell.v != null) {
        loadStr = cell.w !== undefined ? String(cell.w).trim() : String(cell.v).trim();
      }
    }

    // ── Read explicit RPE column ───────────────────────────────────────────
    let rpeVal = null;
    if (columns.rpe >= 0) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: columns.rpe })];
      if (cell && cell.v != null) rpeVal = cell.v;
    }

    // ── Read explicit percentage column ───────────────────────────────────
    let pctVal = null;
    if (columns.percentage >= 0) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: columns.percentage })];
      if (cell && cell.v != null) pctVal = cell.v;
    }

    // Skip entirely empty rows
    const hasAnyData = setsVal != null || repsVal != null || loadStr !== '';
    if (!hasAnyData && !exName) continue;

    // ── Continuation vs. new exercise ─────────────────────────────────────
    const hasValidName = exName && /[a-zA-Z]{2,}/.test(exName); // L176
    const prevGroup = groups.length > 0 ? groups[groups.length - 1] : null;
    const isContinuation =
      (!hasValidName && prevGroup && hasAnyData) ||           // empty name, has data
      (hasValidName && prevGroup && exName === prevGroup.name); // same name as prev

    if (isContinuation && prevGroup) {
      const pres = _buildPrescription(setsVal, repsVal, loadStr, rpeVal, pctVal);
      if (pres) prevGroup.parts.push(pres);
    } else if (hasValidName) {
      const newGroup = { name: exName, parts: [] };
      const pres = _buildPrescription(setsVal, repsVal, loadStr, rpeVal, pctVal);
      if (pres) newGroup.parts.push(pres);
      groups.push(newGroup);
    }
  }

  // L165: every exercise object must include supersetGroup: null
  return groups.map(g => ({
    name: g.name,
    prescription: g.parts.join('; '),
    note: '',
    lifterNote: '',
    loggedWeight: '',
    supersetGroup: null
  }));
}

// ── Adaptive Parser: Internal sub-functions (_ap_ prefix per L075) ────────────

/**
 * _ap_findPrimarySheet(wb, layout)
 * Find the primary training sheet, skipping metadata/helper sheets.
 */
function _ap_findPrimarySheet(wb, layout) {
  const { pattern, sheetLayouts } = layout;
  const META_RE = /homepage|faq|instruction|readme|about|intro|meta|overview|summary|nutrition|movement\s*prep|warm.?up|maxe?s?|calculator|input/i;

  // For horizontal-weeks: prefer sheets classified as such
  if (pattern === 'horizontal-weeks') {
    for (const [sn, pat] of Object.entries(sheetLayouts || {})) {
      if (pat === 'horizontal-weeks' && !META_RE.test(sn) && wb.Sheets[sn]) return sn;
    }
  }

  // For block-based: prefer sheets classified as block-based
  if (pattern === 'block-based') {
    for (const [sn, pat] of Object.entries(sheetLayouts || {})) {
      if (pat === 'block-based' && !META_RE.test(sn) && wb.Sheets[sn]) return sn;
    }
  }

  // For vertical: prefer vertical-classified sheets
  for (const [sn, pat] of Object.entries(sheetLayouts || {})) {
    if (pat === 'vertical' && !META_RE.test(sn) && wb.Sheets[sn] && wb.Sheets[sn]['!ref']) return sn;
  }

  // Fallback: first non-metadata sheet with data
  for (const sn of wb.SheetNames) {
    if (META_RE.test(sn)) continue;
    if (wb.Sheets[sn] && wb.Sheets[sn]['!ref']) return sn;
  }

  return wb.SheetNames[0];
}

/**
 * _ap_inferProgramName(wb, sheetName)
 * Try to find a program name from the first few rows of the primary sheet,
 * or fall back to the sheet name.
 */
function _ap_inferProgramName(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws['!ref']) return sheetName || 'Training Program';

  const range = XLSX.utils.decode_range(ws['!ref']);
  // Scan first 3 rows, first 6 cols for a title-like string
  for (let r = range.s.r; r <= Math.min(range.s.r + 3, range.e.r); r++) {
    for (let c = range.s.c; c <= Math.min(range.s.c + 5, range.e.c); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;
      const v = String(cell.v).trim();
      // A title: 5-80 chars, has real words, not a domain keyword
      if (v.length >= 5 && v.length <= 80 && /[a-zA-Z]{3,}/.test(v) &&
          !/^(exercise|sets?|reps?|week|day|load|weight|rpe|date)$/i.test(v)) {
        return v;
      }
    }
  }
  return sheetName || 'Training Program';
}

/**
 * _ap_extractDaysFromBounds(ws, dayBoundsArr, colMap)
 * Build exercise-populated day objects from pre-computed day boundaries.
 * @returns {Array<{name, exercises}>}
 */
function _ap_extractDaysFromBounds(ws, dayBoundsArr, colMap) {
  const days = [];
  for (let i = 0; i < dayBoundsArr.length; i++) {
    const db = dayBoundsArr[i];
    const rowIndices = [];
    for (let r = db.startRow; r <= db.endRow; r++) rowIndices.push(r);

    const exercises = _handleMultiRowExercise(ws, rowIndices, colMap);
    if (exercises.length > 0) {
      days.push({ name: db.label || ('Day ' + (i + 1)), exercises });
    }
  }
  return days;
}

/**
 * _ap_extractHorizontal(ws, range, weekBoundsForSheet)
 * Extract workout data from a horizontal-weeks layout.
 * Weeks = column groups; day sections = repeated header rows.
 *
 * Column offsets are inferred once from the first week group and
 * reused for all subsequent week groups (they share the same layout).
 *
 * @param {object}  ws                 — SheetJS worksheet
 * @param {object}  range              — XLSX decoded range { s:{r,c}, e:{r,c} }
 * @param {Array}   weekBoundsForSheet — [{ label, columns: [startC, endC] }]
 * @returns {Array<{label, days}>}
 */
function _ap_extractHorizontal(ws, range, weekBoundsForSheet) {
  if (!weekBoundsForSheet || weekBoundsForSheet.length === 0) return null;

  const firstWk = weekBoundsForSheet[0];

  // ── Infer column offsets from the first week's column group ──────────────
  const firstRegion = {
    startRow: range.s.r,
    endRow:   range.e.r,
    startCol: firstWk.columns[0],
    endCol:   firstWk.columns[1]
  };
  const firstHeader = _detectHeaders(ws, firstRegion);
  const colMap = _inferColumns(ws, firstRegion, firstHeader);

  // Relative offsets from each week's start column
  const exOff   = colMap.exercise   >= 0 ? colMap.exercise   - firstWk.columns[0] : 0;
  const setsOff = colMap.sets       >= 0 ? colMap.sets       - firstWk.columns[0] : 1;
  const repsOff = colMap.reps       >= 0 ? colMap.reps       - firstWk.columns[0] : 2;
  const loadOff = colMap.weight     >= 0 ? colMap.weight     - firstWk.columns[0] : 3;
  const rpeOff  = colMap.rpe        >= 0 ? colMap.rpe        - firstWk.columns[0] : -1;
  const pctOff  = colMap.percentage >= 0 ? colMap.percentage - firstWk.columns[0] : -1;

  // ── Find day sections: rows where exercise column has "EXERCISE" keyword ─
  const EXERCISE_KW_RE = /^(exercise|exercises?)$/i;
  const daySections = [];
  let dayNum = 0;
  for (let r = range.s.r; r <= range.e.r; r++) {
    // At least one week group must show "EXERCISE" in this row
    for (const wk of weekBoundsForSheet) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: wk.columns[0] + exOff })];
      if (cell && cell.v && EXERCISE_KW_RE.test(String(cell.v).trim())) {
        dayNum++;
        daySections.push({ headerRow: r, dayNum });
        break;
      }
    }
  }

  if (daySections.length === 0) {
    // ── Hybrid fallback: exercise/sets/reps are shared columns left of week groups ─
    // Pattern: exercise names in col 0, sets/reps in shared cols, only weight repeats per week.
    // Examples: Reddit PPL, Norwegian PPL.
    return _ap_extractHybridHorizontal(ws, range, weekBoundsForSheet);
  }

  // ── Initialize per-week arrays ────────────────────────────────────────────
  const weekData = weekBoundsForSheet.map(wk => ({ label: wk.label, days: [] }));

  // ── Extract per day section ───────────────────────────────────────────────
  for (let d = 0; d < daySections.length; d++) {
    const startDataRow = daySections[d].headerRow + 1;
    const endDataRow   = d + 1 < daySections.length
      ? daySections[d + 1].headerRow - 1
      : range.e.r;

    const rowIndices = [];
    for (let r = startDataRow; r <= endDataRow; r++) rowIndices.push(r);

    const dayLabel = 'Day ' + daySections[d].dayNum;

    // For each week, build its column map and extract exercises
    for (let w = 0; w < weekBoundsForSheet.length; w++) {
      const wkStartCol = weekBoundsForSheet[w].columns[0];
      const weekCols = {
        exercise:   wkStartCol + exOff,
        sets:       wkStartCol + setsOff,
        reps:       wkStartCol + repsOff,
        weight:     loadOff >= 0 ? wkStartCol + loadOff : -1,
        rpe:        rpeOff  >= 0 ? wkStartCol + rpeOff  : -1,
        percentage: pctOff  >= 0 ? wkStartCol + pctOff  : -1,
        isDualLoadRpe: colMap.isDualLoadRpe
      };
      const exercises = _handleMultiRowExercise(ws, rowIndices, weekCols);
      if (exercises.length > 0) {
        weekData[w].days.push({ name: dayLabel, exercises });
      }
    }
  }

  return weekData.filter(w => w.days.length > 0);
}

/**
 * _ap_extractHybridHorizontal(ws, range, weekBoundsForSheet)
 * Hybrid horizontal layout: exercise/sets/reps are shared columns LEFT of all
 * week groups; only weight repeats per week (e.g. Reddit PPL, Norwegian PPL).
 *
 * Detection strategy:
 *   1. Infer shared columns from the region left of all week groups
 *   2. Detect day boundaries within the shared region
 *   3. For each day × week: use shared cols + week's first column as weight
 */
function _ap_extractHybridHorizontal(ws, range, weekBoundsForSheet) {
  if (!weekBoundsForSheet || weekBoundsForSheet.length === 0) return null;

  // Region to the LEFT of the first week group (shared exercise/sets/reps cols)
  const minWeekCol = Math.min(...weekBoundsForSheet.map(w => w.columns[0]));
  if (minWeekCol <= range.s.c) return null; // no room for shared cols

  const sharedRegion = {
    startRow: range.s.r,
    endRow:   range.e.r,
    startCol: range.s.c,
    endCol:   minWeekCol - 1
  };

  // Infer column types within the shared region
  const sharedHeader = _detectHeaders(ws, sharedRegion);
  const sharedCols   = _inferColumns(ws, sharedRegion, sharedHeader);

  if (sharedCols.exercise < 0) return null;

  // Detect day boundaries within the shared region
  const dayBoundsArr = _detectDayBoundaries(ws, sharedRegion, { pattern: 'block-based', sheetLayouts: {}, weekTabNames: [] });
  if (!dayBoundsArr || dayBoundsArr.length === 0) return null;

  const weekData = weekBoundsForSheet.map(wk => ({ label: wk.label, days: [] }));

  for (let d = 0; d < dayBoundsArr.length; d++) {
    const db = dayBoundsArr[d];
    const rowIndices = [];
    for (let r = db.startRow; r <= db.endRow; r++) rowIndices.push(r);
    const dayLabel = db.label || ('Day ' + (d + 1));

    for (let w = 0; w < weekBoundsForSheet.length; w++) {
      const wk = weekBoundsForSheet[w];
      // Shared exercise/sets/reps + this week's first column as weight
      const weekCols = {
        exercise:   sharedCols.exercise,
        sets:       sharedCols.sets,
        reps:       sharedCols.reps >= 0 ? sharedCols.reps : -1,
        weight:     wk.columns[0],  // first col of week group = weight for that week
        rpe:        -1,
        percentage: sharedCols.percentage >= 0 ? sharedCols.percentage : -1,
        isDualLoadRpe: false
      };
      const exercises = _handleMultiRowExercise(ws, rowIndices, weekCols);
      if (exercises.length > 0) {
        weekData[w].days.push({ name: dayLabel, exercises });
      }
    }
  }

  return weekData.filter(w => w.days.length > 0);
}

/**
 * _ap_extractWeekPerTab(wb, layout)
 * Extract workout data from a week-per-tab layout.
 * Each matching tab = one week. Column inference and day detection per tab.
 *
 * @param {object} wb     — SheetJS workbook
 * @param {object} layout — from _classifyLayout
 * @returns {Array<{label, days}>}
 */
function _ap_extractWeekPerTab(wb, layout) {
  const { weekTabNames } = layout;
  if (!weekTabNames || weekTabNames.length === 0) return null;

  const weeks = [];

  for (const sn of weekTabNames) {
    const ws = wb.Sheets[sn];
    if (!ws || !ws['!ref']) continue;

    const matrix = _buildOccupancyMatrix(ws);
    const regions = _detectRegions(matrix, ws);
    if (!regions || regions.length === 0) continue;

    // Use the highest workout-score region
    const primaryRegion = regions.find(r => r.workoutScore >= 0.5) || regions[0];
    const headerInfo = _detectHeaders(ws, primaryRegion);
    const colMap = _inferColumns(ws, primaryRegion, headerInfo);

    if (colMap.exercise < 0) continue;

    const dayBoundsArr = _detectDayBoundaries(ws, primaryRegion, layout);
    const extractedDays = _ap_extractDaysFromBounds(ws, dayBoundsArr, colMap);

    if (extractedDays.length > 0) {
      weeks.push({ label: sn, days: extractedDays });
    }
  }

  return weeks.length > 0 ? weeks : null;
}

/**
 * _adaptiveExtract(wb, regions, layout, headers, columns, dayBounds, weekBounds)
 * Orchestrate full data extraction using the adaptive parser pipeline.
 *
 * Dispatches to the appropriate sub-extractor based on layout pattern.
 * Pre-computed detection results are used when provided (non-null);
 * otherwise computed internally from the workbook.
 *
 * Output shape matches all other parsers (spec L021):
 *   [{ name, format:'adaptive', athleteName, dateRange, maxes, weeks }]
 *   weeks: [{ label, days: [{ name, exercises: [{ name, prescription, ... }] }] }]
 *
 * @param {object} wb         — SheetJS workbook
 * @param {Array}  regions    — workout-like regions from _detectRegions (may have .sheetName)
 * @param {object} layout     — from _classifyLayout
 * @param {object} headers    — { [sheetName]: headerInfo }   (optional, computed if null)
 * @param {object} columns    — { [sheetName]: colMap }       (optional, computed if null)
 * @param {object} dayBounds  — { [sheetName]: dayBoundsArr } (optional, computed if null)
 * @param {Array}  weekBounds — from _detectWeekBoundaries    (optional, computed if null)
 * @returns {Array|null}
 */
function _adaptiveExtract(wb, regions, layout, headers, columns, dayBounds, weekBounds) {
  if (!wb) return null;

  // Compute layout internally if not provided (all params are optional per spec)
  let resolvedLayout = layout;
  let resolvedRegions = regions;
  if (!resolvedLayout) {
    const primarySn = wb.SheetNames.find(sn => wb.Sheets[sn] && wb.Sheets[sn]['!ref']) || wb.SheetNames[0];
    const ws0 = wb.Sheets[primarySn];
    if (ws0) {
      const mat = _buildOccupancyMatrix(ws0);
      resolvedRegions = resolvedRegions || _detectRegions(mat, ws0);
    }
    resolvedLayout = _classifyLayout(wb, resolvedRegions || []);
  }

  const { pattern } = resolvedLayout;

  // Resolve week bounds (compute if not provided)
  const resolvedWeekBounds = weekBounds ||
    _detectWeekBoundaries(wb, resolvedRegions || [], resolvedLayout);

  // ── week-per-tab ────────────────────────────────────────────────────────
  if (pattern === 'week-per-tab') {
    const weekResult = _ap_extractWeekPerTab(wb, resolvedLayout);
    if (!weekResult) return null;
    const primarySn = (resolvedLayout.weekTabNames || [])[0] || wb.SheetNames[0];
    return [{
      name: _ap_inferProgramName(wb, primarySn),
      format: 'adaptive',
      athleteName: '',
      dateRange: '',
      maxes: { squat: null, bench: null, deadlift: null },
      weeks: weekResult
    }];
  }

  // ── Find primary sheet ──────────────────────────────────────────────────
  const primarySn = _ap_findPrimarySheet(wb, resolvedLayout);
  if (!primarySn || !wb.Sheets[primarySn]) return null;
  const ws = wb.Sheets[primarySn];
  if (!ws['!ref']) return null;
  const range = XLSX.utils.decode_range(ws['!ref']);

  const sheetWeekBounds = resolvedWeekBounds.filter(w => w.sheetName === primarySn);

  let weekResult = null;

  if (pattern === 'horizontal-weeks') {
    weekResult = _ap_extractHorizontal(ws, range, sheetWeekBounds);
  } else {
    // vertical or block-based
    // Identify the primary workout region for this sheet
    let primaryRegion = null;
    if (resolvedRegions && resolvedRegions.length > 0) {
      const sheetRegions = resolvedRegions.filter(r => !r.sheetName || r.sheetName === primarySn);
      primaryRegion = sheetRegions.find(r => r.workoutScore >= 0.5) || sheetRegions[0] || resolvedRegions[0];
    }
    if (!primaryRegion) {
      primaryRegion = { startRow: range.s.r, endRow: range.e.r,
                        startCol: range.s.c, endCol: range.e.c };
    }

    const colMap = (columns && columns[primarySn]) ||
      _inferColumns(ws, primaryRegion,
        (headers && headers[primarySn]) || _detectHeaders(ws, primaryRegion));

    const dayBoundsArr = (dayBounds && dayBounds[primarySn]) ||
      _detectDayBoundaries(ws, primaryRegion, resolvedLayout);

    weekResult = _ap_extractVertical(ws, range, colMap, dayBoundsArr, sheetWeekBounds);
  }

  if (!weekResult || weekResult.length === 0) return null;

  return [{
    name: _ap_inferProgramName(wb, primarySn),
    format: 'adaptive',
    athleteName: '',
    dateRange: '',
    maxes: { squat: null, bench: null, deadlift: null },
    weeks: weekResult
  }];
}

/**
 * _ap_extractVertical(ws, range, colMap, dayBoundsArr, weekBoundsArr)
 * Extract workout data from a vertical or block-based layout.
 * Weeks are separated by "WEEK N" label rows; days by blank rows / headers.
 *
 * @param {object} ws            — SheetJS worksheet
 * @param {object} range         — XLSX decoded range
 * @param {object} colMap        — from _inferColumns
 * @param {Array}  dayBoundsArr  — from _detectDayBoundaries
 * @param {Array}  weekBoundsArr — from _detectWeekBoundaries for this sheet
 * @returns {Array<{label, days}>}
 */
function _ap_extractVertical(ws, range, colMap, dayBoundsArr, weekBoundsArr) {
  if (!colMap || colMap.exercise < 0) return null;
  if (!dayBoundsArr || dayBoundsArr.length === 0) return null;

  const weeks = [];

  if (weekBoundsArr && weekBoundsArr.length > 0) {
    // Assign each day boundary to the week that contains its row range
    const weekDayMap = weekBoundsArr.map(wb_e => ({ label: wb_e.label, days: [] }));

    for (const db of dayBoundsArr) {
      let assigned = false;
      for (let wi = 0; wi < weekBoundsArr.length; wi++) {
        const we = weekBoundsArr[wi];
        const wkStart = we.startRow != null ? we.startRow : range.s.r;
        const wkEnd   = we.endRow   != null ? we.endRow   : range.e.r;
        if (db.startRow >= wkStart && db.endRow <= wkEnd) {
          weekDayMap[wi].days.push(db);
          assigned = true;
          break;
        }
      }
      // If a day doesn't fit in any week, assign to last week
      if (!assigned && weekDayMap.length > 0) {
        weekDayMap[weekDayMap.length - 1].days.push(db);
      }
    }

    for (const { label, days } of weekDayMap) {
      const extractedDays = _ap_extractDaysFromBounds(ws, days, colMap);
      if (extractedDays.length > 0) weeks.push({ label, days: extractedDays });
    }
  } else {
    // No explicit week boundaries — treat whole region as one week
    const extractedDays = _ap_extractDaysFromBounds(ws, dayBoundsArr, colMap);
    if (extractedDays.length > 0) weeks.push({ label: 'Week 1', days: extractedDays });
  }

  return weeks.length > 0 ? weeks : null;
}

/**
 * _scoreAdaptiveConfidence(result, detectionContext)
 * Compute a 0.0-1.0 confidence score for an adaptive parse result.
 *
 * Six components (spec §Confidence Scoring):
 *   Exercise recognition  (25%): % of names matching SYNONYM_MAP or lift keywords
 *   Header detection      (20%): headerScore ≥70→1.0, 40→0.7, 10→0.4, <10→0.0
 *   Column mapping        (20%): 4/4 core cols=1.0, 3/4=0.75, 2/4=0.5, 1/4=0.25
 *   Layout confidence     (15%): clean pattern match=1.0, ambiguous=0.5
 *   Data completeness     (10%): % exercises with non-empty prescription
 *   Value validity        (10%): % prescriptions parseable by parseSets/parseSimple
 *
 * @param {Array}  result           — blocks array from _adaptiveExtract
 * @param {object} detectionContext — {
 *   headerScore: number,     // from _detectHeaders.score (0-100)
 *   columnMap: object,       // from _inferColumns
 *   layoutConfidence: number // 0.0-1.0
 * }
 * @returns {number} 0.0-1.0
 */
function _scoreAdaptiveConfidence(result, detectionContext) {
  if (!result || !Array.isArray(result) || result.length === 0) return 0;
  const ctx = detectionContext || {};

  // ── Component 1: Exercise recognition (25%) ────────────────────────────
  let totalEx = 0, recognizedEx = 0;
  const LIFT_KW_RE = /\b(squat|bench|deadlift|press|row|pull|curl|dip|lunge|carry|raise|extension|pushdown|plank|fly|flye|hinge|rdl|sumo|ohp|sldl|rdl)\b/i;
  for (const block of result) {
    for (const week of (block.weeks || [])) {
      for (const day of (week.days || [])) {
        for (const ex of (day.exercises || [])) {
          totalEx++;
          const name = ex.name || '';
          const isKnown = canonicalizeLift(name) != null || LIFT_KW_RE.test(name);
          if (isKnown) recognizedEx++;
        }
      }
    }
  }
  const exerciseScore = totalEx > 0 ? recognizedEx / totalEx : 0;

  // ── Component 2: Header detection (20%) ────────────────────────────────
  const rawHeaderScore = ctx.headerScore || 0;
  const headerScore =
    rawHeaderScore >= 70 ? 1.0 :
    rawHeaderScore >= 40 ? 0.7 :
    rawHeaderScore >= 10 ? 0.4 : 0;

  // ── Component 3: Column mapping (20%) ──────────────────────────────────
  const cm = ctx.columnMap || {};
  const coreFound = [cm.exercise, cm.sets, cm.reps, cm.weight].filter(v => v != null && v >= 0).length;
  const columnScore = coreFound >= 4 ? 1.0 : coreFound === 3 ? 0.75 : coreFound === 2 ? 0.5 : 0.25;

  // ── Component 4: Layout confidence (15%) ───────────────────────────────
  const layoutScore = ctx.layoutConfidence != null ? ctx.layoutConfidence : 0.75;

  // ── Component 5: Data completeness (10%) ───────────────────────────────
  let withPres = 0;
  for (const block of result) {
    for (const week of (block.weeks || [])) {
      for (const day of (week.days || [])) {
        for (const ex of (day.exercises || [])) {
          if (ex.prescription && String(ex.prescription).trim()) withPres++;
        }
      }
    }
  }
  const completenessScore = totalEx > 0 ? withPres / totalEx : 0;

  // ── Component 6: Value validity (10%) ──────────────────────────────────
  let totalPres = 0, validPres = 0;
  for (const block of result) {
    for (const week of (block.weeks || [])) {
      for (const day of (week.days || [])) {
        for (const ex of (day.exercises || [])) {
          const pres = String(ex.prescription || '').trim();
          if (!pres || pres === '-' || pres === '\u2014') continue;
          totalPres++;
          const parsed = parseSets(pres);
          const simple = parseSimple(pres);
          if ((parsed && parsed.length > 0) || simple) validPres++;
        }
      }
    }
  }
  const validityScore = totalPres > 0 ? validPres / totalPres : 0.5;

  // ── Weighted sum ────────────────────────────────────────────────────────
  const confidence =
    exerciseScore     * 0.25 +
    headerScore       * 0.20 +
    columnScore       * 0.20 +
    layoutScore       * 0.15 +
    completenessScore * 0.10 +
    validityScore     * 0.10;

  return Math.min(Math.max(confidence, 0), 1);
}

// ── ADAPTIVE PARSER ORCHESTRATOR ─────────────────────────────────────────────
/**
 * parseAdaptive(wb)
 * Orchestrate the full two-pass adaptive pipeline.
 *   Pass 1: occupancy matrix → region detection → layout classification
 *   Pass 2: column inference → data extraction → confidence scoring
 *
 * Attaches `adaptiveConfidence` (0.0-1.0) to blocks[0].
 * Returns standard block array format (same shape as all other parsers).
 */
function parseAdaptive(wb) {
  try {
    if (!wb || !wb.SheetNames || !wb.SheetNames.length) return null;

    // ── Pass 1: Layout Detection ──────────────────────────────────────────
    // Use first non-empty sheet for matrix / region analysis.
    // _adaptiveExtract handles multi-sheet routing internally.
    const primarySn0 = wb.SheetNames.find(sn => wb.Sheets[sn] && wb.Sheets[sn]['!ref']) || wb.SheetNames[0];
    const ws0 = wb.Sheets[primarySn0];
    if (!ws0 || !ws0['!ref']) return null;

    const matrix = _buildOccupancyMatrix(ws0);
    const regions = _detectRegions(matrix, ws0);
    const layout = _classifyLayout(wb, regions);

    // ── Pass 2: Full Extraction ───────────────────────────────────────────
    const blocks = _adaptiveExtract(wb, regions, layout, null, null, null, null);
    if (!blocks || !blocks.length) return null;

    // ── Confidence Scoring ────────────────────────────────────────────────
    const layoutConfidenceMap = {
      'week-per-tab':     1.0,
      'horizontal-weeks': 1.0,
      'block-based':      0.75,
      'vertical':         0.5
    };
    const layoutConf = layoutConfidenceMap[layout.pattern] || 0.5;

    // Gather header / column context from the primary workout region
    const primaryRegion = regions.find(r => r.workoutScore >= 0.5) || regions[0] || null;
    let detectionContext = { layoutConfidence: layoutConf };
    if (primaryRegion) {
      const headerInfo = _detectHeaders(ws0, primaryRegion);
      const colMap = _inferColumns(ws0, primaryRegion, headerInfo);
      detectionContext = {
        headerScore: headerInfo ? headerInfo.score : 0,
        columnMap: colMap,
        layoutConfidence: layoutConf
      };
    }

    const confidence = _scoreAdaptiveConfidence(blocks, detectionContext);
    blocks[0].adaptiveConfidence = confidence;

    return blocks;
  } catch (e) {
    console.error('[liftlog] parseAdaptive error:', e);
    return null;
  }
}

// ── MODULE EXPORTS (Node.js) / GLOBAL (browser) ──────────────────────────────
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    LIFT_GROUPS, COMP_KEYWORDS, VARIATION_MODIFIERS, SYNONYM_MAP,
    isCompLift, _classifyLiftBase, classifyLift: _classifyLiftBase, canonicalizeLift, _getBaseName,
    _smartExerciseMatch, _extractPrimaryKeyword,
    EXERCISE_DICT, editDistance, spellCorrectWord, spellCorrectExerciseName,
    DAYS, detectFormat, detectF, detectG,
    parseA, parseASheet, parseB, parseBSheet,
    parseWorkbook,
    parseE, parseE_nSuns, parseE_tabular, parseE_phaseGrid, parseE_weekText, parseE_cycleWeek, parseE_parallel531,
    parseF, parseF_stride3, parseF_stride15, parseF_stride6, parseF_pctLoad, parseF_stride1,
    parseG, _gBuildExercise,
    parseCAutoFormat, extractCAutoMeta, parseCAutoSheet,
    parseD, parseDSheet, buildDPrescription,
    deduplicateExerciseNames,
    parseSets, parseSimple, parseRangeSet, parseBarePres, parseWarmupSets,
    parseCSVRows, detectCSVFormat, parseCSVImport,
    detectTexasMethodFmt, parseTexasMethod,
    detectHepburnFmt, parseHepburn,
    detectBulgarianFmt, parseBulgarianMethod,
    detectH, parseH, parseH_coan, parseH_hatch, parseH_edcoan, parseH_hatfield, parseH_deathbench,
    detectL, parseL, _classifyCandito,
    _parseL_standard, _parseL_benchHybrid, _parseL_advancedBench,
    _parseL_advancedDeadlift, _parseL_advancedSquat, _parseL_linear,
    detectM, parseM, _mBuildExercise, _mFmtReps,
    detectN, parseN, _nClassify, _nRound, _nFmtReps,
    detectO, parseO, _oFinalize,
    detectP, parseP,
    detectQ, parseQ,
    detectR, parseR,
    detectS, parseS,
    detectT, parseT,
    _extractMaxesFromSheet,
    detectU, parseU, _parseU_multiSheet, _parseU_singleSheet, _uBuildPrescription,
    detectV, parseV, _vBuildPrescription,
    scoreW, parseW, _w_cleanRows, _w_buildPrescription, _w_computeDateRange,
    _fixParenTypos, _isCircuitPrefix, _isCoachInstruction, _isSectionHeader, _isHeaderTerm, HEADER_BLOCKLIST,
    EXERCISE_ALIASES, _normalizeExerciseName,
    // Adaptive Parser — Session 1
    _resolveMergedCells, _buildOccupancyMatrix, _scoreRegionWorkoutLikeness, _detectRegions,
    // Adaptive Parser — Session 2
    _classifySheetLayout, _classifyLayout, _detectHeaders, _detectDayBoundaries, _detectWeekBoundaries,
    // Adaptive Parser — Session 3
    _scoreColumnForType, _inferColumns, _buildPrescription, _handleMultiRowExercise,
    _ap_findPrimarySheet, _ap_inferProgramName, _ap_extractDaysFromBounds,
    _ap_extractHorizontal, _ap_extractHybridHorizontal, _ap_extractVertical, _ap_extractWeekPerTab,
    _adaptiveExtract, _scoreAdaptiveConfidence,
    // Adaptive Parser — Session 4
    scoreAdaptive, parseAdaptive,
  };
}

