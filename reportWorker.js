// ── reportWorker.js ───────────────────────────────────────────────────────────
// Web Worker: computes training report data off the main thread.
// All constants and functions are copied verbatim from their source files.
// Zero DOM APIs — no document, window, localStorage, getComputedStyle, etc.
//
// SYNC table — functions/constants copied from source files:
// SYNC: KG_TO_LBS                — index.html:15638
// SYNC: LIFT_GROUPS              — parsers.js:7
// SYNC: COMP_KEYWORDS            — parsers.js:16
// SYNC: VARIATION_MODIFIERS      — parsers.js:22
// SYNC: LIFT_GROUP_DISQUALIFIERS — parsers.js:119
// SYNC: SYNONYM_MAP              — parsers.js:150
// SYNC: RPE_TABLE                — index.html:2858
// SYNC: _sanitizeExcelStr         — parsers.js:100
// SYNC: isCompLift               — parsers.js:112
// SYNC: _getBaseName             — parsers.js:107
// SYNC: classifyLift             — parsers.js:134  (_classifyLiftBase renamed)
// SYNC: canonicalizeLift         — parsers.js:386
// SYNC: epley                    — index.html:2850
// SYNC: rpeAdjusted1rm           — index.html:2871
// SYNC: calc1rm                  — index.html:2881
// SYNC: sessionOrder             — index.html:3011
// SYNC: convertWeight            — index.html:4455
// SYNC: displayBlockName         — index.html:4821
// SYNC: calcDOTS                 — index.html:4835
// SYNC: calcWilks                — index.html:4845
// SYNC: _isAmrapPrescription     — index.html:10745
// SYNC: _rpt_classifyTrend       — index.html:10756
// SYNC: _getReportDataWorker     — adapted from index.html:10811

'use strict';

// ── CONSTANTS ─────────────────────────────────────────────────────────────────

const KG_TO_LBS = 2.2046; // USPA standard conversion factor (1 kg = KG_TO_LBS lbs)

const LIFT_GROUPS = {
  squat: ['squat','front squat','zercher','belt squat','box squat','high bar squat','high bar','low bar squat','low bar','pin squat','pause squat','safety bar','ssb','1/4 squat','quarter squat','partial squat',
    'hatfield','walkout'],
  bench: ['bench','close grip bench','close grip press','close grip','larson','larsen','slingshot','incline press','incline bench','decline press','decline bench','floor press','feet up bench','pause bench','touch and go bench','touch and go press','tng bench','tng press','comp bench','board press',
    'spoto','spoto press','dead bench','pin press','cgbp','wgbp','rgbp'],
  deadlift: ['deadlift','sumo deadlift','sumo pull','sumo dead','conventional deadlift','conventional pull','romanian','rdl','rack pull','block pull','deficit deadlift','deficit pull','stiff leg','trap bar','hex bar','snatch grip','pause dead','semi sumo',
    'halting','mat pull','pin pull','sldl','sgdl'],
};

const COMP_KEYWORDS = {
  squat:    ['squat','squats','back squat','back squats','low bar squat','low bar squats','barbell squat','barbell squats','comp squat','competition squat','barbell back squat','barbell low bar squat','barbell comp squat','barbell competition squat'],
  bench:    ['bench','bench press','flat bench','barbell flat bench','barbell bench','barbell bench press','comp bench','comp bench press','barbell comp bench','barbell comp bench press','competition bench','competition bench press','paused bench','paused bench press','pause bench','pause bench press','1 sec bench press','1 sec bench','1 sec pause bench','1 sec paused bench','1 sec paused bench press','1ct bench','1ct bench press','1 count bench','1 count bench press','barbell pause bench','barbell paused bench','barbell paused bench press'],
  deadlift: ['deadlift','dead','barbell deadlift','conventional deadlift','comp deadlift','competition deadlift','conventional pull','comp pull','competition style deadlift','barbell conventional deadlift','barbell comp deadlift','sumo deadlift','sumo pull','sumo','barbell sumo deadlift','barbell sumo','deadlift (primary stance)'],
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
  // Tempo/effort modifiers
  'paused',
  'dead stop','speed','dynamic effort','eccentric','isometric','floating','1.5 rep',
  // Stance/position (sumo deadlift/pull removed — sumo IS a competition deadlift)
  'wide stance','narrow stance','sumo stance','heel elevated','beltless',
  // High bar squat — variation, not comp (some lifters use as comp but default to variation)
  'high bar',
];

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
  '1 sec bench press':    'Bench Press',
  '1 sec bench':          'Bench Press',
  '1 sec pause bench':    'Bench Press',
  '1 sec paused bench':   'Bench Press',
  '1 sec paused bench press': 'Bench Press',
  '1ct bench':            'Bench Press',
  '1ct bench press':      'Bench Press',
  '1 count bench':        'Bench Press',
  '1 count bench press':  'Bench Press',
  'barbell pause bench':  'Bench Press',
  'barbell paused bench': 'Bench Press',
  'barbell paused bench press': 'Bench Press',
  'raw bench':            'Bench Press',

  // ── Sumo Deadlift (comp variant — distinct from conventional) ──
  'sumo deadlift':        'Sumo Deadlift',
  'sumo pull':            'Sumo Deadlift',
  'sumo':                 'Sumo Deadlift',
  'barbell sumo deadlift':'Sumo Deadlift',
  'barbell sumo':         'Sumo Deadlift',

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

  // ── High Bar Squat additions ──
  'high bar back squat':   'High Bar Squat',
  'hb back squat':         'High Bar Squat',
  'barbell high bar squat':'High Bar Squat',

  // ── Nordic Curl synonyms ──
  'nordic curl':           'Nordic Curl',
  'nordic hamstring curl': 'Nordic Curl',

  // ── Chest Supported Row synonyms ──
  'chest supported row':   'Chest Supported Row',
  'chest-supported row':   'Chest Supported Row',

  // ── Cable Fly synonyms ──
  'cable fly':             'Cable Fly',
  'cable flye':            'Cable Fly',
  'cable crossover':       'Cable Fly',

  // ── Incline Dumbbell Press synonyms ──
  'incline dumbbell press':'Incline Dumbbell Press',
  'dumbbell incline press':'Incline Dumbbell Press',
  'incline db press':      'Incline Dumbbell Press',

  // ── Tricep Extension synonyms ──
  'tricep extension':      'Tricep Extension',
  'triceps extension':     'Tricep Extension',

  // ── Calf Raise variations ──
  'standing calf raise':   'Standing Calf Raise',
  'seated calf raise':     'Seated Calf Raise',

  // ── Hip Abduction / Adduction ──
  'hip abduction':         'Hip Abduction',
  'hip adduction':         'Hip Adduction',
};

// RTS RPE chart: maps reps @ RPE to % of 1RM
// RPE_TABLE[reps][rpe] = fraction of 1RM
const RPE_TABLE = {
  1:  {10:1.000, 9.5:0.978, 9:0.955, 8.5:0.939, 8:0.922, 7.5:0.907, 7:0.892, 6.5:0.878, 6:0.863, 5.5:0.849, 5:0.835},
  2:  {10:0.955, 9.5:0.939, 9:0.922, 8.5:0.907, 8:0.892, 7.5:0.878, 7:0.863, 6.5:0.849, 6:0.835, 5.5:0.822, 5:0.808},
  3:  {10:0.922, 9.5:0.907, 9:0.892, 8.5:0.878, 8:0.863, 7.5:0.849, 7:0.835, 6.5:0.822, 6:0.808, 5.5:0.794, 5:0.781},
  4:  {10:0.892, 9.5:0.878, 9:0.863, 8.5:0.849, 8:0.835, 7.5:0.822, 7:0.808, 6.5:0.794, 6:0.781, 5.5:0.768, 5:0.755},
  5:  {10:0.863, 9.5:0.849, 9:0.835, 8.5:0.822, 8:0.808, 7.5:0.794, 7:0.781, 6.5:0.768, 6:0.755, 5.5:0.742, 5:0.730},
  6:  {10:0.835, 9.5:0.822, 9:0.808, 8.5:0.794, 8:0.781, 7.5:0.768, 7:0.755, 6.5:0.742, 6:0.730, 5.5:0.717, 5:0.705},
  7:  {10:0.808, 9.5:0.794, 9:0.781, 8.5:0.768, 8:0.755, 7.5:0.742, 7:0.730, 6.5:0.717, 6:0.705, 5.5:0.693, 5:0.681},
  8:  {10:0.781, 9.5:0.768, 9:0.755, 8.5:0.742, 8:0.730, 7.5:0.717, 7:0.705, 6.5:0.693, 6:0.681, 5.5:0.669, 5:0.658},
  9:  {10:0.755, 9.5:0.742, 9:0.730, 8.5:0.717, 8:0.705, 7.5:0.693, 7:0.681, 6.5:0.669, 6:0.658, 5.5:0.647, 5:0.636},
  10: {10:0.730, 9.5:0.717, 9:0.705, 8.5:0.693, 8:0.681, 7.5:0.669, 7:0.658, 6.5:0.647, 6:0.636, 5.5:0.625, 5:0.614},
};

// ── PURE HELPER FUNCTIONS ─────────────────────────────────────────────────────

// SYNC: _sanitizeExcelStr — parsers.js:100
function _sanitizeExcelStr(s){
  if(!s || typeof s !== 'string') return '';
  return s
    .replace(/[\u2018\u2019\u201A\uFF07]/g, "'")   // smart single quotes → ASCII
    .replace(/[\u201C\u201D\u201E\uFF02]/g, '"')    // smart double quotes → ASCII
    .replace(/[\u2013\u2014]/g, '-')                 // en-dash / em-dash → hyphen
    .replace(/\u00D7/g, 'x')                         // multiplication sign → x
    .replace(/[\u2026]/g, '...')                      // ellipsis → three dots
    .replace(/\s+/g, ' ')                            // collapse all whitespace (incl \u00A0)
    .trim();
}

function isCompLift(name){
  // Strip dedup bracket suffix + set-type qualifiers (backdowns = same lift, lighter sets)
  // + equipment modifiers (belt/wraps/sleeves/straps don't change competition classification)
  const n = _sanitizeExcelStr(name).toLowerCase()
    .replace(/\s*\[.*\]$/, '')
    .replace(/\s+backdowns?\s*$/, '')
    .replace(/\s*w\/?\/?\s*(belt|wraps?|sleeves?|straps?)\s*$/i, '')
    .replace(/\s*with\s+(belt|wraps?|sleeves?|straps?)\s*$/i, '')
    .trim();
  if (VARIATION_MODIFIERS.some(v => n.includes(v))) return false;
  for (const keywords of Object.values(COMP_KEYWORDS)) {
    if (keywords.includes(n)) return true;
  }
  return false;
}

// Helper: extract base name from potentially deduplicated exercise name
function _getBaseName(exName) {
  let name = exName;
  // Strip bracket-based dedup suffixes, e.g. "Barbell Deadlift [1x1(352)]"
  const m = name.match(/^(.+?)\s*\[.*\]$/);
  if (m) name = m[1];
  // Strip set-type qualifiers that describe the same lift for metrics purposes
  // e.g. "Comp Bench Backdowns" → "Comp Bench" (tracks under same lift)
  name = name.replace(/\s+backdowns?\s*$/i, '');
  return name;
}

// Base classifier — returns group name or null. No access to runtime state.
// Renamed from _classifyLiftBase; returns null for non-matches (no custom lifts, no 'other' fallback).
function classifyLift(name){
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

function epley(weight,reps){
  if(!weight||!reps||reps<1) return null;
  if(reps===1) return weight;
  return Math.round(weight*(1+reps/30));
}

function rpeAdjusted1rm(weight, reps, rpe){
  if(!weight||!reps||!rpe) return null;
  const rpeVal=Math.round(parseFloat(rpe)*2)/2; // round to nearest 0.5
  const repRow=RPE_TABLE[Math.min(Math.max(Math.round(reps),1),10)];
  if(!repRow) return null;
  const pct=repRow[rpeVal];
  if(!pct) return null;
  return Math.round(weight/pct);
}

function calc1rm(weight, reps, rpe){
  // Prefer RPE-adjusted if RPE is available, fall back to Epley
  if(rpe&&parseFloat(rpe)>=5){
    const rpeResult=rpeAdjusted1rm(weight,reps,rpe);
    if(rpeResult) return {value:rpeResult, method:'rpe'};
  }
  const epleyResult=epley(weight,reps);
  return epleyResult?{value:epleyResult, method:'epley'}:null;
}

function sessionOrder(label){
  // Try to extract week number: "W1", "W2", "10 Weeks Out", "9 weeks out"
  const wm=label.match(/W(\d+)/i); if(wm) return parseInt(wm[1]);
  const nm=label.match(/(\d+)/); if(nm) return 1000-parseInt(nm[1]); // "10 weeks out" -> earlier = higher number
  return 999;
}

// Convert weight between units
function convertWeight(value, fromUnit, toUnit) {
  if(fromUnit === toUnit || !value) return value;
  return fromUnit === 'lbs' ? value / KG_TO_LBS : value * KG_TO_LBS;
}

function displayBlockName(name) {
  if (!name) return '';
  // Strip trailing timestamp suffix (e.g. "_1772319728896")
  let clean = name.replace(/_\d{10,}$/, '');
  // Clean up underscores from filename-based names
  clean = clean.replace(/_/g, ' ').trim();
  // Truncate to 36 chars
  if (clean.length > 36) clean = clean.slice(0, 34) + '\u2026';
  return clean;
}

// All inputs in kg, total in kg, returns numeric score
function calcDOTS(bwKg, totalKg, gender) {
  if (!bwKg || !totalKg || bwKg <= 0 || totalKg <= 0) return null;
  const m = gender === 'female'
    ? { a:-57.96288,    b:13.6175032,   c:-0.1126655495, d:0.0005158568,  e:-0.0000010706 }
    : { a:-307.75076,   b:24.0900756,   c:-0.1918759221, d:0.0007391293,  e:-0.0000010930 };
  const x = bwKg;
  const coeff = 500 / (m.a + m.b*x + m.c*x*x + m.d*x*x*x + m.e*x*x*x*x);
  return Math.round(coeff * totalKg * 100) / 100;
}

function calcWilks(bwKg, totalKg, gender) {
  if (!bwKg || !totalKg || bwKg <= 0 || totalKg <= 0) return null;
  // Wilks formula — coefficients from Wilks 1998
  const m = gender === 'female'
    ? { a:594.31747775582, b:-27.23842536447, c:0.82112226871, d:-0.00930733913, e:0.00004731582, f:-0.00000009054 }
    : { a:-216.0475144,    b:16.2606339,      c:-0.002388645,  d:-0.00113732,    e:7.01863e-06,   f:-1.291e-08 };
  const x = bwKg;
  const poly = m.a + m.b*x + m.c*x*x + m.d*x*x*x + m.e*x*x*x*x + m.f*x*x*x*x*x;
  return Math.round((500 / poly) * totalKg * 100) / 100;
}

// ── AMRAP prescription detection ───────────────────────────────────────────
function _isAmrapPrescription(pres) {
  if (!pres) return false;
  const p = String(pres);
  if (/amrap/i.test(p)) return true;
  if (/as\s+many\s+reps\s+as\s+possible/i.test(p)) return true;
  if (/(?:^|x|×|\s)MR(?:\s|\(|$)/i.test(p)) return true;
  if (/\d\+(?!\d)/.test(p)) return true; // "3+", "5+", "1+" but not "10"
  return false;
}

// ── Report sub-functions (pure trend classifier — no closure deps) ────────────
function _rpt_classifyTrend(cl) {
  if (!cl || !cl.e1rmByWeek) return { text: 'Insufficient data', color: 'gold' };
  const vals = cl.e1rmByWeek.filter(v => v !== null);
  if (vals.length < 2) return { text: 'Insufficient data', color: 'gold' };
  const rpeVals = cl.rpeArr ? cl.rpeArr.filter(v => v !== null) : [];
  const hasRpe = rpeVals.length >= 2;
  const mid = Math.floor(vals.length / 2);
  const firstHalf = vals.slice(0, mid);
  const secondHalf = vals.slice(mid);
  const avgFirst = firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length;
  const avgSecond = secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length;
  const e1rmDelta = avgSecond - avgFirst;
  const e1rmPct = avgFirst > 0 ? (e1rmDelta / avgFirst) * 100 : 0;
  let regressionWeeks = 0;
  for (let i = 1; i < vals.length; i++) {
    if (vals[i] < vals[i - 1] * 0.98) regressionWeeks++;
  }
  let rpeTrend = 'stable';
  if (hasRpe) {
    const rpeMid = Math.floor(rpeVals.length / 2);
    const rpeFirst = rpeVals.slice(0, rpeMid);
    const rpeSecond = rpeVals.slice(rpeMid);
    const rpeAvgFirst = rpeFirst.reduce((a, b) => a + b, 0) / rpeFirst.length;
    const rpeAvgSecond = rpeSecond.reduce((a, b) => a + b, 0) / rpeSecond.length;
    const rpeDelta = rpeAvgSecond - rpeAvgFirst;
    if (rpeDelta > 0.4) rpeTrend = 'rising';
    else if (rpeDelta < -0.4) rpeTrend = 'falling';
  }
  const e1rmDeltaStr = (e1rmDelta >= 0 ? '+' : '') + Math.round(e1rmDelta);
  const rpeDeltaStr = hasRpe ? ((rpeVals[rpeVals.length-1] - rpeVals[0]) >= 0 ? '+' : '') + (rpeVals[rpeVals.length-1] - rpeVals[0]).toFixed(1) : '';
  const rpeNote = hasRpe && rpeTrend === 'rising' ? ` \u00b7 RPE ${rpeDeltaStr}` : hasRpe && rpeTrend === 'falling' ? ` \u00b7 RPE easing` : '';
  if (e1rmPct > 2 && rpeTrend !== 'rising' && regressionWeeks === 0) {
    return { text: `Steady climb (${e1rmDeltaStr}) \u2014 no dips${rpeNote}`, color: 'green' };
  } else if (e1rmPct > 2 && rpeTrend !== 'rising') {
    return { text: `Strong uptrend (${e1rmDeltaStr})${rpeNote}`, color: 'green' };
  } else if (e1rmPct > 1 && rpeTrend === 'stable') {
    return { text: `Gradual climb (${e1rmDeltaStr}) \u2014 consistent RPE`, color: 'green' };
  } else if (e1rmPct > 1 && rpeTrend === 'rising') {
    return { text: `Grinding higher (${e1rmDeltaStr})${rpeNote}`, color: 'gold' };
  } else if (Math.abs(e1rmPct) <= 1 && rpeTrend === 'stable') {
    return { text: `Maintained (${e1rmDeltaStr}) \u2014 stable effort`, color: 'gold' };
  } else if (Math.abs(e1rmPct) <= 1 && rpeTrend === 'rising') {
    return { text: `Stalling (${e1rmDeltaStr})${rpeNote}`, color: 'amber' };
  } else if (e1rmPct < -1 && rpeTrend === 'rising') {
    return { text: `Regression (${e1rmDeltaStr})${rpeNote}`, color: 'red' };
  } else if (e1rmPct < -1) {
    return { text: `Declining (${e1rmDeltaStr})${rpeNote}`, color: 'red' };
  }
  return { text: `Mixed signals (${e1rmDeltaStr})${rpeNote}`, color: 'gold' };
}

// ── MAIN WORKER FUNCTION ──────────────────────────────────────────────────────
// Adapted from _getReportData(blockId) in index.html:10811.
// Receives a pre-built snapshot object instead of reading from DOM/state.
//
// snapshot shape:
// {
//   blockId, block, blockLogs, displayUnit, athlete,
//   prescriptionMap, setRepsMap, prescribedWeightMap,
//   prevBlock, prevBlockLogs, prevPrescriptionMap, prevSetRepsMap, prevPrescribedWeightMap
// }
function _getReportDataWorker(snapshot) {
  const block = snapshot.block;
  if (!block) return null;
  const weeks = block.weeks || [];

  const blockLogs = snapshot.blockLogs;

  // ── Pre-scan: identify comp lift names per group (mirrors metrics hero cards line 2878) ──
  const compLiftName = { squat: null, bench: null, deadlift: null };
  for (const [key] of Object.entries(blockLogs)) {
    const parts = key.split('||');
    if (parts.length < 5 || !parts[4].startsWith('s')) continue;
    const rawName = parts[3];
    if (!isCompLift(rawName)) continue;
    const g = classifyLift(rawName);
    if (g && !compLiftName[g]) compLiftName[g] = canonicalizeLift(rawName);
  }

  // ── Gather e1RM data by family → exercise → week ─────────────────────────
  const families = { squat: {}, bench: {}, deadlift: {} };

  // ── AMRAP tracking: best AMRAP set per lift family per week ──────────────
  const amrapData = { squat: {}, bench: {}, deadlift: {} };
  // amrapData[group][weekLabel] = { weight, reps, e1rm, exName, prescribedReps }

  // Use pre-built prescription lookup cache from snapshot
  const _presCache = snapshot.prescriptionMap;

  for (const [key, val] of Object.entries(blockLogs)) {
    const parts = key.split('||');
    if (parts.length < 5 || !parts[4].startsWith('s')) continue;
    const exName = canonicalizeLift(parts[3]);
    const group = classifyLift(exName) || classifyLift(parts[3]);
    if (!group || !families[group]) continue;

    const si = parseInt(parts[4].slice(1));
    let weight = null, wUnit = snapshot.displayUnit;
    if (val.actual) { weight = parseFloat(val.actual); wUnit = (val?.unit || 'lbs'); }
    else if (val.done) { weight = snapshot.prescribedWeightMap[key]; }
    if (!weight || isNaN(weight) || weight <= 0) continue;
    weight = convertWeight(weight, wUnit, snapshot.displayUnit);
    const reps = val.actualReps ? parseInt(val.actualReps) : snapshot.setRepsMap[key];
    if (!reps) continue;
    const rpe = val.rpe ? parseFloat(val.rpe) : null;
    const result = calc1rm(weight, reps, rpe);
    if (!result) continue;

    const weekLabel = parts[1];
    if (!families[group][exName]) families[group][exName] = {};
    if (!families[group][exName][weekLabel])
      families[group][exName][weekLabel] = { e1rm: 0, topWeight: 0, topReps: 0, topRpe: null, method: 'epley' };
    const w = families[group][exName][weekLabel];
    if (result.value > w.e1rm) { w.e1rm = result.value; w.topWeight = weight; w.topReps = reps; w.topRpe = rpe; w.method = result.method; }

    // ── Track AMRAP sets ──────────────────────────────────────────────────
    const presKey = parts[1] + '||' + parts[2] + '||' + parts[3];
    const pres = _presCache[presKey];
    if (pres && _isAmrapPrescription(pres) && val.actualReps) {
      const actualReps = parseInt(val.actualReps);
      if (actualReps > 0 && val.done) {
        // Get prescribed reps for this set to compare
        const prescReps = snapshot.setRepsMap[key];
        const prescMin = prescReps ? parseInt(String(prescReps).replace(/\+.*/, '')) : 0;
        if (!amrapData[group][weekLabel] || result.value > amrapData[group][weekLabel].e1rm) {
          amrapData[group][weekLabel] = {
            weight, reps: actualReps, e1rm: result.value, exName,
            prescribedReps: prescMin, exceeded: actualReps > prescMin
          };
        }
      }
    }
  }

  // ── Compliance + volume + RPE tracking ───────────────────────────────────
  let totalDays = 0, completedDays = 0, totalSets = 0, totalVolume = 0;
  let totalRpe = 0, rpeCount = 0;
  const volumeByFamily = { squat: 0, bench: 0, deadlift: 0, other: 0 };
  const volumeByWeekFamily = {}; // weekLabel -> {squat,bench,deadlift,other}
  const rpeByWeek = {}; // weekLabel -> {total, count}

  // Count days with ANY logged set as "attended" (not isDayComplete which requires ALL exercises)
  const daysWithSets = new Set();
  for (const key of Object.keys(blockLogs)) {
    const parts = key.split('||');
    if (parts.length >= 5 && parts[4].startsWith('s') && blockLogs[key].done) {
      daysWithSets.add(parts[1] + '||' + parts[2]); // weekLabel||dayLabel
    }
  }
  for (const week of weeks) {
    for (const day of (week.days || [])) {
      totalDays++;
      if (daysWithSets.has(week.label + '||' + day.name)) completedDays++;
    }
  }

  for (const [key, val] of Object.entries(blockLogs)) {
    const parts = key.split('||');
    if (parts.length < 5 || !parts[4].startsWith('s')) continue;
    if (!val.done) continue;
    const exName = parts[3];
    let weight = val.actual ? parseFloat(val.actual) : snapshot.prescribedWeightMap[key];
    const prescReps = snapshot.setRepsMap[key];
    const actualReps = val.actualReps ? parseInt(val.actualReps) : prescReps;
    if (weight && !isNaN(weight) && weight > 0 && actualReps) {
      const wUnit = val.actual ? (val?.unit || 'lbs') : snapshot.displayUnit;
      weight = convertWeight(weight, wUnit, snapshot.displayUnit);
      totalSets++;
      const vol = weight * actualReps;
      totalVolume += vol;
      const group = classifyLift(exName) || 'other';
      volumeByFamily[group] = (volumeByFamily[group] || 0) + vol;
      const wl = parts[1];
      if (!volumeByWeekFamily[wl]) volumeByWeekFamily[wl] = { squat: 0, bench: 0, deadlift: 0, other: 0 };
      volumeByWeekFamily[wl][group] = (volumeByWeekFamily[wl][group] || 0) + vol;
    }
    // RPE tracking (overall + per-week)
    if (val.rpe) {
      const r = parseFloat(val.rpe);
      if (!isNaN(r)) {
        totalRpe += r; rpeCount++;
        const wl = parts[1];
        if (!rpeByWeek[wl]) rpeByWeek[wl] = { total: 0, count: 0 };
        rpeByWeek[wl].total += r;
        rpeByWeek[wl].count++;
      }
    }
  }

  const compliancePct = totalDays > 0 ? Math.round(completedDays / totalDays * 100) : 0;
  const avgRpe = rpeCount > 0 ? (totalRpe / rpeCount).toFixed(1) : '\u2014';

  // ── Best comp lift e1RM per family ──────────────────────────────────────
  function getBestCompVariation(groupKey) {
    const exercises = families[groupKey];
    const exNames = Object.keys(exercises);
    if (!exNames.length) return null;
    // Prefer the comp lift identified in pre-scan, fall back to highest e1RM variation
    const comp = compLiftName[groupKey];
    const compNames = comp ? exNames.filter(n => n === comp) : [];
    const isVariation = compNames.length === 0;
    let primary;
    if (compNames.length) {
      // If multiple comp-like names, pick the one with the highest best e1RM
      let bestVal = 0;
      primary = compNames[0];
      for (const n of compNames) {
        for (const d of Object.values(exercises[n])) {
          if (d.e1rm > bestVal) { bestVal = d.e1rm; primary = n; }
        }
      }
    } else {
      let bestVal = 0;
      primary = exNames[0];
      for (const n of exNames) {
        for (const d of Object.values(exercises[n])) {
          if (d.e1rm > bestVal) { bestVal = d.e1rm; primary = n; }
        }
      }
    }
    let bestWeek = null, bestE1rm = 0;
    for (const [wl, data] of Object.entries(exercises[primary])) {
      if (data.e1rm > bestE1rm) { bestE1rm = data.e1rm; bestWeek = wl; }
    }
    const allWeeks = Object.keys(exercises[primary]).sort((a, b) => sessionOrder(a) - sessionOrder(b));
    const firstWeek = allWeeks[0];
    const lastWeek = allWeeks[allWeeks.length - 1];
    const startE1rm = exercises[primary][firstWeek]?.e1rm || 0;
    const endE1rm = exercises[primary][lastWeek]?.e1rm || 0;
    const bestData = exercises[primary][bestWeek];
    // Build per-week e1RM array for sparkline
    const weekLabelsAll = weeks.map(w => w.label);
    const e1rmByWeek = weekLabelsAll.map(wl => exercises[primary][wl] ? Math.round(exercises[primary][wl].e1rm) : null);
    // Build per-week RPE array for trend analysis
    const rpeByWeekForLift = {};
    for (const [exN, weekData] of Object.entries(exercises)) {
      for (const [wl, d] of Object.entries(weekData)) {
        if (d.topRpe) {
          if (!rpeByWeekForLift[wl]) rpeByWeekForLift[wl] = { total: 0, count: 0 };
          rpeByWeekForLift[wl].total += d.topRpe;
          rpeByWeekForLift[wl].count++;
        }
      }
    }
    const rpeArr = weekLabelsAll.map(wl => rpeByWeekForLift[wl] ? +(rpeByWeekForLift[wl].total / rpeByWeekForLift[wl].count).toFixed(1) : null);
    return {
      name: primary, isVariation, bestE1rm: Math.round(bestE1rm),
      startE1rm: Math.round(startE1rm), endE1rm: Math.round(endE1rm),
      delta: Math.round(endE1rm - startE1rm),
      bestWeight: bestData?.topWeight, bestReps: bestData?.topReps, bestRpe: bestData?.topRpe,
      e1rmByWeek, rpeArr
    };
  }

  const compLifts = {
    squat: getBestCompVariation('squat'),
    bench: getBestCompVariation('bench'),
    deadlift: getBestCompVariation('deadlift')
  };

  // ── Week labels (needed by generateFlags below) ──────────────────────────
  const weekLabels = weeks.map(w => w.label);

  // ── Per-lift flags generation ────────────────────────────────────────────
  function generateFlags(groupKey, cl) {
    const flags = [];
    if (!cl) return flags;
    const vals = cl.e1rmByWeek.filter(v => v !== null);
    const rpeVals = cl.rpeArr ? cl.rpeArr.filter(v => v !== null) : [];
    const unitLabelShort = snapshot.displayUnit === 'kg' ? 'kg' : 'lb';

    // ── AMRAP flags ──────────────────────────────────────────────────────
    const amrapWeeks = weekLabels.filter(wl => amrapData[groupKey]?.[wl]);
    if (amrapWeeks.length > 0) {
      const amrapSets = amrapWeeks.map(wl => amrapData[groupKey][wl]);

      // Best AMRAP performance
      const best = amrapSets.reduce((a, b) => b.e1rm > a.e1rm ? b : a);
      const bestWeekIdx = amrapWeeks.findIndex(wl => amrapData[groupKey][wl] === best);
      const bestWeekLabel = amrapWeeks[bestWeekIdx];

      // AMRAP progression (week-over-week at same or similar weight)
      if (amrapWeeks.length >= 2) {
        const first = amrapSets[0];
        const last = amrapSets[amrapSets.length - 1];
        const repsDelta = last.reps - first.reps;
        const sameWeight = Math.abs(last.weight - first.weight) < 5;
        if (sameWeight && repsDelta > 0) {
          flags.push({ color: 'green', text: `AMRAP progression: ${Math.round(first.weight)}${unitLabelShort}\u00d7${first.reps} \u2192 \u00d7${last.reps} across block (+${repsDelta} reps)` });
        } else if (sameWeight && repsDelta < 0) {
          flags.push({ color: 'amber', text: `AMRAP reps declining: ${Math.round(first.weight)}${unitLabelShort}\u00d7${first.reps} \u2192 \u00d7${last.reps} across block (${repsDelta} reps)` });
        } else if (!sameWeight) {
          // Different weights — compare e1RM instead
          const e1rmDelta = Math.round(last.e1rm - first.e1rm);
          if (e1rmDelta > 0) {
            flags.push({ color: 'green', text: `AMRAP e1RM up +${e1rmDelta} ${unitLabelShort} across block (${Math.round(first.weight)}\u00d7${first.reps} \u2192 ${Math.round(last.weight)}\u00d7${last.reps})` });
          } else if (e1rmDelta < -5) {
            flags.push({ color: 'amber', text: `AMRAP e1RM down ${e1rmDelta} ${unitLabelShort} across block` });
          }
        }
      } else {
        // Single AMRAP week — just note the performance
        if (best.exceeded) {
          flags.push({ color: 'green', text: `AMRAP set: ${Math.round(best.weight)}${unitLabelShort}\u00d7${best.reps} (prescribed ${best.prescribedReps}+) \u2014 e1RM ${Math.round(best.e1rm)} \u2014 ${bestWeekLabel}` });
        }
      }
    }

    // e1RM trend flag
    if (cl.delta > 0) {
      const rpeStable = rpeVals.length >= 2 && Math.abs(rpeVals[rpeVals.length - 1] - rpeVals[0]) <= 0.5;
      flags.push({ color: 'green', text: `e1RM up +${cl.delta} ${unitLabelShort}${rpeStable ? ' with stable RPE \u2014 strong adaptation signal' : ' across block'}` });
    } else if (cl.delta < 0) {
      flags.push({ color: 'red', text: `e1RM down ${cl.delta} ${unitLabelShort} \u2014 potential overreach` });
    }

    // RPE creep detection
    if (rpeVals.length >= 3) {
      const firstRpe = rpeVals[0];
      const lastRpe = rpeVals[rpeVals.length - 1];
      if (lastRpe - firstRpe >= 1.0) {
        flags.push({ color: 'amber', text: `RPE rose from ${firstRpe} \u2192 ${lastRpe} at comparable loads over block` });
      }
    }

    // Consecutive declining weeks
    if (vals.length >= 3) {
      let consecutiveDecline = 0, maxDecline = 0;
      for (let i = 1; i < vals.length; i++) {
        if (vals[i] < vals[i - 1]) { consecutiveDecline++; maxDecline = Math.max(maxDecline, consecutiveDecline); }
        else consecutiveDecline = 0;
      }
      if (maxDecline >= 2) {
        flags.push({ color: 'red', text: `${maxDecline + 1} consecutive declining e1RM weeks detected` });
      }
    }

    // Session fatigue (within-session e1RM drop) — compute from raw logs
    const fatigueByWeek = computeSessionFatigue(groupKey);
    if (fatigueByWeek.length > 0) {
      const avgFatigue = fatigueByWeek.reduce((a, b) => a + b, 0) / fatigueByWeek.length;
      if (avgFatigue > 7) {
        flags.push({ color: 'red', text: `Session fatigue averaging ${avgFatigue.toFixed(1)}% \u2014 above sustainable threshold` });
      } else if (avgFatigue > 5) {
        flags.push({ color: 'amber', text: `Session fatigue averaging ${avgFatigue.toFixed(1)}% \u2014 approaching threshold` });
      } else if (avgFatigue > 0) {
        flags.push({ color: 'green', text: `Session fatigue averaging ${avgFatigue.toFixed(1)}% \u2014 well within recovery capacity` });
      }
    } else if (vals.length >= 2) {
      // No multi-set fatigue data — count total sets for context (common for deadlift)
      let totalSets = 0, sessionCount = 0;
      const sessKeys = new Set();
      for (const [key, val] of Object.entries(blockLogs)) {
        const parts = key.split('||');
        if (parts.length < 5 || !parts[4].startsWith('s') || !val.done) continue;
        if (classifyLift(parts[3]) !== groupKey) continue;
        totalSets++;
        sessKeys.add(parts[1] + '||' + parts[2]);
      }
      sessionCount = sessKeys.size;
      const avg = sessionCount > 0 ? (totalSets / sessionCount).toFixed(1) : '0';
      flags.push({ color: 'green', text: `Session fatigue: N/A \u2014 ${totalSets} sets across ${sessionCount} sessions (avg ${avg}/session, single-set exercises)` });
    }

    // Flat e1RM with rising RPE = overreach signal
    if (vals.length >= 3 && rpeVals.length >= 3) {
      const e1rmFlat = Math.abs(cl.delta) <= 5;
      const rpeRising = rpeVals[rpeVals.length - 1] - rpeVals[0] >= 0.8;
      if (e1rmFlat && rpeRising) {
        flags.push({ color: 'red', text: 'e1RM flat despite rising effort \u2014 potential overreach', rpeFlag: true });
      }
    }

    return flags;
  }

  // ── Session fatigue: first set e1RM vs last set e1RM per session ────────
  function computeSessionFatigue(groupKey) {
    // Group sets by session (weekLabel + dayName + exName)
    const sessions = {};
    for (const [key, val] of Object.entries(blockLogs)) {
      const parts = key.split('||');
      if (parts.length < 5 || !parts[4].startsWith('s')) continue;
      if (!val.done) continue;
      const exName = parts[3];
      const group = classifyLift(exName);
      if (group !== groupKey) continue;
      const sessionKey = parts[1] + '||' + parts[2] + '||' + exName;
      const si = parseInt(parts[4].slice(1));
      let weight = val.actual ? parseFloat(val.actual) : snapshot.prescribedWeightMap[key];
      if (!weight || isNaN(weight) || weight <= 0) continue;
      const fWUnit = val.actual ? (val?.unit || 'lbs') : snapshot.displayUnit;
      weight = convertWeight(weight, fWUnit, snapshot.displayUnit);
      const reps = val.actualReps ? parseInt(val.actualReps) : snapshot.setRepsMap[key];
      if (!reps) continue;
      const rpe = val.rpe ? parseFloat(val.rpe) : null;
      const result = calc1rm(weight, reps, rpe);
      if (!result) continue;
      if (!sessions[sessionKey]) sessions[sessionKey] = [];
      sessions[sessionKey].push({ si, e1rm: result.value });
    }

    const fatigues = [];
    for (const sets of Object.values(sessions)) {
      if (sets.length < 2) continue;
      sets.sort((a, b) => a.si - b.si);
      const firstE1rm = sets[0].e1rm;
      const lastE1rm = sets[sets.length - 1].e1rm;
      if (firstE1rm > 0) {
        const fatiguePct = ((firstE1rm - lastE1rm) / firstE1rm) * 100;
        if (fatiguePct > 0) fatigues.push(fatiguePct);
      }
    }
    return fatigues;
  }

  // ── Estimated total + DOTS/Wilks ─────────────────────────────────────────
  const estTotal = (compLifts.squat?.bestE1rm || 0) + (compLifts.bench?.bestE1rm || 0) + (compLifts.deadlift?.bestE1rm || 0);
  const athlete = snapshot.athlete;
  let dotsScore = null, wilksScore = null;
  if (athlete?.bodyweight && athlete?.gender && estTotal > 0) {
    const bwKg = athlete.bwUnit === 'kg' ? athlete.bodyweight : athlete.bodyweight / KG_TO_LBS;
    const rptDisplayU = snapshot.displayUnit;
    const totalKg = rptDisplayU === 'kg' ? estTotal : estTotal / KG_TO_LBS;
    dotsScore = calcDOTS(bwKg, totalKg, athlete.gender);
    wilksScore = calcWilks(bwKg, totalKg, athlete.gender);
  }

  // ── Weekly top sets table ────────────────────────────────────────────────
  function getWeeklyTopSet(groupKey, weekLabel) {
    const exercises = families[groupKey];
    let bestComp = null, bestVar = null;
    for (const [exName, weekData] of Object.entries(exercises)) {
      const d = weekData[weekLabel];
      if (!d) continue;
      const isComp = exName === compLiftName[groupKey];
      if (isComp && (!bestComp || d.e1rm > bestComp.e1rm)) {
        bestComp = { ...d, exName, isVariation: false };
      } else if (!isComp && (!bestVar || d.e1rm > bestVar.e1rm)) {
        bestVar = { ...d, exName, isVariation: true };
      }
    }
    const result = bestComp || bestVar;
    // Tag if this week had an AMRAP set for this lift family
    if (result && amrapData[groupKey]?.[weekLabel]) {
      result.isAmrap = true;
      result.amrapReps = amrapData[groupKey][weekLabel].reps;
    }
    return result;
  }

  // Find peak week
  let peakWeekLabel = null, peakWeekTotal = 0;
  for (const wl of weekLabels) {
    const s = getWeeklyTopSet('squat', wl)?.e1rm || 0;
    const b = getWeeklyTopSet('bench', wl)?.e1rm || 0;
    const d = getWeeklyTopSet('deadlift', wl)?.e1rm || 0;
    const total = s + b + d;
    if (total > peakWeekTotal) { peakWeekTotal = total; peakWeekLabel = wl; }
  }

  // ── Per-lift flags ────────────────────────────────────────────────────────
  const allFlags = {
    squat: generateFlags('squat', compLifts.squat),
    bench: generateFlags('bench', compLifts.bench),
    deadlift: generateFlags('deadlift', compLifts.deadlift)
  };
  const generalFlags = [];
  if (compliancePct >= 85) {
    generalFlags.push({ color: 'green', text: `${compliancePct}% compliance \u2014 ${completedDays} of ${totalDays} prescribed sessions completed` });
  } else if (compliancePct >= 70) {
    generalFlags.push({ color: 'amber', text: `${compliancePct}% compliance \u2014 ${completedDays} of ${totalDays} sessions. Consider adherence check-in` });
  } else if (totalDays > 0) {
    generalFlags.push({ color: 'red', text: `${compliancePct}% compliance \u2014 only ${completedDays} of ${totalDays} sessions completed` });
  }

  // ── Volume + RPE chart data ──────────────────────────────────────────────
  const volChartData = {
    squat: weekLabels.map(wl => Math.round((volumeByWeekFamily[wl]?.squat || 0) / 1000)),
    bench: weekLabels.map(wl => Math.round((volumeByWeekFamily[wl]?.bench || 0) / 1000)),
    deadlift: weekLabels.map(wl => Math.round((volumeByWeekFamily[wl]?.deadlift || 0) / 1000)),
    other: weekLabels.map(wl => Math.round((volumeByWeekFamily[wl]?.other || 0) / 1000))
  };
  const weeklyAvgRpe = weekLabels.map(wl => {
    const d = rpeByWeek[wl];
    return d && d.count > 0 ? +(d.total / d.count).toFixed(1) : null;
  });

  const dateStr = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  const athleteName = block.athleteName || athlete?.name || '';

  // ── Status badge computation ────────────────────────────────────────────
  const _liftLabels = { squat: 'Squat', bench: 'Bench', deadlift: 'Deadlift' };
  // Compliance-gated level: if compliance >= 80 and ALL red flags are RPE-related, cap at amber.
  // High compliance + high RPE = intentional peak, not alarm-worthy.
  const _liftLevel = (flags, compPct) => {
    const redFlags = flags.filter(f => f.color === 'red');
    if (redFlags.length === 0) {
      if (flags.some(f => f.color === 'amber')) return 'amber';
      return 'green';
    }
    if (compPct >= 80 && redFlags.every(f => f.rpeFlag)) return 'amber';
    return 'red';
  };
  const liftStatuses = {};
  for (const gk of ['squat', 'bench', 'deadlift']) {
    const cl = compLifts[gk];
    const flags = allFlags[gk] || [];
    const level = _liftLevel(flags, compliancePct);
    const d = cl?.delta || 0;
    const unitLabelShort = snapshot.displayUnit === 'kg' ? 'kg' : 'lb';
    const deltaStr = cl ? ((d >= 0 ? '+' : '') + d + ' ' + unitLabelShort) : '';
    const trend = cl ? _rpt_classifyTrend(cl) : null;
    let summary = '';
    if (level === 'green') summary = 'adapting';
    else if (level === 'amber') summary = flags.find(f => f.color === 'amber')?.text?.split(' \u2014 ')[1] || 'watch';
    else summary = flags.find(f => f.color === 'red')?.text?.split(' \u2014 ')[1] || 'attention needed';
    liftStatuses[gk] = { level, delta: d, deltaStr, summary };
  }
  const redCount = Object.values(liftStatuses).filter(l => l.level === 'red').length;
  const amberCount = Object.values(liftStatuses).filter(l => l.level === 'amber').length;
  let overallStatus = 'green';
  if (redCount >= 2) overallStatus = 'red';
  else if (redCount >= 1 || amberCount >= 2) overallStatus = 'amber';
  else if (amberCount >= 1) overallStatus = 'amber';
  const problemLifts = Object.entries(liftStatuses).filter(([_, l]) => l.level !== 'green').map(([gk]) => _liftLabels[gk]);
  const goodLifts = Object.entries(liftStatuses).filter(([_, l]) => l.level === 'green').map(([gk]) => _liftLabels[gk]);
  let statusText = '';
  if (overallStatus === 'green') {
    statusText = 'All lifts progressing well across the block.';
  } else {
    statusText = '<strong>' + problemLifts.join(' and ') + (problemLifts.length === 1 ? ' needs' : ' need') + ' attention.</strong>';
    // Add detail from the worst flag
    const worstLift = Object.entries(liftStatuses).find(([_, l]) => l.level === 'red') || Object.entries(liftStatuses).find(([_, l]) => l.level === 'amber');
    if (worstLift) {
      const worstFlags = allFlags[worstLift[0]] || [];
      const topFlag = worstFlags.find(f => f.color === 'red') || worstFlags.find(f => f.color === 'amber');
      if (topFlag) statusText += ' ' + topFlag.text + '.';
    }
    if (goodLifts.length) statusText += ' ' + goodLifts.join(' and ') + ' progressing well.';
  }
  const statusBadge = { level: overallStatus, text: statusText, lifts: liftStatuses };

  // ── Block-over-block deltas ─────────────────────────────────────────────
  let blockDeltas = null;
  if (snapshot.prevBlock) {
    const prevSnapshot = {
      blockId: snapshot.prevBlock.id,
      block: snapshot.prevBlock,
      blockLogs: snapshot.prevBlockLogs || {},
      displayUnit: snapshot.displayUnit,
      athlete: snapshot.athlete,
      prescriptionMap: snapshot.prevPrescriptionMap || {},
      setRepsMap: snapshot.prevSetRepsMap || {},
      prescribedWeightMap: snapshot.prevPrescribedWeightMap || {},
      prevBlock: null,
      prevBlockLogs: null,
      prevPrescriptionMap: null,
      prevSetRepsMap: null,
      prevPrescribedWeightMap: null
    };
    const prevData = _getReportDataWorker(prevSnapshot);
    if (prevData) {
      // Count completed sessions per lift family in previous block.
      // A "completed session" = a unique weekLabel||dayName with at least one
      // logged set carrying actual weight data for that lift family.
      const _prevSessions = { squat: new Set(), bench: new Set(), deadlift: new Set() };
      for (const [key, val] of Object.entries(prevData.blockLogs)) {
        const pts = key.split('||');
        if (pts.length < 5 || !pts[4].startsWith('s')) continue;
        if (!val.done) continue;
        const exName = pts[3];
        const grp = classifyLift(canonicalizeLift(exName)) || classifyLift(exName);
        if (!grp || !_prevSessions[grp]) continue;
        // Require actual weight data: either user-entered weight or a done set
        // (done implies prescribed weight was used)
        const hasWeight = (val.actual && parseFloat(val.actual) > 0) || val.done;
        if (hasWeight) _prevSessions[grp].add(pts[1] + '||' + pts[2]);
      }
      blockDeltas = {};
      for (const gk of ['squat', 'bench', 'deadlift']) {
        const currE1rm = compLifts[gk]?.bestE1rm || 0;
        const prevE1rm = prevData.compLifts[gk]?.bestE1rm || 0;
        blockDeltas[gk] = {
          current: currE1rm, previous: prevE1rm,
          delta: currE1rm && prevE1rm ? currE1rm - prevE1rm : null,
          prevBlockName: displayBlockName(prevData.block.name),
          prevCompletedSessions: _prevSessions[gk].size
        };
      }
    }
  }

  return {
    block, weeks, blockLogs, athlete,
    compLiftName, families, amrapData,
    totalDays, completedDays, compliancePct, totalSets, totalVolume,
    volumeByFamily, volumeByWeekFamily, totalRpe, rpeCount, avgRpe, rpeByWeek,
    compLifts, allFlags, generalFlags,
    estTotal, dotsScore, wilksScore,
    weekLabels, peakWeekLabel,
    volChartData, weeklyAvgRpe,
    dateStr, athleteName,
    statusBadge, blockDeltas
  };
}

// ── MESSAGE HANDLER ───────────────────────────────────────────────────────────
self.onmessage = function(e) {
  const result = _getReportDataWorker(e.data);
  self.postMessage(result);
};
