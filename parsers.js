// ── PowerliftingLog — Parser Module ────────────────────────────────────────────
// Extracted from index.html for standalone testing and modularity.
// This file is the single source of truth for all parser logic.
// In production, include via <script src="parsers.js"></script> before the main script.

// ── LIFT GROUPS & CLASSIFICATION ──────────────────────────────────────────────
const LIFT_GROUPS = {
  squat: ['squat','hack squat','goblet squat','front squat','zercher','belt squat','box squat','high bar squat','low bar squat','pin squat','pause squat','safety bar','ssb','1/4 squat','quarter squat','partial squat'],
  bench: ['bench','close grip bench','close grip press','larson','larsen','slingshot','incline press','incline bench','decline press','decline bench','floor press','feet up bench','pause bench','touch and go bench','touch and go press','tng bench','tng press','comp bench','board press'],
  deadlift: ['deadlift','sumo deadlift','sumo pull','sumo dead','conventional deadlift','conventional pull','romanian','rdl','rack pull','block pull','deficit deadlift','deficit pull','stiff leg','trap bar','hex bar','snatch grip','pause dead','semi sumo'],
};

const COMP_KEYWORDS = {
  squat:    ['squat','back squat','low bar squat','barbell squat','comp squat','competition squat'],
  bench:    ['bench','bench press','barbell bench','barbell bench press','comp bench','comp bench press','barbell comp bench','barbell comp bench press','competition bench','competition bench press','paused bench','paused bench press','pause bench','pause bench press'],
  deadlift: ['deadlift','barbell deadlift','conventional deadlift','sumo deadlift','comp deadlift','competition deadlift','conventional pull','sumo pull','comp pull','competition style deadlift'],
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
];

function isCompLift(name){
  const n = name.toLowerCase().trim();
  if (VARIATION_MODIFIERS.some(v => n.includes(v))) return false;
  for (const keywords of Object.values(COMP_KEYWORDS)) {
    if (keywords.includes(n)) return true;
  }
  return false;
}

function classifyLift(name){
  const n=name.toLowerCase();
  for(const [group,keywords] of Object.entries(LIFT_GROUPS)){
    if(keywords.some(k=>n.includes(k))) return group;
  }
  return null;
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
];

function editDistance(a,b){
  const m=a.length,n=b.length;
  const dp=Array.from({length:m+1},(_,i)=>Array.from({length:n+1},(_,j)=>i===0?j:j===0?i:0));
  for(let i=1;i<=m;i++) for(let j=1;j<=n;j++)
    dp[i][j]=a[i-1]===b[j-1]?dp[i-1][j-1]:1+Math.min(dp[i-1][j],dp[i][j-1],dp[i-1][j-1]);
  return dp[m][n];
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

// ── FORMAT DETECTION ──────────────────────────────────────────────────────────
function detectFormat(wb){
  try{
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
    const rows = rawRows.map(r=>r?r.map(c=>(typeof c==='string'&&c.startsWith("'"))?c.slice(1):c):r);

    // Check Format A: multiple day columns in header rows
    for(let i=0;i<Math.min(10,rows.length);i++){
      const row=rows[i]; if(!row) continue;
      const hits=row.filter(c=>c&&DAYS.some(d=>String(c).toLowerCase().trim().startsWith(d)));
      if(hits.length>=2) return 'A';
    }

    // Check Format D: horizontal week layout — "WEEK N, Day N" section headers + repeating Sets/Reps column groups
    // Must check BEFORE Format C since both have split Sets/Reps headers
    // Check multiple sheets since first sheets may be metadata/maxes/instructions
    for (let si = 0; si < Math.min(6, wb.SheetNames.length); si++) {
      const wsD = wb.Sheets[wb.SheetNames[si]];
      const rowsD = XLSX.utils.sheet_to_json(wsD, {header:1, defval:null}).map(r=>r?r.map(c=>(typeof c==='string'&&c.startsWith("'"))?c.slice(1):c):r);
      let dSetsCount = 0, dHasDaySection = false, dSetsHeaderRow = -1;
      for(let i=0;i<Math.min(15,rowsD.length);i++){
        const row=rowsD[i]; if(!row) continue;
        const strs = row.map(c=>String(c||'').toLowerCase().trim());
        const sc = strs.filter(s => s === 'sets').length;
        if(sc >= 2 && dSetsHeaderRow === -1) {
        // Exclude F-style per-week groups where "Exercise" repeats (Rip & Tear)
        const exRepeat = strs.filter(s => s === 'exercise').length;
        if (exRepeat >= 2) continue;
        dSetsCount = sc; dSetsHeaderRow = i;
      }
        if(row.some(c => c != null && /^WEEK\s*\d/i.test(String(c).trim()))) dHasDaySection = true;
      }
      if(dSetsCount >= 2 && dHasDaySection && dSetsHeaderRow >= 0) {
        // Verify data rows below the header have TEXT exercise names in col A, not numbers/decimals/percentages
        // This prevents false positives on Sheiko (numbered indices) and 531 (percentage rows)
        let hasTextExercise = false;
        for(let i = dSetsHeaderRow + 1; i < Math.min(dSetsHeaderRow + 6, rowsD.length); i++){
          const row = rowsD[i]; if(!row) continue;
          const colA = row[0];
          if(colA == null) continue;
          const colAStr = String(colA).trim();
          if(!colAStr) continue;
          // Skip if col A is a number, decimal, or percentage (0.4, 1, 2, 0.65, etc.)
          if(/^\d+\.?\d*$/.test(colAStr)) continue;
          // Skip if it looks like another WEEK header
          if(/^WEEK\s*\d/i.test(colAStr)) continue;
          // Found a text string — this is a real exercise name
          if(/[a-zA-Z]{2,}/.test(colAStr)){ hasTextExercise = true; break; }
        }
        if(hasTextExercise) return 'D';
      }
    }

    // Check Format F: horizontal week columns (531 Wendler, Madcow, Rip & Tear)
    if (detectF(wb)) return 'F';

    // Check Format G: Sheiko numbered exercises
    if (detectG(wb)) return 'G';

    // Check Format E: flat grid / nSuns-like / week-day text structures
    if (detectE(wb)) return 'E';

    // Check Format C: split columns with "Sets"/"Reps"/"Load" headers + "Day N" labels in col A
    let hasSplitHeaders = false;
    let hasDayLabel = false;
    for(let i=0;i<Math.min(40,rows.length);i++){
      const row=rows[i]; if(!row) continue;
      const strs = row.map(c=>String(c||'').toLowerCase().trim());
      if (strs.includes('sets') && strs.includes('reps') && (strs.includes('load') || strs.includes('weight'))) {
        hasSplitHeaders = true;
      }
      if (row[0] && /^day\s*\d/i.test(String(row[0]).trim())) {
        hasDayLabel = true;
      }
      if (hasSplitHeaders && hasDayLabel) return 'C';
    }
  }catch(e){ console.error('[liftlog] detectFormat error:',e); }
  return 'B';
}

// ── FORMAT E DETECTION ──────────────────────────────────────────────────────
function detectE(wb) {
  const SKIP_SHEET = /^(readme|read\s*me|how\s*to|instruction|info|updates?|stats?|save|accessor|start[-\s]?option|inputs?|maxes?|output)$/i;
  for (let si = 0; si < Math.min(6, wb.SheetNames.length); si++) {
    const sn = wb.SheetNames[si];
    if (SKIP_SHEET.test(sn.trim())) continue;
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null})
      .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
    if (!rows || rows.length < 5) continue;

    // E-nSuns: day-of-week name or "Day N:" as standalone row value + many numeric values in data rows below
    // BUT NOT if the rows below have column headers (Sets/Reps/Load/Intensity/Movement) — that's D/B
    for (let i = 0; i < Math.min(50, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let col = 0; col <= 2; col++) {
        const val = row[col]; if (!val) continue;
        const str = String(val).trim().toLowerCase();
        const isDayName = DAYS.some(d => str === d || str.startsWith(d + ' '));
        const isDayLabel = /^day\s*\d/i.test(String(val).trim());
        if (isDayName || isDayLabel) {
          // Must be standalone (not multi-day header like Format A)
          const otherDays = row.filter((c2, ci) => ci !== col && c2 && DAYS.some(d => String(c2).trim().toLowerCase().startsWith(d)));
          if (otherDays.length > 0) continue;
          // Check the next 1-2 rows for column headers (would indicate D/C/B, not E-nSuns)
          let hasColHeaders = false;
          for (let j = i; j < Math.min(i + 3, rows.length); j++) {
            const hr = rows[j]; if (!hr) continue;
            const hstrs = hr.map(c2 => String(c2 || '').toLowerCase().trim());
            if ((hstrs.includes('sets') || hstrs.includes('set')) && (hstrs.includes('reps') || hstrs.includes('rep goal'))) hasColHeaders = true;
            if (hstrs.includes('movement') || hstrs.includes('exercise movement')) hasColHeaders = true;
          }
          if (hasColHeaders) continue;
          // Verify numeric data below (alternating weight/rep pairs → many numbers)
          for (let j = i + 1; j < Math.min(i + 6, rows.length); j++) {
            const dr = rows[j]; if (!dr) continue;
            let numCount = 0;
            for (let c2 = col + 1; c2 < dr.length; c2++) {
              if (typeof dr[c2] === 'number') numCount++;
            }
            if (numCount >= 6) return true;
          }
        }
      }
    }

    // E-weekText: "Week N" labels + SxR text patterns (NxN), WITHOUT "Sets" column headers
    // Also detect sheets with just SxR text patterns (Smolov — no explicit Week labels)
    let hasWeekLabel = false, hasSxR = false, hasSetsColHeader = false;
    let sxrCount = 0, hasBFormatMarker = false;
    for (let i = 0; i < Math.min(30, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < Math.min(row.length, 25); c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        if (/^Week\s*\d/i.test(str)) hasWeekLabel = true;
        if (/^\d+\s*x\s*\d/i.test(str) || /^\d+x\d/i.test(str)) { hasSxR = true; sxrCount++; }
        if (/^sets$/i.test(str) || /^sets\s*x/i.test(str) || /^sets\s*\//i.test(str)) hasSetsColHeader = true;
        // B-format markers: "Record Weight", "Coach Notes", "Sets x RPE"
        const lo = str.toLowerCase();
        if (lo.includes('record') || lo.includes('coach') || (lo.includes('sets') && lo.includes('rpe'))) hasBFormatMarker = true;
      }
    }
    if (hasWeekLabel && hasSxR && !hasSetsColHeader && !hasBFormatMarker) return true;
    // Smolov-like: lots of SxR text without B-format or column headers
    if (sxrCount >= 3 && !hasSetsColHeader && !hasBFormatMarker) return true;

    // E-tabular: column headers including Week + lift/exercise + (Sets or Reps)
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(c => String(c || '').toLowerCase().trim());
      const hasWeek = strs.includes('week');
      const hasLift = strs.some(s => s.includes('primary lift') || (s === 'exercise'));
      const hasSR = strs.includes('sets') || strs.includes('reps');
      if (hasWeek && hasLift && hasSR) return true;
    }

    // E-cycleWeek: "Cycle N" or "Week N NxN" headers + exercise names
    // Also covers Juggernaut: phase headers with "Weight"/"Reps" repeating column groups
    let hasCycle = false, hasExNames = false, hasWeekPres = false, hasPhaseWeek = false;
    let hasWeightRepsGroups = false;
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      for (let c = 0; c < Math.min(row.length, 15); c++) {
        const val = row[c]; if (val == null) continue;
        const str = String(val).trim();
        if (/^Cycle\s*\d/i.test(str)) hasCycle = true;
        if (/^Week\s*\d.*\d+x\d/i.test(str) || /^Week\s*\d.*\d+\/\d+\/\d/i.test(str)) hasWeekPres = true;
        if (/Phase\s*[-–]\s*Week\s*\d/i.test(str) || /^\d+\s*Rep\s*Wave/i.test(str)) hasPhaseWeek = true;
        if (/^(Bench|Squat|Deadlift|Press|Military|Overhead|OHP|Core\s*Lift|OH\s*Press)/i.test(str)) hasExNames = true;
      }
      // Check for repeating Weight/Reps column groups (Juggernaut pattern)
      const strs = row.map(c => String(c || '').toLowerCase().trim());
      const wCols = strs.filter(s => s === 'weight').length;
      const rCols = strs.filter(s => s === 'reps').length;
      if (wCols >= 2 && rCols >= 2) hasWeightRepsGroups = true;
    }
    if ((hasCycle || hasWeekPres) && hasExNames) return true;
    if (hasPhaseWeek && hasExNames) return true;
    if (hasWeightRepsGroups && hasExNames) return true;

    // E-parallel531: Week N as row separators + "Reps" header + multiple exercise column groups
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(c => String(c || '').toLowerCase().trim());
      if (/^week\s*\d/i.test(strs[0] || '')) {
        const repsCols = strs.filter(s => s === 'reps').length;
        if (repsCols >= 2) return true;
      }
    }
  }
  return false;
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
      const c=String(r[j]||'').toLowerCase().trim();
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
          const correctedName=spellCorrectExerciseName(normalizedCs);
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

  return{id:sn,name,dateRange,athleteName,maxes,weeks,format:'A'};
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
  const bId = wb._plFilename ? wb._plFilename.replace(/\.xlsx?$/i,'').trim() : wb.SheetNames[0];
  const blocks=[{id:bId,name:bId||'Training Program',dateRange:'',athleteName,maxes:{},weeks,format:'B'}];
  return blocks;
}

function parseBSheet(sn,rows){
  if(!rows||rows.length===0) return null;
  const days=[]; let curDay=null;
  let colPres=1, colLogged=2;
  const seenDayNames={};
  let weekOffset=0;

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

      const isCircuitRow = /^circuit\s/i.test(normName) || (normName.match(/,/g)||[]).length >= 2;
      if(isCircuitRow){
        const stripped = normName.replace(/^circuit\s+/i,'').trim();
        const parts = stripped.split(/,\s*/);
        if(parts.length >= 2){
          const _circuitIdx = curDay.exercises.length;
          const supersetGroup = { label:`circuit_${_circuitIdx}`, rounds:1, startIdx:_circuitIdx };
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

      curDay.exercises.push({name:normName,prescription:pres,note,lifterNote,loggedWeight:cleanLogged});
    }
  }
  // Fallback: Candito-style — dates as day separators, "Set N" column headers, weight/"xN" pairs
  if(days.length===0){
    let curCDay = null;
    for(let i=0;i<rows.length;i++){
      const row=rows[i]; if(!row) continue;
      const a=row[0];
      // Excel serial date (> 40000) as day separator
      if(typeof a === 'number' && a > 40000 && a < 50000){
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

// ── TEMPLATE MAX EXTRACTION ──────────────────────────────────────────────────
function extractTemplateMaxes(wb){
  const XLSX = typeof require !== 'undefined' ? require('xlsx') : window.XLSX;
  const maxes = { squat: null, bench: null, deadlift: null };

  function isNum(v){ return typeof v === 'number' && isFinite(v) && v > 20; }

  const ABBREVS = { dl:'deadlift', sq:'squat', bp:'bench' };
  function liftCat(name){
    const n = name.toLowerCase().trim().replace(/[:\s]+$/,'');
    if (ABBREVS[n]) return ABBREVS[n];
    return classifyLift(name);
  }

  function assign(lift, val){
    const cat = liftCat(lift);
    if (cat && (cat === 'squat' || cat === 'bench' || cat === 'deadlift')) {
      if (maxes[cat] === null) maxes[cat] = val;
    }
  }

  function getRows(sn, limit){
    const ws = wb.Sheets[sn];
    if (!ws) return [];
    const json = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    return json.slice(0, limit || json.length);
  }

  function done(){ return maxes.squat !== null && maxes.bench !== null && maxes.deadlift !== null; }
  function found(){ return [maxes.squat, maxes.bench, maxes.deadlift].filter(v => v !== null).length >= 2; }

  // Strategy A: Dedicated max/input sheet
  const maxSheetNames = ['max','maxes','input','inputs','start here','1. start here'];
  for (const sn of wb.SheetNames) {
    if (!maxSheetNames.includes(sn.toLowerCase().trim())) continue;
    const rows = getRows(sn, 25);
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (typeof cell !== 'string') continue;
        const cat = liftCat(cell);
        if (!cat || (cat !== 'squat' && cat !== 'bench' && cat !== 'deadlift')) continue;
        // Look right for a number
        for (let cc = c + 1; cc < Math.min(c + 4, row.length); cc++) {
          if (isNum(row[cc])) { assign(cell, row[cc]); break; }
        }
        // Look left for a number (less common)
        if (maxes[cat] === null) {
          for (let cc = c - 1; cc >= Math.max(c - 3, 0); cc--) {
            if (isNum(row[cc])) { assign(cell, row[cc]); break; }
          }
        }
      }
    }
    if (found()) return done() ? maxes : (maxes.squat !== null || maxes.bench !== null || maxes.deadlift !== null) ? maxes : null;
  }
  if (found()) return maxes;

  // Strategy B: Row with "1RM" or "Max" label + lift names in adjacent rows
  for (const sn of wb.SheetNames) {
    const rows = getRows(sn, 15);
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      // B1: "1RM" in col 0 — header row above or lift/value rows below
      const firstStr = String(row[0] || '').toLowerCase().trim();
      if (firstStr === '1rm') {
        if (r > 0 && rows[r - 1]) {
          const hdr = rows[r - 1];
          for (let c = 1; c < Math.min(row.length, hdr.length); c++) {
            if (typeof hdr[c] === 'string' && isNum(row[c])) assign(hdr[c], row[c]);
          }
        }
        for (let rr = r + 1; rr < Math.min(r + 6, rows.length); rr++) {
          const below = rows[rr];
          if (!below) continue;
          if (typeof below[0] === 'string' && isNum(below[1])) assign(below[0], below[1]);
        }
        if (found()) return maxes;
      }
      // B2: "1RM" in non-zero columns — lift names in rows below at col offset
      for (let c = 1; c < row.length; c++) {
        const cell = String(row[c] || '').toLowerCase().trim();
        if (cell !== '1rm') continue;
        for (let rr = r + 1; rr < Math.min(r + 5, rows.length); rr++) {
          const below = rows[rr];
          if (!below) continue;
          // Lift name a few cols before the "1RM" column, value at the "1RM" column
          for (let cc = Math.max(0, c - 4); cc < c; cc++) {
            if (typeof below[cc] === 'string' && liftCat(below[cc]) && isNum(below[c])) {
              assign(below[cc], below[c]); break;
            }
          }
        }
        if (done()) return maxes;
      }
      // B3: "Max" column header — lift name nearby in same row below, value at "Max" col
      for (let c = 1; c < row.length; c++) {
        const cell = String(row[c] || '').toLowerCase().trim();
        if (cell !== 'max') continue;
        for (let rr = r + 1; rr < Math.min(r + 4, rows.length); rr++) {
          const below = rows[rr];
          if (!below) continue;
          // Scan backwards from "Max" col to find the nearest lift name
          for (let cc = c - 1; cc >= Math.max(0, c - 5); cc--) {
            if (typeof below[cc] === 'string' && liftCat(below[cc]) && isNum(below[c])) {
              assign(below[cc], below[c]); break;
            }
          }
        }
        if (done()) return maxes;
      }
    }
  }
  if (found()) return maxes;

  // Strategy C: nSuns inline row with "1rm" or "training max" text
  for (const sn of wb.SheetNames) {
    const rows = getRows(sn, 20);
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      const hasRM = row.some(c => typeof c === 'string' && (/1rm/i.test(c) || /training\s*max/i.test(c)));
      if (!hasRM) continue;
      // Scan for lift name followed by a number
      for (let c = 0; c < row.length - 1; c++) {
        if (typeof row[c] !== 'string') continue;
        const cat = liftCat(row[c]);
        if (!cat || (cat !== 'squat' && cat !== 'bench' && cat !== 'deadlift')) continue;
        for (let cc = c + 1; cc < Math.min(c + 3, row.length); cc++) {
          if (isNum(row[cc])) { assign(row[c], row[cc]); break; }
        }
      }
      if (found()) return maxes;
    }
  }

  return found() ? maxes : null;
}

// ── UNIFIED ENTRY POINT ───────────────────────────────────────────────────────
function parseWorkbook(wb){
  const fmt = detectFormat(wb);
  let blocks;
  if (fmt === 'A') blocks = parseA(wb);
  else if (fmt === 'E') blocks = parseE(wb);
  else if (fmt === 'F') blocks = parseF(wb);
  else if (fmt === 'G') blocks = parseG(wb);
  else if (fmt === 'C') blocks = parseCAutoFormat(wb);
  else if (fmt === 'D') blocks = parseD(wb);
  else blocks = parseB(wb);

  const maxes = extractTemplateMaxes(wb);
  if (maxes && blocks) {
    for (const block of blocks) {
      block.templateMaxes = maxes;
    }
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
          const hstrs = hr.map(c2 => String(c2 || '').toLowerCase().trim());
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
      const strs = row.map(c => String(c || '').toLowerCase().trim());
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
      const strs = row.map(c => String(c || '').toLowerCase().trim());
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
          const hstrs = hr.map(c => String(c || '').toLowerCase().trim());
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
    // Horizontal day layout (Russian Squat style): exercise, SxR, weight in 3-col groups
    const dataStart = dayColumns[0].headerRow + 1;
    const days = [];
    for (const dc of dayColumns) {
      const exercises = [];
      for (let r = dataStart; r < endRow; r++) {
        const row = rows[r]; if (!row) continue;
        // Stop at next Week header
        for (let cc = 0; cc < Math.min(row.length, 3); cc++) {
          if (row[cc] != null && /^Week\s*\d/i.test(String(row[cc]).trim())) return days.length > 0 ? days : [];
        }
        const exName = row[dc.col];
        const sxr = row[dc.col + 1];
        const weight = row[dc.col + 2];
        if (exName == null && sxr == null) continue;
        let name = exName ? String(exName).trim() : '';
        if (!name || /^\d+\.?\d*$/.test(name)) continue;
        if (/^(Week|Day|Warm|REST|TEST|MAX|LB|KG|1\s*RM)/i.test(name)) continue;
        let pres = sxr ? String(sxr).trim() : '';
        if (weight != null && typeof weight === 'number' && weight > 0) {
          pres = pres ? `${pres}(${Math.round(weight * 100) / 100})` : `(${Math.round(weight * 100) / 100})`;
        } else if (weight != null) {
          const ws = String(weight).trim();
          if (ws && !/^\d+\.?\d*$/.test(ws)) pres = pres ? `${pres} ${ws}` : ws;
        }
        exercises.push({ name: name.replace(/\s+/g, ' '), prescription: pres || null, note: '', lifterNote: '', loggedWeight: '' });
      }
      if (exercises.length > 0) days.push({ name: dc.name, exercises });
    }
    return days;
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
      const strs = row.map(c => String(c || '').toLowerCase().trim());
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

// ── FORMAT C PARSER (split Sets/Reps/Load columns with Day N headers) ────────
function parseCAutoFormat(wb) {
  const blocks = [];
  for (const sn of wb.SheetNames) {
    try {
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, defval:null});
      const rows = rawRows.map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
      const week = parseCAutoSheet(sn, rows);
      if (week) blocks.push(week);
    } catch(e) { console.error(`[parseCAutoFormat] Error parsing sheet "${sn}":`, e); }
  }
  if (blocks.length === 0) return [];

  const firstRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, defval:null})
    .map(r => r ? r.map(c => (typeof c === 'string' && c.startsWith("'")) ? c.slice(1) : c) : r);
  const meta = extractCAutoMeta(firstRows);

  const blockName = meta.blockName || wb._plFilename || 'Program';
  const id = 'c_' + blockName.replace(/\s+/g,'_').substring(0,30) + '_' + Date.now();
  return [{
    id, name: blockName, weeks: blocks, athlete: meta.athlete || '',
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
    const strs = row.map(c => String(c || '').toLowerCase().trim());
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
  const SUB_SECTIONS = ['warm up','main exercise','accessories','finisher','abs','recovery','cool down','cardio'];

  for (let i = dataStartRow; i < rows.length; i++) {
    const row = rows[i]; if (!row) continue;
    const colA = String(row[0] || '').trim();
    const exVal = String(row[exIdx] || '').trim();
    const strs = row.map(c => String(c || '').toLowerCase().trim());

    if (/^day\s*\d/i.test(colA)) {
      let dayName = colA;
      if (i + 1 < rows.length && rows[i+1]) {
        const nextA = String(rows[i+1][0] || '').trim();
        if (/^(mon|tue|wed|thu|fri|sat|sun)/i.test(nextA)) {
          dayName = nextA + ' (' + colA + ')';
        }
      }
      curDay = { name: dayName, exercises: [] };
      days.push(curDay);
      continue;
    }

    if (/^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$/i.test(colA)) continue;
    if (strs.includes('sets') && strs.includes('reps')) continue;
    if (!curDay) continue;
    if (!exVal) continue;
    if (SUB_SECTIONS.some(s => exVal.toLowerCase() === s)) continue;

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
      loggedWeight: logStr
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
          const strs = sr.map(c => String(c || '').toLowerCase().trim());
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
      const strs = row.map(c => String(c || '').toLowerCase().trim());
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

    // Find sub-header row with Weight/Reps repeating
    let subHeaderRow = -1;
    let weekColGroups = [];
    for (let i = weekHeaderRow + 1; i < Math.min(weekHeaderRow + 3, rows.length); i++) {
      const row = rows[i]; if (!row) continue;
      const strs = row.map(c => String(c || '').toLowerCase().trim());
      const wCols = [], rCols = [];
      for (let c = 0; c < strs.length; c++) {
        if (strs[c] === 'weight') wCols.push(c);
        if (strs[c] === 'reps') rCols.push(c);
      }
      if (wCols.length >= 2 && rCols.length >= 2) {
        subHeaderRow = i;
        for (let g = 0; g < wCols.length; g++) {
          weekColGroups.push({weightCol: wCols[g], repsCol: rCols[g] !== undefined ? rCols[g] : wCols[g] + 1});
        }
        break;
      }
    }
    if (subHeaderRow === -1 || weekColGroups.length === 0) continue;

    if (!blockName) blockName = sn;

    // Find day sections
    const dayStarts = [];
    for (let i = subHeaderRow + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const colA = String(row[0] || '').trim().toLowerCase();
      if (DAYS.some(d => colA === d || colA.startsWith(d + ' '))) {
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
          if (/[a-zA-Z]{2,}/.test(s) && !/^(Week|Exercise)/i.test(s)) { exName = s; break; }
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
          for (const ri of eg.rows) {
            const row = rows[ri];
            const weight = row[g.weightCol];
            const reps = row[g.repsCol];
            if (weight != null || reps != null) {
              sets.push({weight: typeof weight === 'number' ? weight : null, reps: reps != null ? String(reps) : ''});
            }
          }
          weekSets.push(sets);
        }
        dayExercises.push({name: eg.name, weekSets});
      }
      dayData.push({name: dayStarts[d].name, exercises: dayExercises});
    }

    // Build weeks
    for (let w = 0; w < weekColGroups.length; w++) {
      globalWeekNum++;
      const weekDays = [];
      for (const dd of dayData) {
        const exercises = [];
        for (const ex of dd.exercises) {
          if (w >= ex.weekSets.length) continue;
          const sets = ex.weekSets[w];
          if (sets.length === 0) continue;
          const presParts = [];
          for (const s of sets) {
            if (s.weight != null && s.weight > 0 && s.reps) {
              presParts.push('1x' + s.reps + '(' + (Math.round(s.weight * 10) / 10) + ')');
            } else if (s.reps) {
              presParts.push('1x' + s.reps);
            }
          }
          if (presParts.length === 0) continue;
          exercises.push({name: ex.name, prescription: presParts.join(', '), note: '', lifterNote: '', loggedWeight: ''});
        }
        if (exercises.length > 0) weekDays.push({name: dd.name, exercises});
      }
      if (weekDays.length > 0) {
        allWeeks.push({label: weekPositions[w] ? weekPositions[w].label : 'Week ' + globalWeekNum, days: weekDays});
      }
    }
  }

  if (allWeeks.length === 0) return null;
  return [{name: blockName, weeks: allWeeks}];
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
    const strs = subRow.map(c => String(c || '').toLowerCase().trim());
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
          dayExercises.push({name: exName, prescription: sets.join(', '), note: '', lifterNote: '', loggedWeight: ''});
        }
      }

      if (dayExercises.length > 0) {
        allWeeks.push({label: weekPositions[w].label, days: [{name: 'Day 1', exercises: dayExercises}]});
      }
    }

    if (allWeeks.length > 0) return [{name: sn, weeks: allWeeks}];
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
    let exOff = -1, setsOff = -1, repsOff = -1, weightOff = -1;
    for (let c = base; c < base + stride && c < sr.length; c++) {
      const s = String(sr[c] || '').toLowerCase().trim();
      if (s === 'exercise') exOff = c - base;
      else if (s === 'sets') setsOff = c - base;
      else if (s === 'reps') repsOff = c - base;
      else if (s === 'weight') weightOff = c - base;
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

          let pres = '';
          if (sv > 0 && rr && wt) pres = sv + 'x' + rr + '(' + wt + ')';
          else if (sv > 0 && rr) pres = sv + 'x' + rr;

          const wk = pos.week, dy = pos.day;
          if (!weekMap[wk]) weekMap[wk] = {};
          if (!weekMap[wk][dy]) weekMap[wk][dy] = [];
          weekMap[wk][dy].push({name: exName, prescription: pres, note: '', lifterNote: '', loggedWeight: ''});
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

    if (allWeeks.length > 0) return [{name: sn || 'Training Program', weeks: allWeeks}];
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
      const strs = nextRow.map(c => String(c || '').toLowerCase().trim());
      const loadCols = [];
      let movCol = -1, setCol = -1, repsCol = -1;
      for (let c = 0; c < strs.length; c++) {
        if (strs[c] === 'movement' || strs[c] === 'exercise') movCol = c;
        if (strs[c] === 'set' || strs[c] === 'sets') setCol = c;
        if (strs[c] === 'reps') repsCol = c;
        if (strs[c] === 'load') loadCols.push(c);
      }
      if (loadCols.length < 2 || movCol === -1) continue;

      daySections.push({row: i, name: String(row[0]).trim(), weekPositions, movCol, setCol, repsCol, loadCols, subHeaderRow: i + 1});
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

        const weekData = [];
        for (let w = 0; w < sec.loadCols.length; w++) {
          const load = row[sec.loadCols[w]];
          const wt = typeof load === 'number' && load > 0 ? Math.round(load) : null;
          let pres = '';
          if (sets > 0 && reps && wt) pres = sets + 'x' + reps + '(' + wt + ')';
          else if (sets > 0 && reps) pres = sets + 'x' + reps;
          weekData.push(pres);
        }
        exercises.push({name: exName, weekData});
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
          dayExercises.push({name: ex.name, prescription: ex.weekData[w], note: '', lifterNote: '', loggedWeight: ''});
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
  return [{name: blockName, weeks: allWeeks}];
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
      const strs = row.map(c => String(c || '').toLowerCase().trim());
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
          exercises.push({name: ex.name, prescription: sets.join(', '), note: '', lifterNote: '', loggedWeight: ''});
        }
        if (exercises.length > 0) weekDays.push({name: dd.name, exercises});
      }
      if (weekDays.length > 0) allWeeks.push({label: weekColumns[w].label, days: weekDays});
    }

    if (allWeeks.length > 0) return [{name: sn, weeks: allWeeks}];
  }
  return null;
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
  return [{name: blockName, weeks: allWeeks}];
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
  return {name: ex.name, prescription: parts.join(', '), note: '', lifterNote: '', loggedWeight: ''};
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
            loggedWeight: ''
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
    id: (blockMeta.name||'training_block').replace(/[^a-zA-Z0-9]/g,'_') + '_D',
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
    const strs = row.map(c => String(c || '').toLowerCase().trim());
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
      const strs = row.map(c => String(c || '').toLowerCase().trim());
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
  const restVal = rest != null ? rest : null;

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
function parseSets(pres){
  if(!pres) return [];
  const sets=[]; const pat=/(\d+x)?(\d+)\(([^)]+)\)/g; let m;
  while((m=pat.exec(pres))!==null){
    const mult=m[1]?parseInt(m[1]):1;
    for(let i=0;i<mult;i++) sets.push({reps:parseInt(m[2]),weight:m[3].trim(),isRpe:isNaN(Number(m[3].trim()))});
  }
  return sets;
}

function parseSimple(pres){
  if(!pres) return null;
  const m=pres.match(/^(\d+)\s*x\s*([\d\-]+(?:\s*(?:sec|min|s|m)\w*)?)/i);
  if(!m) return null;
  if(pres.match(/^\d+[\-–]\d+x/i)) return null;
  const sets = parseInt(m[1]);
  if(sets > 20) return null;
  return{sets,reps:m[2].trim()};
}

function parseRangeSet(pres){
  if(!pres) return null;
  const m=pres.match(/^(\d+)[\-–]?(\d*)\s*x\s*([\d\-]+(?:\s*(?:sec|min|s|m)\w*)?)/i);
  if(m){
    const sets=parseInt(m[1]);
    const reps=m[3].trim();
    return{sets,reps};
  }
  const s=pres.match(/^(\d+)[\-–](\d+)\s*sets?/i);
  if(s) return{sets:parseInt(s[1]),reps:'open'};
  return null;
}

function parseBarePres(pres){
  if(!pres) return null;
  const p=pres.trim();
  if(/^\d+$/.test(p)) return{sets:1,reps:p};
  if(/^\d+:\d{2}$/.test(p)) return{sets:1,reps:p};
  return null;
}

function parseWarmupSets(pres){
  if(!pres) return [];
  const p=pres.trim();
  if(!p.includes(',')) return [];
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
  return sets.length>=2 ? sets : [];
}

// ── MODULE EXPORTS (Node.js) / GLOBAL (browser) ──────────────────────────────
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    LIFT_GROUPS, COMP_KEYWORDS, VARIATION_MODIFIERS,
    isCompLift, classifyLift,
    EXERCISE_DICT, editDistance, spellCorrectWord, spellCorrectExerciseName,
    DAYS, detectFormat, detectE, detectF, detectG,
    parseA, parseASheet, parseB, parseBSheet,
    parseWorkbook, extractTemplateMaxes,
    parseE, parseE_nSuns, parseE_tabular, parseE_phaseGrid, parseE_weekText, parseE_cycleWeek, parseE_parallel531,
    parseF, parseF_stride3, parseF_stride15, parseF_stride6, parseF_pctLoad, parseF_stride1,
    parseG, _gBuildExercise,
    parseCAutoFormat, extractCAutoMeta, parseCAutoSheet,
    parseD, parseDSheet, buildDPrescription,
    deduplicateExerciseNames,
    parseSets, parseSimple, parseRangeSet, parseBarePres, parseWarmupSets,
  };
}
