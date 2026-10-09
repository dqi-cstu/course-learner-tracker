/**
 * Learning Progress Tracker v2.24
 * Request-scoped data access helpers.
 *
 * Apps Script normally starts a fresh execution for each server request.
 * These caches therefore avoid duplicate reads only inside one request;
 * they are not persistent cross-request caches and do not create
 * cache-invalidation problems between users.
 */

var REQUEST_ROW_CACHE_ = {};

function resetRequestDataCache_() {
  REQUEST_ROW_CACHE_ = {};
  try{clearSetupStateCache_();}catch(e){}
  try{clearCourseHomeCache_();}catch(e){}
}

function requestSheetCacheKey_(sh) {
  return String(sh.getSheetId());
}

function readRowsCached_(sh) {
  const key = requestSheetCacheKey_(sh);
  if (!REQUEST_ROW_CACHE_[key]) {
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 2 || lastCol < 1) {
      REQUEST_ROW_CACHE_[key] = [];
    } else {
      const vals = sh.getRange(1,1,lastRow,lastCol).getValues();
      const headers = vals[0].map(h => String(h || '').trim());
      REQUEST_ROW_CACHE_[key] = vals.slice(1).map((row,i) => {
        const o = {_row:i+2};
        headers.forEach((h,j) => {
          if (h) o[h] = row[j];
        });
        return o;
      });
    }
  }

  // Return shallow copies so normalization helpers cannot mutate cached rows.
  return REQUEST_ROW_CACHE_[key].map(r => Object.assign({},r));
}

function courseSnapshot_(requested) {
  const wanted = requested || {};
  const out = {};

  function use_(key, tableName) {
    if (wanted[key] === false) return;
    out[key] = rowsToObjects_(table_(tableName));
  }

  if (wanted.students) use_('students', APP.tables.STUDENTS);
  if (wanted.clos) use_('clos', APP.tables.CLOS);
  if (wanted.activities) use_('activities', APP.tables.ACTIVITIES);
  if (wanted.activityClos) use_('activityClos', APP.tables.ACTIVITY_CLOS);
  if (wanted.results) use_('results', APP.tables.RESULTS);
  if (wanted.evidence) use_('evidence', APP.tables.CLO_EVIDENCE);
  if (wanted.attendanceSessions) use_('attendanceSessions', APP.tables.ATTENDANCE_SESSIONS);
  if (wanted.attendance) use_('attendance', APP.tables.ATTENDANCE);
  if (wanted.teamHistory) use_('teamHistory', APP.tables.TEAM_RESULT_MEMBERS);
  if (wanted.policies) use_('policies', APP.tables.CLO_ATTAINMENT_POLICIES);
  if (wanted.weights) use_('weights', APP.tables.CLO_ATTAINMENT_WEIGHTS);
  if (wanted.selfEvaluations) use_('selfEvaluations', APP.tables.CLO_SELF_EVALUATION);
  if (wanted.staff) use_('staff', APP.tables.STAFF);

  return out;
}

function normalizedStudentsFromRows_(rows) {
  return (rows || []).map(r => Object.assign({},r,{Active:bool_(r.Active)}));
}

function normalizedClosFromRows_(rows) {
  return (rows || []).map(r => Object.assign({},r,{Active:bool_(r.Active)}));
}

function normalizedActivitiesFromRows_(rows) {
  return (rows || []).map(r => Object.assign(stripRow_(r),{
    DetailUrl:(typeof safeActivityDetailUrl_==='function'?safeActivityDetailUrl_(r.DetailUrl):String(r.DetailUrl||'')),
    Mode:(typeof normalizeActivityMode_==='function'?normalizeActivityMode_(r.Mode):(String(r.Mode||'').toLowerCase()==='team'?'Team':'Individual')),
    Published:normalizeActivityVisibilityLevel_(r.StudentVisibilityLevel,r.Published)!=='HIDDEN',
    StudentVisibilityLevel:normalizeActivityVisibilityLevel_(r.StudentVisibilityLevel,r.Published),
    AssignedWeek:r.AssignedWeek===''||r.AssignedWeek==null?'':Number(r.AssignedWeek),
    Active:(r.Active===''||r.Active==null?true:bool_(r.Active)),
    CLOEvalMode:String(r.CLOEvalMode||'MANUAL').toUpperCase(),
    CLOAllowOverride:bool_(r.CLOAllowOverride)
  }));
}

function buildCourseIndexes_(snapshot) {
  const idx = {
    activityById:{},
    cloIdsByActivity:{},
    resultsByActivityTarget:{},
    evidenceByActivityCloTarget:{},
    teamByStudentActivity:{},
    attendanceBySessionStudent:{}
  };

  (snapshot.activities || []).forEach(a => {
    idx.activityById[String(a.ActivityID || '')] = a;
  });

  (snapshot.activityClos || []).forEach(m => {
    const aid=String(m.ActivityID || '');
    const cid=String(m.CLOID || '');
    if(!aid || !cid) return;
    if(!idx.cloIdsByActivity[aid]) idx.cloIdsByActivity[aid]=[];
    idx.cloIdsByActivity[aid].push(cid);
  });

  (snapshot.results || []).forEach(r => {
    const key=[
      String(r.ActivityID || ''),
      String(r.TargetType || ''),
      String(r.TargetID || '')
    ].join('|');
    idx.resultsByActivityTarget[key]=r;
  });

  (snapshot.evidence || []).forEach(e => {
    const key=[
      String(e.ActivityID || ''),
      String(e.CLOID || ''),
      String(e.TargetType || ''),
      String(e.TargetID || '')
    ].join('|');
    idx.evidenceByActivityCloTarget[key]=e;
  });

  (snapshot.teamHistory || []).forEach(r => {
    const sk=String(r.StudentKey || '');
    const aid=String(r.ActivityID || '');
    if(!sk || !aid) return;
    const key=sk+'|'+aid;
    if(idx.teamByStudentActivity[key]===undefined){
      idx.teamByStudentActivity[key]=String(r.TeamCode || '');
    }
  });

  (snapshot.attendance || []).forEach(r => {
    const key=String(r.SessionID || '')+'|'+String(r.StudentKey || '');
    idx.attendanceBySessionStudent[key]=r;
  });

  return idx;
}
