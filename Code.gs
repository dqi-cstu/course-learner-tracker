/**
 * Learning Progress Tracker
 * Google Apps Script web app
 *
 * One Apps Script copy = one course offering.
 * New deployments initialize safely from the Web app on first open.
 * setupProject() remains available as an editor fallback and is guarded against reruns.
 */

const APP = {
  name: 'Learning Progress Tracker',
  version: '2.59.3',
  schemaVersion: 15,
  propBoundSheetId: 'LPT_BOUND_SHEET_ID',
  propProjectInitialized: 'LPT_PROJECT_INITIALIZED',
  tables: {
    CONFIG: '_Config',
    STAFF: '_Staff',
    STUDENTS: '_Students',
    CLOS: '_CLOs',
    ACTIVITIES: '_Activities',
    ACTIVITY_CLOS: '_ActivityCLOs',
    ACTIVITY_CLO_SCORE_MAPS: '_ActivityCLOScoreMaps',
    ACTIVITY_TEAMS: '_ActivityTeams',
    ACTIVITY_TEAM_MEMBERS: '_ActivityTeamMembers',
    COURSE_OBJECTIVES: '_CourseObjectives',
    COURSE_SECTIONS: '_CourseSections',
    COURSE_SECTION_SCHEDULES: '_CourseSectionSchedules',
    STUDENT_SECTIONS: '_StudentSections',
    WEEKLY_PLAN: '_WeeklyTeachingPlan',
    WEEKLY_PLAN_CLOS: '_WeeklyTeachingPlanCLOs',
    WEEKLY_PLAN_HISTORY: '_TeachingPlanHistory',
    RESULTS: '_Results',
    CLO_EVIDENCE: '_CLOEvidence',
    ATTENDANCE_SESSIONS: '_AttendanceSessions',
    ATTENDANCE: '_Attendance',
    TEAM_RESULT_MEMBERS: '_TeamResultMembers',
    CLO_ATTAINMENT_POLICIES: '_CLOAttainmentPolicies',
    CLO_ATTAINMENT_WEIGHTS: '_CLOAttainmentWeights',
    CLO_SELF_EVALUATION: '_CLOSelfEvaluation',
    AUDIT: '_AuditLog'
  },
  headers: {
    _Config: ['Key','Value'],
    _Staff: ['StaffKey','Email','DisplayName','Role','Active','AddedAt','StaffTokenHash','TokenIssuedAt','UpdatedAt'],
    _Students: ['StudentKey','StudentId','Name','Email','Alias','TeamCode','Active','AddedAt','StudentTokenHash','TokenIssuedAt'],
    _CLOs: ['CLOID','Code','Description','StudentFriendlyDescription','Sequence','Active'],
    _Activities: ['ActivityID','Title','ShortDescription','DetailUrl','Category','Mode','Purpose','MaxScore','Weight','Published','StudentVisibilityLevel','AssignedWeek','DueDate','Sequence','Active','CreatedAt','UpdatedAt','CLOEvalMode','CLOScoreSource','CLOMaxScore','CLOLevel2Pct','CLOLevel3Pct','CLOLevel4Pct','CLOLevel5Pct','CLOAllowOverride'],
    _ActivityCLOs: ['ActivityID','CLOID'],
    _ActivityCLOScoreMaps: ['ActivityID','CLOID','ScoreSource','MaxScore','UpdatedAt'],
    _ActivityTeams: ['ActivityID','TeamID','TeamName','Sequence','CreatedAt','UpdatedAt'],
    _ActivityTeamMembers: ['ActivityID','TeamID','StudentKey','AddedAt'],
    _CourseObjectives: ['ObjectiveID','Sequence','ObjectiveText','Active','CreatedAt','UpdatedAt'],
    _CourseSections: ['SectionID','SectionCode','Days','StartTime','EndTime','DayTime','Room','Mode','InstructorStaffKey','InstructorName','InstructorEmail','Active','CreatedAt','UpdatedAt'],
    _CourseSectionSchedules: ['ScheduleID','SectionID','ScheduleType','DayOfWeek','ScheduleDate','StartTime','EndTime','Room','ExamType','Sequence','Active','CreatedAt','UpdatedAt'],
    _StudentSections: ['StudentKey','SectionID','Active','CreatedAt','UpdatedAt'],
    _WeeklyTeachingPlan: ['PlanItemID','WeekNo','PlannedDate','Topic','Description','PlannedCLOIDs','PlannedActivityIDs','ActualWeekNo','ActualDate','ActualStatus','ActualNote','TrackChanges','Active','CreatedAt','UpdatedAt'],
    _WeeklyTeachingPlanCLOs: ['PlanCLOID','PlanItemID','CLOID','TeachingMethod','AssessmentMethod','ActivityIDs','Sequence','CreatedAt','UpdatedAt'],
    _TeachingPlanHistory: ['HistoryID','PlanItemID','ChangedAt','ChangedBy','PreviousActualWeekNo','NewActualWeekNo','PreviousStatus','NewStatus','Reason','RecordedReason'],
    _Results: ['ActivityID','TargetType','TargetID','Status','Score','Feedback','UpdatedAt','UpdatedBy'],
    _CLOEvidence: ['ActivityID','TargetType','TargetID','CLOID','Level','Source','Score','MaxScore','UpdatedAt','UpdatedBy'],
    _AttendanceSessions: ['SessionID','SectionID','Title','StartTime','EndTime','Status','Secret','RotateSeconds','LocationRequired','RoomLat','RoomLng','RadiusM','StoreRawLocation','CreatedBy','Published'],
    _Attendance: ['SessionID','StudentKey','CheckedInAt','Status','DistanceM','AccuracyM','LocationVerified','Latitude','Longitude','Source','Remark','EvidenceFileId','EvidenceFileName','EvidenceUrl','UpdatedBy','UpdatedAt'],
    _TeamResultMembers: ['ActivityID','TeamCode','StudentKey','CapturedAt'],
    _CLOAttainmentPolicies: ['CLOID','Policy','Threshold','MinCount','Active','UpdatedBy','UpdatedAt'],
    _CLOAttainmentWeights: ['CLOID','ActivityID','Weight','UpdatedBy','UpdatedAt'],
    _CLOSelfEvaluation: ['StudentKey','EvaluationPoint','CLOID','Rating','Reflection','SubmittedAt','Source','ImportedAt','ImportedBy'],
    _AuditLog: ['Timestamp','Actor','Action','EntityType','EntityID','Details']
  },
  categories: ['Assignment','Quiz','Exam','Lab','Project Milestone','Studio','Presentation','Technical Check','Other'],
  purposes: ['Formative','Assessed'],
  modes: ['Individual','Team'],
  statuses: ['Not Submitted','Submitted','Late','Reviewed','Excused','Missing'],
  cloMin: 1,
  cloMax: 5,
  cloAttainedLevel: 2
};

let REQUEST_STAFF_AUTH_ = null;
let REQUEST_SCHEMA_READY_ = false;

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.initialParams = JSON.stringify((e && e.parameter) || {});
  return t.evaluate()
    .setTitle(APP.name)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function legacyProjectLooksInitialized_(ss) {
  if (!ss) return false;
  const sh = ss.getSheetByName(APP.tables.CONFIG);
  if (!sh || sh.getLastRow() < 2) return false;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(2, Math.max(2, sh.getLastColumn()))).getValues();
  const keys = new Set(values.map(r => String(r[0] || '').trim()).filter(Boolean));
  return keys.has('APP_VERSION') || keys.has('SCHEMA_VERSION') || keys.has('SETUP_COMPLETE');
}

function projectInitializationStatus_() {
  const props = PropertiesService.getScriptProperties();
  const ss = dbOrNull_();
  if (!ss) {
    return { initialized:false, canInitialize:false, spreadsheet:null };
  }

  if (props.getProperty(APP.propProjectInitialized) === 'TRUE') {
    return { initialized:true, canInitialize:false, spreadsheet:ss };
  }

  // Backward compatibility: courses initialized before v2.55.0 do not have
  // the dedicated marker. Recognize their existing _Config values and mark
  // them once without changing course data.
  if (legacyProjectLooksInitialized_(ss)) {
    props.setProperty(APP.propProjectInitialized, 'TRUE');
    return { initialized:true, canInitialize:false, spreadsheet:ss, legacy:true };
  }

  return { initialized:false, canInitialize:true, spreadsheet:ss };
}

function initializeProjectForSpreadsheet_(ss) {
  if (!ss) {
    throw new Error('This version must be used as a container-bound Apps Script attached to a Google Sheet.');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty(APP.propBoundSheetId, ss.getId());

  // Safe rerun guard. Existing initialized trackers are never rebuilt.
  if (props.getProperty(APP.propProjectInitialized) === 'TRUE' || legacyProjectLooksInitialized_(ss)) {
    props.setProperty(APP.propProjectInitialized, 'TRUE');
    return {
      ok:true,
      alreadyInitialized:true,
      spreadsheetUrl:ss.getUrl(),
      spreadsheetId:ss.getId(),
      message:'This Learning Progress Tracker has already been initialized. No setup changes were made.'
    };
  }

  Object.keys(APP.headers).forEach(name => ensureSheet_(ss, name, APP.headers[name]));

  setConfig_('APP_VERSION', APP.version);
  setConfig_('SCHEMA_VERSION', String(APP.schemaVersion));
  setConfig_('COURSE_CODE_TH', getConfig_('COURSE_CODE_TH') || '');
  setConfig_('COURSE_CODE_EN', getConfig_('COURSE_CODE_EN') || getConfig_('COURSE_CODE') || '');
  setConfig_('COURSE_NAME_TH', getConfig_('COURSE_NAME_TH') || '');
  setConfig_('COURSE_NAME_EN', getConfig_('COURSE_NAME_EN') || getConfig_('COURSE_NAME') || '');
  setConfig_('SETUP_COMPLETE', getConfig_('SETUP_COMPLETE') || 'FALSE');
  props.setProperty('LPT_SCHEMA_VERSION', String(APP.schemaVersion));

  const owner = currentEmail_(true);
  if (owner) upsertStaff_(owner, 'OWNER');

  applySystemSheetVisibility_();
  SpreadsheetApp.flush();
  props.setProperty(APP.propProjectInitialized, 'TRUE');
  resetRequestDataCache_();

  return {
    ok:true,
    alreadyInitialized:false,
    spreadsheetUrl:ss.getUrl(),
    spreadsheetId:ss.getId(),
    ownerEmail:owner || '',
    message:'Initialization complete. Continue with Course Setup in the Web app.'
  };
}

function setupProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('This version must be used as a container-bound Apps Script attached to a Google Sheet.');
  }
  return initializeProjectForSpreadsheet_(ss);
}

function initializeProjectFromWeb() {
  const status = projectInitializationStatus_();
  if (status.initialized) {
    return {
      ok:true,
      alreadyInitialized:true,
      spreadsheetUrl:status.spreadsheet ? status.spreadsheet.getUrl() : '',
      message:'This Learning Progress Tracker has already been initialized. No setup changes were made.'
    };
  }
  if (!status.canInitialize || !status.spreadsheet) {
    throw new Error('The bound Google Sheet could not be detected from this deployment. Open the bound Apps Script editor and run setupProject() once as a fallback.');
  }
  return initializeProjectForSpreadsheet_(status.spreadsheet);
}




function ensureHeaders_(sheetName) {
  const db = dbOrNull_();
  if (!db) return;
  const expected = APP.headers[sheetName];
  if (!Array.isArray(expected) || !expected.length) return;

  const sh = ensureSheet_(db, sheetName, expected);
  const lastCol = Math.max(1,sh.getLastColumn());
  const current = sh.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  let changed=false;

  expected.forEach(header => {
    if (!current.includes(header)) {
      const col = sh.getLastColumn()+1;
      sh.getRange(1,col).setValue(header);
      current.push(header);
      changed=true;
    }
  });

  if(changed) resetRequestDataCache_();
}

function storedSchemaVersion_(){
  const raw=String(PropertiesService.getScriptProperties().getProperty('LPT_SCHEMA_VERSION')||'').trim();
  if(/^\d+$/.test(raw))return Number(raw);
  if(raw)return 4;
  return 0;
}
function setStoredSchemaVersion_(version){
  const v=String(Number(version)||0);
  PropertiesService.getScriptProperties().setProperty('LPT_SCHEMA_VERSION',v);
  try{setConfig_('SCHEMA_VERSION',v);}catch(e){}
}
function ensureAllCurrentSheetsAndHeaders_(){
  const seen={};
  Object.keys(APP.tables).forEach(key=>{
    const name=APP.tables[key];
    if(!name||seen[name])return;
    seen[name]=true;
    if(Array.isArray(APP.headers[name]))ensureHeaders_(name);
  });
}
function migrateLegacyTeamsToActivityTeams_(){
  const db=dbOrNull_(); if(!db)return;
  ensureHeaders_(APP.tables.ACTIVITY_TEAMS);
  ensureHeaders_(APP.tables.ACTIVITY_TEAM_MEMBERS);

  const activities=rowsToObjects_(table_(APP.tables.ACTIVITIES)).filter(a=>String(a.Mode||'')==='Team');
  const students=rowsToObjects_(table_(APP.tables.STUDENTS)).filter(s=>bool_(s.Active)&&String(s.TeamCode||'').trim());
  const teamSheet=table_(APP.tables.ACTIVITY_TEAMS);
  const memberSheet=table_(APP.tables.ACTIVITY_TEAM_MEMBERS);
  const existingTeams=rowsToObjects_(teamSheet);

  activities.forEach(activity=>{
    if(existingTeams.some(r=>String(r.ActivityID)===String(activity.ActivityID)))return;
    const names=[...new Set(students.map(s=>String(s.TeamCode||'').trim()).filter(Boolean))].sort();
    names.forEach((name,index)=>{
      const teamId='TEAM-'+Utilities.getUuid();
      appendObject_(teamSheet,{
        ActivityID:activity.ActivityID,TeamID:teamId,TeamName:name,
        Sequence:index+1,CreatedAt:new Date(),UpdatedAt:new Date()
      });
      students.filter(s=>String(s.TeamCode||'').trim()===name).forEach(s=>{
        appendObject_(memberSheet,{
          ActivityID:activity.ActivityID,TeamID:teamId,StudentKey:s.StudentKey,AddedAt:new Date()
        });
      });
    });
  });
}



function migrate7to8_(){
  ensureHeaders_(APP.tables.WEEKLY_PLAN);
  ensureHeaders_(APP.tables.WEEKLY_PLAN_HISTORY);
}

function normalizeActivityVisibilityLevel_(value,published){
  const v=String(value||'').trim().toUpperCase();
  if(['HIDDEN','INFO','SUBMISSION','SCORE','SCORE_CLO'].includes(v))return v;
  return bool_(published)?'SCORE_CLO':'HIDDEN';
}

function migrate8to9_(){
  ensureHeaders_(APP.tables.STAFF);
  ensureHeaders_(APP.tables.ACTIVITIES);
  cleanupLegacyActivityDetailUrls_();

  const staffSh=table_(APP.tables.STAFF);
  rowsToObjects_(staffSh).forEach(r=>{
    if(String(r.StaffKey||'').trim())return;
    updateRowObject_(staffSh,r._row,Object.assign(stripRow_(r),{StaffKey:'STAFF-'+Utilities.getUuid(),UpdatedAt:new Date()}));
  });

  const activitySh=table_(APP.tables.ACTIVITIES);
  rowsToObjects_(activitySh).forEach(r=>{
    const level=normalizeActivityVisibilityLevel_(r.StudentVisibilityLevel,r.Published);
    updateRowObject_(activitySh,r._row,Object.assign(stripRow_(r),{
      StudentVisibilityLevel:level,
      Published:level!=='HIDDEN',
      UpdatedAt:r.UpdatedAt||new Date()
    }));
  });
}

function migrate6to7_(){
  ensureHeaders_(APP.tables.STAFF);
  ensureHeaders_(APP.tables.COURSE_SECTIONS);

  const staffRows=rowsToObjects_(table_(APP.tables.STAFF));
  const staffByEmail={};
  staffRows.forEach(r=>staffByEmail[normalizeEmail_(r.Email)]=r);

  const sectionSheet=table_(APP.tables.COURSE_SECTIONS);
  rowsToObjects_(sectionSheet).forEach(r=>{
    if(r.InstructorStaffKey)return;
    const email=normalizeEmail_(r.InstructorEmail||'');
    const st=email?staffByEmail[email]:null;
    if(st){
      updateRowObject_(sectionSheet,r._row,{
        InstructorStaffKey:String(st.StaffKey||st.Email||''),
        InstructorName:String(st.DisplayName||st.Name||st.Email||'')
      });
    }
  });
}

function migrate5to6_(){
  ensureHeaders_(APP.tables.COURSE_OBJECTIVES);
  ensureHeaders_(APP.tables.COURSE_SECTIONS);

  // Backward-compatible migration of old single-code/single-name values.
  const oldCode=String(getConfig_('COURSE_CODE')||'').trim();
  const oldName=String(getConfig_('COURSE_NAME')||'').trim();

  if(oldCode && !getConfig_('COURSE_CODE_EN') && !getConfig_('COURSE_CODE_TH')){
    setConfig_('COURSE_CODE_EN',oldCode);
  }
  if(oldName && !getConfig_('COURSE_NAME_EN') && !getConfig_('COURSE_NAME_TH')){
    setConfig_('COURSE_NAME_EN',oldName);
  }

  // Preserve legacy compact fields for older code paths.
  const codeEn=String(getConfig_('COURSE_CODE_EN')||'').trim();
  const codeTh=String(getConfig_('COURSE_CODE_TH')||'').trim();
  const nameEn=String(getConfig_('COURSE_NAME_EN')||'').trim();
  const nameTh=String(getConfig_('COURSE_NAME_TH')||'').trim();

  if((codeEn||codeTh) && !oldCode)setConfig_('COURSE_CODE',codeEn||codeTh);
  if((nameEn||nameTh) && !oldName)setConfig_('COURSE_NAME',nameEn||nameTh);
}

function migrate4to5_(){
  ensureAllCurrentSheetsAndHeaders_();
  migrateLegacyTeamsToActivityTeams_();
  cleanupLegacyActivityDetailUrls_();
}
function runSchemaMigrations_(fromVersion){
  let v=Number(fromVersion)||0;
  if(v<4){ensureAllCurrentSheetsAndHeaders_();v=4;}
  if(v<5){migrate4to5_();v=5;}
  if(v<6){migrate5to6_();v=6;}
  if(v<7){migrate6to7_();v=7;}
  if(v<8){migrate7to8_();v=8;}
  if(v<9){migrate8to9_();v=9;}
  if(v<10){migrate9to10_();v=10;}
  if(v<11){migrate10to11_();v=11;}
  if(v<12){migrate11to12_();v=12;}
  if(v<13){migrate12to13_();v=13;}
  if(v<14){migrate13to14_();v=14;}
  if(v<15){migrate14to15_();v=15;}
  setStoredSchemaVersion_(v);
  return v;
}
function ensureCurrentSchema_(){
  if(REQUEST_SCHEMA_READY_)return;
  const current=storedSchemaVersion_();
  if(current<Number(APP.schemaVersion))runSchemaMigrations_(current);
  repairStructuredSectionsV2521_();
  repairStructuredSectionsV2525_();
  repairOrphanStudentSectionsV2526_();
  REQUEST_SCHEMA_READY_=true;
}

function ensureActivityCloEvaluationSchema_() {
  ensureCurrentSchema_();
}

function migrate9to10_(){
  ensureHeaders_(APP.tables.ACTIVITY_CLO_SCORE_MAPS);
  const mapSh=table_(APP.tables.ACTIVITY_CLO_SCORE_MAPS);
  const existing=rowsToObjects_(mapSh);
  listActivities_().forEach(a=>{
    if(String(a.CLOEvalMode||'MANUAL').toUpperCase()!=='SCORE')return;
    const cloIds=activityCloIds_(a.ActivityID);
    if(cloIds.length!==1)return;
    const cloId=cloIds[0];
    if(existing.some(r=>String(r.ActivityID)===String(a.ActivityID)&&String(r.CLOID)===String(cloId)))return;
    const source=String(a.CLOScoreSource||'ACTIVITY_SCORE').toUpperCase()==='SEPARATE_SCORE'?'SEPARATE_SCORE':'ACTIVITY_SCORE';
    appendObject_(mapSh,{ActivityID:a.ActivityID,CLOID:cloId,ScoreSource:source,MaxScore:source==='SEPARATE_SCORE'?a.CLOMaxScore:a.MaxScore,UpdatedAt:new Date()});
  });
}


function migrate10to11_(){
  ensureHeaders_(APP.tables.WEEKLY_PLAN_CLOS);
  const planSh=table_(APP.tables.WEEKLY_PLAN);
  const cloPlanSh=table_(APP.tables.WEEKLY_PLAN_CLOS);
  const existing=rowsToObjects_(cloPlanSh);
  rowsToObjects_(planSh).forEach(plan=>{
    const planItemId=String(plan.PlanItemID||'').trim();
    if(!planItemId)return;
    const cloIds=String(plan.PlannedCLOIDs||'').split('|').map(x=>String(x||'').trim()).filter(Boolean);
    cloIds.forEach((cloId,index)=>{
      if(existing.some(r=>String(r.PlanItemID)===planItemId&&String(r.CLOID)===cloId))return;
      appendObject_(cloPlanSh,{
        PlanCLOID:'PLAN-CLO-'+Utilities.getUuid(),
        PlanItemID:planItemId,
        CLOID:cloId,
        TeachingMethod:'',
        AssessmentMethod:'',
        ActivityIDs:'',
        Sequence:index+1,
        CreatedAt:new Date(),
        UpdatedAt:new Date()
      });
    });
  });
}


function migrate11to12_(){
  ensureHeaders_(APP.tables.ACTIVITIES);
}

function migrate12to13_(){
  ensureHeaders_(APP.tables.STUDENT_SECTIONS);
  ensureHeaders_(APP.tables.ATTENDANCE_SESSIONS);
}

function scheduleTimeText_(value){
  if(value===null||value===undefined||value==='')return '';
  if(value instanceof Date&&!isNaN(value.getTime())){
    return Utilities.formatDate(value,Session.getScriptTimeZone()||'Asia/Bangkok','HH:mm');
  }
  const s=String(value).trim();
  if(!s)return '';
  const m=s.match(/(?:^|\s)([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?(?:\s|$)/);
  if(m&&(s.length>5||/1899|1900|GMT|Indochina/i.test(s)))return String(m[1]).padStart(2,'0')+':'+m[2];
  const plain=s.match(/^([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?$/);
  if(plain)return String(plain[1]).padStart(2,'0')+':'+plain[2];
  return s;
}
function legacyRoomAndExam_(value){
  const s=String(value==null?'':value).trim();
  const marker=s.search(/\s*วันเวลาสอบ\s*/);
  const dedupeRoom=text=>{
    const seen={};return String(text||'').split(',').map(x=>x.trim()).filter(x=>x&&!seen[x]&&(seen[x]=true)).join(', ');
  };
  if(marker<0)return {room:dedupeRoom(s),exam:null};
  const room=dedupeRoom(s.slice(0,marker));
  const examText=s.slice(marker).replace(/^\s*วันเวลาสอบ\s*/,'').trim();
  const em=examText.match(/(?:^|\s)(F|FINAL|M|MIDTERM)?\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+([0-2]?\d:[0-5]\d)\s*[-–]\s*([0-2]?\d:[0-5]\d)(?:\s+ห้อง\s+(.+))?/i);
  if(!em)return {room:room,exam:null};
  let year=Number(em[4]);if(year<100)year+=(year<70?2000:1900);
  const pad=n=>String(n).padStart(2,'0'),code=String(em[1]||'').toUpperCase();
  return {room:room,exam:{examType:(code==='F'||code==='FINAL')?'Final':(code==='M'||code==='MIDTERM')?'Midterm':'Exam',date:year+'-'+pad(em[3])+'-'+pad(em[2]),startTime:scheduleTimeText_(em[5]),endTime:scheduleTimeText_(em[6]),room:String(em[7]||'').trim()}};
}
function sectionCanonicalKey_(value){
  const s=String(value==null?'':value).trim();
  if(/^\d+$/.test(s))return 'N:'+s.replace(/^0+(?=\d)/,'');
  return 'S:'+s;
}
function remapSectionId_(tableName,fromIds,toId){
  const sh=table_(tableName),ids={};(fromIds||[]).forEach(x=>ids[String(x)]=true);
  rowsToObjects_(sh).forEach(r=>{if(ids[String(r.SectionID||'')])updateObjectRow_(sh,r._row,{SectionID:toId});});
}
function repairStructuredSectionsV2521_(){
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('LPT_V2521_SECTION_REPAIR_DONE')==='1')return;
  ensureHeaders_(APP.tables.COURSE_SECTION_SCHEDULES);
  const secSh=table_(APP.tables.COURSE_SECTIONS),scheduleSh=table_(APP.tables.COURSE_SECTION_SCHEDULES),now=new Date();
  const sections=rowsToObjects_(secSh).filter(r=>bool_(r.Active)),groups={};
  sections.forEach(r=>{const k=sectionCanonicalKey_(r.SectionCode);if(!groups[k])groups[k]=[];groups[k].push(r);});
  Object.keys(groups).forEach(k=>{
    const g=groups[k];if(!g.length)return;
    g.sort((a,b)=>String(b.SectionCode||'').length-String(a.SectionCode||'').length||a._row-b._row);
    const keep=g[0],keepId=String(keep.SectionID||'');if(!keepId)return;
    const registrarHint=g.some(r=>/วันเวลาสอบ/.test(String(r.Room||''))||/วันเวลาสอบ/.test(String(r.DayTime||'')));
    let keepCode=String(keep.SectionCode||'').trim();
    if(registrarHint&&/^\d{5}$/.test(keepCode))keepCode='0'+keepCode;
    if(keepCode!==String(keep.SectionCode||'')){updateObjectRow_(secSh,keep._row,{SectionCode:keepCode,UpdatedAt:now});writeSectionCodeText_(secSh,keep._row,keepCode);}
    const dup=g.slice(1),dupIds=dup.map(r=>String(r.SectionID||'')).filter(Boolean);
    if(dupIds.length){
      remapSectionId_(APP.tables.STUDENT_SECTIONS,dupIds,keepId);
      remapSectionId_(APP.tables.ATTENDANCE_SESSIONS,dupIds,keepId);
      remapSectionId_(APP.tables.COURSE_SECTION_SCHEDULES,dupIds,keepId);
      dup.forEach(r=>updateObjectRow_(secSh,r._row,{Active:false,UpdatedAt:now}));
    }
  });
  let scheduleRows=rowsToObjects_(scheduleSh).filter(r=>bool_(r.Active)),examAdds=[];
  scheduleRows.forEach(r=>{
    const roomInfo=legacyRoomAndExam_(r.Room),start=scheduleTimeText_(r.StartTime),end=scheduleTimeText_(r.EndTime);
    updateObjectRow_(scheduleSh,r._row,{StartTime:start,EndTime:end,Room:roomInfo.room,UpdatedAt:now});
    if(roomInfo.exam)examAdds.push({sectionId:String(r.SectionID||''),exam:roomInfo.exam});
  });
  scheduleRows=rowsToObjects_(scheduleSh).filter(r=>bool_(r.Active));
  examAdds.forEach(x=>{
    const e=x.exam,exists=scheduleRows.some(r=>String(r.SectionID||'')===x.sectionId&&String(r.ScheduleType||'').toUpperCase()==='EXAM'&&dateOnly_(r.ScheduleDate)===e.date&&scheduleTimeText_(r.StartTime)===e.startTime&&scheduleTimeText_(r.EndTime)===e.endTime);
    if(!exists){appendObject_(scheduleSh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:x.sectionId,ScheduleType:'EXAM',DayOfWeek:'',ScheduleDate:e.date,StartTime:e.startTime,EndTime:e.endTime,Room:e.room,ExamType:e.examType,Sequence:999,Active:true,CreatedAt:now,UpdatedAt:now});}
  });
  const seen={};rowsToObjects_(scheduleSh).filter(r=>bool_(r.Active)).forEach(r=>{
    const key=[String(r.SectionID||''),String(r.ScheduleType||'CLASS').toUpperCase(),String(r.DayOfWeek||''),dateOnly_(r.ScheduleDate),scheduleTimeText_(r.StartTime),scheduleTimeText_(r.EndTime),String(r.Room||'').trim(),String(r.ExamType||'')].join('|');
    if(seen[key])updateObjectRow_(scheduleSh,r._row,{Active:false,UpdatedAt:now});else seen[key]=true;
  });
  props.setProperty('LPT_V2521_SECTION_REPAIR_DONE','1');
  resetRequestDataCache_();
}
function repairStructuredSectionsV2525_(){
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('LPT_V2525_SECTION_EXAM_REPAIR_DONE')==='1')return;
  ensureHeaders_(APP.tables.COURSE_SECTION_SCHEDULES);
  const scheduleSh=table_(APP.tables.COURSE_SECTION_SCHEDULES),now=new Date();
  let rows=rowsToObjects_(scheduleSh).filter(r=>bool_(r.Active)),adds=[];
  rows.forEach(r=>{
    const info=legacyRoomAndExam_(r.Room),start=scheduleTimeText_(r.StartTime),end=scheduleTimeText_(r.EndTime);
    if(info.room!==String(r.Room||'')||start!==String(r.StartTime||'')||end!==String(r.EndTime||''))updateObjectRow_(scheduleSh,r._row,{Room:info.room,StartTime:start,EndTime:end,UpdatedAt:now});
    if(info.exam)adds.push({sectionId:String(r.SectionID||''),exam:info.exam});
  });
  rows=rowsToObjects_(scheduleSh).filter(r=>bool_(r.Active));
  adds.forEach(x=>{
    const e=x.exam,exists=rows.some(r=>String(r.SectionID||'')===x.sectionId&&String(r.ScheduleType||'').toUpperCase()==='EXAM'&&dateOnly_(r.ScheduleDate)===e.date&&scheduleTimeText_(r.StartTime)===e.startTime&&scheduleTimeText_(r.EndTime)===e.endTime);
    if(!exists)appendObject_(scheduleSh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:x.sectionId,ScheduleType:'EXAM',DayOfWeek:'',ScheduleDate:e.date,StartTime:e.startTime,EndTime:e.endTime,Room:e.room,ExamType:e.examType,Sequence:999,Active:true,CreatedAt:now,UpdatedAt:now});
  });
  // Also scan the legacy Section room because imports made after the v2.52.1 repair flag
  // could still have embedded exam text there.
  const secSh=table_(APP.tables.COURSE_SECTIONS);
  rowsToObjects_(secSh).filter(r=>bool_(r.Active)).forEach(r=>{
    const info=legacyRoomAndExam_(r.Room); if(!info.exam)return;
    updateObjectRow_(secSh,r._row,{Room:info.room,UpdatedAt:now});
    const e=info.exam,sectionId=String(r.SectionID||'');
    const exists=rowsToObjects_(scheduleSh).filter(x=>bool_(x.Active)).some(x=>String(x.SectionID||'')===sectionId&&String(x.ScheduleType||'').toUpperCase()==='EXAM'&&dateOnly_(x.ScheduleDate)===e.date&&scheduleTimeText_(x.StartTime)===e.startTime&&scheduleTimeText_(x.EndTime)===e.endTime);
    if(!exists)appendObject_(scheduleSh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:sectionId,ScheduleType:'EXAM',DayOfWeek:'',ScheduleDate:e.date,StartTime:e.startTime,EndTime:e.endTime,Room:e.room,ExamType:e.examType,Sequence:999,Active:true,CreatedAt:now,UpdatedAt:now});
  });
  props.setProperty('LPT_V2525_SECTION_EXAM_REPAIR_DONE','1');resetRequestDataCache_();
}
function migrate13to14_(){
  ensureHeaders_(APP.tables.COURSE_SECTION_SCHEDULES);
  const scheduleSh=table_(APP.tables.COURSE_SECTION_SCHEDULES);
  const existing=rowsToObjects_(scheduleSh);
  if(existing.length)return;
  const now=new Date();
  rowsToObjects_(table_(APP.tables.COURSE_SECTIONS)).filter(r=>bool_(r.Active)).forEach(sec=>{
    const sectionId=String(sec.SectionID||'').trim();
    if(!sectionId)return;
    const days=String(sec.Days||'').split('|').map(x=>String(x||'').trim()).filter(Boolean);
    const startTime=scheduleTimeText_(sec.StartTime),endTime=scheduleTimeText_(sec.EndTime),roomInfo=legacyRoomAndExam_(sec.Room),room=roomInfo.room;
    if(days.length){
      days.forEach((day,index)=>appendObject_(scheduleSh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:sectionId,ScheduleType:'CLASS',DayOfWeek:day,ScheduleDate:'',StartTime:startTime,EndTime:endTime,Room:room,ExamType:'',Sequence:index+1,Active:true,CreatedAt:now,UpdatedAt:now}));
    }else if(startTime||endTime||room){
      appendObject_(scheduleSh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:sectionId,ScheduleType:'CLASS',DayOfWeek:'',ScheduleDate:'',StartTime:startTime,EndTime:endTime,Room:room,ExamType:'',Sequence:1,Active:true,CreatedAt:now,UpdatedAt:now});
    }
    if(roomInfo.exam){const e=roomInfo.exam;appendObject_(scheduleSh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:sectionId,ScheduleType:'EXAM',DayOfWeek:'',ScheduleDate:e.date,StartTime:e.startTime,EndTime:e.endTime,Room:e.room,ExamType:e.examType,Sequence:999,Active:true,CreatedAt:now,UpdatedAt:now});}
  });
}

function migrate14to15_(){
  // Add optional student-facing explanations while preserving existing CLO IDs and mappings.
  ensureHeaders_(APP.tables.CLOS);
}

function weeklyTeachingPlanCloRows_(){
  ensureHeaders_(APP.tables.WEEKLY_PLAN_CLOS);
  return rowsToObjects_(table_(APP.tables.WEEKLY_PLAN_CLOS));
}

function activityCloScoreMappings_(activityId){
  if(!activityId)return[];
  ensureHeaders_(APP.tables.ACTIVITY_CLO_SCORE_MAPS);
  return rowsToObjects_(table_(APP.tables.ACTIVITY_CLO_SCORE_MAPS)).filter(r=>String(r.ActivityID)===String(activityId)).map(r=>({
    cloId:String(r.CLOID||''),scoreSource:String(r.ScoreSource||'ACTIVITY_SCORE').toUpperCase()==='SEPARATE_SCORE'?'SEPARATE_SCORE':'ACTIVITY_SCORE',
    maxScore:r.MaxScore===''||r.MaxScore==null?'':Number(r.MaxScore)
  }));
}
function normalizeCloScoreMappingsInput_(activity,cloIds){
  const raw=Array.isArray(activity.cloScoreMappings)?activity.cloScoreMappings:[]; const by={};
  raw.forEach(m=>{if(m&&m.cloId)by[String(m.cloId)]=m;});
  return (cloIds||[]).map(cloId=>{const m=by[String(cloId)]||{};const source=String(m.scoreSource||'ACTIVITY_SCORE').toUpperCase()==='SEPARATE_SCORE'?'SEPARATE_SCORE':'ACTIVITY_SCORE';return {cloId:String(cloId),scoreSource:source,maxScore:source==='SEPARATE_SCORE'?(m.maxScore===''||m.maxScore==null?'':Number(m.maxScore)):Number(activity.maxScore)};});
}

function validateScoreBasedCloConfig_(activity, cloIds) {
  const mode=String(activity.cloEvalMode||'MANUAL').toUpperCase(); if(mode!=='SCORE')return;
  if(!Array.isArray(cloIds)||!cloIds.length)throw new Error('Score-based CLO evaluation requires at least one linked CLO.');
  const mappings=normalizeCloScoreMappingsInput_(activity,cloIds);
  mappings.forEach(m=>{if(m.scoreSource==='ACTIVITY_SCORE'){const maxScore=Number(activity.maxScore);if(!Number.isFinite(maxScore)||maxScore<=0)throw new Error('Using Activity Score for CLO evaluation requires a positive Activity Max Score.');}else if(!Number.isFinite(Number(m.maxScore))||Number(m.maxScore)<=0)throw new Error('Each separate CLO score requires a positive maximum score.');});
  const t2=Number(activity.cloLevel2Pct),t3=Number(activity.cloLevel3Pct),t4=Number(activity.cloLevel4Pct),t5=Number(activity.cloLevel5Pct);
  if([t2,t3,t4,t5].some(x=>!Number.isFinite(x)||x<0||x>100))throw new Error('CLO thresholds must be numbers from 0 to 100.');
  if(!(t2<t3&&t3<t4&&t4<t5))throw new Error('CLO thresholds must increase from Level 2 through Level 5.');
}
function deriveCloLevelFromMappedScore_(activity,mapping,activityScore,separateScore){
  if(!activity||String(activity.CLOEvalMode||'MANUAL').toUpperCase()!=='SCORE')return '';
  const source=mapping&&mapping.scoreSource==='SEPARATE_SCORE'?'SEPARATE_SCORE':'ACTIVITY_SCORE'; const maxScore=source==='SEPARATE_SCORE'?Number(mapping.maxScore):Number(activity.MaxScore); const score=source==='SEPARATE_SCORE'?Number(separateScore):Number(activityScore);
  if(!Number.isFinite(maxScore)||maxScore<=0||!Number.isFinite(score))return ''; const pct=score/maxScore*100; const t2=Number(activity.CLOLevel2Pct),t3=Number(activity.CLOLevel3Pct),t4=Number(activity.CLOLevel4Pct),t5=Number(activity.CLOLevel5Pct); if([t2,t3,t4,t5].some(x=>!Number.isFinite(x)))return ''; if(pct>=t5)return 5;if(pct>=t4)return 4;if(pct>=t3)return 3;if(pct>=t2)return 2;return 1;
}
function saveDerivedCloEvidence_(activity,targetType,targetId,activityScore,separateCloScores,staffEmail,requestedLevels){
  if(!activity||String(activity.CLOEvalMode||'MANUAL').toUpperCase()!=='SCORE')return false; const mappings=Array.isArray(activity.cloScoreMappings)?activity.cloScoreMappings:activityCloScoreMappings_(activity.ActivityID); if(!mappings.length)return false; const es=table_(APP.tables.CLO_EVIDENCE); const separate=(separateCloScores&&typeof separateCloScores==='object')?separateCloScores:{};
  mappings.forEach(mapping=>{const cloId=mapping.cloId; const evidenceMax=mapping.scoreSource==='SEPARATE_SCORE'?mapping.maxScore:activity.MaxScore; let evidenceScore=mapping.scoreSource==='SEPARATE_SCORE'?separate[cloId]:activityScore; if(evidenceScore!==''&&evidenceScore!=null&&Number.isFinite(Number(evidenceScore))&&Number.isFinite(Number(evidenceMax))&&Number(evidenceMax)>0)evidenceScore=Math.max(0,Math.min(Number(evidenceMax),Number(evidenceScore))); const derived=deriveCloLevelFromMappedScore_(activity,mapping,activityScore,evidenceScore); if(derived===''){deleteComposite_(es,['ActivityID','TargetType','TargetID','CLOID'],[activity.ActivityID,targetType,targetId,cloId]);return;} const allowOverride=bool_(activity.CLOAllowOverride); const raw=requestedLevels&&requestedLevels[cloId]!==undefined?requestedLevels[cloId]:''; const requested=(raw===''||raw==null)?'':Number(raw); let level=derived,source='SCORE'; if(allowOverride&&requested!==''&&Number.isFinite(requested)){level=clampNum_(requested,APP.cloMin,APP.cloMax,derived);source=Number(level)===Number(derived)?'SCORE':'OVERRIDE';}
    upsertComposite_(es,['ActivityID','TargetType','TargetID','CLOID'],[activity.ActivityID,targetType,targetId,cloId],{ActivityID:activity.ActivityID,TargetType:targetType,TargetID:targetId,CLOID:cloId,Level:level,Source:source,Score:Number.isFinite(Number(evidenceScore))?Number(evidenceScore):'',MaxScore:Number.isFinite(Number(evidenceMax))?Number(evidenceMax):'',UpdatedAt:new Date(),UpdatedBy:staffEmail});
  }); return true;
}

function deleteRowsWhere_(sheet,predicate){
  const rows=rowsToObjects_(sheet).filter(predicate).map(r=>Number(r._row)).filter(Number.isFinite).sort((a,b)=>b-a);
  rows.forEach(row=>sheet.deleteRow(row));
}

function activityTeamData_(activityId){
  const teams=rowsToObjects_(table_(APP.tables.ACTIVITY_TEAMS))
    .filter(r=>String(r.ActivityID)===String(activityId))
    .sort((a,b)=>Number(a.Sequence||0)-Number(b.Sequence||0));
  const members=rowsToObjects_(table_(APP.tables.ACTIVITY_TEAM_MEMBERS))
    .filter(r=>String(r.ActivityID)===String(activityId));
  const byTeam={};
  members.forEach(r=>{if(!byTeam[r.TeamID])byTeam[r.TeamID]=[];byTeam[r.TeamID].push(String(r.StudentKey));});
  return teams.map(t=>({teamId:String(t.TeamID),teamName:String(t.TeamName||t.TeamID),
    sequence:Number(t.Sequence||0),studentKeys:byTeam[t.TeamID]||[]}));
}

function getActivityTeams(activityId){
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const activity=getActivity_(activityId);
  if(!activity)throw new Error('Activity not found.');
  if(String(activity.Mode)!=='Team')throw new Error('Teams are only available for Team activities.');
  const students=listStudentsForStaff_().filter(s=>bool_(s.Active));
  const teams=activityTeamData_(activityId),assigned={};
  teams.forEach(t=>t.studentKeys.forEach(k=>assigned[k]=true));
  const sourceActivities=listActivities_()
    .filter(a=>a.Active!==false&&String(a.Mode)==='Team'&&String(a.ActivityID)!==String(activityId))
    .map(a=>({activityId:a.ActivityID,title:a.Title,teamCount:activityTeamData_(a.ActivityID).length}))
    .filter(a=>a.teamCount>0);
  const hasResults=rowsToObjects_(table_(APP.tables.RESULTS)).some(r=>
    String(r.ActivityID)===String(activityId)&&String(r.TargetType)==='TEAM');
  return {
    activity:serializeActivity_(activity),teams,
    students:students.map(s=>({studentKey:String(s.StudentKey),studentId:String(s.StudentId||''),
      name:String(s.Name||''),alias:String(s.Alias||'')})),
    unassigned:students.filter(s=>!assigned[s.StudentKey]).map(s=>String(s.StudentKey)),
    sourceActivities,hasResults
  };
}

function saveActivityTeams(activityId,payload){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  const activity=getActivity_(activityId);
  if(!activity)throw new Error('Activity not found.');
  if(String(activity.Mode)!=='Team')throw new Error('Teams are only available for Team activities.');
  payload=payload||{};
  const existingResults=rowsToObjects_(table_(APP.tables.RESULTS)).some(r=>
    String(r.ActivityID)===String(activityId)&&String(r.TargetType)==='TEAM');
  if(existingResults&&!bool_(payload.confirmWithResults)){
    throw new Error('This activity already has team results. Confirm team changes explicitly.');
  }
  const allowed=new Set(listStudentsForStaff_().filter(s=>bool_(s.Active)).map(s=>String(s.StudentKey)));
  const seen=new Set(),normalized=[];
  (Array.isArray(payload.teams)?payload.teams:[]).slice(0,300).forEach((raw,index)=>{
    const teamId=clean_(raw.teamId,120)||('TEAM-'+Utilities.getUuid());
    const teamName=clean_(raw.teamName,100)||`Team ${index+1}`;
    const keys=[];
    (Array.isArray(raw.studentKeys)?raw.studentKeys:[]).map(String).forEach(k=>{
      if(!allowed.has(k))return;
      if(seen.has(k))throw new Error('A student cannot belong to more than one team in the same activity.');
      seen.add(k);keys.push(k);
    });
    normalized.push({teamId,teamName,studentKeys:keys});
  });
  const ts=table_(APP.tables.ACTIVITY_TEAMS),ms=table_(APP.tables.ACTIVITY_TEAM_MEMBERS);
  deleteRowsWhere_(ts,r=>String(r.ActivityID)===String(activityId));
  deleteRowsWhere_(ms,r=>String(r.ActivityID)===String(activityId));
  normalized.forEach((t,index)=>{
    appendObject_(ts,{ActivityID:activityId,TeamID:t.teamId,TeamName:t.teamName,
      Sequence:index+1,CreatedAt:new Date(),UpdatedAt:new Date()});
    t.studentKeys.forEach(k=>appendObject_(ms,{ActivityID:activityId,TeamID:t.teamId,StudentKey:k,AddedAt:new Date()}));
  });
  resetRequestDataCache_();
  audit_('SAVE_ACTIVITY_TEAMS','Activity',activityId,{actor:actor.email,teams:normalized.length,assigned:seen.size});
  return getActivityTeams(activityId);
}

function copyActivityTeams(sourceActivityId,targetActivityId,confirmWithResults){
  requireStaff_(['OWNER','INSTRUCTOR']);
  if(String(sourceActivityId)===String(targetActivityId))throw new Error('Choose a different source activity.');
  const source=getActivity_(sourceActivityId),target=getActivity_(targetActivityId);
  if(!source||!target)throw new Error('Source or target activity not found.');
  if(String(source.Mode)!=='Team'||String(target.Mode)!=='Team')throw new Error('Both activities must use Team mode.');
  const sourceTeams=activityTeamData_(sourceActivityId);
  if(!sourceTeams.length)throw new Error('The source activity has no teams to copy.');
  const result=saveActivityTeams(targetActivityId,{
    confirmWithResults:bool_(confirmWithResults),
    teams:sourceTeams.map(t=>({teamName:t.teamName,studentKeys:t.studentKeys}))
  });
  audit_('COPY_ACTIVITY_TEAMS','Activity',targetActivityId,{sourceActivityId,teams:sourceTeams.length});
  return result;
}

function ensureTeamHistorySchema_() {
  const db = dbOrNull_();
  if (!db) return;
  ensureSheet_(db, APP.tables.TEAM_RESULT_MEMBERS, APP.headers[APP.tables.TEAM_RESULT_MEMBERS]);
}

function snapshotExistingTeamResultMemberships_() {
  ensureTeamHistorySchema_();

  const results = rowsToObjects_(table_(APP.tables.RESULTS)).filter(r => r.TargetType === 'TEAM' && r.TargetID);
  if (!results.length) return;

  const students = listStudentsForStaff_().filter(s => s.Active && s.TeamCode);
  const historySheet = table_(APP.tables.TEAM_RESULT_MEMBERS);
  const existing = rowsToObjects_(historySheet);
  const known = {};
  existing.forEach(r => {
    known[String(r.ActivityID) + '|' + String(r.TeamCode) + '|' + String(r.StudentKey)] = true;
  });

  results.forEach(r => {
    students.filter(s => String(s.TeamCode) === String(r.TargetID)).forEach(s => {
      const key = String(r.ActivityID) + '|' + String(r.TargetID) + '|' + String(s.StudentKey);
      if (!known[key]) {
        appendObject_(historySheet, {
          ActivityID:r.ActivityID,
          TeamCode:r.TargetID,
          StudentKey:s.StudentKey,
          CapturedAt:new Date()
        });
        known[key] = true;
      }
    });
  });
}

function captureTeamResultMembers_(activityId, teamId) {
  ensureTeamHistorySchema_();
  let keys=rowsToObjects_(table_(APP.tables.ACTIVITY_TEAM_MEMBERS))
    .filter(r=>String(r.ActivityID)===String(activityId)&&String(r.TeamID)===String(teamId))
    .map(r=>String(r.StudentKey));
  if(!keys.length){
    keys=listStudentsForStaff_().filter(s=>s.Active&&String(s.TeamCode||'')===String(teamId||''))
      .map(s=>String(s.StudentKey));
  }
  if(!keys.length)return;
  const sh=table_(APP.tables.TEAM_RESULT_MEMBERS),existing=rowsToObjects_(sh),known={};
  existing.forEach(r=>known[String(r.ActivityID)+'|'+String(r.TeamCode)+'|'+String(r.StudentKey)]=true);
  keys.forEach(k=>{
    const id=String(activityId)+'|'+String(teamId)+'|'+k;
    if(known[id])return;
    appendObject_(sh,{ActivityID:activityId,TeamCode:teamId,StudentKey:k,CapturedAt:new Date()});
    known[id]=true;
  });
}

function previewTeamAssignments(rows) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  if (!Array.isArray(rows) || !rows.length) throw new Error('No team assignments were provided.');

  const students = listStudentsForStaff_().filter(s => s.Active);
  const byId = {};
  students.forEach(s => {
    byId[String(s.StudentId || '').trim()] = s;
  });

  const normalized = rows.slice(0,3000).map((r,index) => ({
    sourceRow:Number(r.sourceRow || index + 2),
    studentId:clean_(r.studentId,60),
    teamCode:clean_(r.teamCode,100)
  }));

  const groups = {};
  normalized.forEach(r => {
    if (!r.studentId) return;
    if (!groups[r.studentId]) groups[r.studentId] = [];
    groups[r.studentId].push(r);
  });

  const ready = [];
  const problems = [];
  const duplicateIgnored = [];
  const seen = {};

  normalized.forEach(r => {
    const exactKey = r.studentId + '|' + r.teamCode.toLowerCase();
    if (seen[exactKey]) {
      duplicateIgnored.push({
        sourceRow:r.sourceRow,
        studentId:r.studentId,
        teamCode:r.teamCode,
        reason:'Exact duplicate assignment'
      });
      return;
    }
    seen[exactKey] = true;

    const issues = [];
    if (!r.studentId) issues.push('Missing Student ID');

    const student = r.studentId ? byId[r.studentId] : null;
    if (r.studentId && !student) issues.push('Student ID is not registered in this course');

    const sameStudent = r.studentId ? (groups[r.studentId] || []) : [];
    if (sameStudent.length > 1) {
      const teams = [];
      sameStudent.forEach(x => {
        const t = String(x.teamCode || '').toLowerCase();
        if (!teams.includes(t)) teams.push(t);
      });
      if (teams.length > 1) issues.push('Student appears more than once with conflicting team assignments');
    }

    if (issues.length) {
      problems.push({
        sourceRow:r.sourceRow,
        studentId:r.studentId,
        name:student ? student.Name : '',
        currentTeam:student ? (student.TeamCode || '') : '',
        newTeam:r.teamCode,
        reason:issues.join('; ')
      });
      return;
    }

    const currentTeam = student.TeamCode || '';
    if (String(currentTeam) === String(r.teamCode)) {
      duplicateIgnored.push({
        sourceRow:r.sourceRow,
        studentId:r.studentId,
        teamCode:r.teamCode,
        reason:'No change'
      });
      return;
    }

    ready.push({
      sourceRow:r.sourceRow,
      studentKey:student.StudentKey,
      studentId:student.StudentId,
      name:student.Name,
      currentTeam:currentTeam,
      newTeam:r.teamCode
    });
  });

  const uniqueTeams = {};
  students.forEach(s => {
    if (s.TeamCode) uniqueTeams[String(s.TeamCode)] = true;
  });
  ready.forEach(r => {
    if (r.newTeam) uniqueTeams[String(r.newTeam)] = true;
  });

  return {
    ready,
    problems,
    duplicateIgnored,
    summary:{
      sourceRows:normalized.length,
      changes:ready.length,
      problems:problems.length,
      ignored:duplicateIgnored.length,
      resultingTeams:Object.keys(uniqueTeams).length
    }
  };
}

function confirmTeamAssignments(rows) {
  const actor = requireStaff_(['OWNER','INSTRUCTOR']);
  if (!Array.isArray(rows) || !rows.length) throw new Error('There are no validated team changes to apply.');

  // Preserve memberships used by any existing team results before current TeamCode changes.
  snapshotExistingTeamResultMemberships_();

  const sh = table_(APP.tables.STUDENTS);
  const students = rowsToObjects_(sh);
  const byKey = {};
  students.forEach(s => byKey[String(s.StudentKey)] = s);

  let updated = 0;
  const rejected = [];

  rows.slice(0,3000).forEach((r,index) => {
    const key = String(r.studentKey || '').trim();
    const row = byKey[key];
    if (!row || !bool_(row.Active)) {
      rejected.push({
        sourceRow:Number(r.sourceRow || index + 2),
        studentId:r.studentId || '',
        reason:'Student not found or inactive'
      });
      return;
    }

    const newTeam = clean_(r.newTeam,100);
    updateObjectRow_(sh,row._row,{
      StudentKey:row.StudentKey,
      StudentId:row.StudentId,
      Name:row.Name,
      Email:row.Email,
      Alias:row.Alias,
      TeamCode:newTeam,
      Active:bool_(row.Active),
      AddedAt:row.AddedAt || new Date()
    });
    updated++;
  });

  audit_('CONFIRM_TEAM_ASSIGNMENTS','Roster','',{
    actor:actor.email,
    updated,
    rejected:rejected.length
  });

  return {
    ok:true,
    updated,
    rejected,
    students:listStudentsForStaff_()
  };
}

function isCourseConfigured_(){
  if(!dbOrNull_())return false;
  const cfg=configMapFast_();
  return courseProfileComplete_(cfg);
}

function ensureCourseSetupFlag_() {
  const configured = isCourseConfigured_();
  const flag = getConfig_('SETUP_COMPLETE') === 'TRUE';

  if (configured && !flag) {
    setConfig_('SETUP_COMPLETE', 'TRUE');
    audit_('REPAIR_SETUP_FLAG','Course',getConfig_('COURSE_CODE'),{});
    return true;
  }
  return flag || configured;
}


function ensureStaffAuthSchema_() {
  const db = dbOrNull_();
  if (!db) return;
  const sh = db.getSheetByName(APP.tables.STAFF);
  if (!sh) return;

  let headers = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  ['StaffTokenHash','TokenIssuedAt'].forEach(h => {
    if (!headers.includes(h)) {
      sh.getRange(1, sh.getLastColumn()+1).setValue(h);
      headers.push(h);
    }
  });
}

function generateStaffToken_() {
  return [
    Utilities.getUuid().replace(/-/g,''),
    Utilities.getUuid().replace(/-/g,''),
    Utilities.getUuid().replace(/-/g,'')
  ].join('');
}

function hashStaffToken_(token) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token || ''),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => {
    const n = b < 0 ? b + 256 : b;
    return ('0' + n.toString(16)).slice(-2);
  }).join('');
}

function getStaffByToken_(token) {
  ensureStaffAuthSchema_();
  token = String(token || '').trim();
  if (!token || token.length < 64) return null;

  const hash = hashStaffToken_(token);
  const row = rowsToObjects_(table_(APP.tables.STAFF)).find(r =>
    r.StaffTokenHash &&
    String(r.StaffTokenHash) === hash &&
    bool_(r.Active) &&
    ['OWNER','INSTRUCTOR','TA'].includes(String(r.Role))
  );

  return row ? {
    Email: normalizeEmail_(row.Email),
    Role: String(row.Role),
    Active: bool_(row.Active),
    TokenIssuedAt: row.TokenIssuedAt
  } : null;
}

function staffAuthSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = String(props.getProperty('LPT_STAFF_AUTH_SECRET') || '');
  if (!secret) {
    secret = Utilities.base64EncodeWebSafe(Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/=+$/,'');
    props.setProperty('LPT_STAFF_AUTH_SECRET', secret);
  }
  return secret;
}

function staffSessionSign_(payloadB64) {
  const sig = Utilities.computeHmacSha256Signature(String(payloadB64 || ''), staffAuthSecret_(), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/,'');
}

function issueVerifiedStaffSession_(tokenStaff) {
  const now = Date.now();
  const payload = {
    email: normalizeEmail_(tokenStaff.Email),
    role: String(tokenStaff.Role || ''),
    iat: now,
    exp: now + (12 * 60 * 60 * 1000)
  };
  const payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/,'');
  return payloadB64 + '.' + staffSessionSign_(payloadB64);
}

function verifiedStaffSession_(sessionToken, tokenStaff) {
  sessionToken = String(sessionToken || '').trim();
  if (!sessionToken || !tokenStaff) return null;
  const parts = sessionToken.split('.');
  if (parts.length !== 2) return null;
  const payloadB64 = parts[0];
  if (staffSessionSign_(payloadB64) !== parts[1]) return null;
  try {
    const payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString('UTF-8'));
    if (!payload || Number(payload.exp || 0) <= Date.now()) return null;
    if (normalizeEmail_(payload.email) !== normalizeEmail_(tokenStaff.Email)) return null;
    if (String(payload.role || '') !== String(tokenStaff.Role || '')) return null;
    return payload;
  } catch (e) {
    return null;
  }
}

function staffOtpKey_(prefix, staffToken) {
  const hash = hashStaffToken_(String(staffToken || ''));
  return 'LPT_STAFF_' + prefix + '_' + hash.slice(0, 40);
}

function staffOtpHash_(staffToken, email, code) {
  const raw = [hashStaffToken_(staffToken), normalizeEmail_(email), String(code || '').trim()].join('|');
  const sig = Utilities.computeHmacSha256Signature(raw, staffAuthSecret_(), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/,'');
}

function requestStaffEmailVerification(staffToken) {
  const tokenStaff = getStaffByToken_(staffToken);
  if (!tokenStaff) throw new Error('This staff invitation link is invalid, deactivated, or has been replaced.');
  const cache = CacheService.getScriptCache();
  const throttleKey = staffOtpKey_('OTP_THROTTLE', staffToken);
  if (cache.get(throttleKey)) {
    return {ok:true, email:tokenStaff.Email, waitSeconds:60, message:'A verification code was recently sent. Please check your inbox.'};
  }
  const rnd = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + '|' + Date.now(), Utilities.Charset.UTF_8);
  const n = ((rnd[0] & 255) << 24) ^ ((rnd[1] & 255) << 16) ^ ((rnd[2] & 255) << 8) ^ (rnd[3] & 255);
  const code = String(100000 + (Math.abs(n) % 900000));
  const record = JSON.stringify({
    hash: staffOtpHash_(staffToken, tokenStaff.Email, code),
    attempts: 0,
    email: normalizeEmail_(tokenStaff.Email),
    expiresAt: Date.now() + (10 * 60 * 1000)
  });
  cache.put(staffOtpKey_('OTP', staffToken), record, 600);
  cache.put(throttleKey, '1', 60);
  MailApp.sendEmail({
    to: tokenStaff.Email,
    subject: 'Learning Progress Tracker verification code',
    body: 'Your Learning Progress Tracker verification code is ' + code + '.\n\nThis code expires in 10 minutes. If you did not request it, you can ignore this email.',
    htmlBody: '<p>Your Learning Progress Tracker verification code is:</p><p style="font-size:26px;font-weight:700;letter-spacing:4px">' + code + '</p><p>This code expires in 10 minutes. If you did not request it, you can ignore this email.</p>',
    name: 'Learning Progress Tracker'
  });
  return {ok:true, email:tokenStaff.Email, waitSeconds:60};
}

function verifyStaffEmailCode(staffToken, code) {
  const tokenStaff = getStaffByToken_(staffToken);
  if (!tokenStaff) throw new Error('This staff invitation link is invalid, deactivated, or has been replaced.');
  code = String(code || '').replace(/\s+/g,'');
  if (!/^\d{6}$/.test(code)) throw new Error('Enter the 6-digit verification code.');
  const cache = CacheService.getScriptCache();
  const key = staffOtpKey_('OTP', staffToken);
  const raw = cache.get(key);
  if (!raw) throw new Error('The verification code has expired. Request a new code.');
  let record;
  try { record = JSON.parse(raw); } catch(e) { record = null; }
  if (!record || Number(record.expiresAt || 0) <= Date.now()) {
    cache.remove(key);
    throw new Error('The verification code has expired. Request a new code.');
  }
  record.attempts = Number(record.attempts || 0) + 1;
  if (record.attempts > 5) {
    cache.remove(key);
    throw new Error('Too many incorrect attempts. Request a new verification code.');
  }
  if (staffOtpHash_(staffToken, tokenStaff.Email, code) !== String(record.hash || '')) {
    cache.put(key, JSON.stringify(record), Math.max(1, Math.floor((Number(record.expiresAt)-Date.now())/1000)));
    throw new Error('Incorrect verification code. Please try again.');
  }
  cache.remove(key);
  const staffSession = issueVerifiedStaffSession_(tokenStaff);
  audit_('VERIFY_STAFF_EMAIL','Staff',tokenStaff.Email,{role:tokenStaff.Role,method:'EMAIL_OTP'});
  return {ok:true, email:tokenStaff.Email, role:tokenStaff.Role, staffSession:staffSession, expiresHours:12};
}

function assertStaffInviteIdentity_(tokenStaff, staffSession) {
  if (!tokenStaff) throw new Error('This staff invitation/session is invalid or has been revoked.');

  const signedInEmail = currentEmail_(false);
  const expectedEmail = normalizeEmail_(tokenStaff.Email);

  // If Apps Script exposes the active Google identity, require an exact match.
  if (signedInEmail) {
    if (normalizeEmail_(signedInEmail) !== expectedEmail) {
      throw new Error(`This staff access link is issued to ${expectedEmail}. You are currently signed in as ${signedInEmail}. Please switch Google accounts and try again.`);
    }
    return {email:expectedEmail, via:'GOOGLE_IDENTITY'};
  }

  // Execute-as-owner deployments often cannot expose Session.getActiveUser().
  // In that case require a short-lived session obtained by verifying a code
  // delivered to the exact Staff.Email. A bearer invite token alone is never
  // sufficient for OWNER / INSTRUCTOR / TA access.
  if (verifiedStaffSession_(staffSession, tokenStaff)) {
    return {email:expectedEmail, via:'EMAIL_OTP'};
  }

  throw new Error(`EMAIL_VERIFICATION_REQUIRED:${expectedEmail}`);
}

function staffInvitationUrl_(token) {
  const base = ScriptApp.getService().getUrl() || '';
  return base ? `${base}?staffToken=${encodeURIComponent(token)}` : '';
}

function issueStaffInvite_(email) {
  ensureStaffAuthSchema_();
  email = normalizeEmail_(email);
  const sh = table_(APP.tables.STAFF);
  const row = findObject_(sh,'Email',email);
  if (!row) throw new Error('Staff account not found.');

  const token = generateStaffToken_();
  const now = new Date();

  updateObjectRow_(sh,row._row,{
    Email:row.Email,
    Role:row.Role,
    Active:bool_(row.Active),
    AddedAt:row.AddedAt || now,
    StaffTokenHash:hashStaffToken_(token),
    TokenIssuedAt:now
  });

  return {
    email,
    role:row.Role,
    invitationLink:staffInvitationUrl_(token),
    url:staffInvitationUrl_(token),
    tokenIssuedAt:dateIso_(now)
  };
}

function dispatchStaff(action, args, staffToken, staffSession) {
  const tokenStaff = getStaffByToken_(staffToken);
  if (!tokenStaff) throw new Error('This staff invitation/session is invalid or has been revoked.');

  const verifiedIdentity = assertStaffInviteIdentity_(tokenStaff, staffSession);

  REQUEST_STAFF_AUTH_ = {
    email: tokenStaff.Email,
    role: tokenStaff.Role,
    via: verifiedIdentity.via || 'STAFF_INVITE'
  };

  const a = Array.isArray(args) ? args : [];
  const actions = {
    getStaffDashboard: () => getStaffDashboard(),
    getStaffPageData: () => getStaffPageData(...a),
    getDashboardSectionData: () => getDashboardSectionData(...a),
    getCloAttainmentAnalytics: () => getCloAttainmentAnalytics(...a),
    saveCloAttainmentPolicy: () => saveCloAttainmentPolicy(...a),
    getCourseSettings: () => getCourseSettings(),
    getSystemStatus: () => getSystemStatus(),
    getWeeklyTeachingPlan: () => getWeeklyTeachingPlan(...a),
    saveWeeklyTeachingPlan: () => saveWeeklyTeachingPlan(...a),
    updateTeachingPlanExecution: () => updateTeachingPlanExecution(...a),
    getTeachingPlanHistory: () => getTeachingPlanHistory(...a),
    saveCourseSettings: () => saveCourseSettings(...a),
    addCourseStaffProfile: () => addCourseStaffProfile_(...a),
    updateCourseStaffProfile: () => updateCourseStaffProfile_(...a),
    importRoster: () => importRoster(...a),
    saveCLOs: () => saveCLOs(...a),
    saveActivity: () => saveActivity(...a),
    deleteActivity: () => deleteActivity(...a),
    archiveActivity: () => archiveActivity(...a),
    restoreActivity: () => restoreActivity(...a),
    getResultEntryData: () => getResultEntryData(...a),
    getActivityTeams: () => getActivityTeams(...a),
    saveActivityTeams: () => saveActivityTeams(...a),
    copyActivityTeams: () => copyActivityTeams(...a),
    inspectResultSpreadsheet: () => inspectResultSpreadsheet(...a),
    previewGoogleFormImport: () => previewGoogleFormImport(...a),
    previewUploadedResultRows: () => previewUploadedResultRows(...a),
    applyResultImport: () => applyResultImport(...a),
    clearActivityImportedResults: () => clearActivityImportedResults(...a),
    saveResults: () => saveResults(...a),
    createAttendanceSession: () => createAttendanceSession(...a),
    getAttendanceLive: () => getAttendanceLive(...a),
    getAttendanceLiveStatus: () => getAttendanceLiveStatus(...a),
    closeAttendanceSession: () => closeAttendanceSession(...a),
    setManualAttendance: () => setManualAttendance(...a),
    showSystemSheets: () => showSystemSheets(...a),
    hideSystemSheets: () => hideSystemSheets(...a),
    listCourseStaff: () => listCourseStaff(),
    saveCourseStaff: () => saveCourseStaff(...a),
    setCourseStaffActive: () => setCourseStaffActive(...a),
    regenerateStaffInvite: () => regenerateStaffInvite(...a),
    sendStaffAccessEmailBatch: () => sendStaffAccessEmailBatch(...a),
    previewStudentDashboard: () => previewStudentDashboard(...a),
    addStudentManually: () => addStudentManually(...a),
    updateStudentInfo: () => updateStudentInfo(...a),
    previewRosterFile: () => previewRosterFile(...a),
    previewRegistrarRosterMetadata: () => previewRegistrarRosterMetadata(...a),
    confirmRosterFileImport: () => confirmRosterFileImport(...a),
    confirmRosterFileImportBatch: () => confirmRosterFileImportBatch(...a),
    applyRegistrarRosterMetadata: () => applyRegistrarRosterMetadata(...a),
    assignRegistrarStudentsToSection: () => assignRegistrarStudentsToSection(...a),
    exportCourseDataPackage: () => exportCourseDataPackage(...a),
    importCourseTemplatePackage: () => importCourseTemplatePackage(...a),
    previewTqf3DocxImport: () => previewTqf3DocxImport(...a),
    applyTqf3DocxImport: () => applyTqf3DocxImport(...a),
    previewTeamAssignments: () => previewTeamAssignments(...a),
    confirmTeamAssignments: () => confirmTeamAssignments(...a),
    regenerateStudentAccessLink: () => regenerateStudentAccessLink(...a),
    repairActivityVisibility: () => repairActivityVisibility(...a),
    setAttendancePublished: () => setAttendancePublished(...a),
    deleteAttendanceSession: () => deleteAttendanceSession(...a),
    saveAttendanceOverride: () => saveAttendanceOverride(...a),
    clearAttendanceRecord: () => clearAttendanceRecord(...a),
    sendStudentAccessEmail: () => sendStudentAccessEmail(...a),
    getStudentAccessEmailStatus: () => getStudentAccessEmailStatus(...a),
    sendAllStudentAccessEmails: () => sendAllStudentAccessEmails(...a),
    sendStudentAccessEmailBatch: () => sendStudentAccessEmailBatch(...a),
    previewUploadedSelfEvaluationRows: () => previewUploadedSelfEvaluationRows(...a),
    previewGoogleSelfEvaluationImport: () => previewGoogleSelfEvaluationImport(...a),
    applySelfEvaluationImport: () => applySelfEvaluationImport(...a),
    getSelfEvaluationFormStatus: () => getSelfEvaluationFormStatus(...a),
    createSelfEvaluationForm: () => createSelfEvaluationForm(...a)
  };

  if (!actions[action]) throw new Error('Unknown or disallowed staff action.');
  try {
    return actions[action]();
  } finally {
    REQUEST_STAFF_AUTH_ = null;
  }
}


function ensureStudentAuthSchema_() {
  const db = dbOrNull_();
  if (!db) return;
  const sh = ensureSheet_(db, APP.tables.STUDENTS, APP.headers[APP.tables.STUDENTS]);
  const wanted = APP.headers[APP.tables.STUDENTS];
  const current = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
  wanted.forEach(h => {
    if (!current.includes(h)) {
      sh.getRange(1,sh.getLastColumn()+1).setValue(h);
      current.push(h);
    }
  });
}

function newStudentToken_() {
  return Utilities.base64EncodeWebSafe(
    Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()
  ).replace(/=+$/,'');
}

function hashStudentToken_(token) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token || ''),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
}


/* ---------- Fast attendance / burst-traffic cache ---------- */

function attendanceCache_(){
  return CacheService.getScriptCache();
}

function attendanceSessionCacheKey_(sessionId){
  return 'LPT:ATT:SESSION:'+String(sessionId||'');
}

function studentEmailCacheKey_(email){
  return 'LPT:STUDENT:EMAIL:'+normalizeEmail_(email);
}

function studentTokenCacheKey_(hash){
  return 'LPT:STUDENT:TOKEN:'+String(hash||'');
}

function attendanceRecordCacheKey_(sessionId,studentKey){
  return 'LPT:ATT:REC:'+String(sessionId||'')+':'+String(studentKey||'');
}

function cacheStudentForAttendance_(student){
  if(!student)return;
  const slim={
    StudentKey:String(student.StudentKey||''),
    StudentId:String(student.StudentId||''),
    Name:String(student.Name||''),
    Email:String(student.Email||''),
    Alias:String(student.Alias||''),
    Active:bool_(student.Active),
    StudentTokenHash:String(student.StudentTokenHash||'')
  };
  const cache=attendanceCache_();
  const payload=JSON.stringify(slim);
  if(slim.Email)cache.put(studentEmailCacheKey_(slim.Email),payload,300);
  if(slim.StudentTokenHash)cache.put(studentTokenCacheKey_(slim.StudentTokenHash),payload,300);
}

function invalidateStudentAttendanceCache_(student){
  if(!student)return;
  const cache=attendanceCache_();
  if(student.Email)cache.remove(studentEmailCacheKey_(student.Email));
  if(student.StudentTokenHash)cache.remove(studentTokenCacheKey_(student.StudentTokenHash));
}

function getStudentByEmailFast_(email){
  const normalized=normalizeEmail_(email);
  if(!normalized)return null;
  const cache=attendanceCache_();
  const cached=cache.get(studentEmailCacheKey_(normalized));
  if(cached){
    try{
      const student=JSON.parse(cached);
      return student&&student.Active?student:null;
    }catch(e){}
  }
  const student=getStudentByEmail_(normalized);
  if(student&&bool_(student.Active))cacheStudentForAttendance_(student);
  return student;
}

function getStudentByTokenFast_(token){
  if(!token)return null;
  const hash=hashStudentToken_(token);
  const cache=attendanceCache_();
  const cached=cache.get(studentTokenCacheKey_(hash));
  if(cached){
    try{
      const student=JSON.parse(cached);
      return student&&student.Active?student:null;
    }catch(e){}
  }
  const student=getStudentByToken_(token);
  if(student&&bool_(student.Active))cacheStudentForAttendance_(student);
  return student;
}

function resolveStudentFast_(studentToken){
  const email=currentEmail_(false);
  if(email){
    const byEmail=getStudentByEmailFast_(email);
    if(byEmail&&bool_(byEmail.Active))return byEmail;
  }
  const byToken=getStudentByTokenFast_(studentToken);
  if(byToken&&bool_(byToken.Active))return byToken;
  return null;
}

function cacheAttendanceSession_(session){
  if(!session||!session.SessionID)return;
  const slim=stripRow_(session);
  attendanceCache_().put(attendanceSessionCacheKey_(session.SessionID),JSON.stringify(slim),120);
}

function invalidateAttendanceSessionCache_(sessionId){
  attendanceCache_().remove(attendanceSessionCacheKey_(sessionId));
}

function attendanceSessionFast_(sessionId){
  const cache=attendanceCache_();
  const key=attendanceSessionCacheKey_(sessionId);
  const cached=cache.get(key);
  if(cached){
    try{
      const session=JSON.parse(cached);
      if(session&&session.SessionID)return session;
    }catch(e){}
  }
  const session=attendanceSession_(sessionId);
  if(session)cacheAttendanceSession_(session);
  return session;
}

function cacheAttendanceResult_(sessionId,studentKey,result){
  if(!sessionId||!studentKey||!result)return;
  attendanceCache_().put(
    attendanceRecordCacheKey_(sessionId,studentKey),
    JSON.stringify(result),
    300
  );
}

function getCachedAttendanceResult_(sessionId,studentKey){
  const raw=attendanceCache_().get(attendanceRecordCacheKey_(sessionId,studentKey));
  if(!raw)return null;
  try{return JSON.parse(raw);}catch(e){return null;}
}

function invalidateAttendanceResultCache_(sessionId,studentKey){
  attendanceCache_().remove(attendanceRecordCacheKey_(sessionId,studentKey));
}

function getStudentByToken_(token) {
  if (!token) return null;
  ensureStudentAuthSchema_();
  const hash = hashStudentToken_(token);
  return rowsToObjects_(table_(APP.tables.STUDENTS)).find(r =>
    bool_(r.Active) && String(r.StudentTokenHash || '') === hash
  ) || null;
}

function resolveStudent_(studentToken) {
  return resolveStudentFast_(studentToken);
}

function createStudentAccessLink_(studentKey, auditAction) {
  ensureStudentAuthSchema_();

  const sh = table_(APP.tables.STUDENTS);
  const found = findObject_(sh,'StudentKey',studentKey);
  if (!found || !bool_(found.Active)) throw new Error('Student not found or inactive.');

  const raw = newStudentToken_();
  const updated = stripRow_(found);
  updated.StudentTokenHash = hashStudentToken_(raw);
  updated.TokenIssuedAt = new Date();
  updateObjectRow_(sh,found._row,updated);
  invalidateStudentAttendanceCache_(found);
  cacheStudentForAttendance_(updated);

  if(auditAction) audit_(auditAction,'Student',studentKey,{email:found.Email});

  return {
    studentKey:found.StudentKey,
    studentId:found.StudentId,
    name:found.Name,
    email:found.Email,
    alias:found.Alias,
    url:webAppBaseUrl_()
      + '?page=student&studentToken=' + encodeURIComponent(raw)
      + '#studentToken=' + encodeURIComponent(raw)
  };
}

function regenerateStudentAccessLink(studentKey) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  return createStudentAccessLink_(studentKey,'REGENERATE_STUDENT_ACCESS');
}

function studentAccessEmailContent_(student, link) {
  const course=courseConfig_();
  const courseLabel=[course.code,course.name].filter(Boolean).join(' · ');
  const subject=`${course.code || 'Course'} Learning Progress Tracker access`;
  const body=[
    `Hello ${student.Name || 'Student'},`,
    '',
    `You can access your ${courseLabel || 'course'} Learning Progress Tracker using the secure link below:`,
    '',
    link,
    '',
    `Your student codename is: ${student.Alias || ''}`,
    '',
    'This link is personal to you. Please do not share it with anyone else.',
    'Open it once in the browser you normally use for class. The same secure access can also be used when checking attendance if Google identity is unavailable.',
    '',
    'If a new access link is issued later, the previous link will stop working.'
  ].join('\n');
  const html=`<p>Hello ${htmlEscape_(student.Name || 'Student')},</p>`+
    `<p>You can access your <strong>${htmlEscape_(courseLabel || 'course')}</strong> Learning Progress Tracker using the secure link below:</p>`+
    `<p><a href="${htmlEscape_(link)}">Open Student Progress Tracker</a></p>`+
    `<p>Your student codename is: <strong>${htmlEscape_(student.Alias || '')}</strong></p>`+
    `<p>This link is personal to you. Please do not share it with anyone else.</p>`+
    `<p>Open it once in the browser you normally use for class. The same secure access can also be used when checking attendance if Google identity is unavailable.</p>`+
    `<p class="small">If a new access link is issued later, the previous link will stop working.</p>`;
  return {subject:subject,body:body,htmlBody:html};
}

function htmlEscape_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function sendStudentAccessEmail(studentKey) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR']);
  const remaining=MailApp.getRemainingDailyQuota();
  if(remaining < 1) throw new Error('No Apps Script email quota remains today.');

  const link=createStudentAccessLink_(studentKey,'EMAIL_STUDENT_ACCESS_LINK');
  const content=studentAccessEmailContent_({Name:link.name,Alias:link.alias},link.url);
  MailApp.sendEmail({to:link.email,subject:content.subject,body:content.body,htmlBody:content.htmlBody,name:APP.name});
  audit_('SEND_STUDENT_ACCESS_EMAIL','Student',studentKey,{email:link.email,actor:staff.email});
  return {ok:true,email:link.email,alias:link.alias,remainingQuota:MailApp.getRemainingDailyQuota()};
}

function getStudentAccessEmailStatus() {
  requireStaff_(['OWNER','INSTRUCTOR']);
  const students=listStudentsForStaff_().filter(s=>s.Active && s.Email);
  return {activeStudents:students.length,remainingQuota:MailApp.getRemainingDailyQuota()};
}


function sendStudentAccessEmailBatch(studentKeys) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR']);
  const keys=Array.isArray(studentKeys)?studentKeys.slice(0,25).map(String):[];
  if(!keys.length)return {ok:true,sent:0,remainingQuota:MailApp.getRemainingDailyQuota()};

  const byKey={};
  listStudentsForStaff_().filter(s=>s.Active&&s.Email).forEach(s=>byKey[String(s.StudentKey)]=s);
  const students=keys.map(k=>byKey[k]).filter(Boolean);
  const remaining=MailApp.getRemainingDailyQuota();

  if(remaining<students.length){
    throw new Error(`Not enough Apps Script email quota for this batch. Need ${students.length}, remaining today: ${remaining}.`);
  }

  let sent=0;
  students.forEach(st=>{
    const link=createStudentAccessLink_(st.StudentKey,'BULK_EMAIL_STUDENT_ACCESS_LINK');
    const content=studentAccessEmailContent_({Name:st.Name,Alias:st.Alias},link.url);
    MailApp.sendEmail({
      to:st.Email,
      subject:content.subject,
      body:content.body,
      htmlBody:content.htmlBody,
      name:APP.name
    });
    sent++;
  });

  audit_('SEND_STUDENT_ACCESS_EMAIL_BATCH','Roster','',{
    count:sent,actor:staff.email
  });

  return {ok:true,sent,remainingQuota:MailApp.getRemainingDailyQuota()};
}

function sendAllStudentAccessEmails() {
  const staff=requireStaff_(['OWNER','INSTRUCTOR']);
  const students=listStudentsForStaff_().filter(s=>s.Active && s.Email);
  const remaining=MailApp.getRemainingDailyQuota();

  if(!students.length) throw new Error('There are no active students with email addresses.');
  if(remaining < students.length) {
    throw new Error(`Not enough Apps Script email quota to send all links safely. Need ${students.length}, remaining today: ${remaining}. No links were regenerated.`);
  }

  const sent=[];
  students.forEach(st=>{
    const link=createStudentAccessLink_(st.StudentKey,'BULK_EMAIL_STUDENT_ACCESS_LINK');
    const content=studentAccessEmailContent_({Name:st.Name,Alias:st.Alias},link.url);
    MailApp.sendEmail({to:st.Email,subject:content.subject,body:content.body,htmlBody:content.htmlBody,name:APP.name});
    sent.push({studentKey:st.StudentKey,email:st.Email,alias:st.Alias});
  });

  audit_('SEND_ALL_STUDENT_ACCESS_EMAILS','Roster','',{count:sent.length,actor:staff.email});
  return {ok:true,sent:sent.length,remainingQuota:MailApp.getRemainingDailyQuota()};
}

function getBootstrap(params) {
  const initStatus = projectInitializationStatus_();
  const email = currentEmail_(false);
  const effectiveEmail = Session.getEffectiveUser().getEmail() || '';

  if (!initStatus.initialized) {
    return {
      app: appPublicConfig_(),
      state: initStatus.canInitialize ? 'NEEDS_PROJECT_INITIALIZATION' : 'NEEDS_PROJECT_SETUP',
      email: email || effectiveEmail,
      spreadsheetUrl: initStatus.spreadsheet ? initStatus.spreadsheet.getUrl() : '',
      message: initStatus.canInitialize
        ? 'This new tracker needs one-time initialization before Course Setup.'
        : 'The bound Google Sheet could not be detected automatically. Run setupProject() once from the Apps Script editor.'
    };
  }

  const db = initStatus.spreadsheet || dbOrNull_();
  ensureCurrentSchema_();
  const setupComplete = ensureCourseSetupFlag_();
  if (!setupComplete) {
    return {
      app: appPublicConfig_(),
      state: 'NEEDS_COURSE_SETUP',
      email: email || effectiveEmail,
      spreadsheetUrl: db.getUrl()
    };
  }

  const course = courseConfig_();
  let fastStaffData=null;
  const staffData=()=>fastStaffData||(fastStaffData=getFastStaffBootstrapData_());

  const suppliedStaffToken = params ? String(params.staffToken || '') : '';
  if (suppliedStaffToken) {
    const tokenStaff = getStaffByToken_(suppliedStaffToken);
    if (!tokenStaff) {
      return {
        app: appPublicConfig_(),
        state: 'ACCESS_DENIED',
        course,
        message: 'This staff invitation link is invalid, deactivated, or has been replaced.'
      };
    }
    const suppliedStaffSession = params ? String(params.staffSession || '') : '';
    try {
      const verifiedIdentity = assertStaffInviteIdentity_(tokenStaff, suppliedStaffSession);
      return {
        app: appPublicConfig_(),
        state: 'READY',
        role: tokenStaff.Role,
        course,
        email: tokenStaff.Email,
        authMode: 'STAFF_INVITE',
        staffAuthMethod: verifiedIdentity.via || 'STAFF_INVITE',
        data: staffData()
      };
    } catch (err) {
      const message = String(err && err.message || err);
      if (message.indexOf('EMAIL_VERIFICATION_REQUIRED:') === 0) {
        return {
          app: appPublicConfig_(),
          state: 'STAFF_EMAIL_VERIFICATION_REQUIRED',
          course,
          expectedEmail: tokenStaff.Email,
          role: tokenStaff.Role,
          authMode: 'STAFF_INVITE',
          message: 'Verify access using the code sent to your registered staff email.'
        };
      }
      return {
        app: appPublicConfig_(),
        state: 'ACCESS_DENIED',
        course,
        email: currentEmail_(false) || '',
        expectedEmail: tokenStaff.Email,
        authMode: 'STAFF_INVITE',
        message: message
      };
    }
  }

  const explicitStudentRoute=!!(
    params && (
      (params.checkin && params.token) ||
      String(params.page||'').toLowerCase()==='student'
    )
  );

  // A real signed-in staff identity takes priority over a remembered student
  // token unless the user explicitly opened a student/check-in route.
  const identityStaff=email?getStaff_(email):null;
  if(!explicitStudentRoute && identityStaff && identityStaff.Active){
    return {
      app:appPublicConfig_(),
      state:'READY',
      role:identityStaff.Role,
      course,
      email,
      authMode:'GOOGLE_IDENTITY',
      data:staffData()
    };
  }

  const suppliedStudentToken = params ? String(params.studentToken || '').trim() : '';
  const tokenStudent = suppliedStudentToken ? getStudentByToken_(suppliedStudentToken) : null;

  // If a Student Access Link supplied a token but it does not resolve, report
  // that explicitly instead of falling through to the generic identity error.
  if (suppliedStudentToken && !tokenStudent) {
    return {
      app: appPublicConfig_(),
      state: 'ACCESS_DENIED',
      course,
      role: 'STUDENT',
      authMode: 'STUDENT_LINK',
      message: 'This Student Access Link is invalid or has been replaced by a newer link. Ask your instructor for the latest Student Access Link.'
    };
  }

  // Student access-link fallback for accounts whose identity is not exposed by
  // Session.getActiveUser(), or for an explicit student/check-in route.
  if (tokenStudent) {
    if (params && params.checkin && params.token) {
      return {
        app: appPublicConfig_(),
        state: 'CHECKIN',
        role: 'STUDENT',
        course,
        authMode: 'STUDENT_LINK',
        student: {alias: tokenStudent.Alias},
        checkin: {sessionId:String(params.checkin), token:String(params.token)}
      };
    }
    return {
      app: appPublicConfig_(),
      state: 'READY',
      role: 'STUDENT',
      course,
      email: tokenStudent.Email || '',
      authMode: 'STUDENT_LINK'
    };
  }

  if (!email) {
    return {
      app: appPublicConfig_(),
      state: 'IDENTITY_UNAVAILABLE',
      course,
      message: 'Student identity is not available from this Apps Script deployment. Ask your instructor for your secure Student Access Link.'
    };
  }

  const staff = identityStaff || getStaff_(email);
  const student = getStudentByEmail_(email);

  if (params && params.checkin && params.token) {
    if (!student) return {app: appPublicConfig_(), state: 'NOT_ENROLLED', course, message: 'This Google account is not enrolled in this course.'};
    return {
      app: appPublicConfig_(),
      state: 'CHECKIN',
      role: 'STUDENT',
      course,
      student: {alias: student.Alias},
      checkin: {sessionId: String(params.checkin), token: String(params.token)}
    };
  }

  // Optional explicit student route. Identity is still verified server-side.
  // A user can only open the student portal if their signed-in email exists in _Students.
  if (params && String(params.page || '').toLowerCase() === 'student') {
    if (!student || !student.Active) {
      return {
        app: appPublicConfig_(),
        state: 'NOT_ENROLLED',
        course,
        message: 'This Google account is not enrolled in this course.'
      };
    }
    return {
      app: appPublicConfig_(),
      state: 'READY',
      role: 'STUDENT',
      course
    };
  }

  if (staff && staff.Active) {
    return {
      app: appPublicConfig_(),
      state: 'READY',
      role: staff.Role,
      course,
      email,
      data: staffData()
    };
  }

  if (student && student.Active) {
    return {
      app: appPublicConfig_(),
      state: 'READY',
      role: 'STUDENT',
      course
    };
  }

  return {
    app: appPublicConfig_(),
    state: 'NO_ACCESS',
    course,
    message: 'Your Google account is not registered as staff or an enrolled student.'
  };
}

function getBootstrapSafe(params) {
  try {
    const result = getBootstrap(params || {});
    if (!result || typeof result !== 'object') {
      return {
        app: appPublicConfig_(),
        state: 'BOOTSTRAP_ERROR',
        message: 'The application bootstrap returned an empty response. Please reload the deployed web app.'
      };
    }

    // Force the response through JSON serialization before google.script.run
    // returns it. This prevents unsupported/undefined values from turning the
    // whole bootstrap response into null in the browser.
    return JSON.parse(JSON.stringify(result));
  } catch (err) {
    return {
      app: appPublicConfig_(),
      state: 'BOOTSTRAP_ERROR',
      message: err && err.message ? String(err.message) : 'Unable to initialize the Learning Progress Tracker.'
    };
  }
}

function saveCourseSetup(payload) {
  const db = requireDb_();
  const email = currentEmail_(true);
  if (!email) throw new Error('Unable to determine your Google account.');
  upsertStaff_(email, 'OWNER');

  const courseCode = clean_(payload.courseCode, 50);
  const courseName = clean_(payload.courseName, 200);
  const term = clean_(payload.term, 80);
  if (!courseCode || !courseName || !term) throw new Error('Course code, course name, and term are required.');

  setConfig_('COURSE_CODE', courseCode);
  setConfig_('COURSE_NAME', courseName);
  setConfig_('TERM', term);
  setConfig_('ALIAS_PREFIX', clean_(payload.aliasPrefix, 30) || courseCode);
  setConfig_('DEFAULT_ROTATE_SECONDS', String(clampNum_(payload.rotateSeconds, 30, 300, 60)));
  setConfig_('DEFAULT_CHECKIN_MINUTES', String(clampNum_(payload.checkinMinutes, 1, 60, 5)));
  setConfig_('DEFAULT_RADIUS_M', String(clampNum_(payload.radiusM, 10, 2000, 150)));
  setConfig_('DEFAULT_ROOM_LAT', finiteOrBlank_(payload.roomLat));
  setConfig_('DEFAULT_ROOM_LNG', finiteOrBlank_(payload.roomLng));

  const staffRows = Array.isArray(payload.staff) ? payload.staff : [];
  staffRows.forEach(s => {
    const e = normalizeEmail_(s.email);
    if (e) upsertStaff_(e, s.role === 'TA' ? 'TA' : 'INSTRUCTOR');
  });

  setConfig_('SETUP_COMPLETE', 'TRUE');

  const trackerName = renameTrackerSpreadsheet_();
  applySystemSheetVisibility_();

  audit_('COURSE_SETUP', 'Course', courseCode, {term, trackerName});
  return {
    ok:true,
    course:courseConfig_(),
    spreadsheetUrl:db.getUrl(),
    trackerName
  };
}


function listWeeklyTeachingPlan_(){
  const activityMap={};
  listActivities_().forEach(a=>activityMap[String(a.ActivityID)]=a);
  const cloMap={};
  listCLOs_().forEach(c=>cloMap[String(c.CLOID)]=c);
  const cloRowsByPlan={};
  weeklyTeachingPlanCloRows_().forEach(r=>{
    const key=String(r.PlanItemID||'');
    if(!cloRowsByPlan[key])cloRowsByPlan[key]=[];
    cloRowsByPlan[key].push(r);
  });

  return rowsToObjects_(table_(APP.tables.WEEKLY_PLAN))
    .filter(r=>bool_(r.Active))
    .sort((a,b)=>Number(a.WeekNo||0)-Number(b.WeekNo||0))
    .map(r=>{
      const planItemId=String(r.PlanItemID||'');
      const legacyCloIds=String(r.PlannedCLOIDs||'').split('|').filter(Boolean);
      const activityIds=String(r.PlannedActivityIDs||'').split('|').filter(Boolean);
      let cloRows=(cloRowsByPlan[planItemId]||[]).slice().sort((a,b)=>Number(a.Sequence||0)-Number(b.Sequence||0));
      if(!cloRows.length && legacyCloIds.length){
        cloRows=legacyCloIds.map((cloId,index)=>({CLOID:cloId,TeachingMethod:'',AssessmentMethod:'',ActivityIDs:'',Sequence:index+1}));
      }
      const cloPlans=cloRows.map(cp=>{
        const cloId=String(cp.CLOID||'');
        const cpActivityIds=String(cp.ActivityIDs||'').split('|').filter(Boolean);
        const c=cloMap[cloId]||{};
        return {
          planCloId:String(cp.PlanCLOID||''),
          cloId:cloId,
          code:String(c.Code||cloId),
          description:String(c.Description||''),
          teachingMethod:String(cp.TeachingMethod||''),
          assessmentMethod:String(cp.AssessmentMethod||''),
          activityIds:cpActivityIds,
          activities:cpActivityIds.map(id=>activityMap[id]).filter(Boolean).map(a=>({activityId:String(a.ActivityID),title:String(a.Title||'')}))
        };
      });
      const cloIds=cloPlans.map(cp=>cp.cloId);
      return {
        planItemId:planItemId,
        weekNo:Number(r.WeekNo||0),
        plannedDate:dateOnly_(r.PlannedDate),
        topic:String(r.Topic||''),
        description:String(r.Description||''),
        plannedCloIds:cloIds,
        plannedClos:cloIds.map(id=>cloMap[id]).filter(Boolean).map(c=>({cloId:String(c.CLOID),code:String(c.Code||c.CLOID)})),
        cloPlans:cloPlans,
        plannedActivityIds:activityIds,
        plannedActivities:activityIds.map(id=>activityMap[id]).filter(Boolean).map(a=>({activityId:String(a.ActivityID),title:String(a.Title||'') })),
        actualWeekNo:r.ActualWeekNo===''?null:Number(r.ActualWeekNo),
        actualDate:dateOnly_(r.ActualDate),
        actualStatus:String(r.ActualStatus||'Planned'),
        actualNote:String(r.ActualNote||''),
        trackChanges:r.TrackChanges===''?true:bool_(r.TrackChanges)
      };
    });
}

function getWeeklyTeachingPlan(){
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  return {
    items:listWeeklyTeachingPlan_(),
    clos:listCLOs_().map(c=>({cloId:String(c.CLOID),code:String(c.Code||c.CLOID),description:String(c.Description||'')})),
    activities:listActivities_().filter(a=>a.Active!==false).map(a=>({
      activityId:String(a.ActivityID),
      title:String(a.Title||''),
      mode:String(a.Mode||'Individual'),
      published:bool_(a.Published),
      assignedWeek:a.AssignedWeek===''||a.AssignedWeek==null?'':Number(a.AssignedWeek),
      dueDate:a.DueDate instanceof Date ? Utilities.formatDate(a.DueDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm") : String(a.DueDate||'')
    }))
  };
}

function saveWeeklyTeachingPlan(items){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  const sh=table_(APP.tables.WEEKLY_PLAN);
  const cloPlanSh=table_(APP.tables.WEEKLY_PLAN_CLOS);
  const planObjects=[];
  const cloPlanObjects=[];

  (Array.isArray(items)?items:[]).slice(0,60).forEach((item,index)=>{
    const weekNo=Number(item.weekNo);
    if(!Number.isFinite(weekNo)||weekNo<1)return;
    const planItemId=clean_(item.planItemId,160)||('PLAN-'+Utilities.getUuid());
    let cloPlans=Array.isArray(item.cloPlans)?item.cloPlans:[];
    if(!cloPlans.length && Array.isArray(item.plannedCloIds)){
      cloPlans=item.plannedCloIds.map(cloId=>({cloId:cloId,teachingMethod:'',assessmentMethod:'',activityIds:[]}));
    }
    const seen={};
    cloPlans=cloPlans.filter(cp=>{
      const id=String((cp||{}).cloId||'').trim();
      if(!id||seen[id])return false; seen[id]=true; return true;
    });
    const cloIds=cloPlans.map(cp=>String(cp.cloId));
    const now=new Date();
    planObjects.push({
      PlanItemID:planItemId,
      WeekNo:weekNo,
      PlannedDate:clean_(item.plannedDate,80),
      Topic:clean_(item.topic,300),
      Description:clean_(item.description,3000),
      PlannedCLOIDs:cloIds.join('|'),
      PlannedActivityIDs:(Array.isArray(item.plannedActivityIds)?item.plannedActivityIds:[]).map(String).join('|'),
      ActualWeekNo:item.actualWeekNo==null?'':Number(item.actualWeekNo),
      ActualDate:clean_(item.actualDate,80),
      ActualStatus:clean_(item.actualStatus,80)||'Planned',
      ActualNote:clean_(item.actualNote,3000),
      TrackChanges:item.trackChanges===false?'FALSE':'TRUE',
      Active:true,
      CreatedAt:now,
      UpdatedAt:now
    });
    cloPlans.forEach((cp,cpIndex)=>{
      cloPlanObjects.push({
        PlanCLOID:clean_(cp.planCloId,160)||('PLAN-CLO-'+Utilities.getUuid()),
        PlanItemID:planItemId,
        CLOID:String(cp.cloId),
        TeachingMethod:clean_(cp.teachingMethod,3000),
        AssessmentMethod:clean_(cp.assessmentMethod,3000),
        ActivityIDs:(Array.isArray(cp.activityIds)?cp.activityIds:[]).map(String).join('|'),
        Sequence:cpIndex+1,
        CreatedAt:now,
        UpdatedAt:now
      });
    });
  });

  rewriteObjects_(sh,planObjects);
  rewriteObjects_(cloPlanSh,cloPlanObjects);
  resetRequestDataCache_();
  audit_('SAVE_WEEKLY_TEACHING_PLAN','Course','',{actor:actor.email,count:planObjects.length});
  return getWeeklyTeachingPlan();
}

function updateTeachingPlanExecution(planItemId,payload){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  payload=payload||{};
  const sh=table_(APP.tables.WEEKLY_PLAN);
  const row=rowsToObjects_(sh).find(r=>String(r.PlanItemID)===String(planItemId));
  if(!row)throw new Error('Teaching plan item not found.');

  const previousWeek=row.ActualWeekNo===''?null:Number(row.ActualWeekNo);
  const previousStatus=String(row.ActualStatus||'Planned');
  const newWeek=payload.actualWeekNo==null||payload.actualWeekNo===''?null:Number(payload.actualWeekNo);
  const newStatus=clean_(payload.actualStatus,80)||'Planned';
  const note=clean_(payload.actualNote,3000);
  const recordReason=bool_(payload.recordReason);
  const reason=recordReason?clean_(payload.reason,2000):'';

  updateRowObject_(sh,row._row,{
    ActualWeekNo:newWeek==null?'':newWeek,
    ActualDate:clean_(payload.actualDate,80),
    ActualStatus:newStatus,
    ActualNote:note,
    TrackChanges:payload.trackChanges===false?'FALSE':'TRUE',
    UpdatedAt:new Date()
  });

  if(payload.trackChanges!==false && (previousWeek!==newWeek || previousStatus!==newStatus)){
    appendObject_(table_(APP.tables.WEEKLY_PLAN_HISTORY),{
      HistoryID:'HIST-'+Utilities.getUuid(),
      PlanItemID:planItemId,
      ChangedAt:new Date(),
      ChangedBy:actor.email,
      PreviousActualWeekNo:previousWeek==null?'':previousWeek,
      NewActualWeekNo:newWeek==null?'':newWeek,
      PreviousStatus:previousStatus,
      NewStatus:newStatus,
      Reason:reason,
      RecordedReason:recordReason?'TRUE':'FALSE'
    });
  }

  resetRequestDataCache_();
  return getWeeklyTeachingPlan();
}

function getTeachingPlanHistory(planItemId){
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  return rowsToObjects_(table_(APP.tables.WEEKLY_PLAN_HISTORY))
    .filter(r=>String(r.PlanItemID)===String(planItemId))
    .sort((a,b)=>new Date(b.ChangedAt).getTime()-new Date(a.ChangedAt).getTime())
    .map(r=>({
      changedAt:String(r.ChangedAt||''),
      changedBy:String(r.ChangedBy||''),
      previousActualWeekNo:r.PreviousActualWeekNo===''?null:Number(r.PreviousActualWeekNo),
      newActualWeekNo:r.NewActualWeekNo===''?null:Number(r.NewActualWeekNo),
      previousStatus:String(r.PreviousStatus||''),
      newStatus:String(r.NewStatus||''),
      reason:String(r.Reason||''),
      recordedReason:bool_(r.RecordedReason)
    }));
}

function getSystemStatus(){
  requireStaff_(['OWNER','INSTRUCTOR']);
  const db=requireDb_();
  const expected=Number(APP.schemaVersion),actual=storedSchemaVersion_();
  const required=[
    APP.tables.CONFIG,APP.tables.STAFF,APP.tables.STUDENTS,APP.tables.CLOS,
    APP.tables.ACTIVITIES,APP.tables.ACTIVITY_CLOS,
    APP.tables.ACTIVITY_TEAMS,APP.tables.ACTIVITY_TEAM_MEMBERS,
    APP.tables.RESULTS,APP.tables.CLO_EVIDENCE,
    APP.tables.ATTENDANCE_SESSIONS,APP.tables.ATTENDANCE
  ];
  const missing=required.filter(name=>!db.getSheetByName(name));
  return {appVersion:APP.version,schemaVersion:actual,expectedSchemaVersion:expected,
    databaseReady:actual===expected&&!missing.length,missingSheets:missing};
}

function getDiagnostics() {
  const db = dbOrNull_();
  const email = currentEmail_(false);
  return {
    appVersion: APP.version,
    hasDatabase: !!db,
    spreadsheetId: db ? db.getId() : '',
    spreadsheetName: db ? db.getName() : '',
    email: email,
    staff: email ? getStaff_(email) : null,
    setupComplete: db ? getConfig_('SETUP_COMPLETE') : ''
  };
}


function listCourseStaff() {
  requireStaff_(['OWNER','INSTRUCTOR']);
  ensureStaffAuthSchema_();
  return rowsToObjects_(table_(APP.tables.STAFF)).map(r => ({
    staffKey: String(r.StaffKey||r.Email||''),
    displayName: staffDisplayName_(r),
    email: r.Email,
    role: r.Role,
    active: bool_(r.Active),
    addedAt: dateIso_(r.AddedAt),
    inviteEnabled: !!r.StaffTokenHash,
    tokenIssuedAt: dateIso_(r.TokenIssuedAt)
  })).sort((a,b) => {
    const order={OWNER:0,INSTRUCTOR:1,TA:2};
    return (order[a.role]??9)-(order[b.role]??9) || a.email.localeCompare(b.email);
  });
}

function saveCourseStaff(payload) {
  const actor = requireStaff_(['OWNER','INSTRUCTOR']);
  const email = normalizeEmail_(payload.email);
  const role = String(payload.role || '').toUpperCase();

  if (!email) throw new Error('Staff email is required.');
  if (!['INSTRUCTOR','TA'].includes(role)) throw new Error('Role must be INSTRUCTOR or TA.');

  const existing = getStaff_(email);
  if (existing && existing.Role === 'OWNER') {
    throw new Error('The OWNER role cannot be changed from this screen.');
  }

  upsertStaff_(email, role);
  const invite = issueStaffInvite_(email);
  audit_('SAVE_COURSE_STAFF','Staff',email,{role,actor:actor.email,inviteIssued:true});
  return {ok:true, staff:listCourseStaff(), invite};
}


function sendStaffAccessEmailBatch(emails){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  const requested=Array.isArray(emails)?emails.slice(0,25).map(normalizeEmail_).filter(Boolean):[];
  if(!requested.length)return {ok:true,sent:0,failed:[],remainingQuota:MailApp.getRemainingDailyQuota()};

  const all=rowsToObjects_(table_(APP.tables.STAFF)).filter(r=>bool_(r.Active));
  const byEmail={};
  all.forEach(r=>byEmail[normalizeEmail_(r.Email)]=r);
  const staff=requested.map(e=>byEmail[e]).filter(Boolean);

  const remaining=MailApp.getRemainingDailyQuota();
  if(remaining<staff.length){
    throw new Error(`Not enough Apps Script email quota for this batch. Need ${staff.length}, remaining today: ${remaining}.`);
  }

  const course=courseConfig_();
  const courseLabel=[course.code,course.name].filter(Boolean).join(' · ');
  const sent=[],failed=[];

  staff.forEach(st=>{
    try{
      const invite=issueStaffInvite_(st.Email);
      const role=String(st.Role||'INSTRUCTOR');
      const subject=`${course.code||'Course'} Learning Progress Tracker staff access`;
      const body=[
        'Hello,','',
        `You have staff access to ${courseLabel||'the course'} Learning Progress Tracker as ${role}.`,'',
        'Open the secure link below:',
        invite.invitationLink || invite.url || '', '',
        'This link is personal to your staff account. Please do not share it.',
        'If a new secure access link is generated later, the previous link will stop working.'
      ].join('\n');
      const htmlBody=
        `<p>Hello,</p>`+
        `<p>You have staff access to <strong>${htmlEscape_(courseLabel||'the course')}</strong> Learning Progress Tracker as <strong>${htmlEscape_(role)}</strong>.</p>`+
        `<p><a href="${htmlEscape_(invite.invitationLink || invite.url || '')}">Open Staff Learning Progress Tracker</a></p>`+
        `<p>This link is personal to your staff account. Please do not share it.</p>`+
        `<p>If a new secure access link is generated later, the previous link will stop working.</p>`;

      MailApp.sendEmail({to:st.Email,subject,body,htmlBody,name:APP.name});
      sent.push(st.Email);
    }catch(err){
      failed.push({email:st.Email,error:String(err&&err.message||err)});
    }
  });

  audit_('SEND_STAFF_ACCESS_EMAIL_BATCH','Staff','',{
    count:sent.length,failed:failed.length,actor:actor.email
  });

  return {ok:true,sent:sent.length,failed,remainingQuota:MailApp.getRemainingDailyQuota()};
}

function regenerateStaffInvite(email) {
  const actor = requireStaff_(['OWNER','INSTRUCTOR']);
  email = normalizeEmail_(email);
  const existing = getStaff_(email);
  if (!existing) throw new Error('Staff account not found.');

  const invite = issueStaffInvite_(email);
  audit_('REGENERATE_STAFF_INVITE','Staff',email,{actor:actor.email});
  return {ok:true, staff:listCourseStaff(), invite};
}

function setCourseStaffActive(email, active) {
  const actor = requireStaff_(['OWNER','INSTRUCTOR']);
  email = normalizeEmail_(email);
  const sh = table_(APP.tables.STAFF);
  const row = findObject_(sh,'Email',email);
  if (!row) throw new Error('Staff account not found.');
  if (row.Role === 'OWNER') throw new Error('The course OWNER cannot be deactivated.');

  updateObjectRow_(sh,row._row,{
    Email:row.Email,
    Role:row.Role,
    Active:bool_(active),
    AddedAt:row.AddedAt || new Date(),
    StaffTokenHash:row.StaffTokenHash || '',
    TokenIssuedAt:row.TokenIssuedAt || ''
  });

  audit_('SET_COURSE_STAFF_ACTIVE','Staff',email,{active:bool_(active),actor:actor.email});
  return {ok:true, staff:listCourseStaff()};
}



function staffDisplayName_(row){
  return String(row.DisplayName||row.Name||row.Email||'').trim();
}

function listCourseStaffProfiles_(){
  return rowsToObjects_(table_(APP.tables.STAFF))
    .filter(r=>bool_(r.Active))
    .map(r=>({
      staffKey:String(r.StaffKey||r.Email||''),
      displayName:staffDisplayName_(r),
      email:String(r.Email||''),
      role:String(r.Role||'INSTRUCTOR')
    }))
    .sort((a,b)=>a.displayName.localeCompare(b.displayName));
}


function updateCourseStaffProfile(staffKey,payload){
  return updateCourseStaffProfile_(staffKey,payload);
}

function addCourseStaffProfile(payload){
  return addCourseStaffProfile_(payload);
}

function updateCourseStaffProfile_(staffKey,payload){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  payload=payload||{};
  const sh=table_(APP.tables.STAFF);
  const rows=rowsToObjects_(sh);
  const row=rows.find(r=>String(r.StaffKey||r.Email)===String(staffKey));
  if(!row)throw new Error('Staff profile not found.');
  const isOwner=String(row.Role||'').toUpperCase()==='OWNER';
  const displayName=clean_(payload.displayName,250);
  const oldEmail=normalizeEmail_(row.Email||'');
  // The OWNER may edit the display name, but identity and role stay fixed.
  const email=isOwner?oldEmail:normalizeEmail_(payload.email||'');
  const role=isOwner?'OWNER':String(payload.role||row.Role||'INSTRUCTOR').toUpperCase();

  if(!displayName)throw new Error('Enter the staff member name.');
  if(!email)throw new Error('Staff email is required.');
  if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email))throw new Error('Invalid email.');
  if(!isOwner&&!['INSTRUCTOR','TA'].includes(role))throw new Error('Role must be INSTRUCTOR or TA.');

  const duplicate=rows.find(r=>r._row!==row._row && normalizeEmail_(r.Email||'')===email);
  if(duplicate)throw new Error('This email is already used by another course staff member.');

  updateRowObject_(sh,row._row,{
    ...row,
    DisplayName:displayName,
    Email:email,
    Role:role,
    UpdatedAt:new Date()
  });

  // Older schema-8 courses may use the staff email itself as InstructorStaffKey.
  // Keep section references valid when such a staff email is edited.
  if(oldEmail && oldEmail!==email){
    const secSh=table_(APP.tables.COURSE_SECTIONS);
    const sections=rowsToObjects_(secSh);
    sections.filter(sec=>String(sec.InstructorStaffKey||'')===oldEmail || normalizeEmail_(sec.InstructorEmail||'')===oldEmail)
      .forEach(sec=>updateRowObject_(secSh,sec._row,{
        ...sec,
        InstructorStaffKey:String(row.StaffKey||email),
        InstructorName:displayName,
        InstructorEmail:email,
        UpdatedAt:new Date()
      }));
  }else{
    const secSh=table_(APP.tables.COURSE_SECTIONS);
    const sections=rowsToObjects_(secSh);
    sections.filter(sec=>String(sec.InstructorStaffKey||'')===String(row.StaffKey||oldEmail))
      .forEach(sec=>updateRowObject_(secSh,sec._row,{
        ...sec,
        InstructorName:displayName,
        InstructorEmail:email,
        UpdatedAt:new Date()
      }));
  }

  resetRequestDataCache_();
  audit_('UPDATE_STAFF_PROFILE','Staff',staffKey,{actor:actor.email,displayName,email,role});
  return {ok:true,staff:listCourseStaff(),staffProfiles:listCourseStaffProfiles_()};
}

function addCourseStaffProfile_(payload){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  payload=payload||{};
  const displayName=clean_(payload.displayName,250);
  const email=normalizeEmail_(payload.email||'');
  const role=String(payload.role||'INSTRUCTOR').toUpperCase();

  if(!displayName)throw new Error('Enter the instructor name.');
  if(email&&!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){
    throw new Error('Invalid instructor email.');
  }

  const sh=table_(APP.tables.STAFF);
  const rows=rowsToObjects_(sh);
  const existing=email?rows.find(r=>normalizeEmail_(r.Email)===email):null;

  if(existing){
    updateRowObject_(sh,existing._row,{
      DisplayName:displayName,
      Role:role,
      Active:true
    });
    resetRequestDataCache_();
    return {
      staffKey:String(existing.StaffKey||existing.Email||email),
      displayName,
      email:String(existing.Email||email),
      role
    };
  }

  // Allow a profile-only instructor before login access/email is known.
  const staffKey='STAFF-'+Utilities.getUuid();
  const storedEmail=email||('profile-'+Utilities.getUuid()+'@local.invalid');
  appendObject_(sh,{
    StaffKey:staffKey,
    Email:storedEmail,
    DisplayName:displayName,
    Role:role,
    Active:true,
    CreatedAt:new Date(),
    UpdatedAt:new Date()
  });

  resetRequestDataCache_();
  audit_('ADD_STAFF_PROFILE','Staff',staffKey,{
    actor:actor.email,displayName,email:email||''
  });
  return {staffKey,displayName,email,role};
}

function listCourseObjectives_(){
  return rowsToObjects_(table_(APP.tables.COURSE_OBJECTIVES))
    .filter(r=>bool_(r.Active))
    .sort((a,b)=>Number(a.Sequence||0)-Number(b.Sequence||0))
    .map(r=>({
      objectiveId:String(r.ObjectiveID||''),
      sequence:Number(r.Sequence||0),
      text:String(r.ObjectiveText||'')
    }));
}

function courseSectionScheduleRows_(){
  ensureHeaders_(APP.tables.COURSE_SECTION_SCHEDULES);
  return rowsToObjects_(table_(APP.tables.COURSE_SECTION_SCHEDULES)).filter(r=>bool_(r.Active));
}
function sectionSchedulesBySectionId_(){
  const out={};
  courseSectionScheduleRows_().forEach(r=>{
    const sectionId=String(r.SectionID||'');
    if(!out[sectionId])out[sectionId]=[];
    out[sectionId].push({
      scheduleId:String(r.ScheduleID||''),scheduleType:String(r.ScheduleType||'CLASS').toUpperCase()==='EXAM'?'EXAM':'CLASS',
      day:String(r.DayOfWeek||''),date:dateOnly_(r.ScheduleDate),startTime:scheduleTimeText_(r.StartTime),endTime:scheduleTimeText_(r.EndTime),room:legacyRoomAndExam_(r.Room).room,examType:String(r.ExamType||''),sequence:Number(r.Sequence||0)
    });
  });
  Object.keys(out).forEach(k=>out[k].sort((a,b)=>a.sequence-b.sequence));
  return out;
}
function listCourseSections_(){
  const scheduleBy=sectionSchedulesBySectionId_();
  return rowsToObjects_(table_(APP.tables.COURSE_SECTIONS))
    .filter(r=>bool_(r.Active))
    .sort((a,b)=>String(a.SectionCode||'').localeCompare(String(b.SectionCode||'')))
    .map(r=>{
      const sectionId=String(r.SectionID||'');
      const rows=scheduleBy[sectionId]||[];
      let schedules=rows.filter(x=>x.scheduleType==='CLASS');
      const exams=rows.filter(x=>x.scheduleType==='EXAM');
      if(!schedules.length){
        const legacyDays=String(r.Days||'').split('|').filter(Boolean);
        schedules=legacyDays.map((day,index)=>({scheduleId:'',scheduleType:'CLASS',day,startTime:String(r.StartTime||''),endTime:String(r.EndTime||''),room:String(r.Room||''),date:'',examType:'',sequence:index+1}));
      }
      return {
        sectionId,sectionCode:String(r.SectionCode||''),
        days:schedules.map(x=>x.day).filter(Boolean),startTime:String(r.StartTime||''),endTime:String(r.EndTime||''),dayTime:String(r.DayTime||''),room:String(r.Room||''),mode:String(r.Mode||''),
        instructorStaffKey:String(r.InstructorStaffKey||''),instructorName:String(r.InstructorName||''),instructorEmail:String(r.InstructorEmail||''),
        schedules:schedules,exams:exams
      };
    });
}

function replaceCourseObjectives_(items){
  const sh=table_(APP.tables.COURSE_OBJECTIVES);
  const existing=rowsToObjects_(sh);
  existing.slice().sort((a,b)=>b._row-a._row).forEach(r=>sh.deleteRow(r._row));
  (Array.isArray(items)?items:[]).forEach((item,index)=>{
    const text=clean_(item&&item.text!=null?item.text:item,1000);
    if(!text)return;
    appendObject_(sh,{
      ObjectiveID:'OBJ-'+Utilities.getUuid(),
      Sequence:index+1,
      ObjectiveText:text,
      Active:true,
      CreatedAt:new Date(),
      UpdatedAt:new Date()
    });
  });
}

function writeSectionCodeText_(sh,rowNum,sectionCode){
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const col=headers.indexOf('SectionCode')+1;
  if(col>0){
    const cell=sh.getRange(rowNum,col);
    cell.setNumberFormat('@');
    cell.setValue(String(sectionCode==null?'':sectionCode));
  }
}
function replaceSectionSchedules_(sectionId,schedules,exams){
  const sh=table_(APP.tables.COURSE_SECTION_SCHEDULES);
  rowsToObjects_(sh).filter(r=>String(r.SectionID||'')===String(sectionId)).sort((a,b)=>b._row-a._row).forEach(r=>sh.deleteRow(r._row));
  const now=new Date();let seq=1;
  (Array.isArray(schedules)?schedules:[]).forEach(x=>{
    const day=clean_(x&&x.day,30),startTime=clean_(x&&x.startTime,20),endTime=clean_(x&&x.endTime,20),room=clean_(x&&x.room,120);
    if(!day&&!startTime&&!endTime&&!room)return;
    appendObject_(sh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:sectionId,ScheduleType:'CLASS',DayOfWeek:day,ScheduleDate:'',StartTime:startTime,EndTime:endTime,Room:room,ExamType:'',Sequence:seq++,Active:true,CreatedAt:now,UpdatedAt:now});
  });
  (Array.isArray(exams)?exams:[]).forEach(x=>{
    const date=clean_(x&&x.date,30),startTime=clean_(x&&x.startTime,20),endTime=clean_(x&&x.endTime,20),room=clean_(x&&x.room,120),examType=clean_(x&&x.examType,40)||'Exam';
    if(!date&&!startTime&&!endTime&&!room)return;
    appendObject_(sh,{ScheduleID:'SCH-'+Utilities.getUuid(),SectionID:sectionId,ScheduleType:'EXAM',DayOfWeek:'',ScheduleDate:date,StartTime:startTime,EndTime:endTime,Room:room,ExamType:examType,Sequence:seq++,Active:true,CreatedAt:now,UpdatedAt:now});
  });
}
function replaceCourseSections_(items){
  const sh=table_(APP.tables.COURSE_SECTIONS);
  const staff=listCourseStaffProfiles_(),byKey={};staff.forEach(s=>byKey[String(s.staffKey)]=s);
  const existing=rowsToObjects_(sh);
  const active=existing.filter(r=>bool_(r.Active));
  const usedIds={};
  const canonical=v=>sectionCanonicalKey_(String(v==null?'':v));
  (Array.isArray(items)?items:[]).forEach(item=>{
    const sectionCode=clean_(String(item.sectionCode==null?'':item.sectionCode),80);
    let schedules=Array.isArray(item.schedules)?item.schedules:[];
    if(!schedules.length){
      const legacyDays=(Array.isArray(item.days)?item.days:[]).map(x=>clean_(x,20)).filter(Boolean);
      schedules=legacyDays.map(day=>({day,startTime:item.startTime||'',endTime:item.endTime||'',room:item.room||''}));
    }
    const exams=Array.isArray(item.exams)?item.exams:[];
    const mode=clean_(item.mode,40)||'Onsite',staffKey=clean_(item.instructorStaffKey,180),selected=staffKey?byKey[staffKey]:null;
    if(!sectionCode&&!schedules.length&&!exams.length)return;
    const days=[];schedules.forEach(x=>{const d=clean_(x&&x.day,20);if(d&&!days.includes(d))days.push(d);});
    const pairs=schedules.map(x=>[clean_(x&&x.startTime,20),clean_(x&&x.endTime,20)]).filter(x=>x[0]||x[1]);
    const rooms=[];schedules.forEach(x=>{const r=clean_(x&&x.room,120);if(r&&!rooms.includes(r))rooms.push(r);});
    const sameTime=pairs.length&&pairs.every(x=>x[0]===pairs[0][0]&&x[1]===pairs[0][1]);
    const startTime=sameTime?pairs[0][0]:'',endTime=sameTime?pairs[0][1]:'';
    const dayTime=schedules.map(x=>[clean_(x&&x.day,20),[clean_(x&&x.startTime,20),clean_(x&&x.endTime,20)].filter(Boolean).join('–')].filter(Boolean).join(' ')).join(', ');
    const requestedId=clean_(item.sectionId||item.SectionID,180);
    let old=requestedId?active.find(r=>String(r.SectionID||'')===requestedId):null;
    if(!old&&sectionCode)old=active.find(r=>canonical(r.SectionCode)===canonical(sectionCode)&&!usedIds[String(r.SectionID||'')]);
    const now=new Date();
    const sectionId=old?String(old.SectionID||''):'SEC-'+Utilities.getUuid();
    const payload={SectionID:sectionId,SectionCode:sectionCode,Days:days.join('|'),StartTime:startTime,EndTime:endTime,DayTime:dayTime,Room:rooms.join(', '),Mode:mode,InstructorStaffKey:staffKey,InstructorName:selected?selected.displayName:clean_(item.instructorName,250),InstructorEmail:selected?selected.email:normalizeEmail_(item.instructorEmail||''),Active:true,CreatedAt:old&&old.CreatedAt?old.CreatedAt:now,UpdatedAt:now};
    let rowNum;
    if(old){updateObjectRow_(sh,old._row,payload);rowNum=old._row;}else{appendObject_(sh,payload);rowNum=sh.getLastRow();}
    writeSectionCodeText_(sh,rowNum,sectionCode);
    usedIds[sectionId]=true;
    replaceSectionSchedules_(sectionId,schedules,exams);
  });
  // Never silently remove a Section that still has students assigned. This preserves
  // Student -> Section mappings when Course settings are edited and saved.
  const mappings=activeStudentSectionRows_();
  active.forEach(r=>{
    const id=String(r.SectionID||'');if(usedIds[id])return;
    if(mappings.some(m=>String(m.SectionID||'')===id&&bool_(m.Active))){
      throw new Error('Section '+String(r.SectionCode||'')+' still has students assigned. Move or unassign the students before removing this Section.');
    }
    updateObjectRow_(sh,r._row,{Active:false,UpdatedAt:new Date()});
  });
}

function courseProfileComplete_(cfg){
  const hasCode=!!String(cfg.COURSE_CODE_TH||cfg.COURSE_CODE_EN||cfg.COURSE_CODE||'').trim();
  const hasName=!!String(cfg.COURSE_NAME_TH||cfg.COURSE_NAME_EN||cfg.COURSE_NAME||'').trim();
  return hasCode&&hasName;
}

function getCourseSettings() {
  requireStaff_(['OWNER','INSTRUCTOR']);
  const cfg=configMapFast_();
  return {
    courseCodeTh:String(cfg.COURSE_CODE_TH||''),
    courseCodeEn:String(cfg.COURSE_CODE_EN||cfg.COURSE_CODE||''),
    courseNameTh:String(cfg.COURSE_NAME_TH||''),
    courseNameEn:String(cfg.COURSE_NAME_EN||cfg.COURSE_NAME||''),
    totalCredits:String(cfg.TOTAL_CREDITS||cfg.CREDITS||''),
    lectureHours:String(cfg.LECTURE_HOURS||''),
    practiceHours:String(cfg.PRACTICE_HOURS||''),
    selfStudyHours:String(cfg.SELF_STUDY_HOURS||''),
    semester:String(cfg.SEMESTER||''),
    academicYear:String(cfg.ACADEMIC_YEAR||''),
    courseType:String(cfg.COURSE_TYPE||''),
    courseSubtypes:String(cfg.COURSE_SUBTYPES||'').split('|').map(x=>x.trim()).filter(Boolean),
    descriptionTh:String(cfg.DESCRIPTION_TH||''),
    descriptionEn:String(cfg.DESCRIPTION_EN||''),
    prerequisite:String(cfg.PREREQUISITE||''),
    corequisite:String(cfg.COREQUISITE||''),
    consultationChannel:String(cfg.CONSULTATION_CHANNEL||''),
    courseSite:String(cfg.COURSE_SITE||''),
    repositoryUrl:String(cfg.REPOSITORY_URL||''),
    lmsUrl:String(cfg.LMS_URL||''),
    objectives:listCourseObjectives_(),
    sections:listCourseSections_(),
    staffProfiles:listCourseStaffProfiles_(),
    allowExternalStudentEmails:allowExternalStudentEmails_(),
    defaultRadiusM:getConfig_('DEFAULT_RADIUS_M')||150,
    defaultRoomLat:getConfig_('DEFAULT_ROOM_LAT'),
    defaultRoomLng:getConfig_('DEFAULT_ROOM_LNG'),
    profileReady:courseProfileComplete_(cfg)
  };
}

function saveCourseSettings(payload) {
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  payload=payload||{};

  const courseCodeTh=clean_(payload.courseCodeTh,80);
  const courseCodeEn=clean_(payload.courseCodeEn,80);
  const courseNameTh=clean_(payload.courseNameTh,250);
  const courseNameEn=clean_(payload.courseNameEn,250);

  if(!(courseCodeTh||courseCodeEn))throw new Error('Enter at least one course code (Thai or English).');
  if(!(courseNameTh||courseNameEn))throw new Error('Enter at least one course name (Thai or English).');

  const courseSite=clean_(payload.courseSite,1000);
  const repositoryUrl=clean_(payload.repositoryUrl,1000);
  const lmsUrl=clean_(payload.lmsUrl,1000);

  [courseSite,repositoryUrl,lmsUrl].filter(Boolean).forEach(url=>{
    if(!/^https?:\/\/\S+$/i.test(url))throw new Error('Course links must begin with http:// or https://');
  });

  setConfig_('COURSE_CODE_TH',courseCodeTh);
  setConfig_('COURSE_CODE_EN',courseCodeEn);
  setConfig_('COURSE_NAME_TH',courseNameTh);
  setConfig_('COURSE_NAME_EN',courseNameEn);

  // Legacy compatibility for existing code/header.
  setConfig_('COURSE_CODE',courseCodeEn||courseCodeTh);
  setConfig_('COURSE_NAME',courseNameEn||courseNameTh);

  setConfig_('TOTAL_CREDITS',clean_(payload.totalCredits,40));
  setConfig_('CREDITS',clean_(payload.totalCredits,40));
  setConfig_('LECTURE_HOURS',clean_(payload.lectureHours,40));
  setConfig_('PRACTICE_HOURS',clean_(payload.practiceHours,40));
  setConfig_('SELF_STUDY_HOURS',clean_(payload.selfStudyHours,40));
  setConfig_('SEMESTER',clean_(payload.semester,80));
  setConfig_('ACADEMIC_YEAR',clean_(payload.academicYear,80));
  setConfig_('COURSE_TYPE',clean_(payload.courseType,120));

  const subtypes=(Array.isArray(payload.courseSubtypes)?payload.courseSubtypes:[])
    .map(x=>clean_(x,120)).filter(Boolean);
  setConfig_('COURSE_SUBTYPES',subtypes.join('|'));

  setConfig_('DESCRIPTION_TH',clean_(payload.descriptionTh,5000));
  setConfig_('DESCRIPTION_EN',clean_(payload.descriptionEn,5000));
  setConfig_('PREREQUISITE',clean_(payload.prerequisite,1000));
  setConfig_('COREQUISITE',clean_(payload.corequisite,1000));
  setConfig_('CONSULTATION_CHANNEL',clean_(payload.consultationChannel,2000));
  setConfig_('COURSE_SITE',courseSite);
  setConfig_('REPOSITORY_URL',repositoryUrl);
  setConfig_('LMS_URL',lmsUrl);

  replaceCourseObjectives_(payload.objectives||[]);
  replaceCourseSections_(payload.sections||[]);

  setConfig_('DEFAULT_RADIUS_M',finiteOrBlank_(payload.radiusM));
  setConfig_('DEFAULT_ROOM_LAT',finiteOrBlank_(payload.roomLat));
  setConfig_('DEFAULT_ROOM_LNG',finiteOrBlank_(payload.roomLng));
  setConfig_('ALLOW_EXTERNAL_STUDENT_EMAILS',bool_(payload.allowExternalStudentEmails)?'TRUE':'FALSE');
  setConfig_('SETUP_COMPLETE','TRUE');

  resetRequestDataCache_();
  audit_('SAVE_COURSE_SETTINGS','Course','',{
    actor:actor.email,
    courseCodeTh,courseCodeEn,
    semester:payload.semester,
    academicYear:payload.academicYear
  });

  return {ok:true,settings:getCourseSettings(),home:getCourseHomeData_(),setupState:getSetupStateCached_()};
}




function configMapFast_(){
  const out={};
  const rows=rowsToObjects_(table_(APP.tables.CONFIG));
  rows.forEach(r=>{
    const key=String(r.Key||'').trim();
    if(key)out[key]=r.Value;
  });
  return out;
}

function getCourseHomeData_(){
  const cache=CacheService.getScriptCache();
  const key='LPT:COURSE_HOME:'+String(APP.schemaVersion||APP.version||'');
  const cached=cache.get(key);
  if(cached){
    try{return JSON.parse(cached);}catch(e){}
  }

  const clos=listCLOs_().filter(c=>c.Active).sort((a,b)=>Number(a.Sequence)-Number(b.Sequence));
  const staff=rowsToObjects_(table_(APP.tables.STAFF)).filter(r=>bool_(r.Active));
  const cfg=configMapFast_();
  const students=rowsToObjects_(table_(APP.tables.STUDENTS)).filter(r=>bool_(r.Active));

  const out={
    courseCode:String(cfg.COURSE_CODE_EN||cfg.COURSE_CODE_TH||cfg.COURSE_CODE||''),
    courseCodeTh:String(cfg.COURSE_CODE_TH||''),
    courseCodeEn:String(cfg.COURSE_CODE_EN||cfg.COURSE_CODE||''),
    courseName:String(cfg.COURSE_NAME_EN||cfg.COURSE_NAME_TH||cfg.COURSE_NAME||''),
    courseNameTh:String(cfg.COURSE_NAME_TH||''),
    courseNameEn:String(cfg.COURSE_NAME_EN||cfg.COURSE_NAME||''),
    credits:String(cfg.TOTAL_CREDITS||cfg.CREDITS||''),
    semester:String(cfg.SEMESTER||''),
    academicYear:String(cfg.ACADEMIC_YEAR||''),
    schedule:String(cfg.CLASS_SCHEDULE||cfg.SCHEDULE||''),
    room:String(cfg.CLASS_ROOM||cfg.ROOM||''),
    courseSite:String(cfg.COURSE_SITE||cfg.COURSE_URL||''),
    repositoryUrl:String(cfg.REPOSITORY_URL||''),
    lmsUrl:String(cfg.LMS_URL||''),
    courseType:String(cfg.COURSE_TYPE||''),
    courseSubtypes:String(cfg.COURSE_SUBTYPES||'').split('|').map(x=>x.trim()).filter(Boolean),
    descriptionTh:String(cfg.DESCRIPTION_TH||''),
    descriptionEn:String(cfg.DESCRIPTION_EN||''),
    prerequisite:String(cfg.PREREQUISITE||''),
    corequisite:String(cfg.COREQUISITE||''),
    objectives:listCourseObjectives_(),
    sections:listCourseSections_(),
    studentCount:students.length,
    instructors:staff.map(r=>({
      email:String(r.Email||''),
      role:String(r.Role||''),
      name:String(r.Name||r.DisplayName||r.Email||'')
    })),
    clos:clos.map(c=>({
      cloId:String(c.CLOID||''),
      code:String(c.Code||''),
      description:String(c.Description||''),
      sequence:Number(c.Sequence||0)
    }))
  };

  try{cache.put(key,JSON.stringify(out),900);}catch(e){}
  return out;
}

function clearCourseHomeCache_(){
  try{
    CacheService.getScriptCache().remove('LPT:COURSE_HOME:'+String(APP.schemaVersion||APP.version||''));
  }catch(e){}
}


function getFastStaffBootstrapData_(){
  // One small static-data package for the two pages users see first:
  // Course Home and Course Settings. No Results, CLO Evidence, Self Evaluation,
  // or Attendance analytics are touched here.
  const cfg=configMapFast_();

  const cloRows=rowsToObjects_(table_(APP.tables.CLOS))
    .filter(r=>bool_(r.Active))
    .sort((a,b)=>Number(a.Sequence)-Number(b.Sequence));

  const staffRows=rowsToObjects_(table_(APP.tables.STAFF))
    .filter(r=>bool_(r.Active));

  // Count active students using only the Active column rather than materializing
  // the entire roster object set.
  const studentSheet=table_(APP.tables.STUDENTS);
  let activeStudentCount=0;
  const lastRow=studentSheet.getLastRow();
  if(lastRow>1){
    const headers=studentSheet.getRange(1,1,1,studentSheet.getLastColumn()).getValues()[0];
    const activeCol=headers.indexOf('Active')+1;
    if(activeCol>0){
      const vals=studentSheet.getRange(2,activeCol,lastRow-1,1).getValues();
      activeStudentCount=vals.reduce((n,row)=>n+(bool_(row[0])?1:0),0);
    }else{
      activeStudentCount=Math.max(0,lastRow-1);
    }
  }

  const home={
    courseCode:String(cfg.COURSE_CODE||''),
    courseName:String(cfg.COURSE_NAME||''),
    courseNameTh:String(cfg.COURSE_NAME_TH||cfg.COURSE_NAME||''),
    courseNameEn:String(cfg.COURSE_NAME_EN||''),
    credits:String(cfg.CREDITS||cfg.COURSE_CREDITS||''),
    semester:String(cfg.SEMESTER||''),
    academicYear:String(cfg.ACADEMIC_YEAR||''),
    schedule:String(cfg.CLASS_SCHEDULE||cfg.SCHEDULE||''),
    room:String(cfg.CLASS_ROOM||cfg.ROOM||''),
    courseSite:String(cfg.COURSE_SITE||cfg.COURSE_URL||''),
    studentCount:activeStudentCount,
    instructors:staffRows.map(r=>({
      email:String(r.Email||''),
      role:String(r.Role||''),
      name:String(r.Name||r.DisplayName||r.Email||'')
    })),
    clos:cloRows.map(c=>({
      cloId:String(c.CLOID||''),
      code:String(c.Code||''),
      description:String(c.Description||''),
      sequence:Number(c.Sequence||0)
    }))
  };

  const activitySheet=table_(APP.tables.ACTIVITIES);
  let activeActivityCount=0;
  const activityLastRow=activitySheet.getLastRow();
  if(activityLastRow>1){
    const values=activitySheet.getRange(1,1,activityLastRow,activitySheet.getLastColumn()).getValues();
    const headers=values[0].map(String);
    const activeIdx=headers.indexOf('Active');
    for(let i=1;i<values.length;i++){
      const activeVal=activeIdx>=0?values[i][activeIdx]:true;
      if(activeVal==='' || activeVal==null || bool_(activeVal))activeActivityCount++;
    }
  }

  const settings={
    courseCodeTh:String(cfg.COURSE_CODE_TH||''),
    courseCodeEn:String(cfg.COURSE_CODE_EN||cfg.COURSE_CODE||''),
    courseNameTh:String(cfg.COURSE_NAME_TH||''),
    courseNameEn:String(cfg.COURSE_NAME_EN||cfg.COURSE_NAME||''),
    totalCredits:String(cfg.TOTAL_CREDITS||cfg.CREDITS||''),
    lectureHours:String(cfg.LECTURE_HOURS||''),
    practiceHours:String(cfg.PRACTICE_HOURS||''),
    selfStudyHours:String(cfg.SELF_STUDY_HOURS||''),
    semester:String(cfg.SEMESTER||''),
    academicYear:String(cfg.ACADEMIC_YEAR||''),
    courseType:String(cfg.COURSE_TYPE||''),
    courseSubtypes:String(cfg.COURSE_SUBTYPES||'').split('|').map(x=>x.trim()).filter(Boolean),
    descriptionTh:String(cfg.DESCRIPTION_TH||''),
    descriptionEn:String(cfg.DESCRIPTION_EN||''),
    prerequisite:String(cfg.PREREQUISITE||''),
    corequisite:String(cfg.COREQUISITE||''),
    consultationChannel:String(cfg.CONSULTATION_CHANNEL||''),
    courseSite:String(cfg.COURSE_SITE||''),
    repositoryUrl:String(cfg.REPOSITORY_URL||''),
    lmsUrl:String(cfg.LMS_URL||''),
    objectives:listCourseObjectives_(),
    sections:listCourseSections_(),
    staffProfiles:listCourseStaffProfiles_(),
    allowExternalStudentEmails:/^(TRUE|1|YES|Y)$/i.test(String(cfg.ALLOW_EXTERNAL_STUDENT_EMAILS||'')),
    defaultRadiusM:Number(cfg.DEFAULT_RADIUS_M||150),
    defaultRoomLat:String(cfg.DEFAULT_ROOM_LAT||''),
    defaultRoomLng:String(cfg.DEFAULT_ROOM_LNG||''),
    profileReady:courseProfileComplete_(cfg)
  };

  const closReady=cloRows.length>0;
  const studentsReady=activeStudentCount>0;
  const activitiesReady=activeActivityCount>0;
  const sectionsReady=(settings.sections||[]).filter(x=>x&&x.active!==false).length>0;
  const weeklyPlanCount=rowsToObjects_(table_(APP.tables.WEEKLY_PLAN)).filter(r=>bool_(r.Active)).length;
  const weeklyPlanReady=weeklyPlanCount>0;
  const staffReady=staffRows.length>0;
  const onboardingComplete=!!(true&&closReady&&sectionsReady&&staffReady&&studentsReady&&weeklyPlanReady&&activitiesReady);

  return {
    setupState:{
      courseReady:true,
      closReady,
      sectionsReady,
      staffReady,
      studentsReady,
      weeklyPlanReady,
      activitiesReady,
      onboardingComplete,
      canUseCLOs:true,
      canUseRoster:closReady,
      canUseStaff:true,
      canUseActivities:closReady&&studentsReady,
      canEnterResults:studentsReady&&activitiesReady,
      canUseAttendance:studentsReady,
      counts:{
        clos:cloRows.length,
        sections:(settings.sections||[]).filter(x=>x&&x.active!==false).length,
        students:activeStudentCount,
        weeklyPlan:weeklyPlanCount,
        activities:activeActivityCount,
        staff:staffRows.length
      }
    },
    home,
    settings,
    spreadsheetUrl:requireDb_().getUrl()
  };
}

function getStaffShellData_(){
  return getFastStaffBootstrapData_();
}

function getStaffPageData(tab){
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const key=String(tab||'dashboard').toLowerCase();

  if(key==='home') return getFastStaffBootstrapData_();
  if(key==='dashboard') return getDashboardKpiData_();

  if(key==='roster'){
    const students=listStudentsForStaff_().filter(s=>bool_(s.Active));
    const sectionData=studentSectionLookup_();
    return {
      students:students.map(s=>{
        const m=sectionData.byStudent[String(s.StudentKey||'')]||{};
        return {
          studentKey:s.StudentKey,studentId:s.StudentId,name:s.Name,email:s.Email,alias:s.Alias,teamCode:s.TeamCode,active:bool_(s.Active),
          sectionId:String(m.sectionId||''),sectionCode:String(m.sectionCode||'')
        };
      }),
      sections:(sectionData.sections||[]).map(s=>({sectionId:String(s.sectionId||s.SectionID||''),sectionCode:String(s.sectionCode||s.SectionCode||'')})),
      setupState:getSetupStateCached_()
    };
  }

  if(key==='activities'){
    const snap=courseSnapshot_({students:true,activities:true,clos:true,results:true,activityClos:true});
    const students=normalizedStudentsFromRows_(snap.students||[]).filter(s=>s.Active);
    const activities=normalizedActivitiesFromRows_(snap.activities||[])
      .slice().sort(sortSequence_).map(serializeActivity_);
    const clos=normalizedClosFromRows_(snap.clos||[]).filter(c=>c.Active)
      .sort((a,b)=>Number(a.Sequence)-Number(b.Sequence))
      .map(c=>({cloId:c.CLOID,code:c.Code,description:c.Description,sequence:c.Sequence}));
    return {
      students:students.map(s=>({
        studentKey:s.StudentKey,studentId:s.StudentId,name:s.Name,email:s.Email,
        alias:s.Alias,teamCode:s.TeamCode,active:bool_(s.Active)
      })),
      activities,
      clos,
      activityResultSummaries:getActivityResultSummaries_(snap),
      activityStudentVisibility:getActivityStudentVisibility_(snap),
      setupState:getSetupStateCached_()
    };
  }

  if(key==='clos'){
    const snap=courseSnapshot_({clos:true,activities:true,activityClos:true});
    return {
      clos:normalizedClosFromRows_(snap.clos||[]).filter(c=>c.Active)
        .sort((a,b)=>Number(a.Sequence)-Number(b.Sequence))
        .map(c=>({cloId:c.CLOID,code:c.Code,description:c.Description,studentFriendlyDescription:String(c.StudentFriendlyDescription||''),sequence:c.Sequence})),
      activities:normalizedActivitiesFromRows_(snap.activities||[])
        .slice().sort(sortSequence_).map(serializeActivity_),
      setupState:getSetupStateCached_()
    };
  }

  if(key==='attendance'){
    closeExpiredAttendanceSessions_();
    const sessions=rowsToObjects_(table_(APP.tables.ATTENDANCE_SESSIONS))
      .slice().sort((a,b)=>new Date(b.StartTime)-new Date(a.StartTime))
      .slice(0,20).map(s=>sanitizeAttendanceSession_(s));
    return {sessions,sections:listCourseSections_().map(s=>({sectionId:String(s.sectionId||s.SectionID||''),sectionCode:String(s.sectionCode||s.SectionCode||'')})),setupState:getSetupStateCached_()};
  }

  if(key==='staff'){
    const staff=listCourseStaff();
    return {staff,staffProfiles:listCourseStaffProfiles_(),setupState:getSetupStateCached_()};
  }

  return {setupState:getSetupStateCached_(),spreadsheetUrl:requireDb_().getUrl()};
}

function getStaffDashboard() {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  return getStaffDashboardData_();
}




function updateStudentInfo(studentKey,payload){
  const actor=requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureCurrentSchema_();

  const sh=table_(APP.tables.STUDENTS);
  const found=findObject_(sh,'StudentKey',studentKey);
  if(!found)throw new Error('Student not found.');

  const isTa=String(actor.role||'').toUpperCase()==='TA';
  const studentId=isTa?String(found.StudentId||'').trim():clean_(payload&&payload.studentId,60);
  const name=isTa?String(found.Name||'').trim():clean_(payload&&payload.name,200);
  const email=normalizeEmail_(payload&&payload.email);
  const teamCode=isTa?String(found.TeamCode||'').trim():clean_(payload&&payload.teamCode,100);
  const active=isTa?bool_(found.Active):(payload&&payload.active!==undefined?bool_(payload.active):bool_(found.Active));

  const issues=[];
  if(!studentId)issues.push('Student ID is required.');
  if(!name)issues.push('Full name is required.');
  const emailCheck=validateStudentEmail_(email);
  if(!emailCheck.ok)issues.push(emailCheck.error);

  const all=rowsToObjects_(sh);
  if(!isTa&&all.some(s=>String(s.StudentKey)!==String(studentKey)&&String(s.StudentId||'').trim()===studentId)){
    issues.push('This Student ID is already registered to another student.');
  }
  if(all.some(s=>String(s.StudentKey)!==String(studentKey)&&normalizeEmail_(s.Email)===email)){
    issues.push('This email is already registered to another student.');
  }
  if(issues.length)throw new Error([...new Set(issues)].join(' '));

  invalidateStudentAttendanceCache_(found);
  const updatedStudent=Object.assign(stripRow_(found),{
    StudentId:studentId,
    Name:name,
    Email:email,
    TeamCode:teamCode,
    Active:active
  });
  updateObjectRow_(sh,found._row,updatedStudent);
  cacheStudentForAttendance_(updatedStudent);

  audit_(isTa?'UPDATE_STUDENT_EMAIL':'UPDATE_STUDENT_INFO','Student',studentKey,{
    actor:actor.email,studentId,email,teamCode,active,role:actor.role
  });
  return {ok:true,email:email,emailOnly:isTa};
}

function addStudentManually(payload) {
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  ensureCurrentSchema_();

  const studentId=clean_(payload&&payload.studentId,60);
  const name=clean_(payload&&payload.name,200);
  const email=normalizeEmail_(payload&&payload.email);

  const issues=[];
  if(!studentId)issues.push('Student ID is required.');
  if(!name)issues.push('Full name is required.');
  const emailCheck=validateStudentEmail_(email);
  if(!emailCheck.ok)issues.push(emailCheck.error);

  const sh=table_(APP.tables.STUDENTS);
  const existing=rowsToObjects_(sh);
  const existingById=existing.find(s=>String(s.StudentId||'').trim()===studentId);
  const existingByEmail=existing.find(s=>normalizeEmail_(s.Email)===email);

  if(existingById){
    issues.push('This Student ID is already registered in the course.');
  }
  if(existingByEmail){
    issues.push('This email is already registered in the course.');
  }

  if(issues.length)throw new Error([...new Set(issues)].join(' '));

  const usedAliases=new Set(
    existing
      .map(s=>String(s.Alias||'').trim().toLowerCase())
      .filter(Boolean)
  );
  const codename=generateUniqueCodename_(usedAliases);
  const studentKey='STU-'+Utilities.getUuid();
  const now=new Date();

  appendObject_(sh,{
    StudentKey:studentKey,
    StudentId:studentId,
    Name:name,
    Email:email,
    Alias:codename,
    TeamCode:'',
    Active:true,
    AddedAt:now,
    StudentTokenHash:'',
    TokenIssuedAt:''
  });

  audit_('ADD_STUDENT_MANUALLY','Student',studentKey,{
    actor:actor.email,
    studentId:studentId,
    email:email
  });

  return {
    ok:true,
    student:{
      studentKey:studentKey,
      studentId:studentId,
      name:name,
      email:email,
      alias:codename,
      teamCode:'',
      active:true,
      externalEmail:emailCheck.external
    }
  };
}

function previewRosterFile(rows) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  if (!Array.isArray(rows) || !rows.length) throw new Error('No student rows were found in the file.');

  const existing = listStudentsForStaff_();
  const existingById = {};
  const existingByEmail = {};
  const usedAliases = new Set();

  existing.forEach(s => {
    if (s.StudentId) existingById[String(s.StudentId).trim()] = s;
    if (s.Email) existingByEmail[normalizeEmail_(s.Email)] = s;
    if (s.Alias) usedAliases.add(String(s.Alias).trim().toLowerCase());
  });

  const normalized = rows.slice(0, 3000).map((raw, index) => ({
    sourceRow: Number(raw.sourceRow || (index + 2)),
    studentId: clean_(raw.studentId, 60),
    name: clean_(raw.name, 200),
    email: normalizeEmail_(raw.email)
  }));

  const idGroups = {};
  const emailGroups = {};
  normalized.forEach(r => {
    if (r.studentId) (idGroups[r.studentId] || (idGroups[r.studentId] = [])).push(r);
    if (r.email) (emailGroups[r.email] || (emailGroups[r.email] = [])).push(r);
  });

  const ready = [];
  const problems = [];
  const duplicateIgnored = [];
  const unchanged = [];
  const seenExact = new Set();

  normalized.forEach(r => {
    const key = `${r.studentId}|||${r.name.toLowerCase()}|||${r.email}`;
    if (seenExact.has(key)) {
      duplicateIgnored.push({sourceRow:r.sourceRow,studentId:r.studentId,name:r.name,email:r.email,reason:'Exact duplicate row in file'});
      return;
    }
    seenExact.add(key);

    const issues = [];
    if (!r.studentId) issues.push('Missing Student ID');
    if (!r.name) issues.push('Missing full name');
    if (r.email) {
      if(!isValidEmail_(r.email)) issues.push('Invalid email');
      const emailCheck=validateStudentEmail_(r.email);
      if(!emailCheck.ok)issues.push(emailCheck.error);
    }

    const sameId = r.studentId ? (idGroups[r.studentId] || []) : [];
    if (sameId.length > 1) {
      const emails = [...new Set(sameId.map(x => x.email).filter(Boolean))];
      const names = [...new Set(sameId.map(x => x.name.toLowerCase()).filter(Boolean))];
      if (emails.length > 1 || names.length > 1) issues.push('Student ID appears more than once with conflicting information');
    }
    const sameEmail = r.email ? (emailGroups[r.email] || []) : [];
    if (sameEmail.length > 1) {
      const ids = [...new Set(sameEmail.map(x => x.studentId).filter(Boolean))];
      if (ids.length > 1) issues.push('Email appears more than once for different Student IDs');
    }

    const oldById = r.studentId ? existingById[r.studentId] : null;
    const oldByEmail = r.email ? existingByEmail[r.email] : null;

    if (oldById) {
      const oldEmail=normalizeEmail_(oldById.Email);
      if (oldByEmail && String(oldByEmail.StudentKey)!==String(oldById.StudentKey)) {
        issues.push('Email is already registered to another Student ID');
      } else if (r.email && oldEmail && r.email !== oldEmail) {
        issues.push('Email differs from the email already registered for this Student ID');
      }
      if (issues.length) {
        problems.push({sourceRow:r.sourceRow,studentId:r.studentId,name:r.name,email:r.email,reason:[...new Set(issues)].join('; ')});
        return;
      }
      if (r.email && !oldEmail) {
        ready.push({
          action:'updateEmail',sourceRow:r.sourceRow,studentKey:oldById.StudentKey,
          studentId:r.studentId,name:String(oldById.Name||r.name),email:r.email,
          codename:String(oldById.Alias||''),externalEmail:!isInstitutionalStudentEmail_(r.email)
        });
      } else {
        unchanged.push({sourceRow:r.sourceRow,studentId:r.studentId,name:r.name,email:r.email,reason:'Already registered; no new email to apply'});
      }
      return;
    }

    if (oldByEmail) issues.push('Email is already registered to another Student ID');
    if (issues.length) {
      problems.push({sourceRow:r.sourceRow,studentId:r.studentId,name:r.name,email:r.email,reason:[...new Set(issues)].join('; ')});
      return;
    }

    const codename = generateUniqueCodename_(usedAliases);
    usedAliases.add(codename.toLowerCase());
    ready.push({
      action:'insert',sourceRow:r.sourceRow,studentId:r.studentId,name:r.name,email:r.email,codename,
      emailPending:!r.email,externalEmail:r.email?!isInstitutionalStudentEmail_(r.email):false
    });
  });

  const insertCount=ready.filter(x=>x.action==='insert').length;
  const updateEmailCount=ready.filter(x=>x.action==='updateEmail').length;
  const pendingEmailCount=ready.filter(x=>x.action==='insert'&&!x.email).length;
  return {
    ready,problems,duplicateIgnored,unchanged,
    summary:{
      sourceRows:normalized.length,ready:ready.length,newStudents:insertCount,emailUpdates:updateEmailCount,
      pendingEmail:pendingEmailCount,problems:problems.length,duplicateIgnored:duplicateIgnored.length,
      alreadyRegistered:unchanged.length
    }
  };
}

/**
 * Preview exactly what registrar metadata import would change before applying it.
 * This is intentionally read-only and mirrors applyRegistrarRosterMetadata().
 */
function previewRegistrarRosterMetadata(meta) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  meta=meta||{};
  const cfg=configMapFast_();
  const result={course:{changes:[],unchanged:[]},section:null};
  const courseFields=[
    {key:'courseCode',labelTh:'รหัสวิชา',labelEn:'Course code',value:clean_(meta.courseCode,80),current:String(cfg.COURSE_CODE_TH||cfg.COURSE_CODE_EN||cfg.COURSE_CODE||'').trim()},
    {key:'courseName',labelTh:'ชื่อวิชา',labelEn:'Course name',value:clean_(meta.courseName,250),current:String(cfg.COURSE_NAME_TH||cfg.COURSE_NAME_EN||cfg.COURSE_NAME||'').trim()},
    {key:'credits',labelTh:'หน่วยกิต',labelEn:'Credits',value:clean_(meta.credits,40),current:String(cfg.TOTAL_CREDITS||cfg.CREDITS||'').trim()},
    {key:'semester',labelTh:'ภาคการศึกษา',labelEn:'Semester',value:clean_(meta.semester,80),current:String(cfg.SEMESTER||'').trim()},
    {key:'academicYear',labelTh:'ปีการศึกษา',labelEn:'Academic year',value:clean_(meta.academicYear,80),current:String(cfg.ACADEMIC_YEAR||'').trim()}
  ];
  courseFields.forEach(f=>{
    if(!f.value)return;
    const item={key:f.key,labelTh:f.labelTh,labelEn:f.labelEn,value:f.value,current:f.current};
    if(!f.current)result.course.changes.push(item);else result.course.unchanged.push(item);
  });

  const sectionCode=clean_(String(meta.sectionCode==null?'':meta.sectionCode),80);
  if(sectionCode){
    const sh=table_(APP.tables.COURSE_SECTIONS),existing=rowsToObjects_(sh);
    const canonical=s=>String(s||'').trim();
    const digitKey=s=>/^\d+$/.test(canonical(s))?canonical(s).replace(/^0+(?=\d)/,''):'';
    let old=existing.find(r=>canonical(r.SectionCode)===sectionCode);
    if(!old&&digitKey(sectionCode))old=existing.find(r=>digitKey(r.SectionCode)===digitKey(sectionCode));
    const schedules=(Array.isArray(meta.schedules)?meta.schedules:[]).map(x=>({
      day:clean_(x&&x.day,20),startTime:clean_(x&&x.startTime,20),endTime:clean_(x&&x.endTime,20),room:clean_(x&&x.room,120)
    })).filter(x=>x.day||x.startTime||x.endTime||x.room);
    const exams=(Array.isArray(meta.exams)?meta.exams:[]).map(x=>({
      examType:clean_(x&&x.examType,40),date:clean_(x&&x.date,40),startTime:clean_(x&&x.startTime,20),endTime:clean_(x&&x.endTime,20),room:clean_(x&&x.room,120)
    })).filter(x=>x.examType||x.date||x.startTime||x.endTime||x.room);
    const instructors=(Array.isArray(meta.instructors)?meta.instructors:[]).map(x=>clean_(x,250)).filter(Boolean);
    result.section={
      sectionCode:sectionCode,
      action:old?'update':'add',
      currentSectionCode:old?String(old.SectionCode||''):'',
      schedules:schedules,
      exams:exams,
      instructors:instructors,
      mode:old&&old.Mode?String(old.Mode):'Onsite',
      preservesAssignedInstructor:!!(old&&old.InstructorStaffKey)
    };
  }
  return result;
}

/**
 * Apply course/section metadata detected from a Thammasat registrar roster export.
 * This intentionally does not change schema. Section import is an upsert so an
 * instructor can re-upload a refreshed registrar roster without duplicating the section.
 */
function applyRegistrarRosterMetadata(meta, options) {
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);meta=meta||{};options=options||{};const result={ok:true,section:null,course:null};
  if(bool_(options.course)){
    const cfg=configMapFast_(),codeValue=clean_(meta.courseCode,80),nameTh=clean_(meta.courseName,250),credits=clean_(meta.credits,40),semester=clean_(meta.semester,80),academicYear=clean_(meta.academicYear,80);
    if(codeValue&&!String(cfg.COURSE_CODE_TH||cfg.COURSE_CODE_EN||cfg.COURSE_CODE||'').trim()){setConfig_('COURSE_CODE_TH',codeValue);setConfig_('COURSE_CODE',codeValue);}
    if(nameTh&&!String(cfg.COURSE_NAME_TH||cfg.COURSE_NAME_EN||cfg.COURSE_NAME||'').trim()){setConfig_('COURSE_NAME_TH',nameTh);setConfig_('COURSE_NAME',nameTh);}
    if(credits&&!String(cfg.TOTAL_CREDITS||cfg.CREDITS||'').trim()){setConfig_('TOTAL_CREDITS',credits);setConfig_('CREDITS',credits);}
    if(semester&&!String(cfg.SEMESTER||'').trim())setConfig_('SEMESTER',semester);if(academicYear&&!String(cfg.ACADEMIC_YEAR||'').trim())setConfig_('ACADEMIC_YEAR',academicYear);
    result.course={courseCode:codeValue,courseNameTh:nameTh,credits,semester,academicYear};
  }
  if(bool_(options.section)){
    const sectionCode=clean_(String(meta.sectionCode==null?'':meta.sectionCode),80);
    if(sectionCode){
      const schedules=Array.isArray(meta.schedules)?meta.schedules:[],exams=Array.isArray(meta.exams)?meta.exams:[];
      const days=[];schedules.forEach(x=>{const d=clean_(x&&x.day,20);if(d&&!days.includes(d))days.push(d);});
      const pairs=schedules.map(x=>[clean_(x&&x.startTime,20),clean_(x&&x.endTime,20)]).filter(x=>x[0]||x[1]);
      const sameTime=pairs.length&&pairs.every(x=>x[0]===pairs[0][0]&&x[1]===pairs[0][1]);
      const startTime=sameTime?pairs[0][0]:'',endTime=sameTime?pairs[0][1]:'';
      const rooms=[];schedules.forEach(x=>{const r=clean_(x&&x.room,120);if(r&&!rooms.includes(r))rooms.push(r);});
      const dayTime=schedules.map(x=>[clean_(x&&x.day,20),[clean_(x&&x.startTime,20),clean_(x&&x.endTime,20)].filter(Boolean).join('–'),clean_(x&&x.room,120)].filter(Boolean).join(' ')).join(', ');
      const instructorName=(Array.isArray(meta.instructors)?meta.instructors:[]).map(x=>clean_(x,250)).filter(Boolean).join(', ');
      const sh=table_(APP.tables.COURSE_SECTIONS),existing=rowsToObjects_(sh);
      const canonical=s=>String(s||'').trim();const digitKey=s=>/^\d+$/.test(canonical(s))?canonical(s).replace(/^0+(?=\d)/,''):'';
      let old=existing.find(r=>canonical(r.SectionCode)===sectionCode);
      if(!old&&digitKey(sectionCode))old=existing.find(r=>digitKey(r.SectionCode)===digitKey(sectionCode));
      const now=new Date(),payload={SectionID:old?old.SectionID:'SEC-'+Utilities.getUuid(),SectionCode:sectionCode,Days:days.join('|'),StartTime:startTime,EndTime:endTime,DayTime:dayTime,Room:rooms.join(', '),Mode:old&&old.Mode?old.Mode:'Onsite',InstructorStaffKey:old?String(old.InstructorStaffKey||''):'',InstructorName:(old&&old.InstructorStaffKey)?String(old.InstructorName||''):instructorName,InstructorEmail:old?String(old.InstructorEmail||''):'',Active:true,CreatedAt:old&&old.CreatedAt?old.CreatedAt:now,UpdatedAt:now};
      let rowNum;if(old){updateObjectRow_(sh,old._row,payload);rowNum=old._row;}else{appendObject_(sh,payload);rowNum=sh.getLastRow();}
      writeSectionCodeText_(sh,rowNum,sectionCode);replaceSectionSchedules_(String(payload.SectionID),schedules,exams);
      result.section={sectionCode,days,startTime,endTime,room:rooms.join(', '),instructorName,schedules,exams,updated:!!old};
    }
  }
  resetRequestDataCache_();audit_('IMPORT_REGISTRAR_METADATA','Course','',{actor:actor.email,section:!!result.section,course:!!result.course,sectionCode:result.section?result.section.sectionCode:''});return result;
}


function assignRegistrarStudentsToSection(sectionCode,studentIds){ return assignStudentsToSectionByIds(sectionCode,studentIds); }

function confirmRosterFileImportBatch(rows) {
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  if(!Array.isArray(rows)||!rows.length)throw new Error('There are no validated students in this batch.');

  const sheet=table_(APP.tables.STUDENTS);
  const existing=rowsToObjects_(sheet);
  const byId={},byEmail={};
  const usedAliases=new Set();
  existing.forEach(s=>{
    if(s.StudentId)byId[String(s.StudentId).trim()]=s;
    if(s.Email)byEmail[normalizeEmail_(s.Email)]=s;
    if(s.Alias)usedAliases.add(String(s.Alias).trim().toLowerCase());
  });

  let inserted=0,updatedEmail=0;
  const rejected=[];
  rows.slice(0,200).forEach((raw,index)=>{
    const action=String(raw.action||'insert');
    const studentId=clean_(raw.studentId,60);
    const name=clean_(raw.name,200);
    const email=normalizeEmail_(raw.email);
    let codename=clean_(raw.codename,100);
    const issues=[];

    if(!studentId)issues.push('Missing Student ID');
    if(action==='updateEmail'){
      const old=byId[studentId];
      if(!old)issues.push('Student ID is no longer registered');
      if(!email||!isValidEmail_(email))issues.push('Invalid email');
      if(email){
        const emailCheck=validateStudentEmail_(email);if(!emailCheck.ok)issues.push(emailCheck.error);
        const owner=byEmail[email];if(owner&&String(owner.StudentKey)!==String(old&&old.StudentKey))issues.push('Email is already registered to another Student ID');
        const oldEmail=normalizeEmail_(old&&old.Email);if(oldEmail&&oldEmail!==email)issues.push('Existing email differs; update it manually from Edit Student');
      }
      if(issues.length){rejected.push({sourceRow:Number(raw.sourceRow||index+2),studentId,name,email,reason:issues.join('; ')});return;}
      const updated=Object.assign(stripRow_(old),{Email:email});
      updateObjectRow_(sheet,old._row,updated);
      byId[studentId]=Object.assign({},old,{Email:email});byEmail[email]=byId[studentId];
      cacheStudentForAttendance_(updated);
      updatedEmail++;
      return;
    }

    if(!name)issues.push('Missing full name');
    if(email){
      if(!isValidEmail_(email))issues.push('Invalid email');
      const emailCheck=validateStudentEmail_(email);if(!emailCheck.ok)issues.push(emailCheck.error);
    }
    if(byId[studentId])issues.push('Already registered');
    if(email&&byEmail[email])issues.push('Email already registered');
    if(issues.length){rejected.push({sourceRow:Number(raw.sourceRow||index+2),studentId,name,email,reason:issues.join('; ')});return;}

    if(!codename||usedAliases.has(codename.toLowerCase()))codename=generateUniqueCodename_(usedAliases);
    const newStudent={
      StudentKey:'STU-'+Utilities.getUuid(),StudentId:studentId,Name:name,Email:email,Alias:codename,
      TeamCode:'',Active:true,AddedAt:new Date(),StudentTokenHash:'',TokenIssuedAt:''
    };
    appendObject_(sheet,newStudent);
    usedAliases.add(codename.toLowerCase());byId[studentId]=newStudent;if(email)byEmail[email]=newStudent;inserted++;
  });

  audit_('CONFIRM_ROSTER_FILE_IMPORT_BATCH','Roster','',{actor:actor.email,inserted,updatedEmail,rejected:rejected.length});
  return {ok:true,inserted,updatedEmail,rejected};
}

function confirmRosterFileImport(rows) {
  return confirmRosterFileImportBatch(rows);
}

function allowExternalStudentEmails_(){
  return bool_(getConfig_('ALLOW_EXTERNAL_STUDENT_EMAILS'));
}
function isInstitutionalStudentEmail_(email){
  return /@(?:dome|sci)\.tu\.ac\.th$/i.test(String(email||'').trim());
}
function validateStudentEmail_(email){
  const value=normalizeEmail_(email);
  if(!value)return {ok:false,error:'Email is required.',external:false};
  if(!isValidEmail_(value))return {ok:false,error:'Invalid email format.',external:false};
  const external=!isInstitutionalStudentEmail_(value);
  if(external&&!allowExternalStudentEmails_()){
    return {ok:false,error:'Student email must use @dome.tu.ac.th or @sci.tu.ac.th unless external student emails are enabled in Course Settings.',external:true};
  }
  return {ok:true,error:'',external:external};
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function generateUniqueCodename_(usedAliases) {
  const adjectives = [
    'Brave','Calm','Clever','Crimson','Crystal','Golden','Green','Hidden',
    'Iron','Mighty','Mystic','Noble','Quiet','Silver','Silent','Swift',
    'Azure','Bright','Gentle','Steady','Bold','Wise','Vivid','True'
  ];
  const nouns = [
    'Falcon','Phoenix','Tiger','Lion','Eagle','Wolf','Dragon','Serpent',
    'River','Forest','Mountain','Shield','Spear','Crown','Storm','Thunder',
    'Ocean','Hawk','Bear','Comet','Orbit','Nova','Cedar','Lotus'
  ];

  const used = usedAliases instanceof Set ? usedAliases : new Set();
  const total = adjectives.length * nouns.length;

  // Randomized starting point, then deterministic scan guarantees uniqueness
  // until all adjective/noun combinations are exhausted.
  const seed = Math.floor(Math.random() * total);
  for (let i = 0; i < total; i++) {
    const n = (seed + i) % total;
    const alias = `${adjectives[Math.floor(n / nouns.length)]} ${nouns[n % nouns.length]}`;
    if (!used.has(alias.toLowerCase())) return alias;
  }

  // Fallback for very large courses.
  let n = 1;
  while (used.has(`student ${n}`)) n++;
  return `Student ${n}`;
}

function importRoster(rows) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  if (!Array.isArray(rows) || !rows.length) throw new Error('No roster rows supplied.');

  const sheet = table_(APP.tables.STUDENTS);
  const existing = rowsToObjects_(sheet);
  const byEmail = {};
  const byStudentId = {};
  existing.forEach(r => {
    if (r.Email) byEmail[normalizeEmail_(r.Email)] = r;
    if (r.StudentId) byStudentId[String(r.StudentId)] = r;
  });

  let nextAlias = nextAliasNumber_(existing);
  let inserted = 0, updated = 0, skipped = 0;

  rows.slice(0, 2000).forEach(raw => {
    const studentId = clean_(raw.studentId, 60);
    const name = clean_(raw.name, 200);
    const email = normalizeEmail_(raw.email);
    const teamCode = clean_(raw.teamCode, 60);
    if (!studentId || !email) { skipped++; return; }

    const old = byEmail[email] || byStudentId[studentId];
    if (old) {
      updateObjectRow_(sheet, old._row, {
        StudentKey: old.StudentKey,
        StudentId: studentId,
        Name: name,
        Email: email,
        Alias: old.Alias || aliasFor_(nextAlias++),
        TeamCode: teamCode,
        Active: true,
        AddedAt: old.AddedAt || new Date()
      });
      updated++;
    } else {
      appendObject_(sheet, {
        StudentKey: 'STU-' + Utilities.getUuid(),
        StudentId: studentId,
        Name: name,
        Email: email,
        Alias: aliasFor_(nextAlias++),
        TeamCode: teamCode,
        Active: true,
        AddedAt: new Date()
      });
      inserted++;
    }
  });

  audit_('IMPORT_ROSTER', 'Roster', '', {inserted,updated,skipped});
  const students = listStudentsForStaff_();
  return {
    ok:true,
    appVersion:APP.version,
    inserted,
    updated,
    skipped,
    students:Array.isArray(students)?students:[]
  };
}

function saveCLOs(cloRows) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  if (!Array.isArray(cloRows)) throw new Error('Invalid CLO data.');
  const sh = table_(APP.tables.CLOS);
  const existing = rowsToObjects_(sh);
  const byCode = {};
  existing.forEach(c => byCode[String(c.Code).trim().toLowerCase()] = c);

  const next = [];
  cloRows.filter(x => clean_(x.code,50)).slice(0,50).forEach((x,i) => {
    const code = clean_(x.code,50);
    const old = byCode[code.toLowerCase()];
    next.push({
      CLOID: old ? old.CLOID : 'CLO-' + Utilities.getUuid(),
      Code: code,
      Description: clean_(x.description,1000),
      StudentFriendlyDescription: clean_(x.studentFriendlyDescription!==undefined?x.studentFriendlyDescription:(old?old.StudentFriendlyDescription:''),1500),
      Sequence: i + 1,
      Active: true
    });
  });

  rewriteObjects_(sh, next);
  audit_('SAVE_CLOS', 'CLO', '', {count:next.length});
  return {ok:true, clos:listCLOs_()};
}


function decodeActivityDetailUrlEntities_(value){
  let out=String(value||'');
  for(let i=0;i<4;i++){
    const next=out
      .replace(/&lt;/gi,'<')
      .replace(/&gt;/gi,'>')
      .replace(/&#0*60;/gi,'<')
      .replace(/&#x0*3c;/gi,'<')
      .replace(/&#0*62;/gi,'>')
      .replace(/&#x0*3e;/gi,'>')
      .replace(/&quot;/gi,'\"')
      .replace(/&#0*39;/gi,"'")
      .replace(/&amp;/gi,'&');
    if(next===out)break;
    out=next;
  }
  return out;
}

function safeActivityDetailUrl_(value){
  const raw=String(value||'').trim();
  if(!raw)return '';
  const decoded=decodeActivityDetailUrlEntities_(raw).trim();
  // Discard raw or HTML-encoded fragments left by the old activity-form template bug.
  if(/<(?:span|label|select|option|input|div|textarea|button)\b/i.test(decoded))return '';
  // A detail link must be one complete HTTP(S) URL, not mixed text/markup.
  if(!/^https?:\/\/[^\s<>\"']+$/i.test(decoded))return '';
  return decoded;
}

function cleanActivityDetailUrl_(value) {
  const raw=clean_(value,1000);
  if(!raw)return '';
  const url=safeActivityDetailUrl_(raw);
  if(!url){
    throw new Error('Activity Detail Link must be a complete URL beginning with http:// or https://');
  }
  return url;
}

function cleanupLegacyActivityDetailUrls_(){
  const props=PropertiesService.getScriptProperties();
  // New key intentionally re-runs cleanup for courses where the v2.30 migration
  // already completed before HTML-encoded malformed values were detected.
  const key='LPT_MIGRATION_DETAILURL_V2462';
  if(props.getProperty(key)==='DONE')return;
  try{
    const sh=table_(APP.tables.ACTIVITIES);
    rowsToObjects_(sh).forEach(r=>{
      const raw=String(r.DetailUrl||'').trim();
      if(!raw)return;
      const safe=safeActivityDetailUrl_(raw);
      if(!safe){
        updateObjectRow_(sh,r._row,Object.assign(stripRow_(r),{DetailUrl:''}));
      }else if(safe!==raw){
        updateObjectRow_(sh,r._row,Object.assign(stripRow_(r),{DetailUrl:safe}));
      }
    });
    resetRequestDataCache_();
    props.setProperty(key,'DONE');
  }catch(e){
    console.warn('Legacy activity DetailUrl cleanup skipped:',e);
  }
}

function normalizeActivityMode_(value) {
  const raw = String(value == null ? '' : value).trim();
  const key = raw.toLowerCase();
  if (key === 'team' || raw === 'กลุ่ม' || raw === 'ทีม') return 'Team';
  if (key === 'individual' || raw === 'รายบุคคล' || raw === 'บุคคล') return 'Individual';
  return 'Individual';
}

function normalizeActivityPurpose_(value) {
  const raw = String(value == null ? '' : value).trim();
  const key = raw.toLowerCase();
  if (key === 'assessed' || key === 'assessment' || key === 'summative' || raw === 'ประเมินผล') return 'Assessed';
  if (key === 'formative' || raw === 'เพื่อพัฒนา') return 'Formative';
  return APP.purposes.includes(raw) ? raw : 'Formative';
}

function saveActivity(activity) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  const sh = table_(APP.tables.ACTIVITIES);
  const now = new Date();
  const id = clean_(activity.activityId,100) || 'ACT-' + Utilities.getUuid();
  const mode = normalizeActivityMode_(activity.mode);
  const purpose = normalizeActivityPurpose_(activity.purpose);
  const category = APP.categories.includes(activity.category) ? activity.category : 'Other';
  const maxScore = activity.maxScore === '' || activity.maxScore == null ? '' : clampNum_(activity.maxScore,0,100000,0);
  const weight = activity.weight === '' || activity.weight == null ? '' : clampNum_(activity.weight,0,100,0);
  const title = clean_(activity.title,200);
  if (!title) throw new Error('Activity title is required.');
  const shortDescription = clean_(activity.shortDescription,800);
  const detailUrl = cleanActivityDetailUrl_(activity.detailUrl);
  const selected = Array.isArray(activity.cloIds) ? activity.cloIds : [];
  const visibilityLevel=normalizeActivityVisibilityLevel_(activity.studentVisibilityLevel,activity.published);
  validateScoreBasedCloConfig_(activity, selected);
  const cloScoreMappings=normalizeCloScoreMappingsInput_(activity,selected);

  // ActivityID is a logical primary key. Older course files may contain duplicate
  // rows from earlier edit/import flows, so resolve all matching rows explicitly
  // instead of updating only the first cached match.
  const foundRows = rowsToObjects_(sh).filter(r => String(r.ActivityID||'') === String(id));
  const found = foundRows.length ? foundRows[0] : null;
  const obj = {
    ActivityID:id,
    Title:title,
    ShortDescription:shortDescription,
    DetailUrl:detailUrl,
    Category:category,
    Mode:mode,
    Purpose:purpose,
    MaxScore:maxScore,
    Weight:weight,
    Published:visibilityLevel!=='HIDDEN',
    StudentVisibilityLevel:visibilityLevel,
    AssignedWeek: activity.assignedWeek===''||activity.assignedWeek==null?'':clampNum_(activity.assignedWeek,1,60,''),
    DueDate: clean_(activity.dueDate,30),
    Sequence: clampNum_(activity.sequence,1,10000,nextActivitySequence_()),
    Active:true,
    CreatedAt: found ? found.CreatedAt : now,
    UpdatedAt: now,
    CLOEvalMode:String(activity.cloEvalMode||'MANUAL').toUpperCase()==='SCORE'?'SCORE':'MANUAL',
    CLOScoreSource:String(activity.cloScoreSource||'ACTIVITY_SCORE').toUpperCase()==='SEPARATE_SCORE'?'SEPARATE_SCORE':'ACTIVITY_SCORE',
    CLOMaxScore:activity.cloMaxScore===''||activity.cloMaxScore==null?'':clampNum_(activity.cloMaxScore,0,100000,0),
    CLOLevel2Pct:activity.cloLevel2Pct===''?'':clampNum_(activity.cloLevel2Pct,0,100,50),
    CLOLevel3Pct:activity.cloLevel3Pct===''?'':clampNum_(activity.cloLevel3Pct,0,100,65),
    CLOLevel4Pct:activity.cloLevel4Pct===''?'':clampNum_(activity.cloLevel4Pct,0,100,80),
    CLOLevel5Pct:activity.cloLevel5Pct===''?'':clampNum_(activity.cloLevel5Pct,0,100,90),
    CLOAllowOverride:bool_(activity.cloAllowOverride)
  };
  if (found) {
    updateObjectRow_(sh, found._row, obj);
    // Remove duplicate rows with the same ActivityID, if any. References elsewhere
    // use ActivityID rather than row number, so keeping one canonical row is safe.
    foundRows.slice(1).map(r=>Number(r._row)).filter(Number.isFinite).sort((a,b)=>b-a).forEach(rowNum=>sh.deleteRow(rowNum));
    if(foundRows.length>1) resetRequestDataCache_();
  } else {
    appendObject_(sh,obj);
  }
  SpreadsheetApp.flush();

  const mapSh = table_(APP.tables.ACTIVITY_CLOS);
  const maps = rowsToObjects_(mapSh).filter(r => r.ActivityID !== id);
  rewriteObjects_(mapSh, maps.map(stripRow_));
  selected.forEach(cloId => appendObject_(mapSh,{ActivityID:id,CLOID:String(cloId)}));

  const scoreMapSh=table_(APP.tables.ACTIVITY_CLO_SCORE_MAPS);
  rewriteObjects_(scoreMapSh,rowsToObjects_(scoreMapSh).filter(r=>String(r.ActivityID)!==String(id)).map(stripRow_));
  if(obj.CLOEvalMode==='SCORE')cloScoreMappings.forEach(m=>appendObject_(scoreMapSh,{ActivityID:id,CLOID:m.cloId,ScoreSource:m.scoreSource,MaxScore:m.scoreSource==='SEPARATE_SCORE'?m.maxScore:maxScore,UpdatedAt:now}));

  audit_('SAVE_ACTIVITY','Activity',id,{title,mode,purpose,published:obj.Published,studentVisibilityLevel:visibilityLevel});
  // Activity, CLO-link and score-map writes above invalidate the request-local snapshot.
  // Clear it before returning the saved record so the client receives the canonical persisted value.
  resetRequestDataCache_();
  const savedActivity=getActivity_(id);
  if(!savedActivity) throw new Error('Activity was written but could not be read back. Please reload and try again.');
  if(String(savedActivity.Purpose||'')!==String(purpose)){
    throw new Error('Activity purpose was not persisted correctly. Expected '+purpose+' but found '+String(savedActivity.Purpose||'')+'.');
  }
  return {ok:true, activityId:id, activity:serializeActivity_(savedActivity)};
}


function repairActivityVisibility(activityId) {
  const staff = requireStaff_(['OWNER','INSTRUCTOR']);
  ensureCurrentSchema_();

  const sh = table_(APP.tables.ACTIVITIES);
  const found = findObject_(sh,'ActivityID',activityId);
  if (!found) throw new Error('Activity not found.');

  const obj = stripRow_(found);
  obj.Published = true;
  obj.StudentVisibilityLevel = 'SCORE_CLO';
  obj.Active = true;
  obj.UpdatedAt = new Date();
  updateObjectRow_(sh,found._row,obj);

  audit_('REPAIR_ACTIVITY_VISIBILITY','Activity',activityId,{actor:staff.email});
  return {
    ok:true,
    activity:getActivity_(activityId),
    visibility:getActivityStudentVisibility_()[activityId]
  };
}

function archiveActivity(activityId) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  const sh = table_(APP.tables.ACTIVITIES);
  const found = findObject_(sh,'ActivityID',activityId);
  if (!found) throw new Error('Activity not found.');

  updateObjectRow_(sh, found._row, Object.assign(stripRow_(found), {
    Active:false,
    Published:false,
    StudentVisibilityLevel:'HIDDEN',
    UpdatedAt:new Date()
  }));

  audit_('ARCHIVE_ACTIVITY','Activity',activityId,{title:found.Title});
  return {ok:true, activities:listActivities_()};
}

function restoreActivity(activityId) {
  requireStaff_(['OWNER','INSTRUCTOR']);
  const sh = table_(APP.tables.ACTIVITIES);
  const found = findObject_(sh,'ActivityID',activityId);
  if (!found) throw new Error('Activity not found.');

  updateObjectRow_(sh, found._row, Object.assign(stripRow_(found), {
    Active:true,
    UpdatedAt:new Date()
  }));

  audit_('RESTORE_ACTIVITY','Activity',activityId,{title:found.Title});
  return {ok:true, activities:listActivities_()};
}

// Backward-compatible alias from earlier versions.
// This performs an archive, not a destructive delete.
function deleteActivity(activityId) {
  return archiveActivity(activityId);
}

function getResultEntryData(activityId) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const act = getActivity_(activityId);
  if (!act) throw new Error('Activity not found.');

  const students = listStudentsForStaff_().filter(s=>s.Active);
  let targets;
  if (act.Mode === 'Team') {
    const teams=activityTeamData_(activityId);
    targets=teams.map(t=>({targetType:'TEAM',targetId:t.teamId,label:t.teamName,memberCount:t.studentKeys.length}));
  } else {
    targets = students.map(s => ({targetType:'STUDENT', targetId:s.StudentKey, label:`${s.StudentId} — ${s.Name}`}));
  }

  const resultRows = rowsToObjects_(table_(APP.tables.RESULTS)).filter(r=>r.ActivityID===activityId);
  const evidenceRows = rowsToObjects_(table_(APP.tables.CLO_EVIDENCE)).filter(r=>r.ActivityID===activityId);
  const resultMap = {};
  resultRows.forEach(r => resultMap[r.TargetType+'|'+r.TargetID]=r);
  const evMap = {};
  evidenceRows.forEach(r => {
    const k=r.TargetType+'|'+r.TargetID;
    if(!evMap[k]) evMap[k]={levels:{},sources:{},scores:{},maxScores:{}};
    evMap[k].levels[r.CLOID]=r.Level;
    evMap[k].sources[r.CLOID]=r.Source||'MANUAL';
    evMap[k].scores[r.CLOID]=r.Score===''||r.Score==null?'':r.Score;
    evMap[k].maxScores[r.CLOID]=r.MaxScore===''||r.MaxScore==null?'':r.MaxScore;
  });
  const cloIds = activityCloIds_(activityId);
  const clos = listCLOs_().filter(c=>cloIds.includes(c.CLOID));

  return {
    activity:serializeActivity_(act),
    clos,
    targets:targets.map(t=>{
      const k=t.targetType+'|'+t.targetId, r=resultMap[k]||{};
      return Object.assign({},t,{
        hasSavedResult:!!resultMap[k],
        hasEvidence:!!evMap[k],
        status:r.Status||'Not Submitted',
        score:r.Score===''||r.Score==null?'':r.Score,
        feedback:r.Feedback||'',
        cloLevels:evMap[k]?evMap[k].levels:{},
        cloSources:evMap[k]?evMap[k].sources:{},
        cloScores:evMap[k]?evMap[k].scores:{},
        cloMaxScores:evMap[k]?evMap[k].maxScores:{}
      });
    })
  };
}


function inspectResultSpreadsheet(spreadsheetRef) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const id = spreadsheetIdFromRef_(spreadsheetRef);
  if (!id) throw new Error('Enter a valid Google Spreadsheet URL or ID.');

  const ss = SpreadsheetApp.openById(id);
  const sheets = ss.getSheets().map(sh => {
    const lastCol = sh.getLastColumn();
    const headers = lastCol > 0 ? sh.getRange(1,1,1,lastCol).getDisplayValues()[0] : [];
    return {
      name: sh.getName(),
      headers,
      rowCount: Math.max(0, sh.getLastRow()-1)
    };
  });

  return {
    spreadsheetId:id,
    spreadsheetName:ss.getName(),
    sheets
  };
}

function previewGoogleFormImport(activityId, options) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const act = getActivity_(activityId);
  if (!act || !act.Active) throw new Error('Activity not found or archived.');

  const id = spreadsheetIdFromRef_(options.spreadsheetRef);
  if (!id) throw new Error('Enter a valid Google Spreadsheet URL or ID.');
  const ss = SpreadsheetApp.openById(id);

  const sheetName = clean_(options.sheetName,200);
  const sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sh) throw new Error('Response sheet not found.');
  if (sh.getLastRow() < 2) return {activity:serializeActivity_(act), rows:[], summary:{matched:0,unmatched:0,duplicates:0,total:0}, headers:[]};

  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0].map(x=>String(x).trim());
  const records = values.slice(1).map((row,i)=>{
    const obj={_sourceRow:i+2};
    headers.forEach((h,j)=>obj[h]=row[j]);
    return obj;
  });

  const matchColumn = clean_(options.matchColumn,200) || detectHeader_(headers,['Email Address','Email','E-mail','Student Email','Student ID','StudentID']);
  const scoreColumn = clean_(options.scoreColumn,200) || detectHeader_(headers,['Score','Points','Total Score','คะแนน']);
  const timestampColumn = clean_(options.timestampColumn,200) || detectHeader_(headers,['Timestamp','Submission Time','Submitted at','เวลา']);
  const feedbackColumn = clean_(options.feedbackColumn,200);
  const cloScoreColumns=(options.cloScoreColumns&&typeof options.cloScoreColumns==='object')?options.cloScoreColumns:{};
  const matchBy = options.matchBy === 'studentId' ? 'studentId' : 'email';
  const attemptPolicy = ['latest','highest','first'].includes(options.attemptPolicy) ? options.attemptPolicy : 'latest';

  if (!matchColumn || !headers.includes(matchColumn)) throw new Error('Select the column used to identify students.');
  if (!scoreColumn || !headers.includes(scoreColumn)) throw new Error('Select the Activity score column.');
  activityCloScoreMappings_(activityId).filter(m=>m.scoreSource==='SEPARATE_SCORE').forEach(m=>{const col=clean_(cloScoreColumns[m.cloId],200);if(!col||!headers.includes(col))throw new Error('Select a score column for each separately scored CLO.');});

  return buildImportPreview_(act, records, {
    matchBy, matchColumn, scoreColumn, cloScoreColumns, timestampColumn, feedbackColumn, attemptPolicy
  });
}

function previewUploadedResultRows(activityId, payload) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const act = getActivity_(activityId);
  if (!act || !act.Active) throw new Error('Activity not found or archived.');

  const records = Array.isArray(payload.records) ? payload.records.slice(0,3000) : [];
  if (!records.length) throw new Error('No result rows were found in the uploaded file.');

  return buildImportPreview_(act, records, {
    matchBy: payload.matchBy === 'studentId' ? 'studentId' : 'email',
    matchColumn: clean_(payload.matchColumn,200),
    scoreColumn: clean_(payload.scoreColumn,200),
    cloScoreColumns:(payload.cloScoreColumns&&typeof payload.cloScoreColumns==='object')?payload.cloScoreColumns:{},
    timestampColumn: clean_(payload.timestampColumn,200),
    feedbackColumn: clean_(payload.feedbackColumn,200),
    attemptPolicy: ['latest','highest','first'].includes(payload.attemptPolicy) ? payload.attemptPolicy : 'latest'
  });
}

function applyResultImport(activityId, previewRows) {
  const staff = requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const act = getActivity_(activityId);
  if (!act || !act.Active) throw new Error('Activity not found or archived.');
  act.cloScoreMappings=activityCloScoreMappings_(activityId);
  if (!Array.isArray(previewRows)) throw new Error('Invalid import rows.');

  const rs = table_(APP.tables.RESULTS);
  let imported = 0;

  previewRows.slice(0,3000).forEach(row => {
    if (!row.matched || !row.targetId) return;

    let score = row.score;
    if (score === '' || score == null) score = '';
    else {
      score = Number(score);
      if (!Number.isFinite(score)) score = '';
      if (score !== '' && act.MaxScore !== '' && act.MaxScore != null) {
        score = Math.max(0, Math.min(Number(act.MaxScore), score));
      }
    }

    upsertComposite_(rs,['ActivityID','TargetType','TargetID'],
      [activityId,row.targetType,row.targetId],{
        ActivityID:activityId,
        TargetType:row.targetType,
        TargetID:row.targetId,
        Status:row.status || 'Submitted',
        Score:score,
        Feedback:clean_(row.feedback,2000),
        UpdatedAt:new Date(),
        UpdatedBy:staff.email
      });
    if (row.targetType === 'TEAM') captureTeamResultMembers_(activityId,row.targetId);
    saveDerivedCloEvidence_(act,row.targetType,row.targetId,score,row.cloScores||{},staff.email,null);
    imported++;
  });

  audit_('IMPORT_RESULTS','Activity',activityId,{imported});
  return {ok:true, imported};
}

function clearActivityImportedResults(activityId) {
  const staff = requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const act = getActivity_(activityId);
  if (!act || !act.Active) throw new Error('Activity not found or archived.');

  const resultsSheet = table_(APP.tables.RESULTS);
  const evidenceSheet = table_(APP.tables.CLO_EVIDENCE);
  const teamMembersSheet = table_(APP.tables.TEAM_RESULT_MEMBERS);

  const resultCount = rowsToObjects_(resultsSheet).filter(r=>String(r.ActivityID)===String(activityId)).length;
  const evidenceCount = rowsToObjects_(evidenceSheet).filter(r=>String(r.ActivityID)===String(activityId)).length;
  const snapshotCount = rowsToObjects_(teamMembersSheet).filter(r=>String(r.ActivityID)===String(activityId)).length;

  deleteRowsWhere_(resultsSheet, r=>String(r.ActivityID)===String(activityId));
  deleteRowsWhere_(evidenceSheet, r=>String(r.ActivityID)===String(activityId));
  deleteRowsWhere_(teamMembersSheet, r=>String(r.ActivityID)===String(activityId));
  resetRequestDataCache_();
  audit_('CLEAR_ACTIVITY_RESULTS','Activity',activityId,{results:resultCount,cloEvidence:evidenceCount,teamResultMembers:snapshotCount,actor:staff.email});
  return {ok:true, resultsCleared:resultCount, evidenceCleared:evidenceCount};
}

function buildImportPreview_(act, records, options) {
  const students = listStudentsForStaff_().filter(s=>s.Active);
  const byEmail = {};
  const byStudentId = {};
  students.forEach(s=>{
    if(s.Email) byEmail[normalizeEmail_(s.Email)] = s;
    if(s.StudentId) byStudentId[String(s.StudentId).trim()] = s;
  });

  if (!options.matchColumn) throw new Error('Select a student matching column.');
  if (!options.scoreColumn) throw new Error('Select the Activity score column.');
  const allScoreMappings=activityCloScoreMappings_(act.ActivityID);
  const separateMappings=allScoreMappings.filter(m=>m.scoreSource==='SEPARATE_SCORE');
  separateMappings.forEach(m=>{if(!options.cloScoreColumns||!options.cloScoreColumns[m.cloId])throw new Error('Select a score column for each separately scored CLO.');});
  const currentResults={};rowsToObjects_(table_(APP.tables.RESULTS)).filter(r=>String(r.ActivityID)===String(act.ActivityID)).forEach(r=>currentResults[String(r.TargetType)+'|'+String(r.TargetID)]=r);
  const currentEvidence={};rowsToObjects_(table_(APP.tables.CLO_EVIDENCE)).filter(r=>String(r.ActivityID)===String(act.ActivityID)).forEach(r=>currentEvidence[String(r.TargetType)+'|'+String(r.TargetID)+'|'+String(r.CLOID)]=r);
  const sameValue_=(a,b)=>{if((a===''||a==null)&&(b===''||b==null))return true;const na=Number(a),nb=Number(b);if(Number.isFinite(na)&&Number.isFinite(nb))return Math.abs(na-nb)<1e-9;return String(a??'')===String(b??'');};

  const grouped = {};
  let duplicateCount = 0;

  records.forEach((rec,index)=>{
    const rawMatch = String(rec[options.matchColumn] ?? '').trim();
    if (!rawMatch) return;

    const key = options.matchBy === 'studentId' ? rawMatch : normalizeEmail_(rawMatch);
    if(!grouped[key]) grouped[key]=[];
    else duplicateCount++;

    grouped[key].push({
      rec,
      sourceIndex:index,
      score:parseScore_(rec[options.scoreColumn]),
      cloScores:Object.fromEntries(separateMappings.map(m=>[m.cloId,parseScore_(rec[(options.cloScoreColumns||{})[m.cloId]])])),
      timestamp:parseImportDate_(options.timestampColumn ? rec[options.timestampColumn] : '')
    });
  });

  const chosen = [];
  Object.keys(grouped).forEach(key=>{
    const attempts=grouped[key];
    let item=attempts[0];

    if(options.attemptPolicy==='highest'){
      item=attempts.reduce((best,x)=>{
        const bs=Number.isFinite(best.score)?best.score:-Infinity;
        const xs=Number.isFinite(x.score)?x.score:-Infinity;
        return xs>bs?x:best;
      },attempts[0]);
    } else if(options.attemptPolicy==='latest'){
      item=attempts.reduce((best,x)=>{
        const bt=best.timestamp?best.timestamp.getTime():best.sourceIndex;
        const xt=x.timestamp?x.timestamp.getTime():x.sourceIndex;
        return xt>=bt?x:best;
      },attempts[0]);
    }

    chosen.push({key,item,attempts:attempts.length});
  });

  let matched=0,unmatched=0;
  const rows=chosen.map(({key,item,attempts})=>{
    const st = options.matchBy==='studentId' ? byStudentId[key] : byEmail[key];
    const parsed = Number.isFinite(item.score) ? item.score : '';
    const parsedClos={};separateMappings.forEach(m=>{const v=(item.cloScores||{})[m.cloId];parsedClos[m.cloId]=Number.isFinite(v)?v:'';});
    const feedback = options.feedbackColumn ? clean_(item.rec[options.feedbackColumn],2000) : '';

    if(!st){
      unmatched++;
      return {
        matched:false,
        sourceIdentity:key,
        studentLabel:'',
        targetType:'',
        targetId:'',
        score:parsed,
        cloScores:parsedClos,
        status:'Submitted',
        feedback,
        attempts
      };
    }

    matched++;
    const targetType='STUDENT',targetId=st.StudentKey;
    const cur=currentResults[targetType+'|'+String(targetId)];
    let unchanged=!!cur&&String(cur.Status||'')==='Submitted'&&sameValue_(cur.Score,parsed)&&String(cur.Feedback||'')===String(feedback||'');
    if(unchanged&&String(act.CLOEvalMode||'MANUAL')==='SCORE'){
      allScoreMappings.forEach(m=>{
        if(!unchanged)return;
        const evidenceScore=m.scoreSource==='SEPARATE_SCORE'?parsedClos[m.cloId]:parsed;
        const expectedLevel=deriveCloLevelFromMappedScore_(act,m,parsed,evidenceScore);
        const ev=currentEvidence[targetType+'|'+String(targetId)+'|'+String(m.cloId)];
        if(expectedLevel===''){if(ev)unchanged=false;return;}
        if(!ev||String(ev.Source||'')==='OVERRIDE'||Number(ev.Level)!==Number(expectedLevel)||!sameValue_(ev.Score,evidenceScore))unchanged=false;
      });
    }
    return {
      matched:true,
      unchanged:unchanged,
      sourceIdentity:key,
      studentLabel:`${st.StudentId} — ${st.Name}`,
      alias:st.Alias,
      targetType:targetType,
      targetId:targetId,
      score:parsed,
      cloScores:parsedClos,
      status:'Submitted',
      feedback,
      attempts
    };
  });

  const unchanged=rows.filter(r=>r.matched&&r.unchanged).length;
  return {
    activity:serializeActivity_(act),
    rows,
    summary:{matched,unmatched,duplicates:duplicateCount,total:chosen.length,unchanged:unchanged,changedMatched:matched-unchanged},
    headers:records.length?Object.keys(records[0]).filter(k=>!k.startsWith('_')):[],
    mapping:options
  };
}

function spreadsheetIdFromRef_(ref) {
  const s=String(ref||'').trim();
  if(!s) return '';
  const m=s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if(m) return m[1];
  return /^[a-zA-Z0-9-_]{20,}$/.test(s) ? s : '';
}

function detectHeader_(headers,candidates) {
  const lower={};
  headers.forEach(h=>lower[String(h).toLowerCase()]=h);
  for(const c of candidates){
    const found=lower[String(c).toLowerCase()];
    if(found) return found;
  }
  return '';
}

function parseScore_(value) {
  if(value===''||value==null) return NaN;
  if(typeof value==='number') return Number(value);
  const s=String(value).trim();
  const first=s.match(/-?\d+(?:\.\d+)?/);
  return first ? Number(first[0]) : NaN;
}

function parseImportDate_(value) {
  if(!value) return null;
  const d=new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}

function saveResults(activityId, rows) {
  const staff = requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const act = getActivity_(activityId);
  if (!act) throw new Error('Activity not found.');
  act.cloScoreMappings=activityCloScoreMappings_(activityId);
  if (!Array.isArray(rows)) throw new Error('Invalid results.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rs = table_(APP.tables.RESULTS);
    const es = table_(APP.tables.CLO_EVIDENCE);
    const allowedCLOs = activityCloIds_(activityId);
    rows.slice(0,2000).forEach(row => {
      const targetType = row.targetType === 'TEAM' ? 'TEAM' : 'STUDENT';
      const targetId = clean_(row.targetId,120);
      if (!targetId) return;
      const status = APP.statuses.includes(row.status) ? row.status : 'Not Submitted';
      let score = row.score;
      if (score === '' || score == null) score = '';
      else {
        score = Number(score);
        if (!Number.isFinite(score)) score = '';
        if (score !== '' && act.MaxScore !== '' && act.MaxScore != null && Number(act.MaxScore) >= 0) score = Math.max(0, Math.min(Number(act.MaxScore), score));
      }
      upsertComposite_(rs, ['ActivityID','TargetType','TargetID'], [activityId,targetType,targetId], {
        ActivityID:activityId, TargetType:targetType, TargetID:targetId,
        Status:status, Score:score, Feedback:clean_(row.feedback,2000),
        UpdatedAt:new Date(), UpdatedBy:staff.email
      });

      if (targetType === 'TEAM') {
        captureTeamResultMembers_(activityId,targetId);
      }

      const levels = row.cloLevels || {};
      const handledByScore = saveDerivedCloEvidence_(act,targetType,targetId,score,row.cloScores||{},staff.email,levels);
      if (!handledByScore) {
        allowedCLOs.forEach(cloId => {
          const raw = levels[cloId];
          if (raw === '' || raw == null) {
            deleteComposite_(es, ['ActivityID','TargetType','TargetID','CLOID'], [activityId,targetType,targetId,cloId]);
            return;
          }
          const level = clampNum_(raw, APP.cloMin, APP.cloMax, APP.cloMin);
          upsertComposite_(es, ['ActivityID','TargetType','TargetID','CLOID'], [activityId,targetType,targetId,cloId], {
            ActivityID:activityId,TargetType:targetType,TargetID:targetId,CLOID:cloId,
            Level:level,Source:'MANUAL',Score:'',MaxScore:'',UpdatedAt:new Date(),UpdatedBy:staff.email
          });
        });
      }
    });
  } finally {
    lock.releaseLock();
  }

  audit_('SAVE_RESULTS','Activity',activityId,{count:rows.length});
  return {ok:true};
}


function attendanceDisplayKey_(s) {
  if (!s || !s.SessionID || !s.Secret) return '';
  const bytes = Utilities.computeHmacSha256Signature(String(s.SessionID), String(s.Secret));
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/,'').substring(0,32);
}

function getAttendanceDisplayData(sessionId, displayKey) {
  const s = attendanceSession_(sessionId);
  if (!s) throw new Error('Attendance session not found.');

  const expected = attendanceDisplayKey_(s);
  if (!expected || String(displayKey || '') !== expected) {
    throw new Error('Invalid attendance presentation link.');
  }

  autoCloseSessionIfExpired_(s);
  const fresh = attendanceSession_(sessionId);
  const open = fresh && fresh.Status === 'OPEN';
  const token = open ? attendanceToken_(fresh,new Date()) : '';
  const url = open ? webAppBaseUrl_()+`?checkin=${encodeURIComponent(fresh.SessionID)}&token=${encodeURIComponent(token)}` : '';

  return {
    session:sanitizeAttendanceSession_(fresh),
    checkinUrl:url,
    serverNow:new Date().toISOString()
  };
}

function createAttendanceSession(options) {
  const staff = requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const title = clean_(options.title,200) || 'Class Check-in';
  const sectionId=clean_(options.sectionId,120);
  const minutes = clampNum_(options.minutes,1,60,Number(getConfig_('DEFAULT_CHECKIN_MINUTES')||5));
  const rotate = clampNum_(options.rotateSeconds,30,300,Number(getConfig_('DEFAULT_ROTATE_SECONDS')||60));
  const locationRequired = bool_(options.locationRequired);
  const roomLat = finiteOrBlank_(options.roomLat);
  const roomLng = finiteOrBlank_(options.roomLng);
  const radiusM = clampNum_(options.radiusM,10,2000,Number(getConfig_('DEFAULT_RADIUS_M')||150));
  if (locationRequired && (roomLat === '' || roomLng === '')) throw new Error('Room latitude and longitude are required when location verification is enabled.');

  const start = new Date();
  const end = new Date(start.getTime()+minutes*60000);
  const id = 'ATT-' + Utilities.getUuid();
  const sessionRecord={
    SessionID:id,SectionID:sectionId,Title:title,StartTime:start,EndTime:end,Status:'OPEN',
    Secret:Utilities.getUuid()+Utilities.getUuid(),RotateSeconds:rotate,
    LocationRequired:locationRequired,RoomLat:roomLat,RoomLng:roomLng,RadiusM:radiusM,
    StoreRawLocation:bool_(options.storeRawLocation),CreatedBy:staff.email,Published:false
  };
  appendObject_(table_(APP.tables.ATTENDANCE_SESSIONS),sessionRecord);
  cacheAttendanceSession_(sessionRecord);
  audit_('CREATE_ATTENDANCE','AttendanceSession',id,{title,minutes,locationRequired,sectionId});
  return getAttendanceLive(id);
}


function ensureAttendanceRecordSchema_() {
  ensureHeaders_(APP.tables.ATTENDANCE_SESSIONS);
  ensureHeaders_(APP.tables.ATTENDANCE);
}

function attendanceEvidenceFolder_() {
  const props=PropertiesService.getScriptProperties();
  const saved=props.getProperty('LPT_ATTENDANCE_EVIDENCE_FOLDER_ID');
  if(saved){
    try{return DriveApp.getFolderById(saved);}catch(e){props.deleteProperty('LPT_ATTENDANCE_EVIDENCE_FOLDER_ID');}
  }

  const course=safeSheetNamePart_(getConfig_('COURSE_CODE')||'Course');
  const term=safeSheetNamePart_(getConfig_('TERM')||'Term');
  const folder=DriveApp.createFolder(`${course}_${term}_Attendance_Evidence`);
  props.setProperty('LPT_ATTENDANCE_EVIDENCE_FOLDER_ID',folder.getId());
  return folder;
}

function saveAttendanceEvidencePdf_(studentId, sessionId, filePayload) {
  if(!filePayload || !filePayload.base64) return null;

  const name=clean_(filePayload.name,180);
  const mime=String(filePayload.mimeType||'').toLowerCase();
  if(mime!=='application/pdf' && !/\.pdf$/i.test(name)) {
    throw new Error('Attendance evidence must be a PDF file.');
  }

  const bytes=Utilities.base64Decode(String(filePayload.base64));
  if(bytes.length > 10*1024*1024) throw new Error('Attendance evidence PDF must be 10 MB or smaller.');

  const safeStudent=safeSheetNamePart_(studentId||'student');
  const safeSession=safeSheetNamePart_(sessionId||'attendance');
  const safeName=safeSheetNamePart_(name.replace(/\.pdf$/i,''))||'evidence';
  const finalName=`${safeSession}_${safeStudent}_${safeName}.pdf`;

  const blob=Utilities.newBlob(bytes,'application/pdf',finalName);
  const file=attendanceEvidenceFolder_().createFile(blob);

  return {
    fileId:file.getId(),
    fileName:file.getName(),
    url:file.getUrl()
  };
}

function trashAttendanceEvidenceFile_(fileId) {
  if(!fileId) return;
  try{DriveApp.getFileById(String(fileId)).setTrashed(true);}catch(e){}
}

function setAttendancePublished(sessionId, published) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR']);
  ensureAttendanceRecordSchema_();
  const sh=table_(APP.tables.ATTENDANCE_SESSIONS);
  const found=findObject_(sh,'SessionID',sessionId);
  if(!found) throw new Error('Attendance session not found.');

  updateObjectRow_(sh,found._row,Object.assign(stripRow_(found),{
    Published:bool_(published)
  }));
  invalidateAttendanceSessionCache_(sessionId);
  audit_(bool_(published)?'PUBLISH_ATTENDANCE':'UNPUBLISH_ATTENDANCE','AttendanceSession',sessionId,{actor:staff.email});
  return {ok:true,published:bool_(published)};
}

function deleteAttendanceSession(sessionId) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR']);
  ensureAttendanceRecordSchema_();

  const sessionSh=table_(APP.tables.ATTENDANCE_SESSIONS);
  const found=findObject_(sessionSh,'SessionID',sessionId);
  if(!found) throw new Error('Attendance session not found.');

  const attendanceSh=table_(APP.tables.ATTENDANCE);
  const rows=rowsToObjects_(attendanceSh);
  rows.filter(r=>String(r.SessionID)===String(sessionId)).forEach(r=>{
    trashAttendanceEvidenceFile_(r.EvidenceFileId);
  });

  const remaining=rows.filter(r=>String(r.SessionID)!==String(sessionId)).map(stripRow_);
  rewriteObjects_(attendanceSh,remaining);
  sessionSh.deleteRow(found._row);

  audit_('DELETE_ATTENDANCE_SESSION','AttendanceSession',sessionId,{actor:staff.email});
  return {ok:true};
}

function clearAttendanceRecord(sessionId, studentKey, deleteEvidence) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureAttendanceRecordSchema_();
  const sh=table_(APP.tables.ATTENDANCE);
  const found=findComposite_(sh,['SessionID','StudentKey'],[sessionId,studentKey]);
  if(!found) return {ok:true,alreadyClear:true};

  if(bool_(deleteEvidence)) trashAttendanceEvidenceFile_(found.EvidenceFileId);
  deleteComposite_(sh,['SessionID','StudentKey'],[sessionId,studentKey]);
  invalidateAttendanceResultCache_(sessionId,studentKey);

  audit_('CLEAR_ATTENDANCE_RECORD','AttendanceSession',sessionId,{
    studentKey:studentKey,
    deleteEvidence:bool_(deleteEvidence),
    actor:staff.email
  });
  return {ok:true};
}

function saveAttendanceOverride(sessionId, studentKey, payload) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureAttendanceRecordSchema_();

  const allowed=['Present','Late','Excused','Absent'];
  const status=String((payload&&payload.status)||'');
  if(!allowed.includes(status)) throw new Error('Invalid attendance status.');

  const student=listStudentsForStaff_().find(s=>String(s.StudentKey)===String(studentKey));
  if(!student) throw new Error('Student not found.');

  const session=attendanceSession_(sessionId);
  if(!session) throw new Error('Attendance session not found.');

  const sh=table_(APP.tables.ATTENDANCE);
  const found=findComposite_(sh,['SessionID','StudentKey'],[sessionId,studentKey]);
  const old=found?stripRow_(found):{};

  let evidenceFileId=old.EvidenceFileId||'';
  let evidenceFileName=old.EvidenceFileName||'';
  let evidenceUrl=old.EvidenceUrl||'';

  if(payload&&payload.removeEvidence){
    trashAttendanceEvidenceFile_(evidenceFileId);
    evidenceFileId=''; evidenceFileName=''; evidenceUrl='';
  }

  if(payload&&payload.file&&payload.file.base64){
    const saved=saveAttendanceEvidencePdf_(student.StudentId,sessionId,payload.file);
    if(evidenceFileId) trashAttendanceEvidenceFile_(evidenceFileId);
    evidenceFileId=saved.fileId;
    evidenceFileName=saved.fileName;
    evidenceUrl=saved.url;
  }

  const now=new Date();
  upsertComposite_(sh,['SessionID','StudentKey'],[sessionId,studentKey],{
    SessionID:sessionId,
    StudentKey:studentKey,
    CheckedInAt:old.CheckedInAt||now,
    Status:status,
    DistanceM:old.DistanceM||'',
    AccuracyM:old.AccuracyM||'',
    LocationVerified:old.LocationVerified||false,
    Latitude:old.Latitude||'',
    Longitude:old.Longitude||'',
    Source:'MANUAL:'+staff.email,
    Remark:clean_((payload&&payload.remark)||'',1000),
    EvidenceFileId:evidenceFileId,
    EvidenceFileName:evidenceFileName,
    EvidenceUrl:evidenceUrl,
    UpdatedBy:staff.email,
    UpdatedAt:now
  });

  invalidateAttendanceResultCache_(sessionId,studentKey);
  audit_('SAVE_ATTENDANCE_OVERRIDE','AttendanceSession',sessionId,{
    studentKey:studentKey,
    status:status,
    hasEvidence:!!evidenceFileId,
    actor:staff.email
  });

  return {ok:true};
}

function getAttendanceLive(sessionId) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureAttendanceRecordSchema_();
  const s = attendanceSession_(sessionId);
  if (!s) throw new Error('Attendance session not found.');
  autoCloseSessionIfExpired_(s);
  const fresh = attendanceSession_(sessionId);
  const token = fresh.Status === 'OPEN' ? attendanceToken_(fresh,new Date()) : '';
  const url = fresh.Status === 'OPEN' ? webAppBaseUrl_()+`?checkin=${encodeURIComponent(fresh.SessionID)}&token=${encodeURIComponent(token)}` : '';
  const snap=courseSnapshot_({attendance:true,students:true});
  let students=normalizedStudentsFromRows_(snap.students).filter(s=>s.Active);
  const attendanceSectionId=String(fresh.SectionID||'');
  if(attendanceSectionId){const allowed=new Set(activeStudentSectionRows_().filter(r=>String(r.SectionID||'')===attendanceSectionId).map(r=>String(r.StudentKey||'')));students=students.filter(st=>allowed.has(String(st.StudentKey||'')));}
  const checked={};
  (snap.attendance||[]).forEach(a=>{
    if(String(a.SessionID||'')===String(sessionId)) checked[String(a.StudentKey||'')]=a;
  });
  const displayKey = attendanceDisplayKey_(fresh);
  const displayUrl = displayKey
    ? webAppBaseUrl_()+`?page=attendance-display&display=1&sessionId=${encodeURIComponent(fresh.SessionID)}&displayKey=${encodeURIComponent(displayKey)}`
    : '';
  return {
    session:sanitizeAttendanceSession_(fresh),
    displayUrl:displayUrl,
    token,
    checkinUrl:url,
    serverNow:new Date().toISOString(),
    roster:students.map(st=>{
      const a=checked[st.StudentKey];
      return {
        studentKey:st.StudentKey,studentId:st.StudentId,name:st.Name,alias:st.Alias,
        checkedIn:!!a && ['Present','Late'].includes(String(a.Status||'')),
        attempted:!!a && !['Present','Late','Excused','Absent'].includes(String(a.Status||'')),
        recorded:!!a,
        status:a?a.Status:'Absent',
        checkedInAt:a?dateIso_(a.CheckedInAt):'',
        distanceM:a&&a.DistanceM!==''?a.DistanceM:'',
        accuracyM:a&&a.AccuracyM!==''?a.AccuracyM:'',
        locationVerified:a?bool_(a.LocationVerified):false,
        remark:a?(a.Remark||''):'',
        evidenceFileName:a?(a.EvidenceFileName||''):'',
        evidenceUrl:a?(a.EvidenceUrl||''):'',
        updatedBy:a?(a.UpdatedBy||''):'',
        updatedAt:a?dateIso_(a.UpdatedAt):''
      };
    })
  };
}

function getAttendanceLiveStatus(sessionId){
  requireStaff_(['OWNER','INSTRUCTOR','TA']);

  let s=attendanceSessionFast_(sessionId);
  if(!s)throw new Error('Attendance session not found.');

  if(String(s.Status||'')==='OPEN'&&new Date(s.EndTime).getTime()<Date.now()){
    const sheetSession=attendanceSession_(sessionId);
    if(sheetSession){
      autoCloseSessionIfExpired_(sheetSession);
      invalidateAttendanceSessionCache_(sessionId);
      s=attendanceSessionFast_(sessionId);
    }
  }

  const rows=rowsToObjects_(table_(APP.tables.ATTENDANCE));
  const records=[];
  rows.forEach(a=>{
    if(String(a.SessionID||'')!==String(sessionId))return;
    records.push({
      studentKey:String(a.StudentKey||''),
      checkedIn:['Present','Late'].includes(String(a.Status||'')),
      attempted:!['Present','Late','Excused','Absent'].includes(String(a.Status||'')),
      recorded:true,
      status:String(a.Status||'Absent'),
      checkedInAt:dateIso_(a.CheckedInAt),
      distanceM:a.DistanceM!==''?a.DistanceM:'',
      accuracyM:a.AccuracyM!==''?a.AccuracyM:'',
      locationVerified:bool_(a.LocationVerified),
      remark:a.Remark||'',
      evidenceFileName:a.EvidenceFileName||'',
      evidenceUrl:a.EvidenceUrl||'',
      updatedBy:a.UpdatedBy||'',
      updatedAt:dateIso_(a.UpdatedAt)
    });
  });

  const token=String(s.Status||'')==='OPEN'?attendanceToken_(s,new Date()):'';
  const url=String(s.Status||'')==='OPEN'
    ? webAppBaseUrl_()+`?checkin=${encodeURIComponent(s.SessionID)}&token=${encodeURIComponent(token)}`
    : '';

  return {
    session:sanitizeAttendanceSession_(s),
    token,
    checkinUrl:url,
    serverNow:new Date().toISOString(),
    records
  };
}

function closeAttendanceSession(sessionId) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const sh=table_(APP.tables.ATTENDANCE_SESSIONS), found=findObject_(sh,'SessionID',sessionId);
  if(!found) throw new Error('Attendance session not found.');
  updateObjectRow_(sh,found._row,Object.assign(stripRow_(found),{Status:'CLOSED',EndTime:new Date()}));
  invalidateAttendanceSessionCache_(sessionId);
  audit_('CLOSE_ATTENDANCE','AttendanceSession',sessionId,{});
  return {ok:true};
}

function getCheckinInfo(sessionId, token, studentToken) {
  return getCheckinInfoFast(sessionId,token,studentToken);
}

function getCheckinInfoFast(sessionId,token,studentToken){
  const student=resolveStudentFast_(studentToken);
  if(!student||!bool_(student.Active)){
    throw new Error('Unable to identify an enrolled student. Open your Student Access Link once on this browser, then scan the QR again.');
  }

  let s=attendanceSessionFast_(sessionId);
  if(!s)throw new Error('Check-in session not found.');

  if(String(s.Status||'')==='OPEN'&&new Date(s.EndTime).getTime()<Date.now()){
    // Expiry is rare; only then touch the session sheet.
    const sheetSession=attendanceSession_(sessionId);
    if(sheetSession){
      autoCloseSessionIfExpired_(sheetSession);
      invalidateAttendanceSessionCache_(sessionId);
      s=attendanceSessionFast_(sessionId);
    }
  }

  if(String(s.Status||'')!=='OPEN'){
    return {open:false,title:s.Title,locationRequired:bool_(s.LocationRequired),message:'Check-in has closed.'};
  }
  if(!validAttendanceToken_(s,token,new Date())){
    return {open:false,title:s.Title,locationRequired:bool_(s.LocationRequired),message:'This QR code has expired. Scan the current QR code.'};
  }

  return {
    open:true,
    title:s.Title,
    locationRequired:bool_(s.LocationRequired),
    radiusM:bool_(s.LocationRequired)?Number(s.RadiusM||0):null
  };
}

function checkIn(sessionId, token, location, studentToken) {
  return checkInFast(sessionId,token,location,studentToken);
}

function checkInFast(sessionId,token,location,studentToken){
  const student=resolveStudentFast_(studentToken);
  if(!student||!bool_(student.Active)){
    throw new Error('Unable to identify an enrolled student. Open your Student Access Link once on this browser, then scan the QR again.');
  }

  let s=attendanceSessionFast_(sessionId);
  if(!s)throw new Error('Check-in session not found.');

  if(String(s.Status||'')==='OPEN'&&new Date(s.EndTime).getTime()<Date.now()){
    const sheetSession=attendanceSession_(sessionId);
    if(sheetSession){
      autoCloseSessionIfExpired_(sheetSession);
      invalidateAttendanceSessionCache_(sessionId);
      s=attendanceSessionFast_(sessionId);
    }
  }

  if(String(s.Status||'')!=='OPEN')throw new Error('Check-in has closed.');
  if(String(s.SectionID||'')&&!studentBelongsToSection_(student.StudentKey,String(s.SectionID||'')))throw new Error('This attendance session is for another Section.');
  if(!validAttendanceToken_(s,token,new Date()))throw new Error('This QR code has expired. Scan the current QR code.');

  // Fast idempotency path. A successful prior request can be returned directly
  // without reading Sheets again (important when a browser retries after a
  // response/network failure).
  const cached=getCachedAttendanceResult_(sessionId,student.StudentKey);
  if(cached&&['Present','Late','Excused'].includes(String(cached.attendanceStatus||''))){
    cached.already=true;
    cached.message='Your attendance is already recorded as '+String(cached.attendanceStatus||'Present')+'.';
    return cached;
  }

  const sh=table_(APP.tables.ATTENDANCE);
  const existing=findComposite_(sh,['SessionID','StudentKey'],[sessionId,student.StudentKey]);

  if(existing&&['Present','Late','Excused'].includes(String(existing.Status||''))){
    const result={
      ok:true,already:true,
      message:'Your attendance is already recorded as '+String(existing.Status||'Present')+'.',
      attendanceStatus:String(existing.Status||'Present'),
      checkedInAt:dateIso_(existing.CheckedInAt),
      locationRequired:bool_(s.LocationRequired),
      locationVerified:bool_(existing.LocationVerified),
      locationReason:'',
      distanceM:existing.DistanceM!==''?existing.DistanceM:'',
      accuracyM:existing.AccuracyM!==''?existing.AccuracyM:'',
      radiusM:bool_(s.LocationRequired)?Number(s.RadiusM||0):null,
      canRetry:false
    };
    cacheAttendanceResult_(sessionId,student.StudentKey,result);
    return result;
  }

  const loc=validateLocation_(s,location||{});
  const locationRequired=bool_(s.LocationRequired);
  let attendanceStatus='Present';
  let message='Check-in recorded.';

  if(locationRequired&&!loc.verified){
    if(loc.reason==='OUTSIDE_RADIUS'){
      attendanceStatus='Outside Allowed Area';
      message='Check-in attempt recorded, but your location is outside the allowed area. You are not marked Present.';
    }else{
      attendanceStatus='Location Not Verified';
      message='Check-in attempt recorded, but your location could not be verified. You are not marked Present.';
    }
  }

  const now=new Date();
  const raw=bool_(s.StoreRawLocation);
  const record={
    SessionID:sessionId,
    StudentKey:student.StudentKey,
    CheckedInAt:now,
    Status:attendanceStatus,
    DistanceM:loc.distanceM,
    AccuracyM:loc.accuracyM,
    LocationVerified:loc.verified,
    Latitude:raw?loc.lat:'',
    Longitude:raw?loc.lng:'',
    Source:'QR',
    Remark:existing?(existing.Remark||''):'',
    EvidenceFileId:existing?(existing.EvidenceFileId||''):'',
    EvidenceFileName:existing?(existing.EvidenceFileName||''):'',
    EvidenceUrl:existing?(existing.EvidenceUrl||''):'',
    UpdatedBy:'',
    UpdatedAt:now
  };

  if(existing)updateObjectRow_(sh,existing._row,record);
  else appendObject_(sh,record);

  // Successful QR check-in is already evidenced by the Attendance record.
  // Avoid a second synchronous _AuditLog write during the high-concurrency burst.
  const result={
    ok:true,already:false,message,
    attendanceStatus,
    checkedInAt:now.toISOString(),
    locationRequired,
    locationVerified:loc.verified,
    locationReason:loc.reason||'',
    distanceM:loc.distanceM,
    accuracyM:loc.accuracyM,
    radiusM:locationRequired?Number(s.RadiusM||0):null,
    canRetry:locationRequired&&!loc.verified
  };
  cacheAttendanceResult_(sessionId,student.StudentKey,result);
  return result;
}

function setManualAttendance(sessionId, studentKey, status) {
  return saveAttendanceOverride(sessionId,studentKey,{status:status,remark:''});
}




/* ---------- Generated Google Forms for CLO self-evaluation ---------- */
function selfEvaluationFormConfigPrefix_(point) {
  point = normalizeSelfEvaluationPoint_(point);
  return 'SELF_EVAL_' + point + '_';
}

function selfEvaluationPointLabel_(point) {
  point = normalizeSelfEvaluationPoint_(point);
  return point === 'START' ? 'Start of term' : point === 'MIDTERM' ? 'Midterm' : 'Final';
}

function selfEvaluationFormRecord_(point) {
  point = normalizeSelfEvaluationPoint_(point);
  const prefix = selfEvaluationFormConfigPrefix_(point);
  const formId = String(getConfig_(prefix + 'FORM_ID') || '').trim();
  const responseSpreadsheetId = String(getConfig_(prefix + 'RESPONSE_SPREADSHEET_ID') || '').trim();
  const createdAt = String(getConfig_(prefix + 'CREATED_AT') || '').trim();
  let formUrl = '', editUrl = '', responseUrl = '', available = false, title = '';
  if (formId) {
    try {
      const form = FormApp.openById(formId);
      formUrl = form.getPublishedUrl();
      editUrl = form.getEditUrl();
      title = form.getTitle();
      available = true;
    } catch (e) {}
  }
  if (responseSpreadsheetId) {
    try { responseUrl = SpreadsheetApp.openById(responseSpreadsheetId).getUrl(); } catch (e) {}
  }
  return {
    point: point,
    label: selfEvaluationPointLabel_(point),
    exists: !!formId,
    available: available,
    formId: formId,
    formUrl: formUrl,
    editUrl: editUrl,
    responseSpreadsheetId: responseSpreadsheetId,
    responseUrl: responseUrl,
    createdAt: createdAt,
    title: title
  };
}

function getSelfEvaluationFormStatus() {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureCurrentSchema_();
  return {
    forms: ['START','MIDTERM','END'].map(selfEvaluationFormRecord_),
    clos: listCLOs_().filter(c => c.Active).sort((a,b) => Number(a.Sequence)-Number(b.Sequence))
      .map(c => ({cloId:c.CLOID, code:c.Code, description:c.Description, studentFriendlyDescription:String(c.StudentFriendlyDescription||'')}))
  };
}

function tryMoveGeneratedSelfEvaluationFiles_(fileIds) {
  try {
    const db = dbOrNull_();
    if (!db) return;
    const parents = DriveApp.getFileById(db.getId()).getParents();
    if (!parents.hasNext()) return;
    const folder = parents.next();
    (fileIds || []).filter(Boolean).forEach(id => {
      try { DriveApp.getFileById(id).moveTo(folder); } catch (e) {}
    });
  } catch (e) {}
}

function createSelfEvaluationForm(point, replaceExisting) {
  const staff = requireStaff_(['OWNER','INSTRUCTOR']);
  ensureCurrentSchema_();
  point = normalizeSelfEvaluationPoint_(point);
  const existing = selfEvaluationFormRecord_(point);
  if (existing.exists && existing.available && !replaceExisting) {
    return {ok:true, created:false, existing:true, form:existing};
  }

  const clos = listCLOs_().filter(c => c.Active).sort((a,b) => Number(a.Sequence)-Number(b.Sequence));
  if (!clos.length) throw new Error('Create and save at least one CLO before creating a self-evaluation form.');

  const cfg = courseConfig_();
  const courseLabel = [String(cfg.code||'').trim(), String(cfg.name||'').trim()].filter(Boolean).join(' — ') || 'Course';
  const pointLabel = selfEvaluationPointLabel_(point);
  const title = courseLabel + ' — CLO Self-Evaluation (' + pointLabel + ')';
  const form = FormApp.create(title);
  form.setDescription(
    'ประเมินตนเองตาม Course Learning Outcomes (CLOs) ของรายวิชา\n' +
    'ช่วงการประเมิน: ' + pointLabel + '\n\n' +
    'ระดับ 1–5 ใช้เพื่อสะท้อนการรับรู้ความสามารถของตนเอง และเป็นหลักฐานทางอ้อม (indirect evidence) ไม่ใช่คะแนนผลสัมฤทธิ์โดยตรง'
  );
  try { form.setConfirmationMessage('ขอบคุณค่ะ ระบบได้รับแบบประเมินตนเองของคุณแล้ว'); } catch (e) {}
  // Self-evaluation forms should verify the respondent's Google account,
  // allow correction after submission, and accept only one response per account.
  try { form.setCollectEmail(true); } catch (e) {}
  try { form.setAllowResponseEdits(true); } catch (e) {}
  try { form.setLimitOneResponsePerUser(true); } catch (e) {}

  form.addTextItem().setTitle('Student ID').setRequired(true)
    .setHelpText('กรอกรหัสนักศึกษาให้ตรงกับที่ลงทะเบียนในรายวิชา');
  form.addTextItem().setTitle('Name').setRequired(false);
  form.addTextItem().setTitle('Section').setRequired(false);

  const selfEvaluationChoices = [
    '1 — ยังไม่บรรลุ ทำไม่ได้ แม้จะได้รับคำแนะนำแล้ว',
    '2 — ผ่านขั้นต่ำ ทำได้บ้างแต่ยังไม่ครบถ้วนและยังต้องมีแนะนำอยู่มาก',
    '3 — ดี ทำได้โอเค แต่อาจต้องมีคำแนะนำเล็กน้อย',
    '4 — ดีมาก ทำได้ถูกต้องเกือบทั้งหมดด้วยตนเอง',
    '5 — ดีเยี่ยม ทำได้ถูกต้องครบถ้วนด้วยตนเองและประยุกต์ได้'
  ];
  clos.forEach(c => {
    const item = form.addMultipleChoiceItem()
      .setTitle(String(c.Code || 'CLO') + '  ' + String(c.Description || ''))
      .setChoiceValues(selfEvaluationChoices)
      .setRequired(true);
    const friendly = clean_(c.StudentFriendlyDescription,1500);
    if (friendly) {
      try { item.setHelpText(friendly); } catch (e) {}
    }
  });
  // Reflection / course feedback is tailored to the evaluation period.
  if (point === 'START') {
    form.addParagraphTextItem()
      .setTitle('Reflection / สิ่งที่อยากสะท้อนเพิ่มเติม')
      .setHelpText('เมื่อดูผลลัพธ์การเรียนรู้ของรายวิชาแล้ว คุณคิดว่าตนเองมีพื้นฐานอะไรอยู่แล้ว เรื่องใดที่ยังไม่มั่นใจ และจะเตรียมหรือปรับตัวอย่างไรเพื่อให้เรียนวิชานี้ได้ดี')
      .setRequired(false);
  } else if (point === 'MIDTERM') {
    form.addParagraphTextItem()
      .setTitle('Reflection / สิ่งที่อยากสะท้อนเพิ่มเติม')
      .setHelpText('เมื่อเรียนมาถึงกลางภาค คุณคิดว่าตนเองพัฒนาเรื่องใดได้ดีขึ้นแล้ว เรื่องใดที่ยังไม่มั่นใจหรือเป็นอุปสรรค และจะปรับวิธีการเรียนของตนเองอย่างไรในช่วงที่เหลือของภาคการศึกษา')
      .setRequired(false);
  } else if (point === 'END') {
    const agreementChoices = [
      '1 — ไม่เห็นด้วยอย่างยิ่ง',
      '2 — ไม่เห็นด้วย',
      '3 — ปานกลาง',
      '4 — เห็นด้วย',
      '5 — เห็นด้วยอย่างยิ่ง'
    ];
    const ratingQuestions = [
      'ฉันเข้าใจว่ารายวิชานี้คาดหวังให้ฉันเรียนรู้หรือสามารถทำอะไรได้เมื่อจบรายวิชา',
      'กิจกรรมการเรียนรู้ในรายวิชาช่วยให้ฉันพัฒนาความรู้และทักษะตามผลลัพธ์การเรียนรู้ที่กำหนดไว้',
      'งาน แบบฝึกหัด โครงงาน หรือการสอบ สอดคล้องกับสิ่งที่รายวิชาต้องการให้ฉันเรียนรู้',
      'Feedback ที่ได้รับช่วยให้ฉันรู้ว่าตนเองทำได้ดีตรงไหนและควรปรับปรุงอะไร',
      'ปริมาณและช่วงเวลาของงานหรือการประเมินเหมาะสมกับการเรียนรู้ในรายวิชานี้',
      'เอกสาร เครื่องมือ ระบบ ห้องปฏิบัติการ หรือทรัพยากรที่ใช้ในรายวิชาสนับสนุนการเรียนรู้ได้เพียงพอ',
      'เมื่อมีปัญหาในการเรียน ฉันรู้ว่าจะขอความช่วยเหลือหรือคำแนะนำจากผู้สอน ผู้ช่วยสอน หรือช่องทางที่เกี่ยวข้องได้อย่างไร',
      'เมื่อจบรายวิชานี้ ฉันสามารถอธิบายได้ว่าตนเองพัฒนาความรู้หรือความสามารถอะไรขึ้นบ้าง'
    ];
    ratingQuestions.forEach((q, i) => {
      form.addMultipleChoiceItem()
        .setTitle(String(i + 1) + '. ' + q)
        .setChoiceValues(agreementChoices)
        .setRequired(true);
    });
    const openQuestions = [
      'สิ่งใดในรายวิชานี้ช่วยให้คุณเรียนรู้ได้ดีที่สุด เพราะอะไร?',
      'อะไรเป็นอุปสรรคต่อการเรียนรู้ของคุณมากที่สุดในรายวิชานี้?',
      'หากรายวิชานี้เปิดอีกครั้ง คุณเสนอให้ปรับอะไรหนึ่งเรื่องเพื่อช่วยให้นักศึกษาเรียนรู้ได้ดีขึ้น?',
      'หากต้องแนะนำรุ่นน้องที่กำลังจะเรียนวิชานี้ คุณอยากบอกอะไรเกี่ยวกับวิชานี้ และควรเตรียมตัวอย่างไรเพื่อให้เรียนได้ดี?'
    ];
    openQuestions.forEach((q, i) => {
      form.addParagraphTextItem()
        .setTitle(String(i + 9) + '. ' + q)
        .setRequired(false);
    });
  }

  // Google Apps Script exposes collect-email / edit-response / one-response settings,
  // but not the Forms UI's "Send responders a copy: Always" switch. Install a
  // form-submit trigger so generated self-evaluation forms provide the same behavior.
  try {
    ScriptApp.newTrigger('sendSelfEvaluationResponseCopy')
      .forForm(form)
      .onFormSubmit()
      .create();
  } catch (e) {
    console.warn('Unable to install self-evaluation response-copy trigger: ' + e);
  }

  const responseSs = SpreadsheetApp.create(title + ' — Responses');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSs.getId());
  SpreadsheetApp.flush();

  tryMoveGeneratedSelfEvaluationFiles_([form.getId(), responseSs.getId()]);

  const prefix = selfEvaluationFormConfigPrefix_(point);
  setConfig_(prefix + 'FORM_ID', form.getId());
  setConfig_(prefix + 'RESPONSE_SPREADSHEET_ID', responseSs.getId());
  setConfig_(prefix + 'CREATED_AT', new Date().toISOString());
  setConfig_(prefix + 'CREATED_BY', staff.email);
  setConfig_('APP_VERSION', APP.version);
  audit_('CREATE_SELF_EVALUATION_FORM','CLOSelfEvaluation',point,{
    actor:staff.email, formId:form.getId(), responseSpreadsheetId:responseSs.getId(), replaced:!!existing.exists, cloCount:clos.length
  });
  resetRequestDataCache_();

  return {ok:true, created:true, existing:false, replaced:!!existing.exists, form:selfEvaluationFormRecord_(point)};
}

/* ---------- CLO self-evaluation (indirect evidence) ---------- */

/**
 * Sends the respondent a copy of a generated self-evaluation response.
 * This is the programmatic equivalent of the Google Forms UI option
 * "Send responders a copy of their response: Always", which FormApp does not expose.
 */
function sendSelfEvaluationResponseCopy(e) {
  try {
    if (!e || !e.response || !e.source) return;
    const email = clean_(e.response.getRespondentEmail(), 320);
    if (!email) return;

    const formTitle = clean_(e.source.getTitle(), 500) || 'CLO Self-Evaluation';
    const submittedAt = e.response.getTimestamp();
    const rows = [];
    (e.response.getItemResponses() || []).forEach(ir => {
      const item = ir.getItem();
      const title = clean_(item && item.getTitle ? item.getTitle() : '', 1200);
      let answer = ir.getResponse();
      if (Array.isArray(answer)) answer = answer.join(', ');
      answer = clean_(answer, 4000);
      if (title) rows.push(title + '\n' + (answer || '—'));
    });

    const subject = 'สำเนาคำตอบ: ' + formTitle;
    const body = [
      'นี่คือสำเนาคำตอบแบบประเมินตนเองที่คุณส่ง',
      'แบบฟอร์ม: ' + formTitle,
      'ส่งเมื่อ: ' + Utilities.formatDate(submittedAt || new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'd MMM yyyy HH:mm'),
      '',
      rows.join('\n\n'),
      '',
      'หากแบบฟอร์มอนุญาตให้แก้ไขคำตอบ คุณสามารถใช้ลิงก์แก้ไขที่ Google Forms แสดงหลังส่งแบบฟอร์มได้'
    ].join('\n');

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: APP.name
    });
  } catch (err) {
    console.error('sendSelfEvaluationResponseCopy failed: ' + err);
  }
}

function normalizeSelfEvaluationPoint_(v) {
  const raw=String(v||'').trim();
  if(!raw) throw new Error('Evaluation period is blank.');

  const s=raw.toUpperCase()
    .replace(/[()\[\]{}]/g,' ')
    .replace(/[\s_-]+/g,' ')
    .trim();

  // Exact normalized English values.
  if(['START','BEGIN','BEGINNING','PRE','START OF TERM','BEGINNING OF TERM','PRE TERM'].includes(s)) return 'START';
  if(['MID','MIDTERM','MID TERM','MIDDLE','MIDDLE OF TERM'].includes(s)) return 'MIDTERM';
  if(['END','FINAL','ENDTERM','END TERM','END OF TERM','POST','POST TERM'].includes(s)) return 'END';

  // Bilingual / descriptive labels such as:
  // "ก่อนเรียน (Start)", "กลางภาค (Midterm)", "หลังเรียน (End)".
  if(/(^|\s)(START|BEGIN|BEGINNING|PRE)(\s|$)/.test(s) ||
     /ก่อนเรียน|ก่อนเริ่มเรียน|ต้นภาค|เริ่มภาค/.test(raw)) return 'START';

  if(/(^|\s)(MID|MIDTERM|MIDDLE)(\s|$)/.test(s) ||
     /กลางภาค|กลางเทอม|ระหว่างภาค/.test(raw)) return 'MIDTERM';

  if(/(^|\s)(END|FINAL|POST)(\s|$)/.test(s) ||
     /หลังเรียน|ปลายภาค|ปลายเทอม|สิ้นภาค|จบภาค/.test(raw)) return 'END';

  throw new Error('Evaluation period must map to Start, Midterm, or End of term.');
}

function selfEvaluationPointOrder_(p) {
  return p==='START'?1:(p==='MIDTERM'?2:(p==='END'?3:99));
}

function studentMatchIndexes_() {
  const students=listStudentsForStaff_().filter(s=>s.Active);
  const byEmail={},byId={};
  students.forEach(s=>{
    const email=String(s.Email||'').trim().toLowerCase();
    const sid=String(s.StudentId||'').trim();
    if(email)byEmail[email]=s;
    if(sid)byId[sid]=s;
  });
  return {students,byEmail,byId};
}


function parseSelfEvaluationRating_(rawValue) {
  if(rawValue === null || rawValue === undefined || rawValue === '') return null;

  if(typeof rawValue === 'number'){
    return (rawValue>=1 && rawValue<=5) ? rawValue : null;
  }

  const s=String(rawValue).trim();
  if(!s) return null;

  // Plain value: "1", "2.0", etc.
  const direct=Number(s);
  if(Number.isFinite(direct) && direct>=1 && direct<=5) return direct;

  // Google Form multiple-choice / dropdown labels often look like:
  // "1 — ยังไม่บรรลุ ..."
  // "3 - ทำได้..."
  // "5: เชี่ยวชาญ..."
  // Accept a leading 1-5 only when it is followed by a normal separator,
  // whitespace, or the end of the string.
  const m=s.match(/^([1-5])(?:\s*(?:[-–—:.)]|$)|\s+)/);
  if(m) return Number(m[1]);

  return null;
}

function buildSelfEvaluationPreview_(records, options) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureCurrentSchema_();

  const matchBy=options.matchBy==='studentId'?'studentId':'email';
  const matchColumn=clean_(options.matchColumn,200);
  const periodColumn=clean_(options.periodColumn,200);
  const periodFilterRaw=clean_(options.periodFilter,40);
  const periodFilter=periodFilterRaw?normalizeSelfEvaluationPoint_(periodFilterRaw):'';
  const fixedPoint=periodColumn?'':normalizeSelfEvaluationPoint_(options.evaluationPoint);
  const reflectionColumn=clean_(options.reflectionColumn,200);
  const timestampColumn=clean_(options.timestampColumn,200);
  const source=String(options.source||'IMPORT').toUpperCase();
  const cloColumns=options.cloColumns||{};

  if(!matchColumn)throw new Error('Select the student identifier column.');

  const clos=listCLOs_().filter(c=>c.Active).sort((a,b)=>Number(a.Sequence)-Number(b.Sequence));
  const selectedClos=clos.filter(c=>clean_(cloColumns[c.CLOID],200));
  if(!selectedClos.length)throw new Error('Map at least one CLO to a rating column.');

  const idx=studentMatchIndexes_();
  const rows=[];
  let matchedStudents=0,unmatchedRows=0,validRatings=0,invalidRatings=0,invalidPeriods=0,filteredOutRows=0;
  const periodStats={
    START:{rows:0,students:{},ratings:0},
    MIDTERM:{rows:0,students:{},ratings:0},
    END:{rows:0,students:{},ratings:0}
  };

  (records||[]).slice(0,3000).forEach((rec,i)=>{
    const rawIdentity=String(rec[matchColumn]||'').trim();
    const lookup=matchBy==='studentId' ? idx.byId[rawIdentity] : idx.byEmail[rawIdentity.toLowerCase()];

    let point=fixedPoint;
    let periodRaw='';
    if(periodColumn){
      periodRaw=String(rec[periodColumn]||'').trim();
      try{point=normalizeSelfEvaluationPoint_(periodRaw);}catch(e){point='';invalidPeriods++;}
      if(point && periodFilter && point!==periodFilter){filteredOutRows++;return;}
    }

    if(!lookup){
      unmatchedRows++;
      rows.push({
        sourceRow:Number(rec._sourceRow||i+2),sourceIdentity:rawIdentity,matched:false,
        studentKey:'',studentLabel:'',ratings:[],reflection:reflectionColumn?String(rec[reflectionColumn]||''):'',
        submittedAt:timestampColumn?String(rec[timestampColumn]||''):'',evaluationPoint:point,
        periodRaw:periodRaw,periodValid:!!point,source:source
      });
      return;
    }

    const ratings=[];
    selectedClos.forEach(c=>{
      const col=clean_(cloColumns[c.CLOID],200);
      const raw=String(rec[col] == null ? '' : rec[col]).trim();
      if(raw==='')return;
      const rating=parseSelfEvaluationRating_(raw);
      if(rating!==null){ratings.push({cloId:c.CLOID,code:c.Code,rating:rating,rawLabel:raw});validRatings++;}
      else{ratings.push({cloId:c.CLOID,code:c.Code,rating:'',invalidRaw:raw});invalidRatings++;}
    });

    matchedStudents++;
    if(point){
      periodStats[point].rows++;
      periodStats[point].students[String(lookup.StudentKey)]=true;
      periodStats[point].ratings+=ratings.filter(r=>r.rating!=='').length;
    }

    rows.push({
      sourceRow:Number(rec._sourceRow||i+2),sourceIdentity:rawIdentity,matched:true,
      studentKey:lookup.StudentKey,studentLabel:String(lookup.StudentId||'')+' — '+String(lookup.Name||''),
      ratings:ratings,reflection:reflectionColumn?String(rec[reflectionColumn]||''):'',
      submittedAt:timestampColumn?String(rec[timestampColumn]||''):'',evaluationPoint:point,
      periodRaw:periodRaw,periodValid:!!point,source:source
    });
  });

  const periods={};
  ['START','MIDTERM','END'].forEach(p=>periods[p]={
    rows:periodStats[p].rows,
    students:Object.keys(periodStats[p].students).length,
    ratings:periodStats[p].ratings
  });

  return {
    evaluationPoint:periodColumn?(periodFilter||'MIXED'):fixedPoint,
    periodColumn:periodColumn,
    periodFilter:periodFilter,
    clos:selectedClos.map(c=>({cloId:c.CLOID,code:c.Code,description:c.Description})),
    rows:rows,
    periods:periods,
    summary:{matchedStudents,unmatchedRows,validRatings,invalidRatings,invalidPeriods,filteredOutRows,totalRows:rows.length}
  };
}

function previewUploadedSelfEvaluationRows(payload) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const records=Array.isArray(payload.records)?payload.records:[];
  if(!records.length)throw new Error('No self-evaluation rows were found in the uploaded file.');
  return buildSelfEvaluationPreview_(records,{
    evaluationPoint:payload.evaluationPoint,
    periodColumn:payload.periodColumn,
    periodFilter:payload.periodFilter,
    matchBy:payload.matchBy,
    matchColumn:payload.matchColumn,
    cloColumns:payload.cloColumns||{},
    reflectionColumn:payload.reflectionColumn,
    timestampColumn:payload.timestampColumn,
    source:'CSV'
  });
}

function previewGoogleSelfEvaluationImport(options) {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const id=spreadsheetIdFromRef_(options.spreadsheetRef);
  if(!id)throw new Error('Enter a valid Google Spreadsheet URL or ID.');

  const ss=SpreadsheetApp.openById(id);
  const sheetName=clean_(options.sheetName,200);
  const sh=sheetName?ss.getSheetByName(sheetName):ss.getSheets()[0];
  if(!sh)throw new Error('Response sheet not found.');
  if(sh.getLastRow()<2){
    return {evaluationPoint:options.periodColumn?(options.periodFilter||'MIXED'):normalizeSelfEvaluationPoint_(options.evaluationPoint),periodFilter:options.periodFilter||'',clos:[],rows:[],periods:{START:{rows:0,students:0,ratings:0},MIDTERM:{rows:0,students:0,ratings:0},END:{rows:0,students:0,ratings:0}},summary:{matchedStudents:0,unmatchedRows:0,validRatings:0,invalidRatings:0,invalidPeriods:0,filteredOutRows:0,totalRows:0}};
  }

  const values=sh.getDataRange().getDisplayValues();
  const headers=values[0].map(x=>String(x).trim());
  const records=values.slice(1).map((row,i)=>{
    const obj={_sourceRow:i+2};
    headers.forEach((h,j)=>obj[h]=row[j]);
    return obj;
  });

  return buildSelfEvaluationPreview_(records,{
    evaluationPoint:options.evaluationPoint,
    periodColumn:options.periodColumn,
    periodFilter:options.periodFilter,
    matchBy:options.matchBy,
    matchColumn:options.matchColumn,
    cloColumns:options.cloColumns||{},
    reflectionColumn:options.reflectionColumn,
    timestampColumn:options.timestampColumn,
    source:'GOOGLE_FORM'
  });
}

function applySelfEvaluationImport(previewRows, evaluationPoint) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureCurrentSchema_();
  if(!Array.isArray(previewRows))throw new Error('Invalid self-evaluation preview rows.');

  const fallbackPoint=String(evaluationPoint||'').toUpperCase()==='MIXED'?'':normalizeSelfEvaluationPoint_(evaluationPoint);
  const sh=table_(APP.tables.CLO_SELF_EVALUATION);
  const now=new Date();
  const existing=rowsToObjects_(sh).map(stripRow_);
  const byKey={};
  existing.forEach((r,i)=>{
    const key=[String(r.StudentKey||''),String(r.EvaluationPoint||''),String(r.CLOID||'')].join('|');
    if(key!=='||')byKey[key]=i;
  });

  let importedRatings=0,skippedInvalidPeriod=0;
  const studentsImported={},periodCounts={START:0,MIDTERM:0,END:0};

  previewRows.slice(0,3000).forEach(row=>{
    if(!row.matched||!row.studentKey)return;
    let point='';
    try{point=row.evaluationPoint?normalizeSelfEvaluationPoint_(row.evaluationPoint):fallbackPoint;}catch(e){point='';}
    if(!point){skippedInvalidPeriod++;return;}

    const reflection=clean_(row.reflection||'',2000);
    const parsedSubmittedAt=row.submittedAt?new Date(row.submittedAt):now;
    const safeSubmittedAt=isNaN(parsedSubmittedAt.getTime())?now:parsedSubmittedAt;
    const source=clean_(row.source||'IMPORT',40);
    let rowImported=0;

    (row.ratings||[]).forEach(r=>{
      const rating=parseSelfEvaluationRating_(r.rating!==''&&r.rating!=null?r.rating:r.rawLabel);
      if(!r.cloId||rating===null)return;
      const key=[String(row.studentKey),point,String(r.cloId)].join('|');
      const obj={StudentKey:row.studentKey,EvaluationPoint:point,CLOID:r.cloId,Rating:rating,
        Reflection:reflection,SubmittedAt:safeSubmittedAt,Source:source,ImportedAt:now,ImportedBy:staff.email};
      if(byKey[key]!==undefined)existing[byKey[key]]=obj;
      else{byKey[key]=existing.length;existing.push(obj);}
      importedRatings++;rowImported++;
      studentsImported[String(row.studentKey)]=true;
    });
    if(rowImported)periodCounts[point]+=rowImported;
  });

  rewriteObjects_(sh,existing);
  audit_('IMPORT_CLO_SELF_EVALUATION','CLOSelfEvaluation',fallbackPoint||'MIXED',{
    students:Object.keys(studentsImported).length,ratings:importedRatings,periodCounts,actor:staff.email
  });
  return {ok:true,evaluationPoint:fallbackPoint||'MIXED',students:Object.keys(studentsImported).length,
    ratings:importedRatings,periodCounts:periodCounts,skippedInvalidPeriod:skippedInvalidPeriod};
}

function selfEvaluationStats_(rows) {
  const points=['START','MIDTERM','END'];
  const out={};
  points.forEach(p=>{
    const vals=(rows||[]).filter(r=>String(r.EvaluationPoint||'')===p)
      .map(r=>Number(r.Rating)).filter(v=>Number.isFinite(v));
    const n=vals.length;
    const average=n?vals.reduce((a,b)=>a+b,0)/n:null;
    let sd=null;
    if(n>1){
      const ss=vals.reduce((sum,v)=>sum+Math.pow(v-average,2),0);
      sd=Math.sqrt(ss/(n-1));
    }else if(n===1){
      sd=0;
    }
    out[p]={count:n,average:average,sd:sd};
  });
  out.changeStartEnd=(out.START.average!=null&&out.END.average!=null)?out.END.average-out.START.average:null;
  return out;
}

/* ---------- CLO attainment analytics ---------- */

function defaultCloAttainmentPolicy_() {
  return {Policy:'LATEST',Threshold:3,MinCount:2,Active:true};
}

function getCloAttainmentPolicy_(cloId) {
  ensureCurrentSchema_();
  const sh=table_(APP.tables.CLO_ATTAINMENT_POLICIES);
  const found=findObject_(sh,'CLOID',cloId);
  const d=defaultCloAttainmentPolicy_();
  if(!found) return Object.assign({CLOID:cloId},d);
  return {
    CLOID:cloId,
    Policy:String(found.Policy||d.Policy).toUpperCase(),
    Threshold:Number(found.Threshold||d.Threshold),
    MinCount:Number(found.MinCount||d.MinCount),
    Active:found.Active===''?true:bool_(found.Active)
  };
}

function listCloAttainmentWeights_(cloId) {
  ensureCurrentSchema_();
  return rowsToObjects_(table_(APP.tables.CLO_ATTAINMENT_WEIGHTS))
    .filter(r=>String(r.CLOID||'')===String(cloId||''))
    .map(r=>({CLOID:r.CLOID,ActivityID:r.ActivityID,Weight:Number(r.Weight)||0}));
}

function saveCloAttainmentPolicy(cloId, payload) {
  const staff=requireStaff_(['OWNER','INSTRUCTOR']);
  ensureCurrentSchema_();
  const policy=String((payload&&payload.policy)||'LATEST').toUpperCase();
  const allowed=['LATEST','AVERAGE','WEIGHTED','MIN_COUNT'];
  if(!allowed.includes(policy)) throw new Error('Invalid CLO aggregation policy.');

  const threshold=Number((payload&&payload.threshold)!=null?payload.threshold:3);
  if(!(threshold>=1 && threshold<=5)) throw new Error('Threshold must be between 1 and 5.');

  const minCount=Math.max(1,Math.floor(Number((payload&&payload.minCount)||2)));
  const now=new Date();

  upsertComposite_(table_(APP.tables.CLO_ATTAINMENT_POLICIES),['CLOID'],[cloId],{
    CLOID:cloId,Policy:policy,Threshold:threshold,MinCount:minCount,Active:true,
    UpdatedBy:staff.email,UpdatedAt:now
  });

  // Replace weights for this CLO
  const wSh=table_(APP.tables.CLO_ATTAINMENT_WEIGHTS);
  const all=rowsToObjects_(wSh);
  const keep=all.filter(r=>String(r.CLOID||'')!==String(cloId||'')).map(stripRow_);
  const weights=Array.isArray(payload&&payload.weights)?payload.weights:[];
  weights.forEach(w=>{
    const activityId=String(w.activityId||'').trim();
    const weight=Number(w.weight)||0;
    if(activityId && weight>0){
      keep.push({CLOID:cloId,ActivityID:activityId,Weight:weight,UpdatedBy:staff.email,UpdatedAt:now});
    }
  });
  rewriteObjects_(wSh,keep);

  audit_('SAVE_CLO_ATTAINMENT_POLICY','CLO',cloId,{policy,threshold,minCount,actor:staff.email});
  return {ok:true};
}

function evaluateCloStudent_(studentKey, cloId) {
  const maps=rowsToObjects_(table_(APP.tables.ACTIVITY_CLOS))
    .filter(m=>String(m.CLOID||'')===String(cloId||''));
  const evidence=rowsToObjects_(table_(APP.tables.CLO_EVIDENCE));
  const activities=listActivities_();
  const activityById={};
  activities.forEach(a=>activityById[String(a.ActivityID||'')]=a);

  const points=evidence
    .filter(e=>String(e.CLOID||'')===String(cloId||'') &&
      String(e.TargetType||'')==='STUDENT' &&
      String(e.TargetID||'')===String(studentKey||''))
    .map(e=>{
      const a=activityById[String(e.ActivityID||'')]||{};
      return {
        activityId:String(e.ActivityID||''),
        title:a.Title||String(e.ActivityID||''),
        sequence:Number(a.Sequence)||0,
        level:Number(e.Level)||0
      };
    })
    .filter(p=>p.level>=1 && p.level<=5)
    .sort((a,b)=>a.sequence-b.sequence);

  const policy=getCloAttainmentPolicy_(cloId);
  const weights=listCloAttainmentWeights_(cloId);
  const weightByActivity={};
  weights.forEach(w=>weightByActivity[String(w.ActivityID)]=Number(w.Weight)||0);

  let aggregate=null, attained=false;
  if(points.length){
    if(policy.Policy==='LATEST'){
      aggregate=points[points.length-1].level;
      attained=aggregate>=policy.Threshold;
    }else if(policy.Policy==='AVERAGE'){
      aggregate=points.reduce((s,p)=>s+p.level,0)/points.length;
      attained=aggregate>=policy.Threshold;
    }else if(policy.Policy==='WEIGHTED'){
      let num=0,den=0;
      points.forEach(p=>{
        const w=weightByActivity[p.activityId]||0;
        if(w>0){num+=p.level*w;den+=w;}
      });
      if(den>0){
        aggregate=num/den;
        attained=aggregate>=policy.Threshold;
      }
    }else if(policy.Policy==='MIN_COUNT'){
      aggregate=points.filter(p=>p.level>=policy.Threshold).length;
      attained=aggregate>=policy.MinCount;
    }
  }

  return {studentKey,cloId,policy,points,aggregate,attained};
}

function getCloAttainmentAnalytics() {
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  ensureCurrentSchema_();

  // Load each table ONCE. The previous implementation called evaluateCloStudent_()
  // for every student x CLO, and that function reread several sheets each time.
  const snap=courseSnapshot_({
    clos:true,students:true,activityClos:true,activities:true,
    evidence:true,policies:true,weights:true,selfEvaluations:true
  });
  const clos=normalizedClosFromRows_(snap.clos).filter(c=>c.Active)
    .sort((a,b)=>Number(a.Sequence)-Number(b.Sequence));
  const students=normalizedStudentsFromRows_(snap.students).filter(s=>s.Active);
  const maps=snap.activityClos||[];
  const activities=normalizedActivitiesFromRows_(snap.activities||[]);
  const evidence=snap.evidence||[];
  const policies=snap.policies||[];
  const weights=snap.weights||[];
  const selfEvaluations=snap.selfEvaluations||[];

  const activityById={};
  activities.forEach(a=>activityById[String(a.ActivityID||'')]=a);

  const policyByClo={};
  policies.forEach(r=>{
    const cid=String(r.CLOID||'');
    if(!cid)return;
    policyByClo[cid]={
      CLOID:cid,
      Policy:String(r.Policy||'LATEST').toUpperCase(),
      Threshold:Number(r.Threshold||3),
      MinCount:Number(r.MinCount||2),
      Active:r.Active===''?true:bool_(r.Active)
    };
  });

  const weightsByClo={};
  weights.forEach(r=>{
    const cid=String(r.CLOID||'');
    const aid=String(r.ActivityID||'');
    if(!cid||!aid)return;
    if(!weightsByClo[cid])weightsByClo[cid]={};
    weightsByClo[cid][aid]=Number(r.Weight)||0;
  });

  // evidenceByCloStudent[CLOID][StudentKey] = ordered evidence points
  const evidenceByCloStudent={};
  evidence.forEach(e=>{
    if(String(e.TargetType||'')!=='STUDENT')return;
    const cid=String(e.CLOID||'');
    const sk=String(e.TargetID||'');
    const aid=String(e.ActivityID||'');
    const level=Number(e.Level)||0;
    if(!cid||!sk||level<1||level>5)return;

    if(!evidenceByCloStudent[cid])evidenceByCloStudent[cid]={};
    if(!evidenceByCloStudent[cid][sk])evidenceByCloStudent[cid][sk]=[];

    const a=activityById[aid]||{};
    evidenceByCloStudent[cid][sk].push({
      activityId:aid,
      title:a.Title||aid,
      sequence:Number(a.Sequence)||0,
      level:level
    });
  });

  Object.keys(evidenceByCloStudent).forEach(cid=>{
    Object.keys(evidenceByCloStudent[cid]).forEach(sk=>{
      evidenceByCloStudent[cid][sk].sort((a,b)=>a.sequence-b.sequence);
    });
  });

  const linkedByClo={};
  maps.forEach(m=>{
    const cid=String(m.CLOID||'');
    const aid=String(m.ActivityID||'');
    if(!cid||!aid)return;
    if(!linkedByClo[cid])linkedByClo[cid]=[];
    if(!linkedByClo[cid].includes(aid))linkedByClo[cid].push(aid);
  });

  function policyFor_(cloId){
    return policyByClo[cloId] || {
      CLOID:cloId,Policy:'LATEST',Threshold:3,MinCount:2,Active:true
    };
  }

  function evaluate_(points, policy, weightMap){
    let aggregate=null, attained=false;
    if(!points.length)return {aggregate:null,attained:false};

    if(policy.Policy==='LATEST'){
      aggregate=points[points.length-1].level;
      attained=aggregate>=policy.Threshold;
    }else if(policy.Policy==='AVERAGE'){
      aggregate=points.reduce((s,p)=>s+p.level,0)/points.length;
      attained=aggregate>=policy.Threshold;
    }else if(policy.Policy==='WEIGHTED'){
      let num=0,den=0;
      points.forEach(p=>{
        const w=Number((weightMap||{})[p.activityId]||0);
        if(w>0){num+=p.level*w;den+=w;}
      });
      if(den>0){
        aggregate=num/den;
        attained=aggregate>=policy.Threshold;
      }
    }else if(policy.Policy==='MIN_COUNT'){
      aggregate=points.filter(p=>p.level>=policy.Threshold).length;
      attained=aggregate>=policy.MinCount;
    }

    return {aggregate:aggregate,attained:attained};
  }

  const selfEvalByCloStudent={};
  selfEvaluations.forEach(r=>{
    const cid=String(r.CLOID||'');
    const sk=String(r.StudentKey||'');
    const point=String(r.EvaluationPoint||'').toUpperCase();
    const rating=Number(r.Rating);
    if(!cid||!sk||!['START','MIDTERM','END'].includes(point)||!Number.isFinite(rating))return;
    if(!selfEvalByCloStudent[cid])selfEvalByCloStudent[cid]={};
    if(!selfEvalByCloStudent[cid][sk])selfEvalByCloStudent[cid][sk]={};
    selfEvalByCloStudent[cid][sk][point]={
      rating:rating,
      reflection:String(r.Reflection||''),
      submittedAt:dateIso_(r.SubmittedAt)
    };
  });

  const summary=clos.map(c=>{
    const cloId=String(c.CLOID||'');
    const policy=policyFor_(cloId);
    const weightMap=weightsByClo[cloId]||{};
    const linkedIds=(linkedByClo[cloId]||[]).slice();
    linkedIds.sort((x,y)=>{
      const ax=activityById[x]||{}, ay=activityById[y]||{};
      return Number(ax.Sequence||0)-Number(ay.Sequence||0);
    });

    const linkedActivities=linkedIds.map(id=>{
      const a=activityById[id]||{};
      return {activityId:id,title:a.Title||id,sequence:a.Sequence||0};
    });

    const studentRows=students.map(s=>{
      const sk=String(s.StudentKey||'');
      const points=((evidenceByCloStudent[cloId]||{})[sk]||[]);
      const result=evaluate_(points,policy,weightMap);
      return {
        studentKey:sk,
        studentId:s.StudentId,
        name:s.Name,
        alias:s.Alias,
        teamCode:s.TeamCode||'',
        points:points,
        aggregate:result.aggregate,
        attained:result.attained,
        selfRatings:(selfEvalByCloStudent[cloId]||{})[sk]||{}
      };
    });

    const withEvidence=studentRows.filter(r=>r.points.length>0);
    const attainedCount=withEvidence.filter(r=>r.attained).length;

    let avgLevel=null;
    if(withEvidence.length){
      const vals=withEvidence.map(r=>{
        if(policy.Policy==='MIN_COUNT'){
          return r.points.reduce((s,p)=>s+p.level,0)/r.points.length;
        }
        return r.aggregate==null?null:Number(r.aggregate);
      }).filter(v=>v!=null && !isNaN(v));
      avgLevel=vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:null;
    }

    return {
      cloId:cloId,
      code:c.Code,
      description:c.Description,
      policy:policy,
      linkedActivities:linkedActivities,
      weights:Object.keys(weightMap).map(aid=>({
        CLOID:cloId,ActivityID:aid,Weight:weightMap[aid]
      })),
      studentRows:studentRows,
      evidenceStudents:withEvidence.length,
      attainedCount:attainedCount,
      notAttainedCount:Math.max(0,withEvidence.length-attainedCount),
      attainmentPct:withEvidence.length?attainedCount*100/withEvidence.length:null,
      avgLevel:avgLevel,
      selfEvaluation:selfEvaluationStats_(
        selfEvaluations.filter(r=>String(r.CLOID||'')===cloId)
      )
    };
  });

  // Serialize here so client always receives plain JSON-safe values.
  return JSON.parse(JSON.stringify({summary:summary}));
}

/* ---------- Student / staff views ---------- */


function studentDashboardJsonSafe_(student) {
  const dashboard=getStudentDashboard_(student);
  return JSON.parse(JSON.stringify(dashboard));
}

function previewStudentDashboard(studentKey) {
  const actor = requireStaff_(['OWNER','INSTRUCTOR']);
  const key = String(studentKey || '').trim();
  if (!key) throw new Error('Student key is required.');

  const student = listStudentsForStaff_().find(s =>
    String(s.StudentKey || '') === key && s.Active
  );
  if (!student) throw new Error('Student not found or inactive.');

  audit_('PREVIEW_STUDENT_PORTAL','Student',key,{actor:actor.email});
  // Use the exact same JSON-safe payload shape as the real Student Portal.
  return studentDashboardJsonSafe_(student);
}

function getStudentDashboard(studentToken) {
  const st=resolveStudent_(studentToken);
  if(!st) throw new Error('Unable to identify an enrolled student.');
  return studentDashboardJsonSafe_(st);
}

function studentFlagTrue_(v) {
  if (v === true || v === 1) return true;
  const s = String(v == null ? '' : v).trim().toUpperCase();
  return ['TRUE','1','YES','Y','PUBLISHED','ACTIVE'].includes(s);
}

function listStudentVisibleActivities_(snapshot) {
  ensureCurrentSchema_();
  const rawRows=snapshot && snapshot.activities
    ? snapshot.activities
    : rowsToObjects_(table_(APP.tables.ACTIVITIES));

  return normalizedActivitiesFromRows_(rawRows)
    .filter(a => {
      const raw=rawRows.find(r=>String(r.ActivityID||'')===String(a.ActivityID||''))||{};
      const level=normalizeActivityVisibilityLevel_(a.StudentVisibilityLevel,a.Published);
      let active=bool_(a.Active);
      if(!active && (raw.Active==='' || raw.Active==null)) active=true;
      return level!=='HIDDEN' && active;
    })
    .sort(sortSequence_);
}


function getStudentCourseworkVisibilityDebug_() {
  const raw = rowsToObjects_(table_(APP.tables.ACTIVITIES));
  return raw.map(a => ({
    activityId:a.ActivityID,
    title:a.Title,
    publishedRaw:a.Published,
    activeRaw:a.Active,
    published:studentFlagTrue_(a.Published),
    active:(a.Active === '' || a.Active == null) ? true : studentFlagTrue_(a.Active)
  }));
}

function getStudentDashboard_(student) {
  ensureCurrentSchema_();

  const snap=courseSnapshot_({
    activities:true,results:true,evidence:true,activityClos:true,clos:true,
    teamHistory:true,attendanceSessions:true,attendance:true,selfEvaluations:true
  });

  const activities=listStudentVisibleActivities_(snap);
  const clos=normalizedClosFromRows_(snap.clos||[])
    .filter(c=>c.Active)
    .sort((a,b)=>Number(a.Sequence)-Number(b.Sequence));

  const idx=buildCourseIndexes_(snap);
  const studentKey=String(student.StudentKey||'');
  const individualTarget='STUDENT|'+studentKey;

  function teamFor_(activityId){
    const key=studentKey+'|'+String(activityId||'');
    if(idx.teamByStudentActivity[key]!==undefined){
      return idx.teamByStudentActivity[key];
    }
    return String(student.TeamCode||'');
  }

  function resultFor_(activityId){
    const aid=String(activityId||'');
    const direct=idx.resultsByActivityTarget[aid+'|STUDENT|'+studentKey];
    if(direct)return direct;
    const team=teamFor_(aid);
    return team?idx.resultsByActivityTarget[aid+'|TEAM|'+team]||null:null;
  }

  function evidenceFor_(activityId,cloId){
    const aid=String(activityId||'');
    const cid=String(cloId||'');
    const direct=idx.evidenceByActivityCloTarget[aid+'|'+cid+'|STUDENT|'+studentKey];
    if(direct)return direct;
    const team=teamFor_(aid);
    return team?idx.evidenceByActivityCloTarget[aid+'|'+cid+'|TEAM|'+team]||null:null;
  }

  const activityViews=activities.map(a=>{
    const activityId=String(a.ActivityID||'');
    const result=resultFor_(activityId);
    const cloIds=idx.cloIdsByActivity[activityId]||[];
    const cloLevels={};

    cloIds.forEach(cid=>{
      const e=evidenceFor_(activityId,cid);
      if(e && e.Level!=='' && e.Level!=null)cloLevels[cid]=Number(e.Level);
    });

    const visibilityLevel=normalizeActivityVisibilityLevel_(a.StudentVisibilityLevel,a.Published);
    const canSeeSubmission=['SUBMISSION','SCORE','SCORE_CLO'].includes(visibilityLevel);
    const canSeeScore=visibilityLevel==='SCORE'||visibilityLevel==='SCORE_CLO';
    const canSeeClo=visibilityLevel==='SCORE_CLO';

    return {
      activityId:activityId,
      title:a.Title||'',
      shortDescription:a.ShortDescription||'',
      detailUrl:safeActivityDetailUrl_(a.DetailUrl),
      category:a.Category||'',
      mode:a.Mode||'Individual',
      purpose:a.Purpose||'Formative',
      maxScore:a.MaxScore,
      weight:a.Weight,
      dueDate:a.DueDate,
      status:canSeeSubmission?(result?(result.Status||'Not Submitted'):'Not Submitted'):'',
      score:canSeeScore&&result&&result.Score!==undefined&&result.Score!==null?result.Score:'',
      feedback:canSeeScore&&result?(result.Feedback||''):'',
      cloLevels:canSeeClo?cloLevels:{},
      visibilityLevel:visibilityLevel,
      canSeeSubmission:canSeeSubmission,
      canSeeScore:canSeeScore,
      canSeeClo:canSeeClo,
      sequence:a.Sequence
    };
  });

  const trajectories={};
  clos.forEach(c=>trajectories[String(c.CLOID)]=[]);
  activityViews.forEach(a=>{
    Object.keys(a.cloLevels||{}).forEach(cid=>{
      if(!trajectories[cid])trajectories[cid]=[];
      trajectories[cid].push({
        activityId:a.activityId,title:a.title,level:a.cloLevels[cid],sequence:a.sequence
      });
    });
  });
  Object.keys(trajectories).forEach(cid=>{
    trajectories[cid].sort((a,b)=>Number(a.sequence)-Number(b.sequence));
  });

  const attendanceSessions=(snap.attendanceSessions||[])
    .filter(s=>bool_(s.Published) && new Date(s.StartTime).getTime()<=Date.now())
    .sort((a,b)=>new Date(a.StartTime)-new Date(b.StartTime));

  const attendance=attendanceSessions.map(s=>{
    const a=idx.attendanceBySessionStudent[String(s.SessionID||'')+'|'+studentKey];
    return {
      title:s.Title,
      date:dateIso_(s.StartTime),
      status:a?a.Status:'Absent'
    };
  });

  const matchedResultCount=activityViews.filter(a=>
    a.score!=='' || !['Not Submitted','Missing'].includes(String(a.status||''))
  ).length;

  const selfByClo={};
  (snap.selfEvaluations||[]).forEach(r=>{
    if(String(r.StudentKey||'')!==studentKey)return;
    const cid=String(r.CLOID||'');
    if(!cid)return;
    if(!selfByClo[cid])selfByClo[cid]=[];
    selfByClo[cid].push({
      point:String(r.EvaluationPoint||''),
      rating:Number(r.Rating)||0,
      reflection:String(r.Reflection||''),
      submittedAt:dateIso_(r.SubmittedAt)
    });
  });
  Object.keys(selfByClo).forEach(cid=>{
    selfByClo[cid].sort((a,b)=>selfEvaluationPointOrder_(a.point)-selfEvaluationPointOrder_(b.point));
  });

  return {
    alias:student.Alias,
    teamCode:student.TeamCode||'',
    activities:activityViews,
    clos:clos.map(c=>({
      cloId:c.CLOID,
      code:c.Code,
      description:c.Description,
      trajectory:trajectories[String(c.CLOID)]||[],
      selfEvaluation:selfByClo[String(c.CLOID)]||[]
    })),
    attendance:attendance,
    debug:{
      studentKey:studentKey,
      visibleActivityCount:activities.length,
      activityViewCount:activityViews.length,
      matchedResultCount:matchedResultCount,
      rawResultRowsForStudent:(snap.results||[]).filter(r=>
        String(r.TargetType||'')==='STUDENT' &&
        String(r.TargetID||'')===studentKey
      ).length
    }
  };
}

function serializeActivity_(a) {
  return {
    ActivityID:a.ActivityID,
    Title:a.Title,
    ShortDescription:String(a.ShortDescription||''),
    DetailUrl:safeActivityDetailUrl_(a.DetailUrl),
    Category:a.Category,
    Mode:a.Mode,
    Purpose:a.Purpose,
    MaxScore:a.MaxScore,
    Weight:a.Weight,
    Published:normalizeActivityVisibilityLevel_(a.StudentVisibilityLevel,a.Published)!=='HIDDEN',
    StudentVisibilityLevel:normalizeActivityVisibilityLevel_(a.StudentVisibilityLevel,a.Published),
    AssignedWeek:a.AssignedWeek===''||a.AssignedWeek==null?'':Number(a.AssignedWeek),
    DueDate:a.DueDate instanceof Date ? Utilities.formatDate(a.DueDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm") : String(a.DueDate || ''),
    Sequence:Number(a.Sequence)||0,
    Active:bool_(a.Active),
    CLOEvalMode:String(a.CLOEvalMode||'MANUAL').toUpperCase(),
    CLOScoreSource:String(a.CLOScoreSource||'ACTIVITY_SCORE').toUpperCase()==='SEPARATE_SCORE'?'SEPARATE_SCORE':'ACTIVITY_SCORE',
    CLOMaxScore:a.CLOMaxScore===''||a.CLOMaxScore==null?'':Number(a.CLOMaxScore),
    CLOLevel2Pct:a.CLOLevel2Pct===''||a.CLOLevel2Pct==null?50:Number(a.CLOLevel2Pct),
    CLOLevel3Pct:a.CLOLevel3Pct===''||a.CLOLevel3Pct==null?65:Number(a.CLOLevel3Pct),
    CLOLevel4Pct:a.CLOLevel4Pct===''||a.CLOLevel4Pct==null?80:Number(a.CLOLevel4Pct),
    CLOLevel5Pct:a.CLOLevel5Pct===''||a.CLOLevel5Pct==null?90:Number(a.CLOLevel5Pct),
    CLOAllowOverride:bool_(a.CLOAllowOverride),
    cloScoreMappings:a.ActivityID?activityCloScoreMappings_(a.ActivityID):[],
    cloIds:a.ActivityID?activityCloIds_(a.ActivityID):[]
  };
}



const COURSE_EXPORT_GROUPS_ = {
  courseProfile:[APP.tables.CONFIG,APP.tables.COURSE_OBJECTIVES],
  clos:[APP.tables.CLOS,APP.tables.CLO_ATTAINMENT_POLICIES,APP.tables.CLO_ATTAINMENT_WEIGHTS],
  weeklyPlan:[APP.tables.WEEKLY_PLAN,APP.tables.WEEKLY_PLAN_CLOS],
  activities:[APP.tables.ACTIVITIES,APP.tables.ACTIVITY_CLOS,APP.tables.ACTIVITY_CLO_SCORE_MAPS],
  sections:[APP.tables.COURSE_SECTIONS,APP.tables.COURSE_SECTION_SCHEDULES],
  staff:[APP.tables.STAFF],
  students:[APP.tables.STUDENTS,APP.tables.STUDENT_SECTIONS],
  results:[APP.tables.RESULTS],
  cloEvidence:[APP.tables.CLO_EVIDENCE],
  attendance:[APP.tables.ATTENDANCE_SESSIONS,APP.tables.ATTENDANCE],
  teams:[APP.tables.ACTIVITY_TEAMS,APP.tables.ACTIVITY_TEAM_MEMBERS,APP.tables.TEAM_RESULT_MEMBERS],
  selfEvaluation:[APP.tables.CLO_SELF_EVALUATION],
  teachingHistory:[APP.tables.WEEKLY_PLAN_HISTORY],
  audit:[APP.tables.AUDIT]
};

function exportSafeCell_(value){
  if(value instanceof Date)return Utilities.formatDate(value,Session.getScriptTimeZone()||'Asia/Bangkok',"yyyy-MM-dd'T'HH:mm:ssXXX");
  if(value===null||value===undefined)return '';
  return value;
}
function exportSafeHeaders_(tableName){
  const all=(APP.headers[tableName]||[]).slice();
  const blocked={StaffTokenHash:true,TokenIssuedAt:true,StudentTokenHash:true,Secret:true};
  return all.filter(h=>!blocked[h]);
}
function exportTable_(tableName){
  const sh=table_(tableName);
  const headers=exportSafeHeaders_(tableName);
  const actualHeaders=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(v=>String(v||'').trim());
  const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,Math.max(1,sh.getLastColumn())).getValues():[];
  const indices=headers.map(h=>actualHeaders.indexOf(h));
  return {name:tableName,headers:headers,rows:rows.map(r=>indices.map(i=>exportSafeCell_(i>=0?r[i]:'')))};
}
function exportConfigTable_(templateMode){
  const sh=table_(APP.tables.CONFIG);
  const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,2).getValues():[];
  const blockedPattern=/(TOKEN|SECRET|HASH)/i;
  const offeringKeys={APP_VERSION:true,SCHEMA_VERSION:true,SETUP_COMPLETE:true,SEMESTER:true,ACADEMIC_YEAR:true};
  const out=rows.filter(r=>{
    const k=String(r[0]||'').trim();
    if(!k||blockedPattern.test(k))return false;
    if(templateMode&&offeringKeys[k])return false;
    return true;
  }).map(r=>[exportSafeCell_(r[0]),exportSafeCell_(r[1])]);
  return {name:APP.tables.CONFIG,headers:['Key','Value'],rows:out};
}

function exportSheetFromObjects_(name, headers, rows){
  return {
    name:name,
    headers:headers,
    rows:(rows||[]).map(obj=>headers.map(h=>exportSafeCell_(obj&&obj[h]!==undefined?obj[h]:'')))
  };
}

function buildObeExportSheets_(){
  const cfg=configMapFast_();
  const objectives=rowsToObjects_(table_(APP.tables.COURSE_OBJECTIVES))
    .filter(r=>r.Active===''||bool_(r.Active))
    .sort((a,b)=>Number(a.Sequence||0)-Number(b.Sequence||0));
  const clos=rowsToObjects_(table_(APP.tables.CLOS))
    .filter(r=>r.Active===''||bool_(r.Active))
    .sort((a,b)=>Number(a.Sequence||0)-Number(b.Sequence||0));
  const policies=rowsToObjects_(table_(APP.tables.CLO_ATTAINMENT_POLICIES));
  const policyByClo={};policies.forEach(r=>policyByClo[String(r.CLOID||'')]=r);
  const students=rowsToObjects_(table_(APP.tables.STUDENTS));
  const studentByKey={};students.forEach(r=>studentByKey[String(r.StudentKey||'')]=r);
  const analytics=getCloAttainmentAnalytics();
  const summary=Array.isArray(analytics&&analytics.summary)?analytics.summary:[];

  const definitionRows=[];
  definitionRows.push({Type:'Course',Code:String(cfg.COURSE_CODE_TH||cfg.COURSE_CODE_EN||cfg.COURSE_CODE||''),Description:String(cfg.COURSE_NAME_TH||cfg.COURSE_NAME_EN||cfg.COURSE_NAME||''),Policy:'',Threshold:'',MinCount:'',Sequence:''});
  objectives.forEach(o=>definitionRows.push({Type:'Objective',Code:String(o.ObjectiveID||''),Description:String(o.ObjectiveText||''),Policy:'',Threshold:'',MinCount:'',Sequence:Number(o.Sequence||0)}));
  clos.forEach(c=>{const p=policyByClo[String(c.CLOID||'')]||{};definitionRows.push({Type:'CLO',Code:String(c.Code||c.CLOID||''),Description:String(c.Description||''),StudentFriendlyDescription:String(c.StudentFriendlyDescription||''),Policy:String(p.Policy||'LATEST'),Threshold:p.Threshold===''?'':Number(p.Threshold||3),MinCount:p.MinCount===''?'':Number(p.MinCount||2),Sequence:Number(c.Sequence||0)});});

  const studentResultRows=[];
  const courseSummaryRows=[];
  const selfRows=[];
  summary.forEach(item=>{
    const pol=item.policy||{};
    courseSummaryRows.push({CLO:item.code||item.cloId,Description:item.description||'',EvidenceStudents:Number(item.evidenceStudents||0),AttainedCount:Number(item.attainedCount||0),NotAttainedCount:Number(item.notAttainedCount||0),AttainmentPct:item.attainmentPct==null?'':Number(item.attainmentPct),AverageLevel:item.avgLevel==null?'':Number(item.avgLevel),Policy:String(pol.Policy||''),Threshold:pol.Threshold==null?'':Number(pol.Threshold),MinCount:pol.MinCount==null?'':Number(pol.MinCount)});
    (item.studentRows||[]).forEach(r=>{
      const st=studentByKey[String(r.studentKey||'')]||{};
      studentResultRows.push({StudentID:String(r.studentId||st.StudentId||''),Name:String(r.name||st.Name||''),CLO:item.code||item.cloId,Aggregate:r.aggregate==null?'':Number(r.aggregate),Attained:r.points&&r.points.length?!!r.attained:'',EvidenceCount:(r.points||[]).length,Policy:String(pol.Policy||''),Threshold:pol.Threshold==null?'':Number(pol.Threshold)});
      const sr=r.selfRatings||{};
      const start=sr.START||{}, mid=sr.MIDTERM||{}, end=sr.END||{};
      if(start.rating!==undefined||mid.rating!==undefined||end.rating!==undefined){
        selfRows.push({StudentID:String(r.studentId||st.StudentId||''),Name:String(r.name||st.Name||''),CLO:item.code||item.cloId,Start:start.rating===undefined?'':Number(start.rating),Midterm:mid.rating===undefined?'':Number(mid.rating),Final:end.rating===undefined?'':Number(end.rating),StartReflection:String(start.reflection||''),MidtermReflection:String(mid.reflection||''),FinalReflection:String(end.reflection||''),StartSubmittedAt:String(start.submittedAt||''),MidtermSubmittedAt:String(mid.submittedAt||''),FinalSubmittedAt:String(end.submittedAt||'')});
      }
    });
  });

  const activities=rowsToObjects_(table_(APP.tables.ACTIVITIES));
  const activityById={};activities.forEach(a=>activityById[String(a.ActivityID||'')]=a);
  const cloById={};clos.forEach(c=>cloById[String(c.CLOID||'')]=c);
  const evidenceRows=rowsToObjects_(table_(APP.tables.CLO_EVIDENCE)).map(e=>{
    const a=activityById[String(e.ActivityID||'')]||{};
    const c=cloById[String(e.CLOID||'')]||{};
    const st=String(e.TargetType||'')==='STUDENT'?(studentByKey[String(e.TargetID||'')]||{}):{};
    return {StudentID:String(st.StudentId||''),Name:String(st.Name||''),TargetType:String(e.TargetType||''),TargetID:String(e.TargetID||''),CLO:String(c.Code||e.CLOID||''),ActivityID:String(e.ActivityID||''),ActivityTitle:String(a.Title||''),EvidenceType:String(e.Source||''),Level:e.Level===''?'':Number(e.Level),Score:e.Score===''?'':Number(e.Score),MaxScore:e.MaxScore===''?'':Number(e.MaxScore),EvidenceReference:String(a.DetailUrl||e.ActivityID||''),UpdatedAt:e.UpdatedAt||''};
  });

  return [
    exportSheetFromObjects_('Course & CLO Definition',['Type','Code','Description','Policy','Threshold','MinCount','Sequence'],definitionRows),
    exportSheetFromObjects_('Student CLO Results',['StudentID','Name','CLO','Aggregate','Attained','EvidenceCount','Policy','Threshold'],studentResultRows),
    exportSheetFromObjects_('Course CLO Summary',['CLO','Description','EvidenceStudents','AttainedCount','NotAttainedCount','AttainmentPct','AverageLevel','Policy','Threshold','MinCount'],courseSummaryRows),
    exportSheetFromObjects_('Self-Evaluation Summary',['StudentID','Name','CLO','Start','Midterm','Final','StartReflection','MidtermReflection','FinalReflection','StartSubmittedAt','MidtermSubmittedAt','FinalSubmittedAt'],selfRows),
    exportSheetFromObjects_('CLO Evidence Reference',['StudentID','Name','TargetType','TargetID','CLO','ActivityID','ActivityTitle','EvidenceType','Level','Score','MaxScore','EvidenceReference','UpdatedAt'],evidenceRows)
  ];
}

function exportCourseDataPackage(options){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  ensureCurrentSchema_();
  options=options||{};
  let mode=String(options.mode||'FULL_BACKUP').toUpperCase();
  if(mode==='FULL_ARCHIVE')mode='FULL_BACKUP'; // backward compatibility
  const templateMode=mode==='COURSE_TEMPLATE';
  const obeMode=mode==='LEARNING_RESULTS_OBE';
  let groups=Array.isArray(options.groups)?options.groups.filter(g=>COURSE_EXPORT_GROUPS_[g]):[];
  if(!groups.length)groups=templateMode?['courseProfile','clos','weeklyPlan','activities']:(obeMode?['courseProfile','clos','students','results','cloEvidence','selfEvaluation']:Object.keys(COURSE_EXPORT_GROUPS_));
  const names=[];
  groups.forEach(g=>(COURSE_EXPORT_GROUPS_[g]||[]).forEach(n=>{if(!names.includes(n))names.push(n);}));
  const sheets=[];
  names.forEach(name=>sheets.push(name===APP.tables.CONFIG?exportConfigTable_(templateMode):exportTable_(name)));
  if(obeMode)buildObeExportSheets_().forEach(sh=>sheets.push(sh));
  const cfg=configMapFast_();
  const manifest={
    format:'LPT_COURSE_EXPORT',formatVersion:1,exportType:templateMode?'COURSE_TEMPLATE':(obeMode?'LEARNING_RESULTS_OBE':(mode==='CUSTOM'?'CUSTOM':'FULL_BACKUP')),
    appVersion:APP.version,schemaVersion:APP.schemaVersion,exportedAt:new Date().toISOString(),
    courseCode:String(cfg.COURSE_CODE_TH||cfg.COURSE_CODE_EN||cfg.COURSE_CODE||''),
    courseName:String(cfg.COURSE_NAME_TH||cfg.COURSE_NAME_EN||cfg.COURSE_NAME||''),groups:groups
  };
  audit_('EXPORT_COURSE_DATA','Course','',{actor:actor.email,mode:manifest.exportType,groups:groups.join(',')});
  return {manifest:manifest,sheets:sheets};
}

function importCourseTemplatePackage(pkg){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  ensureCurrentSchema_();
  pkg=pkg||{};
  const manifest=pkg.manifest||{};
  if(String(manifest.format||'')!=='LPT_COURSE_EXPORT')throw new Error('This file is not a Course Learning Tracker export.');
  if(String(manifest.exportType||'')!=='COURSE_TEMPLATE')throw new Error('Import requires a Course Template export.');
  if(Number(manifest.formatVersion||0)!==1)throw new Error('Unsupported Course Template format version.');
  const dependentTables=[APP.tables.RESULTS,APP.tables.CLO_EVIDENCE,APP.tables.ATTENDANCE,APP.tables.ACTIVITY_TEAMS,APP.tables.ACTIVITY_TEAM_MEMBERS,APP.tables.TEAM_RESULT_MEMBERS,APP.tables.CLO_SELF_EVALUATION];
  const hasDependentData=dependentTables.some(name=>{const sh=table_(name);return sh.getLastRow()>1;});
  if(hasDependentData)throw new Error('Course Template import is blocked because this tracker already contains results, attendance, teams, or other learner evidence. Import the template into a fresh/new-term tracker before recording learner evidence.');
  const incoming=Array.isArray(pkg.sheets)?pkg.sheets:[];
  const allowed=[APP.tables.CONFIG,APP.tables.COURSE_OBJECTIVES,APP.tables.CLOS,APP.tables.CLO_ATTAINMENT_POLICIES,APP.tables.CLO_ATTAINMENT_WEIGHTS,APP.tables.WEEKLY_PLAN,APP.tables.WEEKLY_PLAN_CLOS,APP.tables.ACTIVITIES,APP.tables.ACTIVITY_CLOS,APP.tables.ACTIVITY_CLO_SCORE_MAPS];
  const imported=[];
  incoming.forEach(data=>{
    const name=String(data&&data.name||'');
    if(!allowed.includes(name))return;
    const rows=Array.isArray(data.rows)?data.rows:[];
    const headers=Array.isArray(data.headers)?data.headers.map(String):[];
    if(name===APP.tables.CONFIG){
      const preserve={SEMESTER:true,ACADEMIC_YEAR:true,APP_VERSION:true,SCHEMA_VERSION:true,SETUP_COMPLETE:true};
      const keyIdx=headers.indexOf('Key'),valIdx=headers.indexOf('Value');
      rows.forEach(r=>{
        const key=clean_(keyIdx>=0?r[keyIdx]:'',120);
        if(!key||preserve[key]||/(TOKEN|SECRET|HASH)/i.test(key))return;
        setConfig_(key,valIdx>=0?r[valIdx]:'');
      });
      imported.push(name);return;
    }
    const expected=APP.headers[name]||[];
    if(!headers.length||!expected.length)return;
    const sh=table_(name);
    if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,Math.max(1,sh.getLastColumn())).clearContent();
    if(rows.length){
      const mapped=rows.map(r=>expected.map(h=>{
        const idx=headers.indexOf(h);return idx>=0?exportSafeCell_(r[idx]):'';
      }));
      sh.getRange(2,1,mapped.length,expected.length).setValues(mapped);
    }
    imported.push(name);
  });
  setConfig_('APP_VERSION',APP.version);setConfig_('SCHEMA_VERSION',String(APP.schemaVersion));
  resetRequestDataCache_();
  audit_('IMPORT_COURSE_TEMPLATE','Course','',{actor:actor.email,sourceVersion:String(manifest.appVersion||''),tables:imported.join(',')});
  return {ok:true,imported:imported,sourceVersion:String(manifest.appVersion||''),sourceCourseCode:String(manifest.courseCode||'')};
}

function setupStateCacheKey_(){return 'LPT:SETUP_STATE:'+String(APP.schemaVersion||APP.version||'');}

function clearSetupStateCache_(){
  try{CacheService.getScriptCache().remove(setupStateCacheKey_());}catch(e){}
}

function getSetupStateCached_(){
  const cache=CacheService.getScriptCache();
  const key=setupStateCacheKey_();
  const raw=cache.get(key);
  if(raw){
    try{return JSON.parse(raw);}catch(e){}
  }
  const state=getSetupState_();
  try{cache.put(key,JSON.stringify(state),120);}catch(e){}
  return state;
}

function getSetupState_() {
  const clos = listCLOs_().filter(c=>c.Active);
  const students = listStudentsForStaff_().filter(s=>s.Active);
  const activities = listActivities_().filter(a=>a.Active);
  const staff = rowsToObjects_(table_(APP.tables.STAFF)).filter(s=>bool_(s.Active));
  const sections = listCourseSections_().filter(s=>s&&s.active!==false);
  const weeklyPlan = rowsToObjects_(table_(APP.tables.WEEKLY_PLAN)).filter(r=>bool_(r.Active));

  const courseReady = isCourseConfigured_();
  const closReady = clos.length > 0;
  const sectionsReady = sections.length > 0;
  const staffReady = staff.length > 0;
  const studentsReady = students.length > 0;
  const weeklyPlanReady = weeklyPlan.length > 0;
  const activitiesReady = activities.length > 0;
  const onboardingComplete = !!(courseReady && closReady && sectionsReady && staffReady && studentsReady && weeklyPlanReady && activitiesReady);

  return {
    courseReady,
    closReady,
    sectionsReady,
    staffReady,
    studentsReady,
    weeklyPlanReady,
    activitiesReady,
    onboardingComplete,

    canUseCLOs: courseReady,
    canUseRoster: courseReady && closReady,
    canUseStaff: courseReady,
    canUseActivities: courseReady && closReady && studentsReady,
    canEnterResults: courseReady && studentsReady && activitiesReady,
    canUseAttendance: courseReady && studentsReady,

    counts: {
      clos: clos.length,
      sections: sections.length,
      students: students.length,
      weeklyPlan: weeklyPlan.length,
      activities: activities.length,
      staff: staff.length
    }
  };
}



function getActivityStudentVisibility_(snapshot) {
  const snap=snapshot || courseSnapshot_({activities:true});
  const rawRows=snap.activities||[];
  const normalized=normalizedActivitiesFromRows_(rawRows);
  const visible={};
  listStudentVisibleActivities_(snap).forEach(a=>visible[String(a.ActivityID)]=true);

  const rawById={};
  rawRows.forEach(r=>rawById[String(r.ActivityID||'')]=r);

  const out={};
  normalized.forEach(a=>{
    const raw=rawById[String(a.ActivityID||'')]||{};
    out[a.ActivityID]={
      published:bool_(a.Published),
      active:bool_(a.Active) || raw.Active==='' || raw.Active==null,
      visibleToStudents:!!visible[String(a.ActivityID)],
      publishedRaw:String(raw.Published==null?'':raw.Published),
      activeRaw:String(raw.Active==null?'':raw.Active)
    };
  });
  return out;
}

function getActivityResultSummaries_(snapshot) {
  const snap=snapshot || courseSnapshot_({students:true,activities:true,results:true});
  const students=normalizedStudentsFromRows_(snap.students||[]).filter(s=>s.Active);
  const activities=normalizedActivitiesFromRows_(snap.activities||[]);
  const results=snap.results||[];
  const out={};

  const teamSet={};
  students.forEach(s=>{if(s.TeamCode)teamSet[String(s.TeamCode)]=true;});
  const teamTotal=Object.keys(teamSet).length;

  const agg={};
  results.forEach(r=>{
    const aid=String(r.ActivityID||'');
    if(!aid)return;
    if(!agg[aid])agg[aid]={resultRows:0,submitted:0,scored:0,scoreSum:0};
    const a=agg[aid];
    a.resultRows++;
    const status=String(r.Status||'');
    if(status && !['Not Submitted','Missing'].includes(status))a.submitted++;
    if(r.Score!=='' && r.Score!=null && Number.isFinite(Number(r.Score))){
      a.scored++;
      a.scoreSum+=Number(r.Score);
    }
  });

  activities.forEach(a=>{
    const x=agg[String(a.ActivityID)]||{resultRows:0,submitted:0,scored:0,scoreSum:0};
    out[a.ActivityID]={
      totalTargets:a.Mode==='Team'?teamTotal:students.length,
      resultRows:x.resultRows,
      submitted:x.submitted,
      scored:x.scored,
      average:x.scored?Math.round((x.scoreSum/x.scored)*100)/100:''
    };
  });

  return out;
}


function dashboardPct_(value){
  const n=Number(value);
  return Number.isFinite(n)?Math.round(n*10)/10:null;
}

function buildDashboardInsights_(snap, students, activities, clos, summaries) {
  const activeActivities=(activities||[]).filter(a=>a.Active!==false);
  const publishedActivities=activeActivities.filter(a=>bool_(a.Published));

  // ---------- Activity health ----------
  const activityAttention=[];
  publishedActivities.forEach(a=>{
    const s=(summaries||{})[a.ActivityID]||{};
    const total=Number(s.totalTargets||0);
    const submitted=Number(s.submitted||0);
    const missing=Math.max(0,total-submitted);
    if(total>0 && missing>0){
      activityAttention.push({
        type:'ACTIVITY_MISSING',
        activityId:a.ActivityID,
        title:missing+' student'+(missing===1?'':'s')+' not submitted / missing',
        detail:a.Title,
        count:missing,
        total:total
      });
    }
  });
  activityAttention.sort((a,b)=>b.count-a.count);

  // ---------- CLO attainment ----------
  const maps=snap.activityClos||[];
  const evidence=snap.evidence||[];
  const policies=snap.policies||[];
  const weights=snap.weights||[];

  const activityById={};
  activeActivities.forEach(a=>activityById[String(a.ActivityID||'')]=a);

  const linkedByClo={};
  maps.forEach(m=>{
    const cid=String(m.CLOID||'');
    const aid=String(m.ActivityID||'');
    if(!cid||!aid)return;
    if(!linkedByClo[cid])linkedByClo[cid]=[];
    if(!linkedByClo[cid].includes(aid))linkedByClo[cid].push(aid);
  });

  const policyByClo={};
  policies.forEach(r=>{
    const cid=String(r.CLOID||'');
    if(!cid)return;
    policyByClo[cid]={
      Policy:String(r.Policy||'LATEST').toUpperCase(),
      Threshold:Number(r.Threshold||3),
      MinCount:Number(r.MinCount||2)
    };
  });

  const weightByClo={};
  weights.forEach(r=>{
    const cid=String(r.CLOID||'');
    const aid=String(r.ActivityID||'');
    if(!cid||!aid)return;
    if(!weightByClo[cid])weightByClo[cid]={};
    weightByClo[cid][aid]=Number(r.Weight)||0;
  });

  const evidenceByCloStudent={};
  evidence.forEach(e=>{
    if(String(e.TargetType||'')!=='STUDENT')return;
    const cid=String(e.CLOID||'');
    const sk=String(e.TargetID||'');
    const aid=String(e.ActivityID||'');
    const level=Number(e.Level)||0;
    if(!cid||!sk||level<1||level>5)return;
    if(!evidenceByCloStudent[cid])evidenceByCloStudent[cid]={};
    if(!evidenceByCloStudent[cid][sk])evidenceByCloStudent[cid][sk]=[];
    const a=activityById[aid]||{};
    evidenceByCloStudent[cid][sk].push({
      activityId:aid,
      sequence:Number(a.Sequence)||0,
      level:level
    });
  });

  Object.keys(evidenceByCloStudent).forEach(cid=>{
    Object.keys(evidenceByCloStudent[cid]).forEach(sk=>{
      evidenceByCloStudent[cid][sk].sort((a,b)=>a.sequence-b.sequence);
    });
  });

  function evaluateDashboardClo_(points,policy,weightMap){
    if(!points.length)return null;
    let aggregate=null,attained=false;
    if(policy.Policy==='LATEST'){
      aggregate=points[points.length-1].level;
      attained=aggregate>=policy.Threshold;
    }else if(policy.Policy==='AVERAGE'){
      aggregate=points.reduce((s,p)=>s+p.level,0)/points.length;
      attained=aggregate>=policy.Threshold;
    }else if(policy.Policy==='WEIGHTED'){
      let num=0,den=0;
      points.forEach(p=>{
        const w=Number((weightMap||{})[p.activityId]||0);
        if(w>0){num+=p.level*w;den+=w;}
      });
      if(den>0){
        aggregate=num/den;
        attained=aggregate>=policy.Threshold;
      }
    }else if(policy.Policy==='MIN_COUNT'){
      aggregate=points.filter(p=>p.level>=policy.Threshold).length;
      attained=aggregate>=policy.MinCount;
    }
    return {aggregate:aggregate,attained:attained};
  }

  const selfEvalByCloPointStudent={};
  (snap.selfEvaluations||[]).forEach(r=>{
    const cid=String(r.CLOID||''),sk=String(r.StudentKey||''),point=String(r.EvaluationPoint||'').toUpperCase(),rating=Number(r.Rating);
    if(!cid||!sk||!['START','MIDTERM','END'].includes(point)||rating<1||rating>5)return;
    if(!selfEvalByCloPointStudent[cid])selfEvalByCloPointStudent[cid]={START:{},MIDTERM:{},END:{}};
    selfEvalByCloPointStudent[cid][point][sk]=rating;
  });

  const cloSnapshot=(clos||[]).map(c=>{
    const cid=String(c.CLOID||c.cloId||'');
    const policy=policyByClo[cid]||{Policy:'LATEST',Threshold:3,MinCount:2};
    const byStudent=evidenceByCloStudent[cid]||{};
    let withEvidence=0,attained=0;
    const studentMeanLevels=[];

    (students||[]).forEach(st=>{
      const points=byStudent[String(st.StudentKey||st.studentKey||'')]||[];
      const result=evaluateDashboardClo_(points,policy,weightByClo[cid]||{});
      if(result){
        withEvidence++;
        if(result.attained)attained++;
      }
      // Class radar mirrors the student radar: first average each student's
      // assessed-evidence levels for this CLO, then average across students.
      // Students without evidence are excluded rather than treated as zero.
      if(points.length){
        const mean=points.reduce((sum,p)=>sum+Number(p.level||0),0)/points.length;
        if(mean>=1&&mean<=5)studentMeanLevels.push(mean);
      }
    });

    const classAverageLevel=studentMeanLevels.length
      ? Math.round((studentMeanLevels.reduce((sum,v)=>sum+v,0)/studentMeanLevels.length)*100)/100
      : null;
    const targetLevel=(policy.Policy==='LATEST'||policy.Policy==='AVERAGE'||policy.Policy==='WEIGHTED')
      ? Number(policy.Threshold||0)
      : null;
    const linkedCount=(linkedByClo[cid]||[]).length;
    const selfPointRows=(selfEvalByCloPointStudent[cid]||{START:{},MIDTERM:{},END:{}});
    const selfEvaluation={};
    ['START','MIDTERM','END'].forEach(point=>{
      const vals=Object.keys(selfPointRows[point]||{}).map(sk=>Number(selfPointRows[point][sk])).filter(v=>v>=1&&v<=5);
      selfEvaluation[point]={average:vals.length?Math.round((vals.reduce((sum,v)=>sum+v,0)/vals.length)*100)/100:null,count:vals.length};
    });
    return {
      cloId:cid,
      code:c.Code||c.code||cid,
      description:c.Description||c.description||'',
      studentFriendlyDescription:c.StudentFriendlyDescription||c.studentFriendlyDescription||'',
      linkedActivities:linkedCount,
      evidenceStudents:withEvidence,
      attainedCount:attained,
      attainmentPct:withEvidence?dashboardPct_(attained*100/withEvidence):null,
      classAverageLevel:classAverageLevel,
      targetLevel:(targetLevel>=1&&targetLevel<=5)?targetLevel:null,
      selfEvaluation:selfEvaluation
    };
  });

  const assessedClos=cloSnapshot.filter(c=>c.evidenceStudents>0);
  const overallCloPct=assessedClos.length
    ? dashboardPct_(assessedClos.reduce((s,c)=>s+Number(c.attainmentPct||0),0)/assessedClos.length)
    : null;

  const cloAttention=cloSnapshot
    .filter(c=>c.linkedActivities===0 || c.evidenceStudents===0)
    .map(c=>({
      type:'CLO_NO_EVIDENCE',
      cloId:c.cloId,
      code:c.code,
      title:c.code+(c.linkedActivities===0?' · No evidence activity':' · No student evidence yet'),
      detail:c.description||''
    }));

  // ---------- Attendance health ----------
  const now=Date.now();
  const allSessions=snap.attendanceSessions||[];
  const attendanceRows=snap.attendance||[];

  const attendanceRowsBySession={};
  attendanceRows.forEach(a=>{
    const sid=String(a.SessionID||'');
    if(!attendanceRowsBySession[sid])attendanceRowsBySession[sid]=[];
    attendanceRowsBySession[sid].push(a);
  });

  function dashboardSessionCompleted_(s){
    const status=String(s.Status||'').toUpperCase();
    if(status==='CLOSED')return true;
    const end=s.EndTime?new Date(s.EndTime).getTime():NaN;
    return Number.isFinite(end)&&end<=now;
  }

  const publishedSessions=allSessions.filter(s=>{
    const sid=String(s.SessionID||'');
    return bool_(s.Published)
      && dashboardSessionCompleted_(s)
      && (attendanceRowsBySession[sid]||[]).length>0;
  });

  const completedUnpublishedSessions=allSessions.filter(s=>{
    const sid=String(s.SessionID||'');
    return !bool_(s.Published)
      && dashboardSessionCompleted_(s)
      && (attendanceRowsBySession[sid]||[]).length>0;
  });
  const unpublishedSessionCount=completedUnpublishedSessions.length;

  const attendanceByKey={};
  attendanceRows.forEach(a=>{
    attendanceByKey[String(a.SessionID||'')+'|'+String(a.StudentKey||'')]=a;
  });

  let attendedUnits=0,attendanceUnits=0,excusedUnits=0;
  const absenceCountByStudent={};

  publishedSessions.forEach(session=>{
    (students||[]).forEach(st=>{
      const sk=String(st.StudentKey||st.studentKey||'');
      const row=attendanceByKey[String(session.SessionID||'')+'|'+sk];
      const status=String(row?row.Status:'Absent');

      if(status==='Excused'){
        excusedUnits++;
        return;
      }

      attendanceUnits++;
      if(status==='Present'||status==='Late'){
        attendedUnits++;
      }else{
        absenceCountByStudent[sk]=(absenceCountByStudent[sk]||0)+1;
      }
    });
  });

  const repeatedAbsenceStudents=Object.keys(absenceCountByStudent)
    .filter(sk=>absenceCountByStudent[sk]>=2).length;

  const attendancePct=attendanceUnits?dashboardPct_(attendedUnits*100/attendanceUnits):null;

  const attendanceAttention=[];
  if(repeatedAbsenceStudents>0){
    attendanceAttention.push({
      type:'REPEATED_ABSENCE',
      count:repeatedAbsenceStudents,
      title:repeatedAbsenceStudents+' student'+(repeatedAbsenceStudents===1?'':'s')+' absent from 2+ sessions',
      detail:publishedSessions.length+' completed published attendance session'+(publishedSessions.length===1?'':'s')+' included'
    });
  }
  if(unpublishedSessionCount>0){
    attendanceAttention.push({
      type:'UNPUBLISHED_ATTENDANCE',
      count:unpublishedSessionCount,
      title:unpublishedSessionCount+' completed attendance session'+(unpublishedSessionCount===1?' is':'s are')+' unpublished',
      detail:'Publish completed sessions when students should be able to see them.'
    });
  }

  // ---------- Recent activities ----------
  const recentActivities=activeActivities
    .slice()
    .sort((a,b)=>Number(b.Sequence||0)-Number(a.Sequence||0))
    .slice(0,4)
    .map(a=>{
      const s=(summaries||{})[a.ActivityID]||{};
      return {
        activityId:a.ActivityID,
        title:a.Title,
        shortDescription:a.ShortDescription||'',
        detailUrl:safeActivityDetailUrl_(a.DetailUrl),
        category:a.Category,
        purpose:a.Purpose,
        published:bool_(a.Published),
        scored:Number(s.scored||0),
        submitted:Number(s.submitted||0),
        total:Number(s.totalTargets||0),
        average:s.average,
        maxScore:a.MaxScore,
        dueDate:a.DueDate||''
      };
    });

  return {
    studentsCount:(students||[]).length,
    activitiesCount:activeActivities.length,
    publishedActivitiesCount:publishedActivities.length,
    activitiesNeedingResults:activityAttention.length,
    cloOverallPct:overallCloPct,
    assessedCloCount:assessedClos.length,
    totalCloCount:(clos||[]).length,
    attendancePct:attendancePct,
    publishedAttendanceSessions:publishedSessions.length,
    excusedAttendanceUnits:excusedUnits,
    needsAttention:activityAttention.slice(0,3)
      .concat(cloAttention.slice(0,2))
      .concat(attendanceAttention)
      .slice(0,7),
    cloSnapshot:cloSnapshot,
    recentActivities:recentActivities
  };
}


function getDashboardKpiData_(){
  const snap=courseSnapshot_({
    students:true,activities:true,results:true,clos:true,
    evidence:true,attendanceSessions:true,attendance:true
  });

  const students=normalizedStudentsFromRows_(snap.students||[]).filter(s=>s.Active);
  const activities=normalizedActivitiesFromRows_(snap.activities||[]).filter(a=>a.Active!==false);
  const publishedActivities=activities.filter(a=>bool_(a.Published));
  const clos=normalizedClosFromRows_(snap.clos||[]).filter(c=>c.Active);

  const summaries=getActivityResultSummaries_(snap);
  let activitiesNeedingResults=0;
  publishedActivities.forEach(a=>{
    const s=summaries[a.ActivityID]||{};
    if(Number(s.totalTargets||0)>Number(s.scored||0))activitiesNeedingResults++;
  });

  // CLO overall uses the existing insight builder so the KPI is a real value.
  const insights=buildDashboardInsights_(snap,students,activities,clos,summaries);

  return {
    setupState:getSetupStateCached_(),
    dashboardKpis:{
      studentsCount:students.length,
      activitiesCount:activities.length,
      publishedActivitiesCount:publishedActivities.length,
      activitiesNeedingResults:activitiesNeedingResults,
      cloOverallPct:insights.cloOverallPct,
      assessedCloCount:insights.assessedCloCount,
      totalCloCount:clos.length,
      attendancePct:insights.attendancePct,
      publishedAttendanceSessions:insights.publishedAttendanceSessions
    }
  };
}

function getDashboardSectionData(section){
  requireStaff_(['OWNER','INSTRUCTOR','TA']);
  const snap=courseSnapshot_({
    students:true,activities:true,results:true,clos:true,activityClos:true,
    evidence:true,policies:true,weights:true,attendanceSessions:true,attendance:true,selfEvaluations:true
  });
  const students=normalizedStudentsFromRows_(snap.students||[]).filter(s=>s.Active);
  const activities=normalizedActivitiesFromRows_(snap.activities||[]).filter(a=>a.Active!==false);
  const clos=normalizedClosFromRows_(snap.clos||[]).filter(c=>c.Active);
  const summaries=getActivityResultSummaries_(snap);
  const insights=buildDashboardInsights_(snap,students,activities,clos,summaries);
  const key=String(section||'').toLowerCase();

  if(key==='attention')return {needsAttention:insights.needsAttention||[]};
  if(key==='clo')return {cloSnapshot:insights.cloSnapshot||[]};
  if(key==='activities')return {recentActivities:insights.recentActivities||[]};
  return {};
}



function getStaffDashboardData_() {
  ensureCurrentSchema_();
  closeExpiredAttendanceSessions_();
  const snap=courseSnapshot_({
    students:true,activities:true,clos:true,results:true,
    attendanceSessions:true,attendance:true,staff:true,
    activityClos:true,evidence:true,policies:true,weights:true
  });

  const students=normalizedStudentsFromRows_(snap.students).filter(s=>s.Active);
  const normalizedActivities=normalizedActivitiesFromRows_(snap.activities);
  const activities=normalizedActivities.slice().sort(sortSequence_).map(serializeActivity_);
  const clos=normalizedClosFromRows_(snap.clos).filter(c=>c.Active)
    .sort((a,b)=>Number(a.Sequence)-Number(b.Sequence));
  const sessions=(snap.attendanceSessions||[])
    .slice().sort((a,b)=>new Date(b.StartTime)-new Date(a.StartTime))
    .slice(0,20).map(s=>sanitizeAttendanceSession_(s));
  const staff=(snap.staff||[])
    .map(s=>({email:s.Email,role:s.Role,active:bool_(s.Active),addedAt:dateIso_(s.AddedAt)}));

  const summaries=getActivityResultSummaries_(snap);

  return {
    students:students.map(s=>({
      studentKey:s.StudentKey,studentId:s.StudentId,name:s.Name,email:s.Email,
      alias:s.Alias,teamCode:s.TeamCode,active:bool_(s.Active)
    })),
    activities:activities,
    activityResultSummaries:summaries,
    activityStudentVisibility:getActivityStudentVisibility_(snap),
    clos:clos.map(c=>({
      cloId:c.CLOID,code:c.Code,description:c.Description,sequence:c.Sequence
    })),
    sessions:sessions,
    staff:staff,
    dashboardInsights:buildDashboardInsights_(snap,students,normalizedActivities,clos,summaries),
    setupState:getSetupStateCached_(),
    spreadsheetUrl:requireDb_().getUrl()
  };
}

/* ---------- Auth ---------- */

function currentEmail_(allowEffectiveFallback) {
  const active=normalizeEmail_(Session.getActiveUser().getEmail());
  if(active) return active;
  return allowEffectiveFallback ? normalizeEmail_(Session.getEffectiveUser().getEmail()) : '';
}
function requireStaff_(roles) {
  if (REQUEST_STAFF_AUTH_) {
    if (!roles.includes(REQUEST_STAFF_AUTH_.role)) {
      throw new Error('You do not have permission for this action.');
    }
    const latest = getStaff_(REQUEST_STAFF_AUTH_.email);
    if (!latest || !latest.Active || latest.Role !== REQUEST_STAFF_AUTH_.role) {
      throw new Error('This staff access has been revoked or changed.');
    }
    return {email:REQUEST_STAFF_AUTH_.email,role:REQUEST_STAFF_AUTH_.role,via:REQUEST_STAFF_AUTH_.via||'STAFF_INVITE'};
  }

  const email=currentEmail_(false);
  if(!email) throw new Error('Unable to identify your Google account. External staff must use their secure invitation link.');
  const s=getStaff_(email);
  if(!s || !s.Active || !roles.includes(s.Role)) throw new Error('You do not have permission for this action.');
  return {email,role:s.Role,via:'GOOGLE_IDENTITY'};
}
function getStaff_(email) {
  if(!dbOrNull_()) return null;
  const r=rowsToObjects_(table_(APP.tables.STAFF)).find(x=>normalizeEmail_(x.Email)===normalizeEmail_(email));
  return r?{
    Email:r.Email,Role:r.Role,Active:bool_(r.Active),
    StaffTokenHash:r.StaffTokenHash || '',
    TokenIssuedAt:r.TokenIssuedAt || ''
  }:null;
}
function upsertStaff_(email,role) {
  ensureStaffAuthSchema_();
  const sh=table_(APP.tables.STAFF);
  const normalized=normalizeEmail_(email);
  const old=findObject_(sh,'Email',normalized);
  upsertComposite_(sh,['Email'],[normalized],{
    Email:normalized,
    Role:role,
    Active:true,
    AddedAt:old ? (old.AddedAt || new Date()) : new Date(),
    StaffTokenHash:old ? (old.StaffTokenHash || '') : '',
    TokenIssuedAt:old ? (old.TokenIssuedAt || '') : ''
  });
}
function getStudentByEmail_(email) {
  if(!dbOrNull_()) return null;
  const r=rowsToObjects_(table_(APP.tables.STUDENTS)).find(x=>normalizeEmail_(x.Email)===normalizeEmail_(email));
  return r?Object.assign(r,{Active:bool_(r.Active)}):null;
}

/* ---------- Attendance internals ---------- */


function closeExpiredAttendanceSessions_(){
  ensureAttendanceRecordSchema_();
  const sh=table_(APP.tables.ATTENDANCE_SESSIONS);
  const rows=rowsToObjects_(sh);
  const now=Date.now();
  let changed=0;

  rows.forEach(r=>{
    if(String(r.Status||'').toUpperCase()!=='OPEN')return;
    const end=r.EndTime?new Date(r.EndTime):null;
    if(!end||isNaN(end.getTime())||end.getTime()>now)return;

    updateObjectRow_(sh,r._row,Object.assign(stripRow_(r),{
      Status:'CLOSED'
    }));
    changed++;
  });

  if(changed)resetRequestDataCache_();
  return changed;
}

function attendanceSession_(id){ return findObject_(table_(APP.tables.ATTENDANCE_SESSIONS),'SessionID',id); }
function autoCloseSessionIfExpired_(s){
  if(s.Status==='OPEN' && new Date(s.EndTime).getTime()<Date.now()){
    const sh=table_(APP.tables.ATTENDANCE_SESSIONS);
    updateObjectRow_(sh,s._row,Object.assign(stripRow_(s),{Status:'CLOSED'}));
    invalidateAttendanceSessionCache_(s.SessionID);
  }
}
function attendanceBucket_(s,when){
  const sec=Math.max(30,Number(s.RotateSeconds)||60);
  return Math.floor(when.getTime()/1000/sec);
}
function attendanceTokenForBucket_(s,bucket){
  const raw=Utilities.computeHmacSha256Signature(String(bucket),String(s.Secret));
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/,'').substring(0,22);
}
function attendanceToken_(s,when){ return attendanceTokenForBucket_(s,attendanceBucket_(s,when)); }
function validAttendanceToken_(s,token,when){
  const b=attendanceBucket_(s,when);
  return token===attendanceTokenForBucket_(s,b) || token===attendanceTokenForBucket_(s,b-1);
}
function webAppBaseUrl_(){
  return ScriptApp.getService().getUrl() || '';
}

function validateLocation_(s,loc){
  const lat=Number(loc.lat),lng=Number(loc.lng),accuracy=Math.max(0,Number(loc.accuracy)||0);
  if(!bool_(s.LocationRequired)){
    if(Number.isFinite(lat)&&Number.isFinite(lng)&&s.RoomLat!==''&&s.RoomLng!==''){
      const d=haversineM_(Number(s.RoomLat),Number(s.RoomLng),lat,lng);
      return {
        verified:d<=Number(s.RadiusM)+(Math.min(accuracy,100)),
        reason:d<=Number(s.RadiusM)+(Math.min(accuracy,100))?'VERIFIED':'OUTSIDE_RADIUS',
        distanceM:Math.round(d),accuracyM:Math.round(accuracy),lat,lng
      };
    }
    return {
      verified:false,
      reason:'NOT_AVAILABLE',
      distanceM:'',
      accuracyM:accuracy?Math.round(accuracy):'',
      lat:Number.isFinite(lat)?lat:'',
      lng:Number.isFinite(lng)?lng:''
    };
  }

  if(!Number.isFinite(lat)||!Number.isFinite(lng)){
    return {
      verified:false,
      reason:'NOT_AVAILABLE',
      message:'Location could not be verified.',
      distanceM:'',
      accuracyM:''
    };
  }

  const d=haversineM_(Number(s.RoomLat),Number(s.RoomLng),lat,lng);
  const allowance=Number(s.RadiusM)+Math.min(accuracy,100);
  const verified=d<=allowance;
  return {
    verified:verified,
    reason:verified?'VERIFIED':'OUTSIDE_RADIUS',
    message:verified?'':'Your reported location is outside the allowed classroom area.',
    distanceM:Math.round(d),
    accuracyM:Math.round(accuracy),
    lat,lng
  };
}
function haversineM_(lat1,lon1,lat2,lon2){
  const R=6371000,toRad=x=>x*Math.PI/180;
  const p1=toRad(lat1),p2=toRad(lat2),dp=toRad(lat2-lat1),dl=toRad(lon2-lon1);
  const a=Math.sin(dp/2)**2+Math.cos(p1)*Math.cos(p2)*Math.sin(dl/2)**2;
  return 2*R*Math.asin(Math.sqrt(a));
}
function sanitizeAttendanceSession_(s){
  const sectionId=String(s.SectionID||'');
  const sec=sectionId?listCourseSections_().find(x=>String(x.sectionId||x.SectionID||'')===sectionId):null;
  return {
    sessionId:s.SessionID,sectionId:sectionId,sectionCode:sec?String(sec.sectionCode||sec.SectionCode||''):'',title:s.Title,startTime:dateIso_(s.StartTime),endTime:dateIso_(s.EndTime),
    status:s.Status,rotateSeconds:Number(s.RotateSeconds)||60,locationRequired:bool_(s.LocationRequired),
    roomLat:s.RoomLat,roomLng:s.RoomLng,radiusM:s.RadiusM,published:bool_(s.Published)
  };
}


function renameTrackerSpreadsheet_() {
  const ss = requireDb_();
  const courseCode = safeSheetNamePart_(getConfig_('COURSE_CODE')) || 'Course';
  const term = safeSheetNamePart_(getConfig_('TERM')) || 'Term';
  const desiredName = `${courseCode}_${term}_Learning_Progress_Tracker`;
  ss.rename(desiredName);
  return desiredName;
}

function safeSheetNamePart_(value) {
  return String(value || '')
    .trim()
    .replace(/[\\/:*?"<>|#%{}\[\]]+/g, '-')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^[_\-.]+|[_\-.]+$/g, '')
    .substring(0, 80);
}

function applySystemSheetVisibility_() {
  const ss = requireDb_();
  const systemNames = Object.values(APP.tables);

  systemNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    try {
      sh.hideSheet();
    } catch (e) {
      // Google Sheets requires at least one visible sheet.
    }
  });

  let home = ss.getSheetByName('Home');
  if (!home) {
    home = ss.insertSheet('Home', 0);
  }

  home.showSheet();
  home.clear();

  home.getRange('A1').setValue(APP.name);
  home.getRange('A2').setValue('This spreadsheet stores the course tracker data. Use the web app for normal operation.');
  home.getRange('A4').setValue('Course');
  home.getRange('B4').setValue(getConfig_('COURSE_CODE') || 'Not configured');
  home.getRange('A5').setValue('Term');
  home.getRange('B5').setValue(getConfig_('TERM') || 'Not configured');
  home.getRange('A7').setValue('System sheets are hidden automatically.');
  home.getRange('A1:B1').setFontWeight('bold');
  home.autoResizeColumns(1, 2);

  ss.setActiveSheet(home);
}

function showSystemSheets() {
  requireStaff_(['OWNER','INSTRUCTOR']);
  const ss = requireDb_();
  Object.values(APP.tables).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.showSheet();
  });
  return {ok:true};
}

function hideSystemSheets() {
  requireStaff_(['OWNER','INSTRUCTOR']);
  applySystemSheetVisibility_();
  return {ok:true};
}


/* ---------- Table helpers ---------- */

function dbOrNull_(){
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(APP.propBoundSheetId);

  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      props.deleteProperty(APP.propBoundSheetId);
    }
  }

  // Fallback is useful only when called from the bound editor itself.
  // Once setupProject() runs, deployed web requests use the stored ID above.
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      props.setProperty(APP.propBoundSheetId, active.getId());
      return active;
    }
  } catch (e) {}

  return null;
}
function requireDb_(){
  const db = dbOrNull_();
  if(!db) throw new Error('This Apps Script must be bound to a Google Sheet.');
  return db;
}
function table_(name){
  if(!name) throw new Error('Missing system table configuration.');
  const sh=requireDb_().getSheetByName(name);
  if(!sh) throw new Error('Missing system table '+name);
  return sh;
}
function ensureSheet_(ss,name,headers){
  let sh=ss.getSheetByName(name);
  if(!sh){
    sh=ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    resetRequestDataCache_();
  }else if(sh.getLastRow()===0){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    resetRequestDataCache_();
  }
  sh.setFrozenRows(1);
  return sh;
}
function rowsToObjects_(sh){
  return readRowsCached_(sh);
}
function appendObject_(sh,obj){
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h=>obj[h]!==undefined?obj[h]:''));
  resetRequestDataCache_();
}
function updateObjectRow_(sh,rowNum,obj){
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.getRange(rowNum,1,1,headers.length).setValues([headers.map(h=>obj[h]!==undefined?obj[h]:'')]);
  resetRequestDataCache_();
}
function updateRowObject_(sh,rowNum,obj){return updateObjectRow_(sh,rowNum,obj);}
function rewriteObjects_(sh,objs){
  const lastRow=sh.getLastRow();
  const lastCol=sh.getLastColumn();
  if(lastRow>1) sh.getRange(2,1,lastRow-1,lastCol).clearContent();

  if(objs.length){
    const headers=sh.getRange(1,1,1,lastCol).getValues()[0];
    sh.getRange(2,1,objs.length,headers.length).setValues(
      objs.map(o=>headers.map(h=>o[h]!==undefined?o[h]:''))
    );
  }
  resetRequestDataCache_();
}
function clearDataRows_(sh){
  if(sh.getLastRow()>1){
    sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
    resetRequestDataCache_();
  }
}
function findObject_(sh,key,value){
  return rowsToObjects_(sh).find(r=>String(r[key])===String(value))||null;
}
function findComposite_(sh,keys,vals){
  return rowsToObjects_(sh).find(r=>keys.every((k,i)=>String(r[k])===String(vals[i])))||null;
}
function upsertComposite_(sh,keys,vals,obj){
  const f=findComposite_(sh,keys,vals);
  if(f) updateObjectRow_(sh,f._row,obj);
  else appendObject_(sh,obj);
}
function deleteComposite_(sh,keys,vals){
  const f=findComposite_(sh,keys,vals);
  if(f){
    sh.deleteRow(f._row);
    resetRequestDataCache_();
  }
}
function repairOrphanStudentSectionsV2526_(){
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('LPT_V2526_STUDENT_SECTION_REPAIR_DONE')==='1')return;
  ensureHeaders_(APP.tables.STUDENT_SECTIONS);
  const sectionRows=rowsToObjects_(table_(APP.tables.COURSE_SECTIONS));
  const activeSections=sectionRows.filter(r=>bool_(r.Active));
  const activeIds={};activeSections.forEach(r=>activeIds[String(r.SectionID||'')]=r);
  const activeByCode={};activeSections.forEach(r=>{activeByCode[sectionCanonicalKey_(r.SectionCode)]=r;});
  const oldIdToCode={};
  // Recover the original Section code from audit records created when registrar
  // mappings were assigned. This lets us repair mappings orphaned by older
  // versions that recreated SectionIDs on every Course save.
  try{
    rowsToObjects_(table_(APP.tables.AUDIT)).forEach(a=>{
      if(String(a.Action||'')!=='ASSIGN_STUDENT_SECTION')return;
      let d={};try{d=JSON.parse(String(a.Details||'{}'));}catch(e){}
      const code=String(d.sectionCode||'').trim(),id=String(a.EntityID||'').trim();
      if(id&&code)oldIdToCode[id]=code;
    });
  }catch(e){}
  const sh=table_(APP.tables.STUDENT_SECTIONS),now=new Date();let repaired=0;
  rowsToObjects_(sh).forEach(r=>{
    if(!bool_(r.Active))return;
    const oldId=String(r.SectionID||'');
    if(activeIds[oldId])return;
    const code=oldIdToCode[oldId];
    const target=code?activeByCode[sectionCanonicalKey_(code)]:null;
    if(target){updateObjectRow_(sh,r._row,{SectionID:String(target.SectionID||''),UpdatedAt:now});repaired++;}
  });
  props.setProperty('LPT_V2526_STUDENT_SECTION_REPAIR_DONE','1');
  if(repaired)resetRequestDataCache_();
}
function stripRow_(o){const x=Object.assign({},o);delete x._row;return x;}
function setConfig_(key,value){
  const sh=table_(APP.tables.CONFIG);
  upsertComposite_(sh,['Key'],[key],{Key:key,Value:value});
}
function getConfig_(key){
  if(!dbOrNull_())return'';
  const f=findObject_(table_(APP.tables.CONFIG),'Key',key);return f?String(f.Value):'';
}
function courseConfig_(){
  const semester=String(getConfig_('SEMESTER')||'').trim();
  let year=String(getConfig_('ACADEMIC_YEAR')||'').trim();
  const rawTerm=String(getConfig_('TERM')||'').trim();
  const dateYear=year.match(/^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+\w{3}\s+\d{1,2}\s+(\d{4})\s+/i);
  if(dateYear)year=dateYear[1];
  let term='';
  if(semester&&year)term=`${semester}/${year}`;
  else if(year)term=year;
  else if(rawTerm&&!/GMT[+-]\d{4}/.test(rawTerm))term=rawTerm;
  return {code:getConfig_('COURSE_CODE'),name:getConfig_('COURSE_NAME'),term:term,aliasPrefix:getConfig_('ALIAS_PREFIX')};
}
function appPublicConfig_(){
  return {name:APP.name,version:APP.version,categories:APP.categories,purposes:APP.purposes,modes:APP.modes,statuses:APP.statuses,cloMin:APP.cloMin,cloMax:APP.cloMax,cloAttainedLevel:APP.cloAttainedLevel};
}
function listStudentsForStaff_(){return normalizedStudentsFromRows_(rowsToObjects_(table_(APP.tables.STUDENTS)));}
function activeStudentSectionRows_(){
  ensureHeaders_(APP.tables.STUDENT_SECTIONS);
  return rowsToObjects_(table_(APP.tables.STUDENT_SECTIONS)).filter(r=>bool_(r.Active));
}
function studentSectionLookup_(){
  const sections=listCourseSections_();
  const sectionById={};sections.forEach(s=>sectionById[String(s.sectionId||s.SectionID||'')]=s);
  const byStudent={};
  activeStudentSectionRows_().forEach(r=>{
    const s=sectionById[String(r.SectionID||'')];
    if(s)byStudent[String(r.StudentKey||'')]={sectionId:String(r.SectionID||''),sectionCode:String(s.sectionCode||s.SectionCode||'')};
  });
  return {byStudent,sections};
}
function studentBelongsToSection_(studentKey,sectionId){
  if(!sectionId)return true;
  return activeStudentSectionRows_().some(r=>String(r.StudentKey||'')===String(studentKey)&&String(r.SectionID||'')===String(sectionId));
}
function assignStudentsToSectionByIds(sectionCode,studentIds){
  const actor=requireStaff_(['OWNER','INSTRUCTOR']);
  ensureCurrentSchema_();
  sectionCode=clean_(sectionCode,80);
  const ids=[...new Set((Array.isArray(studentIds)?studentIds:[]).map(x=>clean_(x,60)).filter(Boolean))];
  if(!sectionCode||!ids.length)return {ok:true,assigned:0,unchanged:0,moved:0};
  const digitKey=s=>/^\d+$/.test(String(s||'').trim())?String(s||'').trim().replace(/^0+(?=\d)/,''):'';
  let section=listCourseSections_().find(s=>String(s.sectionCode||s.SectionCode||'').trim()===sectionCode);
  if(!section&&digitKey(sectionCode))section=listCourseSections_().find(s=>digitKey(s.sectionCode||s.SectionCode||'')===digitKey(sectionCode));
  if(!section)throw new Error('Section '+sectionCode+' was not found. Import/update the Section first.');
  const sectionId=String(section.sectionId||section.SectionID||'');
  const students=listStudentsForStaff_();
  const byId={};students.forEach(s=>{if(s.StudentId)byId[String(s.StudentId).trim()]=s;});
  const sh=table_(APP.tables.STUDENT_SECTIONS);
  const mappings=rowsToObjects_(sh);
  let assigned=0,unchanged=0,moved=0;const now=new Date();
  ids.forEach(id=>{
    const st=byId[id];if(!st)return;
    const active=mappings.filter(r=>String(r.StudentKey||'')===String(st.StudentKey)&&bool_(r.Active));
    const same=active.find(r=>String(r.SectionID||'')===sectionId);
    if(same){unchanged++;return;}
    active.forEach(r=>{updateObjectRow_(sh,r._row,Object.assign(stripRow_(r),{Active:false,UpdatedAt:now}));moved++;});
    const prior=mappings.find(r=>String(r.StudentKey||'')===String(st.StudentKey)&&String(r.SectionID||'')===sectionId);
    if(prior)updateObjectRow_(sh,prior._row,Object.assign(stripRow_(prior),{Active:true,UpdatedAt:now}));
    else appendObject_(sh,{StudentKey:st.StudentKey,SectionID:sectionId,Active:true,CreatedAt:now,UpdatedAt:now});
    assigned++;
  });
  audit_('ASSIGN_STUDENT_SECTION','Section',sectionId,{actor:actor.email,sectionCode,assigned,unchanged,moved});
  resetRequestDataCache_();
  return {ok:true,sectionId,sectionCode,assigned,unchanged,moved};
}
function listCLOs_(){return normalizedClosFromRows_(rowsToObjects_(table_(APP.tables.CLOS)));}
function listActivities_(){return normalizedActivitiesFromRows_(rowsToObjects_(table_(APP.tables.ACTIVITIES)));}
function getActivity_(id){return listActivities_().find(a=>a.ActivityID===id)||null;}
function activityCloIds_(activityId){return rowsToObjects_(table_(APP.tables.ACTIVITY_CLOS)).filter(r=>r.ActivityID===activityId).map(r=>r.CLOID);}
function nextActivitySequence_(){const a=listActivities_();return a.length?Math.max(...a.map(x=>Number(x.Sequence)||0))+1:1;}
function aliasFor_(n){
  const prefix=getConfig_('ALIAS_PREFIX')||getConfig_('COURSE_CODE')||'COURSE';
  return `${prefix}-S${String(n).padStart(3,'0')}`;
}
function nextAliasNumber_(students){
  let max=0;students.forEach(s=>{const m=String(s.Alias||'').match(/S(\d+)$/);if(m)max=Math.max(max,Number(m[1]));});return max+1;
}
function audit_(action,type,id,details){
  try{appendObject_(table_(APP.tables.AUDIT),{Timestamp:new Date(),Actor:currentEmail_(true),Action:action,EntityType:type,EntityID:id,Details:JSON.stringify(details||{})});}catch(e){}
}
function sortSequence_(a,b){return (Number(a.Sequence)||0)-(Number(b.Sequence)||0);}
function normalizeEmail_(v){return String(v||'').trim().toLowerCase();}
function clean_(v,max){return String(v==null?'':v).trim().substring(0,max||500);}
function bool_(v){return v===true||String(v).toUpperCase()==='TRUE'||String(v)==='1';}
function clampNum_(v,min,max,def){const n=Number(v);return Number.isFinite(n)?Math.max(min,Math.min(max,n)):def;}
function finiteOrBlank_(v){if(v===''||v==null)return'';const n=Number(v);return Number.isFinite(n)?n:'';}
function dateOnly_(v){
 if(!v)return'';
 if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
 const raw=String(v||'').trim();
 if(/^\d{4}-\d{2}-\d{2}$/.test(raw))return raw;
 const d=new Date(raw);
 return Number.isNaN(d.getTime())?'':Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
}
function dateIso_(v){if(!v)return'';const d=v instanceof Date?v:new Date(v);return Number.isNaN(d.getTime())?'':d.toISOString();}


// ============================================================================
// TQF3 DOCX IMPORT — v2.57.0
// Parses common TU TQF3/Detailed Course Specification DOCX structures and
// presents a selective preview before any data is applied.
// ============================================================================
function tqf3XmlDecode_(s){
  return String(s||'')
    .replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"')
    .replace(/&apos;/g,"'").replace(/&amp;/g,'&')
    .replace(/&#(\d+);/g,function(_,n){return String.fromCharCode(Number(n)||0);})
    .replace(/&#x([0-9a-f]+);/gi,function(_,n){return String.fromCharCode(parseInt(n,16)||0);});
}
function tqf3CellText_(xml){
  var paras=String(xml||'').match(/<w:p\b[\s\S]*?<\/w:p>/g)||[];
  var chunks=paras.map(function(p){
    var out='';
    var re=/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g,m;
    while((m=re.exec(p)))out+=tqf3XmlDecode_(m[1]);
    return out.trim();
  }).filter(Boolean);
  if(!chunks.length){
    var out='',re=/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g,m;
    while((m=re.exec(String(xml||''))))out+=tqf3XmlDecode_(m[1]);
    return out.trim();
  }
  return chunks.join('\n').replace(/[ \t]+\n/g,'\n').trim();
}
function tqf3DocxStructure_(base64,fileName){
  var bytes=Utilities.base64Decode(String(base64||''));
  if(!bytes.length)throw new Error('ไฟล์ว่างหรืออ่านไม่ได้');
  if(bytes.length>8*1024*1024)throw new Error('ไฟล์ TQF3 ใหญ่เกิน 8 MB');
  // DOCX is an OOXML ZIP package. Apps Script Utilities.unzip() expects a ZIP blob;
  // using the DOCX MIME type can make valid .docx files fail to unzip.
  if(bytes.length<4 || bytes[0]!==80 || bytes[1]!==75){
    throw new Error('ไฟล์นี้ไม่มีโครงสร้าง DOCX/ZIP ที่ถูกต้อง กรุณาเลือกไฟล์ .docx ที่บันทึกจาก Microsoft Word หรือ Google Docs');
  }
  var blob=Utilities.newBlob(bytes,'application/zip',clean_(fileName,200)||'TQF3.docx');
  var entries;
  try{entries=Utilities.unzip(blob);}catch(e){
    throw new Error('ไม่สามารถอ่านโครงสร้าง DOCX ได้ กรุณาบันทึกไฟล์เป็น .docx ใหม่แล้วลองอีกครั้ง');
  }
  var doc=entries.find(function(b){return String(b.getName()).replace(/\\/g,'/')==='word/document.xml';});
  if(!doc)throw new Error('ไม่พบ word/document.xml ในไฟล์ DOCX');
  var xml=doc.getDataAsString('UTF-8');
  var tables=[];
  (xml.match(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)||[]).forEach(function(tbl){
    var rows=[];
    (tbl.match(/<w:tr\b[\s\S]*?<\/w:tr>/g)||[]).forEach(function(tr){
      var cells=[];
      (tr.match(/<w:tc\b[\s\S]*?<\/w:tc>/g)||[]).forEach(function(tc){cells.push(tqf3CellText_(tc));});
      if(cells.length)rows.push(cells);
    });
    if(rows.length)tables.push(rows);
  });
  var paras=(xml.match(/<w:p\b[\s\S]*?<\/w:p>/g)||[]).map(tqf3CellText_).filter(Boolean);
  return {tables:tables,text:paras.join('\n'),paragraphs:paras};
}
function tqf3Norm_(s){return String(s||'').replace(/\u00a0/g,' ').replace(/[ \t]+/g,' ').replace(/\s*\n\s*/g,' ').trim();}
function tqf3FindTable_(tables,pred,last){
  var found=null;
  (tables||[]).forEach(function(t){if(pred(t)){if(!found||last)found=t;}});
  return found;
}
function tqf3ParseTimeRange_(s){
  var m=String(s||'').match(/(\d{1,2})[.:](\d{2})\s*[-–]\s*(\d{1,2})[.:](\d{2})/);
  if(!m)return {startTime:'',endTime:''};
  function hh(x){return String(Number(x)).padStart(2,'0');}
  return {startTime:hh(m[1])+':'+m[2],endTime:hh(m[3])+':'+m[4]};
}
function tqf3ParseSections_(tables){
  var t=tqf3FindTable_(tables,function(x){return x[0]&&tqf3Norm_(x[0][0])==='Section'&&x[0].some(function(c){return /วัน\/เวลา/.test(c);});},false);
  if(!t)return [];
  var days=['จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์','อาทิตย์'];
  return t.slice(1).map(function(r){
    var code=tqf3Norm_(r[0]),dt=tqf3Norm_(r[1]),roomMode=tqf3Norm_(r[2]),instructor=tqf3Norm_(r[3]);
    if(!code)return null;
    var tm=tqf3ParseTimeRange_(dt),matched=days.filter(function(d){return dt.indexOf(d)>=0;});
    var rm=roomMode.split(/\s*\/\s*/);var room=rm[0]||'',mode=rm[1]||'Onsite';
    return {sectionCode:code,schedules:matched.map(function(d){return {day:d,startTime:tm.startTime,endTime:tm.endTime,room:room};}),exams:[],mode:mode,instructorName:instructor,instructorEmail:''};
  }).filter(Boolean);
}
function tqf3ParseClos_(tables){
  var t=tqf3FindTable_(tables,function(x){return x[0]&&x[0].some(function(c){return /ผลลัพธ์การเรียนรู้รายวิชา CLOs/i.test(c);});},false);
  if(!t)return [];
  return t.slice(1).map(function(r){var code=tqf3Norm_(r[0]),desc=tqf3Norm_(r[1]);return /^CLO\d+$/i.test(code)&&desc?{code:code.toUpperCase(),description:desc}:null;}).filter(Boolean);
}
function tqf3ParseCloMethods_(tables){
  var t=tqf3FindTable_(tables,function(x){return x[0]&&x[0].some(function(c){return /วิธีการสอน/.test(c);})&&x[0].some(function(c){return /Key Assessment/i.test(c);});},false);
  var out={};if(!t)return out;
  t.slice(1).forEach(function(r){var first=tqf3Norm_(r[0]),m=first.match(/^(CLO\d+)/i);if(!m)return;out[m[1].toUpperCase()]={teachingMethod:tqf3Norm_(r[1]),assessmentMethod:tqf3Norm_(r[2]),keyAssessment:tqf3Norm_(r[3]),minimum:tqf3Norm_(r[4]),evidence:tqf3Norm_(r[5])};});
  return out;
}
function tqf3ParseWeekly_(tables,cloMethods){
  var t=tqf3FindTable_(tables,function(x){return x[0]&&x[0].some(function(c){return /สัปดาห์/.test(c);})&&x[0].some(function(c){return /หัวข้อ/.test(c);})&&x[0].some(function(c){return /กิจกรรมการเรียนรู้/.test(c);});},false);
  if(!t)return [];
  return t.slice(1).map(function(r){
    var week=Number(tqf3Norm_(r[0]));if(!Number.isFinite(week)||week<1)return null;
    var topicCell=String(r[1]||'').trim();var lines=topicCell.split(/\n+/).map(function(x){return x.trim();}).filter(Boolean);
    var compactTopic=tqf3Norm_(topicCell),bulletAt=compactTopic.indexOf('•');var topic=lines[0]||compactTopic;if(bulletAt>0)topic=compactTopic.substring(0,bulletAt).trim();var description=lines.slice(1).join('\n');
    if(!description&&bulletAt>0)description=compactTopic.substring(bulletAt).trim();
    var activityText=tqf3Norm_(r[2]);var codes=(String(r[3]||'').match(/CLO\d+/gi)||[]).map(function(x){return x.toUpperCase();});
    return {weekNo:week,topic:topic,description:description,learningActivities:activityText,cloCodes:codes,cloPlans:codes.map(function(code){var cm=cloMethods[code]||{};return {cloCode:code,teachingMethod:cm.teachingMethod||activityText,assessmentMethod:cm.assessmentMethod||'',activityTitles:[]};})};
  }).filter(Boolean);
}
function tqf3CategoryFromTitle_(title){
  var s=String(title||'').toLowerCase();
  if(/exam|สอบ|แบบทดสอบ/.test(s))return 'Exam';
  if(/quiz/.test(s))return 'Quiz';
  if(/lab/.test(s))return 'Lab';
  if(/presentation|demonstration|สาธิต|นำเสนอ/.test(s))return 'Presentation';
  if(/project|โครงงาน|proposal/.test(s))return 'Project Milestone';
  if(/assignment|งาน/.test(s))return 'Assignment';
  return 'Other';
}
function tqf3ParseActivities_(tables){
  var t=tqf3FindTable_(tables,function(x){return x[0]&&x[0].some(function(c){return /งานประเมิน/.test(c);})&&x[0].some(function(c){return /น้ำหนักคะแนน/.test(c);})&&x[0].some(function(c){return /CLO ที่วัด/.test(c);});},true);
  if(!t)return [];
  return t.slice(1).map(function(r){
    var title=tqf3Norm_(r[0]);if(!title||/^รวม\s*100/i.test(title))return null;
    var weightM=tqf3Norm_(r[1]).match(/([\d.]+)\s*%/),cloCodes=(String(r[2]||'').match(/CLO\d+/gi)||[]).map(function(x){return x.toUpperCase();});
    var mode=/กลุ่ม/.test(String(r[3]||''))?'Team':'Individual';var due=tqf3Norm_(r[6]);var wk=due.match(/สัปดาห์ที่\s*(\d+)/);
    return {title:title,shortDescription:tqf3Norm_(r[5]),category:tqf3CategoryFromTitle_(title),mode:mode,purpose:'Assessed',weight:weightM?Number(weightM[1]):'',maxScore:'',assignedWeek:wk?Number(wk[1]):'',dueDate:'',cloCodes:cloCodes,keyAssessment:/[✓☒]/.test(String(r[4]||'')),scheduleText:due};
  }).filter(Boolean);
}
function tqf3ParseStaff_(tables){
  var t=tqf3FindTable_(tables,function(x){return x[0]&&x[0].some(function(c){return /ชื่อ-นามสกุล/.test(c);})&&x[0].some(function(c){return /อาจารย์ผู้สอน/.test(c);});},false);
  if(!t)return [];
  return t.slice(1).map(function(r){var name=tqf3Norm_(r[0]),contact=tqf3Norm_(r[1]),m=contact.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);return name?{displayName:name,email:m?m[0].toLowerCase():'',role:'INSTRUCTOR'}:null;}).filter(Boolean);
}
function tqf3ParseProfile_(text){
  var p={};var t=String(text||'');var m;
  m=t.match(/รหัสวิชาและชื่อวิชาภาษาไทย\s*([^\s]+)\s+([^\n]+)/);if(m){p.courseCodeTh=tqf3Norm_(m[1]);p.courseNameTh=tqf3Norm_(m[2]);}
  m=t.match(/รหัสวิชาและชื่อวิชาภาษาอังกฤษ\s*([^\s]+)\s+([^\n]+)/);if(m){p.courseCodeEn=tqf3Norm_(m[1]);p.courseNameEn=tqf3Norm_(m[2]);}
  m=t.match(/จำนวนหน่วยกิต\s*([\d.]+)\s*\(\s*(\d+)\s*-\s*(\d+)\s*-\s*(\d+)\s*\)/);if(m){p.totalCredits=m[1];p.lectureHours=m[2];p.practiceHours=m[3];p.selfStudyHours=m[4];}
  m=t.match(/ภาค\/ปีการศึกษา\s*([^\s/]+)\s*\/\s*([^\s\n]+)/);if(m){p.semester=tqf3Norm_(m[1]);p.academicYear=tqf3Norm_(m[2]);}
  m=t.match(/เงื่อนไขวิชาบังคับก่อน[\s\S]*?\n\s*([^\n]+?)\n\s*2\./);if(m)p.prerequisite=tqf3Norm_(m[1]);
  m=t.match(/3\.\s*ช่องทางที่อาจารย์ให้คำปรึกษา[\s\S]*?\n([^\n]+(?:\n(?!4\.).+)*)\n\s*4\./);if(m)p.consultationChannel=tqf3Norm_(m[1]);
  var d=t.match(/5\.\s*คำอธิบายรายวิชา[\s\S]*?ภาษาไทย\s*([\s\S]*?)\s*English\s*([\s\S]*?)\s*6\.\s*วัตถุประสงค์/);if(d){p.descriptionTh=tqf3Norm_(d[1]);p.descriptionEn=tqf3Norm_(d[2]);}
  return p;
}
function tqf3ParseObjectives_(text){
  var t=String(text||''),m=t.match(/6\.\s*วัตถุประสงค์ของรายวิชา[\s\S]*?\n([\s\S]*?)\n\s*7\./);if(!m)return [];
  var block=m[1],re=/(\d+)\)\s*([\s\S]*?)(?=\n\s*\d+\)|$)/g,x,out=[];
  while((x=re.exec(block))){var v=tqf3Norm_(x[2]);if(v)out.push(v);}
  if(!out.length)out=block.split(/\n+/).map(tqf3Norm_).filter(Boolean).slice(0,20);
  return out;
}
function tqf3ParsePolicy_(tables,text){
  var t=tqf3FindTable_(tables,function(x){return x[0]&&x[0].some(function(c){return /เกณฑ์ที่กำหนด/.test(c);});},false),out={thresholdLevel:null,scorePct:null,courseTargetPct:null,combination:'',note:''};
  if(t){t.slice(1).forEach(function(r){var k=tqf3Norm_(r[0]),v=tqf3Norm_(r[1]);if(/เกณฑ์ผ่านระดับนักศึกษา/.test(k)){var a=v.match(/ร้อยละ\s*(\d+)/),b=v.match(/ระดับ\s*(\d+)\s*จาก\s*5/);if(a)out.scorePct=Number(a[1]);if(b)out.thresholdLevel=Number(b[1]);}if(/เป้าหมายระดับรายวิชา/.test(k)){var c=v.match(/ร้อยละ\s*(\d+)/);if(c)out.courseTargetPct=Number(c[1]);}if(/วิธีรวมผล/.test(k))out.combination=v;});}
  if(out.combination&&!/LATEST|AVERAGE|WEIGHTED|MIN_COUNT/i.test(out.combination))out.note='วิธีรวมผลจาก TQF3 อาจไม่ตรงกับ policy ที่ระบบรองรับ โปรดตรวจสอบหลังนำเข้า';
  return out;
}
function tqf3PreviewStatus_(exists,same){return !exists?'new':(same?'same':'existing');}
function previewTqf3DocxImport(filePayload){
  requireStaff_(['OWNER','INSTRUCTOR']);ensureCurrentSchema_();filePayload=filePayload||{};
  var st=tqf3DocxStructure_(filePayload.base64,filePayload.fileName),profile=tqf3ParseProfile_(st.text),objectives=tqf3ParseObjectives_(st.text),clos=tqf3ParseClos_(st.tables),methods=tqf3ParseCloMethods_(st.tables),sections=tqf3ParseSections_(st.tables),staff=tqf3ParseStaff_(st.tables),weekly=tqf3ParseWeekly_(st.tables,methods),activities=tqf3ParseActivities_(st.tables),policy=tqf3ParsePolicy_(st.tables,st.text);
  var cur=getCourseSettings(),curClos=listCLOs_(),curSec=listCourseSections_(),curStaff=listCourseStaffProfiles_(),curWeek=listWeeklyTeachingPlan_(),curActs=listActivities_().filter(function(a){return bool_(a.Active);});
  var fields=[['courseCodeTh','รหัสวิชา (TH)'],['courseCodeEn','Course Code'],['courseNameTh','ชื่อวิชา (TH)'],['courseNameEn','Course Name'],['totalCredits','หน่วยกิต'],['lectureHours','Lecture'],['practiceHours','Practice/Lab'],['selfStudyHours','Self-study'],['semester','ภาคการศึกษา'],['academicYear','ปีการศึกษา'],['descriptionTh','คำอธิบายรายวิชา (TH)'],['descriptionEn','Course Description'],['prerequisite','Prerequisite'],['consultationChannel','ช่องทางให้คำปรึกษา']];
  var courseItems=fields.filter(function(f){return profile[f[0]]!==undefined&&String(profile[f[0]]).trim()!=='';}).map(function(f){var old=String(cur[f[0]]||''),val=String(profile[f[0]]||'');return {key:'course.'+f[0],field:f[0],label:f[1],value:val,current:old,status:tqf3PreviewStatus_(!!old,old===val)};});
  var objItems=objectives.map(function(v,i){var exists=(cur.objectives||[]).some(function(o){return tqf3Norm_(o.text)===tqf3Norm_(v);});return {key:'objective.'+i,label:'วัตถุประสงค์ '+(i+1),value:v,status:exists?'same':'new'};});
  var cloItems=clos.map(function(x){var old=curClos.find(function(c){return String(c.Code||'').toUpperCase()===x.code;});return Object.assign({key:'clo.'+x.code,label:x.code,status:tqf3PreviewStatus_(!!old,old&&tqf3Norm_(old.Description)===tqf3Norm_(x.description))},x);});
  var sectionItems=sections.map(function(x){var old=curSec.find(function(c){return String(c.sectionCode||'')===x.sectionCode;});return Object.assign({key:'section.'+x.sectionCode,label:'Section '+x.sectionCode,status:tqf3PreviewStatus_(!!old,false)},x);});
  var staffItems=staff.map(function(x,i){var old=curStaff.find(function(c){return x.email&&String(c.email||'').toLowerCase()===x.email;});return Object.assign({key:'staff.'+(x.email||i),label:x.displayName,status:old?'same':'new'},x);});
  var weekItems=weekly.map(function(x){var old=curWeek.find(function(c){return Number(c.weekNo)===Number(x.weekNo);});return Object.assign({key:'week.'+x.weekNo,label:'Week '+x.weekNo+' · '+x.topic,status:tqf3PreviewStatus_(!!old,false)},x);});
  var actItems=activities.map(function(x,i){var old=curActs.find(function(c){return tqf3Norm_(c.Title).toLowerCase()===tqf3Norm_(x.title).toLowerCase();});return Object.assign({key:'activity.'+i,label:x.title,status:tqf3PreviewStatus_(!!old,false)},x);});
  var warnings=[];if(!clos.length)warnings.push('ไม่พบตาราง CLO ในรูปแบบที่ระบบรู้จัก');if(!weekly.length)warnings.push('ไม่พบตารางแผนรายสัปดาห์ในรูปแบบที่ระบบรู้จัก');if(policy.note)warnings.push(policy.note);
  return {ok:true,fileName:clean_(filePayload.fileName,200),parsed:{profile:profile,objectives:objectives,clos:clos,sections:sections,staff:staff,weekly:weekly,activities:activities,policy:policy},groups:[{id:'course',label:'ข้อมูลรายวิชา',items:courseItems},{id:'objectives',label:'วัตถุประสงค์รายวิชา',items:objItems},{id:'clos',label:'CLO',items:cloItems},{id:'sections',label:'Section และตารางเรียน',items:sectionItems},{id:'staff',label:'ผู้สอน',items:staffItems},{id:'weekly',label:'แผนรายสัปดาห์',items:weekItems},{id:'activities',label:'กิจกรรม / Key Assessment',items:actItems}],policy:policy,warnings:warnings};
}
function tqf3Selected_(selected,key){return Array.isArray(selected)&&selected.indexOf(key)>=0;}
function tqf3MergeCourseConfig_(preview,selected,overwrite){
  var p=preview.parsed.profile||{},cur=getCourseSettings(),map={courseCodeTh:'COURSE_CODE_TH',courseCodeEn:'COURSE_CODE_EN',courseNameTh:'COURSE_NAME_TH',courseNameEn:'COURSE_NAME_EN',totalCredits:'TOTAL_CREDITS',lectureHours:'LECTURE_HOURS',practiceHours:'PRACTICE_HOURS',selfStudyHours:'SELF_STUDY_HOURS',semester:'SEMESTER',academicYear:'ACADEMIC_YEAR',descriptionTh:'DESCRIPTION_TH',descriptionEn:'DESCRIPTION_EN',prerequisite:'PREREQUISITE',consultationChannel:'CONSULTATION_CHANNEL'};
  Object.keys(map).forEach(function(field){if(!tqf3Selected_(selected,'course.'+field))return;var old=String(cur[field]||'');if(old&&!overwrite)return;setConfig_(map[field],clean_(p[field],field.indexOf('description')>=0?5000:2000));});
  var codeEn=String(getConfig_('COURSE_CODE_EN')||''),codeTh=String(getConfig_('COURSE_CODE_TH')||''),nameEn=String(getConfig_('COURSE_NAME_EN')||''),nameTh=String(getConfig_('COURSE_NAME_TH')||'');
  setConfig_('COURSE_CODE',codeEn||codeTh);setConfig_('COURSE_NAME',nameEn||nameTh);if(getConfig_('TOTAL_CREDITS'))setConfig_('CREDITS',getConfig_('TOTAL_CREDITS'));
}
function applyTqf3DocxImport(preview,selected,options){
  var actor=requireStaff_(['OWNER','INSTRUCTOR']);ensureCurrentSchema_();preview=preview||{};selected=Array.isArray(selected)?selected:[];options=options||{};var overwrite=bool_(options.overwriteExisting),applied=[];
  tqf3MergeCourseConfig_(preview,selected,overwrite);
  if(selected.some(function(k){return k.indexOf('objective.')===0;})){
    var current=listCourseObjectives_().map(function(o){return o.text;}),incoming=preview.parsed.objectives||[];
    incoming.forEach(function(v,i){if(!tqf3Selected_(selected,'objective.'+i))return;if(!current.some(function(x){return tqf3Norm_(x)===tqf3Norm_(v);}))current.push(v);});replaceCourseObjectives_(current.map(function(x){return {text:x};}));applied.push('objectives');
  }
  if(selected.some(function(k){return k.indexOf('staff.')===0;})){
    (preview.parsed.staff||[]).forEach(function(x,i){var key='staff.'+(x.email||i);if(!tqf3Selected_(selected,key))return;var current=listCourseStaffProfiles_().find(function(s){return x.email&&String(s.email||'').toLowerCase()===String(x.email).toLowerCase();});if(!current)addCourseStaffProfile_({displayName:x.displayName,email:x.email,role:x.role||'INSTRUCTOR'});});applied.push('staff');
  }
  if(selected.some(function(k){return k.indexOf('section.')===0;})){
    var secs=listCourseSections_().slice();(preview.parsed.sections||[]).forEach(function(x){var key='section.'+x.sectionCode;if(!tqf3Selected_(selected,key))return;var idx=secs.findIndex(function(s){return String(s.sectionCode||'')===String(x.sectionCode);});if(idx<0)secs.push(x);else if(overwrite)secs[idx]=Object.assign({},secs[idx],x,{sectionId:secs[idx].sectionId});});replaceCourseSections_(secs);applied.push('sections');
  }
  if(selected.some(function(k){return k.indexOf('clo.')===0;})){
    var rows=listCLOs_().filter(function(c){return bool_(c.Active);}).map(function(c){return {code:String(c.Code||''),description:String(c.Description||'')};});(preview.parsed.clos||[]).forEach(function(x){if(!tqf3Selected_(selected,'clo.'+x.code))return;var idx=rows.findIndex(function(r){return r.code.toUpperCase()===x.code;});if(idx<0)rows.push({code:x.code,description:x.description});else if(overwrite)rows[idx].description=x.description;});saveCLOs(rows);applied.push('clos');
  }
  var cloByCode={};listCLOs_().forEach(function(c){cloByCode[String(c.Code||'').toUpperCase()]=String(c.CLOID||'');});
  if(selected.some(function(k){return k.indexOf('activity.')===0;})){
    var currentActs=listActivities_().filter(function(a){return bool_(a.Active);});(preview.parsed.activities||[]).forEach(function(x,i){if(!tqf3Selected_(selected,'activity.'+i))return;var old=currentActs.find(function(a){return tqf3Norm_(a.Title).toLowerCase()===tqf3Norm_(x.title).toLowerCase();});if(old&&!overwrite)return;var cloIds=(x.cloCodes||[]).map(function(c){return cloByCode[c];}).filter(Boolean);saveActivity({activityId:old?String(old.ActivityID||''):'',title:x.title,shortDescription:x.shortDescription||'',detailUrl:old?String(old.DetailUrl||''):'',category:x.category,mode:x.mode,purpose:'Assessed',maxScore:old?old.MaxScore:'',weight:x.weight===''?(old?old.Weight:''):x.weight,published:old?bool_(old.Published):false,studentVisibilityLevel:old?String(old.StudentVisibilityLevel||'HIDDEN'):'HIDDEN',assignedWeek:x.assignedWeek,dueDate:old?dateOnly_(old.DueDate):'',sequence:old?Number(old.Sequence||0):nextActivitySequence_(),cloIds:cloIds,cloEvalMode:old?String(old.CLOEvalMode||'MANUAL'):'MANUAL',cloScoreSource:old?String(old.CLOScoreSource||'ACTIVITY_SCORE'):'ACTIVITY_SCORE'});});applied.push('activities');
  }
  if(selected.some(function(k){return k.indexOf('week.')===0;})){
    var currentWeeks=listWeeklyTeachingPlan_().slice();(preview.parsed.weekly||[]).forEach(function(x){if(!tqf3Selected_(selected,'week.'+x.weekNo))return;var idx=currentWeeks.findIndex(function(w){return Number(w.weekNo)===Number(x.weekNo);});if(idx>=0&&!overwrite)return;var obj={planItemId:idx>=0?currentWeeks[idx].planItemId:'',weekNo:x.weekNo,plannedDate:'',topic:x.topic,description:x.description,cloPlans:(x.cloPlans||[]).map(function(cp){return {cloId:cloByCode[cp.cloCode]||'',teachingMethod:cp.teachingMethod||x.learningActivities||'',assessmentMethod:cp.assessmentMethod||'',activityIds:[]};}).filter(function(cp){return cp.cloId;}),plannedActivityIds:[],actualWeekNo:null,actualDate:'',actualStatus:'Planned',actualNote:'',trackChanges:true};if(idx>=0)currentWeeks[idx]=obj;else currentWeeks.push(obj);});currentWeeks.sort(function(a,b){return Number(a.weekNo)-Number(b.weekNo);});saveWeeklyTeachingPlan(currentWeeks);applied.push('weekly');
  }
  if(tqf3Selected_(selected,'policy.threshold')&&preview.policy&&preview.policy.thresholdLevel){listCLOs_().filter(function(c){return bool_(c.Active);}).forEach(function(c){var old=getCloAttainmentPolicy_(c.CLOID);saveCloAttainmentPolicy(c.CLOID,{policy:old.Policy,threshold:Number(preview.policy.thresholdLevel),minCount:old.MinCount});});applied.push('policyThreshold');}
  setConfig_('SETUP_COMPLETE','TRUE');resetRequestDataCache_();clearSetupStateCache_();audit_('IMPORT_TQF3_DOCX','Course','',{actor:actor.email,fileName:String(preview.fileName||''),selected:selected.length,overwrite:overwrite,applied:applied.join(',')});
  return {ok:true,applied:applied,settings:getCourseSettings(),setupState:getSetupState_()};
}
