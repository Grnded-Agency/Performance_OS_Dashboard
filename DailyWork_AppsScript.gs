/************************************************************************
 * Daily Client Work — backend for the Performance OS task board.
 *
 * Lives inside your existing Performance OS Google Sheet. It keeps two
 * HIDDEN tabs ("TasksDB" and "TasksPeople") that your supervisors never
 * open — they only use the dashboard board. The board reads and writes
 * through this script over JSONP (no CORS headaches on GitHub Pages).
 *
 * SETUP (one time):
 *   1. Open your Performance OS Google Sheet.
 *   2. Extensions ▸ Apps Script.  Delete anything there, paste ALL of this.
 *   3. Save.  Run ▸ setup  (authorize when Google asks).  It creates the
 *      two tabs and seeds the roster (Juan, Alyssa).
 *   4. Deploy ▸ New deployment ▸ type "Web app".
 *        Execute as:  Me
 *        Who has access:  Anyone
 *      Deploy, copy the Web app URL (ends in /exec).
 *   5. In your dashboard HTML (Performance_OS_Dashboard.html), find the line
 *        var DCW_API = "";      (top of the "DAILY CLIENT WORK" module, near
 *      the bottom of the file) and paste the URL between the quotes. Commit
 *      to GitHub Pages. Done — the "Daily Work" tab goes live.
 *   6. (Optional) Right-click each of the TasksDB / TasksPeople tabs ▸
 *      "Hide sheet" so they stay out of the way.
 *
 * If you ever edit this script, re-deploy with Deploy ▸ Manage deployments ▸
 * edit ▸ Version: New version, so the change goes live.
 ************************************************************************/

var TASKS_TAB  = "TasksDB";
var PEOPLE_TAB = "TasksPeople";
var SEED_PEOPLE = ["Juan", "Alyssa"];
var HEADERS = ["id","num","createdAt","updatedAt","status","client","type",
               "title","details","assignee","assignedBy","priority","reviewJson"];

/* ---- run this once from the editor ---- */
function setup(){ ensureSheets_(); }

function ensureSheets_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t = ss.getSheetByName(TASKS_TAB);
  if(!t){ t = ss.insertSheet(TASKS_TAB); }
  if(t.getLastRow() === 0){ t.getRange(1,1,1,HEADERS.length).setValues([HEADERS]); }
  var p = ss.getSheetByName(PEOPLE_TAB);
  if(!p){ p = ss.insertSheet(PEOPLE_TAB); p.getRange(1,1).setValue("name"); }
  if(p.getLastRow() < 2){
    for(var i=0;i<SEED_PEOPLE.length;i++){ p.appendRow([SEED_PEOPLE[i]]); }
  }
  return { ss:ss, tasks:t, people:p };
}

/* ================= HTTP entry points ================= */
function doGet(e){ return handle_(e); }
function doPost(e){ return handle_(e); }

function handle_(e){
  var params = (e && e.parameter) ? e.parameter : {};
  var cb = params.callback || "";
  var action = params.action || "list";
  var out;
  try{
    if(action === "list"){
      out = listAll_();
    }else{
      var payload = params.payload ? JSON.parse(params.payload) : {};
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try{
        if(action === "create")         out = createTask_(payload.task || {});
        else if(action === "update")    out = updateTask_(payload.id, payload.patch || {});
        else if(action === "delete")    out = deleteTask_(payload.id);
        else if(action === "addPerson") out = addPerson_(payload.name);
        else                            out = { ok:false, error:"unknown action" };
      } finally { lock.releaseLock(); }
    }
  }catch(err){
    out = { ok:false, error:String(err) };
  }
  return reply_(out, cb);
}

function reply_(obj, cb){
  var json = JSON.stringify(obj);
  if(cb){
    return ContentService.createTextOutput(cb + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= data ops ================= */
function listAll_(){
  var s = ensureSheets_();
  var tasks = [];
  var t = s.tasks;
  var last = t.getLastRow();
  if(last >= 2){
    var vals = t.getRange(2,1,last-1,HEADERS.length).getValues();
    for(var i=0;i<vals.length;i++){
      var row = vals[i];
      if(!row[0]) continue;                     // skip blank id
      var obj = {};
      for(var c=0;c<HEADERS.length;c++){ obj[HEADERS[c]] = row[c]; }
      obj.review = obj.reviewJson || "";        // board parses this
      delete obj.reviewJson;
      tasks.push(obj);
    }
  }
  var people = [];
  var p = s.people, pl = p.getLastRow();
  if(pl >= 2){
    var pv = p.getRange(2,1,pl-1,1).getValues();
    for(var j=0;j<pv.length;j++){ if(pv[j][0]) people.push(String(pv[j][0])); }
  }
  return { ok:true, tasks:tasks, people:people };
}

function findRow_(t, id){
  var last = t.getLastRow();
  if(last < 2) return -1;
  var ids = t.getRange(2,1,last-1,1).getValues();
  for(var i=0;i<ids.length;i++){ if(String(ids[i][0]) === String(id)) return i+2; }
  return -1;
}

function nextNum_(t){
  var last = t.getLastRow();
  if(last < 2) return 1;
  var nums = t.getRange(2,2,last-1,1).getValues();
  var max = 0;
  for(var i=0;i<nums.length;i++){ var n = Number(nums[i][0])||0; if(n>max) max=n; }
  return max + 1;
}

function createTask_(task){
  var s = ensureSheets_();
  var t = s.tasks;
  var now = Date.now();
  var id = Utilities.getUuid();
  var rec = {
    id:id, num:nextNum_(t), createdAt:now, updatedAt:now, status:"todo",
    client:task.client||"", type:task.type||"", title:task.title||"",
    details:task.details||"", assignee:task.assignee||"",
    assignedBy:task.assignedBy||"", priority:task.priority||"med", reviewJson:""
  };
  var rowArr = HEADERS.map(function(h){ return rec[h]; });
  t.appendRow(rowArr);
  return { ok:true, id:id, num:rec.num };
}

function updateTask_(id, patch){
  var s = ensureSheets_();
  var t = s.tasks;
  var r = findRow_(t, id);
  if(r < 0) return { ok:false, error:"not found" };
  var range = t.getRange(r,1,1,HEADERS.length);
  var row = range.getValues()[0];
  var idx = {}; for(var c=0;c<HEADERS.length;c++){ idx[HEADERS[c]] = c; }
  var editable = ["status","client","type","title","details","assignee","priority"];
  for(var k=0;k<editable.length;k++){
    var key = editable[k];
    if(patch.hasOwnProperty(key)) row[idx[key]] = patch[key];
  }
  if(patch.hasOwnProperty("review")){
    row[idx["reviewJson"]] = patch.review ? JSON.stringify(patch.review) : "";
  }
  row[idx["updatedAt"]] = Date.now();
  range.setValues([row]);
  return { ok:true };
}

function deleteTask_(id){
  var s = ensureSheets_();
  var t = s.tasks;
  var r = findRow_(t, id);
  if(r < 0) return { ok:false, error:"not found" };
  t.deleteRow(r);
  return { ok:true };
}

function addPerson_(name){
  if(!name) return { ok:false };
  var s = ensureSheets_();
  var p = s.people, last = p.getLastRow();
  if(last >= 2){
    var pv = p.getRange(2,1,last-1,1).getValues();
    for(var i=0;i<pv.length;i++){ if(String(pv[i][0]) === String(name)) return { ok:true, dup:true }; }
  }
  p.appendRow([name]);
  return { ok:true };
}
