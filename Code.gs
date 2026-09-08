/**
 * Akshardham Associates — "daily expences" sheet connector (Kota + Jhalawar routing).
 *
 * Paste this INTO your existing sheet:
 *   Extensions -> Apps Script -> delete sample -> paste all -> Deploy ->
 *   New deployment -> Web app -> Execute as: Me, Who has access: Anyone -> Deploy -> copy URL.
 *   Put that URL into index.html:  const SCRIPT_URL = "...";
 *
 * Routing:  office = "jhalawar" -> tab "jhalawar Office"   |   else -> tab "kota office".
 * The Salary tab is never touched.
 * It finds the Date/Amount/Balance columns BY HEADER NAME, so it won't write to the wrong column,
 * appends below the last dated row, auto-fills S.No and the running Balance (prev +/- amount).
 */
var TABS = { kota: 'kota office', jhalawar: 'jhalawar Office' };

function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tabName = (d.office === 'jhalawar') ? TABS.jhalawar : TABS.kota;
    var sheet = ss.getSheetByName(tabName);
    if(!sheet) return json({ok:false, error:'Tab not found: ' + tabName});

    // --- locate header row + columns by name (robust to column position) ---
    var lastRowAll = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    var scan = sheet.getRange(1, 1, Math.min(20, lastRowAll || 1), lastCol).getValues();
    var want = { sno:'s.no', date:'date', desc:'descrip', exp:'expen', amt:'amount', head:'head', remark:'remark', bal:'balance' };
    var headerRow = 0, col = {};
    for(var r=0; r<scan.length; r++){
      var row = scan[r].map(function(x){ return String(x).trim().toLowerCase(); });
      if(row.indexOf('date') > -1 && row.indexOf('amount') > -1){
        headerRow = r + 1;
        for(var k in want){
          for(var c=0; c<row.length; c++){ if(row[c].indexOf(want[k]) > -1){ col[k] = c + 1; break; } }
        }
        break;
      }
    }
    if(!headerRow || !col.date || !col.amt) return json({ok:false, error:'Header row not found in ' + tabName});

    // --- last dated row (skip trailing blank/total rows) ---
    var lastRow = headerRow, nData = lastRowAll - headerRow;
    if(nData > 0){
      var dates = sheet.getRange(headerRow + 1, col.date, nData, 1).getValues();
      for(var i=0; i<dates.length; i++){ if(String(dates[i][0]).trim() !== '') lastRow = headerRow + 1 + i; }
    }

    var amt = Number(d.amount) || 0;

    // SHARED cash pool across BOTH offices (cash comes from Krishan Sir via Kota, both offices spend it).
    // Running balance is stored in one script property so Kota + Jhalawar draw from the same pool.
    // ponytail: property is the single source of truth for form entries; if you add a row BY HAND,
    //           re-run setStartingCash() with the correct current cash so the next form entry stays right.
    var props = PropertiesService.getScriptProperties();
    var prevBal = Number(props.getProperty('cashBalance')) || 0;
    var newBal = prevBal + (d.type === 'cashin' ? amt : -amt);
    props.setProperty('cashBalance', String(newBal));

    var prevSno = col.sno && lastRow > headerRow ? Number(sheet.getRange(lastRow, col.sno).getValue()) || 0 : 0;

    var newRow = lastRow + 1;
    function put(c, v){ if(c) sheet.getRange(newRow, c).setValue(v); }

    // employee/source/remark -> Remark column if the tab has one (Jhalawar), else into Description ( ... )
    var extra = [d.employee, d.source, d.remark].filter(function(x){ return x && String(x).trim(); }).join(' — ');

    put(col.sno,  prevSno + 1);
    put(col.date, d.date);
    if(col.remark){ put(col.desc, d.item || ''); put(col.remark, extra); }
    else          { put(col.desc, (d.item || '') + (extra ? ' (' + extra + ')' : '')); }
    put(col.exp,  d.type === 'cashin' ? 'Credit' : (d.category || ''));
    put(col.amt,  amt);
    if(col.head)  put(col.head, d.type === 'cashin' ? 'Credit' : (tabName === TABS.kota ? 'HD Office' : 'Jhalawar Office'));
    put(col.bal,  newBal);

    return json({ ok:true, tab:tabName, balance:newBal });
  } catch(err){
    return json({ ok:false, error:String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/**
 * RUN THIS ONCE to set the starting cash (the actual total cash in hand right now).
 * 1. Put today's real cash amount below.  2. Save.  3. Select "setStartingCash" in the toolbar -> Run.
 * After this, every form entry (Kota or Jhalawar) updates the shared balance automatically.
 */
function setStartingCash(){
  var CURRENT_CASH = 1113;   // <-- aaj ki actual cash (Kota tab ka Closing Balance). Badle to yahan update karo.
  PropertiesService.getScriptProperties().setProperty('cashBalance', String(CURRENT_CASH));
  Logger.log('Starting cash set to ' + CURRENT_CASH);
}
