/** SiteGauge review-request worker. Source-controlled copy; deploy into bound Apps Script.
 * Uses the existing Clients, Requests, Activity, Settings and Suppression tabs.
 * Default remains dry-run. Do not enable until authorized test succeeds.
 */
const SG_REVIEW_TABS={clients:'Clients',requests:'Requests',activity:'Activity',settings:'Settings',suppression:'Suppression'};
function sgReviewSheet_(ss,name){const sh=ss.getSheetByName(name);if(!sh)throw new Error('Missing tab: '+name);return sh;}
function sgReviewSetting_(ss,key,fallback){const rows=sgReviewSheet_(ss,SG_REVIEW_TABS.settings).getDataRange().getValues();const row=rows.find(r=>String(r[0]).trim()===key);return row?row[1]:fallback;}
function sgReviewEmail_(v){return String(v||'').trim().toLowerCase();}
function sgReviewSuppressed_(ss,clientId,email){
  const sh=sgReviewSheet_(ss,SG_REVIEW_TABS.suppression);
  if(sh.getLastRow()<2)return false;
  return sh.getRange(2,1,sh.getLastRow()-1,2).getValues().some(r=>
    sgReviewEmail_(r[1])===email&&(String(r[0]).trim()===String(clientId).trim()||String(r[0]).trim()==='*'));
}
function runReviewRequests(){
  const lock=LockService.getScriptLock();if(!lock.tryLock(10000))return {status:'LOCKED'};
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const dryRun=String(sgReviewSetting_(ss,'Dry run','TRUE')).toUpperCase()!=='FALSE';
    const cap=Math.max(0,Math.min(100,Number(sgReviewSetting_(ss,'Daily send limit',20))||0));
    const req=sgReviewSheet_(ss,SG_REVIEW_TABS.requests);
    const activity=sgReviewSheet_(ss,SG_REVIEW_TABS.activity);
    const data=req.getDataRange().getValues();
    const clients=sgReviewSheet_(ss,SG_REVIEW_TABS.clients).getDataRange().getValues().slice(1);
    const byId=new Map(clients.map(r=>[String(r[0]).trim(),r]));
    const tz=ss.getSpreadsheetTimeZone();
    const today=Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
    let sentToday=data.slice(1).filter(r=>r[7] instanceof Date&&Utilities.formatDate(r[7],tz,'yyyy-MM-dd')===today&&String(r[6]).trim()==='Sent').length;
    const result={mode:dryRun?'DRY RUN':'LIVE',eligible:0,sent:0,skipped:0,errors:0};
    const log=(id,client,action,detail)=>activity.appendRow([new Date(),id,client,action,detail]);
    for(let i=1;i<data.length&&sentToday<cap;i++){
      const r=data[i], [id,clientId,name,rawEmail,consent,eligible,status]=r;
      if(String(status).trim()!=='Pending')continue;
      const email=sgReviewEmail_(rawEmail), client=byId.get(String(clientId).trim());
      if(!id||!client||String(client[4]).trim()!=='Active'||String(consent).trim().toUpperCase()!=='YES'||String(eligible).trim().toUpperCase()!=='YES'||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){result.skipped++;continue;}
      if(sgReviewSuppressed_(ss,clientId,email)){result.skipped++;log(id,clientId,'SUPPRESSED','Recipient opted out');continue;}
      const url=String(client[3]||'').trim();
      if(!/^https:\/\/(?:g\.page|www\.google\.com|google\.com|maps\.google\.com)\//i.test(url)){result.skipped++;log(id,clientId,'SKIPPED','Invalid or missing Google review URL');continue;}
      const duplicate=data.slice(1).some((other,j)=>j!==i-1&&String(other[1]).trim()===String(clientId).trim()&&sgReviewEmail_(other[3])===email&&['Sent','Sending'].includes(String(other[6]).trim()));
      if(duplicate){result.skipped++;log(id,clientId,'SKIPPED','Previous request exists');continue;}
      result.eligible++;
      if(dryRun){log(id,clientId,'DRY_RUN','Eligible; no email sent');continue;}
      try{
        // Reserve before sending to avoid automatic duplicate delivery on partial failures.
        req.getRange(i+1,7).setValue('Sending');SpreadsheetApp.flush();
        const business=String(client[1]).trim();
        MailApp.sendEmail({to:email,subject:'How was your experience with '+business+'?',body:'Hi '+String(name||'there')+',\n\nThank you for choosing '+business+'. We welcome your honest feedback, whatever your experience. If you would like to leave a public Google review, you can do so here:\n'+url+'\n\nThank you,\n'+business+'\n\nTo opt out of future review invitations, reply to this email and let us know.',name:business,replyTo:'derek@derekverbrugge.com'});
        req.getRange(i+1,7,1,3).setValues([['Sent',new Date(),'']]);sentToday++;result.sent++;log(id,clientId,'SENT','Invitation sent');
      }catch(e){req.getRange(i+1,7).setValue('Error');req.getRange(i+1,9).setValue(String(e).slice(0,300));result.errors++;log(id,clientId,'ERROR','Manual review required: '+String(e).slice(0,300));}
    }
    return result;
  }finally{lock.releaseLock();}
}
function installDailyTrigger(){
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='runReviewRequests').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runReviewRequests').timeBased().everyDays(1).atHour(10).inTimezone('America/New_York').create();
}
