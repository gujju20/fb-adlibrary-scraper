'use strict';
const $ = id => document.getElementById(id);
const busy = state => ['queued', 'running', 'stopping'].includes(state.status);
let selected = localStorage.getItem('studio-job');
let current = null;
let receivedAt = 0;
let logSignature = '';
let pollCount = 0;
const titles = {queued:'Waiting for a worker', starting:'Opening the browser', loading:'Discovering matching ads', scraping:'Reading ad details', saving:'Saving your workbook', assets:'Collecting images & videos', finished:'Files are ready', idle:'Ready'};
const fmt = seconds => { if (!Number.isFinite(seconds)) return '—'; seconds=Math.max(0,Math.round(seconds)); return seconds<60 ? `${seconds}s` : seconds<3600 ? `${Math.floor(seconds/60)}m ${seconds%60}s` : `${Math.floor(seconds/3600)}h ${Math.floor(seconds%3600/60)}m`; };
function error(message) { $('error').textContent=message; $('error').hidden=!message; }
async function api(url, options={}) {
  const response=await fetch(url, {credentials:'same-origin', ...options, headers:{'Content-Type':'application/json', ...options.headers}});
  const data=await response.json().catch(()=>({error:'The server could not complete this request. Please retry.'}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}
function selectedPhase(job) { return job.assets.status!=='idle' ? job.assets : job.scrape; }
function tick() {
  if (!current) return;
  const phase=selectedPhase(current), extra=busy(phase) ? (Date.now()-receivedAt)/1000 : 0;
  $('elapsed').textContent=fmt(phase.elapsed_seconds+extra);
  const estimate=phase.eta_seconds==null ? null : phase.eta_seconds-extra;
  $('eta').textContent=!busy(phase) ? 'Finished' : phase.phase==='loading' ? 'Total unknown' : estimate>0 ? `~${fmt(estimate)}` : phase.completed>=3 ? 'Updating…' : 'Calculating…';
  $('eta-label').textContent=current.assets.status!=='idle' ? 'MEDIA PHASE ESTIMATE' : 'DETAIL PHASE ESTIMATE';
}
function download(url) { const link=document.createElement('a'); link.href=url; link.download=''; document.body.append(link); link.click(); link.remove(); }
function render(job) {
  current=job; receivedAt=Date.now();
  $('empty').hidden=true; $('job-content').hidden=false;
  const phase=selectedPhase(job), running=busy(job.scrape)||busy(job.assets);
  $('status-pill').textContent=phase.status.toUpperCase();
  $('job-country').textContent=`${job.country} · ${job.limit ? `up to ${job.limit} ads` : 'all matching ads'}`;
  $('job-title').textContent=busy(phase) ? titles[phase.phase]||'Working on your export' : ({complete:'Your export is ready',partial:'Ready with some incomplete data',cancelled:'Stopped — saved files below',failed:'Export could not finish',interrupted:'Export interrupted'})[phase.status]||'Ready';
  $('job-message').textContent=phase.message||'';
  const incomplete=!busy(job.scrape) && job.scrape.status!=='complete';
  const warning=incomplete ? 'The scrape did not fully finish. Downloads contain only the data that was saved.' : job.scrape.warnings>0 ? `${job.scrape.warnings} ads have incomplete details. See the workbook for per-ad status.` : '';
  $('scrape-warning').textContent=warning; $('scrape-warning').hidden=!warning;
  if(incomplete && job.assets.status==='complete'){$('job-title').textContent='ZIP ready ? scrape incomplete';$('status-pill').textContent='PARTIAL';}
  $('expiry').textContent=`Files eligible for cleanup after ${new Date(job.expires_at*1000).toLocaleString()}.`;
  $('stop').hidden=!running; $('stop').disabled=phase.status==='stopping';
  $('ad-count').textContent=job.scrape.ads??job.scrape.loaded??'—';
  $('phase-label').textContent=phase.phase==='loading' ? `${phase.loaded||0} ad IDs discovered · total not yet known` : phase.total!=null ? `${phase.completed||0} / ${phase.total} ${job.assets.status!=='idle'?'media URLs':'listing cards'}` : titles[phase.phase]||phase.status;
  if (phase.percent!=null && !['saving','loading','starting'].includes(phase.phase)) {
    $('progress').value=phase.percent; $('percent').textContent=`${phase.percent}%`;
  } else if (busy(phase)) { $('progress').removeAttribute('value'); $('percent').textContent='—'; }
  else { $('progress').value=phase.status==='complete'?100:0; $('percent').textContent=phase.status==='complete'?'100%':'—'; }
  const base=`/api/jobs/${job.id}/download/`;
  $('xlsx').classList.toggle('disabled',!job.xlsx_ready); $('xlsx').setAttribute('aria-disabled',String(!job.xlsx_ready));
  if(job.xlsx_ready) $('xlsx').href=base+'xlsx'; else $('xlsx').removeAttribute('href');
  $('xlsx-note').textContent=job.xlsx_ready ? (job.scrape.status==='complete'?'Download .xlsx':'Download partial .xlsx') : 'Available after scraping';
  $('assets').disabled=!job.xlsx_ready||busy(job.assets);
  $('asset-title').textContent=job.zip_ready?'Download assets ZIP':busy(job.assets)?'Preparing your ZIP…':'Download all assets';
  $('asset-note').textContent=job.assets.total!=null ? `${job.assets.downloaded||0} saved · ${job.assets.failed||0} failed · ${job.assets.missing||0} missing` : 'Images & videos in one ZIP';
  $('zip-link').hidden=!job.zip_ready; $('zip-link').href=base+'zip';
  $('diagnostics').hidden=!job.diagnostics_ready; $('diagnostics').href=base+'diagnostics';
  $('save-log').hidden=false; $('save-log').href=base+'logs';
  const signature=job.id+JSON.stringify(job.logs);
  if (signature!==logSignature) {
    const box=$('logs'), atBottom=box.scrollHeight-box.clientHeight-box.scrollTop<55;
    box.replaceChildren();
    for(const item of job.logs){const row=document.createElement('div');row.className='log-line';const stamp=document.createElement('span');stamp.className='log-time';stamp.textContent=new Date(item.time*1000).toLocaleTimeString();const text=document.createElement('span');text.className='log-text';text.textContent=item.text;row.append(stamp,text);box.append(row);}
    if(atBottom||!logSignature) box.scrollTop=box.scrollHeight;
    logSignature=signature;
  }
  if(job.zip_ready && sessionStorage.getItem('studio-download')===job.id){sessionStorage.removeItem('studio-download');download(base+'zip');}
  tick();
}
async function history() {
  const jobs=await api('/api/jobs');
  $('history').replaceChildren();
  if(!jobs.length){const p=document.createElement('p');p.className='muted';p.textContent='No exports yet.';$('history').append(p);}
  for(const job of jobs){const button=document.createElement('button');button.className='history-item'+(selected===job.id?' selected':'');const name=document.createElement('strong');name.textContent=job.country;const note=document.createElement('small');note.textContent=`${job.scrape.status} · ${new Date(job.created*1000).toLocaleString()}`;button.append(name,note);button.onclick=async()=>{selected=job.id;localStorage.setItem('studio-job',selected);logSignature='';try{render(await api(`/api/jobs/${selected}`));await history();}catch(exc){error(exc.message);}};$('history').append(button);}
  if(!selected && jobs.length){selected=jobs[0].id;localStorage.setItem('studio-job',selected);}
}
async function poll(){
  try{
    if(pollCount++%5===0) await history();
    if(selected){const id=selected;const job=await api(`/api/jobs/${id}`);if(id===selected)render(job);}
  }catch(exc){error(exc.message);}
  finally{setTimeout(poll,1500);}
}
$('export-form').addEventListener('submit',async event=>{
  event.preventDefault();error('');$('start').disabled=true;
  try{const job=await api('/api/jobs',{method:'POST',body:JSON.stringify({url:$('url').value.trim(),limit:Number($('limit').value)})});selected=job.id;localStorage.setItem('studio-job',selected);logSignature='';render(job);await history();}
  catch(exc){error(exc.message);}finally{$('start').disabled=false;}
});
$('assets').addEventListener('click',async()=>{
  if(!current)return;error('');
  if(current.zip_ready){download(`/api/jobs/${current.id}/download/zip`);return;}
  const id=current.id;$('assets').disabled=true;
  try{sessionStorage.setItem('studio-download',id);const job=await api(`/api/jobs/${id}/assets`,{method:'POST',body:'{}'});if(selected===id)render(job);}
  catch(exc){sessionStorage.removeItem('studio-download');error(exc.message);$('assets').disabled=false;}
});
$('stop').addEventListener('click',async()=>{if(!current)return;const id=current.id;$('stop').disabled=true;try{const job=await api(`/api/jobs/${id}/cancel`,{method:'POST',body:'{}'});if(selected===id)render(job);}catch(exc){error(exc.message);$('stop').disabled=false;}});
$('url').addEventListener('input',()=>{
  let country='';try{country=new URL($('url').value.trim()).searchParams.get('country')?.toUpperCase()||'';}catch{}
  let name=country;try{if(country && country!=='ALL')name=new Intl.DisplayNames(['en'],{type:'region'}).of(country);}catch{}
  $('country-label').textContent=country==='ALL'?'All countries · worldwide':country?name:'Country follows your link';
  $('scope-description').textContent=country==='ALL'?'Scrape and ZIP include this worldwide search.':'Use country=ALL for worldwide search results.';
});
setInterval(tick,1000);poll();
