// 統計表を計算してレンダリングする
// cutoffDate: オプションの Date または日付文字列。指定した場合、その日（含む）までの集計を行う
function renderStatTable(cutoffDate){
  const store = readStore();
  if(store.current < 0 || !store.sheets[store.current]){
    document.getElementById('statTableContainer').innerHTML = '';
    return;
  }
  const sheet = store.sheets[store.current];
  const members = sheet.members;
  const daysCount = daysInMonth(sheet.year, sheet.month);

  // 締め日までの勤務時間および出勤/休の集計のみ行う
  const cutoff = cutoffDate ? (cutoffDate instanceof Date ? cutoffDate : new Date(cutoffDate)) : new Date();
  const statYear = cutoff.getFullYear();
  const statMonth = cutoff.getMonth();
  const statDate = cutoff.getDate();
  const statDaysCount = Math.min(statDate, daysCount); // 選択日（含む）までの集計日数

  // 統計表ヘッダ
  // まず各メンバーごとの統計を先に計算して配列に格納する（後でヘッダの表示を決めるため）
  const memberStats = members.map(member=>{
    // 今月の1日から締め日までの出勤/休/未入力の集計
    let restDays = 0, workDays = 0, halfRestDays = 0, notFilledDays = 0, totalWorkHours = 0;
    let workHoursArr = [];
    for(let d=0; d<statDaysCount; d++){
      const day = member.days[d];
      let isHalfRest = false;
      const checkin = day.values && day.values[0];
      const checkout = day.values && day.values[1];
      if(day.type === '休'){ restDays++; continue; }
      if((checkin === '13:00') || (checkout === '12:00')) isHalfRest = true;
      if(isHalfRest) halfRestDays++;
      if(checkin && checkout){
        workDays++;
        let wh = 0;
        const [inH,inM] = checkin.split(':').map(x=>parseInt(x,10));
        const [outH,outM] = checkout.split(':').map(x=>parseInt(x,10));
        if(isHalfRest){
          if(checkout === '12:00') wh = (12*60 - (inH*60+inM))/60;
          else if(checkin === '13:00') wh = (outH*60+outM - (13*60))/60;
        } else {
          wh = (outH*60+outM - inH*60-inM)/60 - 1;
        }
        if(wh < 0) wh = 0;
        workHoursArr.push(wh);
        totalWorkHours += wh;
      } else {
        notFilledDays++;
      }
    }

    // 月全体（整月）の統計
    let restDays_full = 0, halfRestDays_full = 0, notFilledDays_full = 0, totalWorkHours_full = 0, workDays_days = 0;
    for(let d=0; d<daysCount; d++){
      const day = member.days[d];
      if(day.type === '休'){ restDays_full++; continue; }
      const checkin = day.values && day.values[0];
      const checkout = day.values && day.values[1];
      let isHalf = false;
      if((checkin === '13:00') || (checkout === '12:00')) isHalf = true;
      if(isHalf) halfRestDays_full++;
      if(checkin && checkout){
        workDays_days++;
        const [inH,inM] = checkin.split(':').map(x=>parseInt(x,10));
        const [outH,outM] = checkout.split(':').map(x=>parseInt(x,10));
        let wh = 0;
        if(isHalf){
          if(checkout === '12:00') wh = (12*60 - (inH*60+inM))/60;
          else if(checkin === '13:00') wh = (outH*60+outM - (13*60))/60;
        } else {
          wh = (outH*60+outM - inH*60-inM)/60 - 1;
        }
        if(wh < 0) wh = 0;
        totalWorkHours_full += wh;
      } else {
        notFilledDays_full++;
      }
    }
    const workDays_full = daysCount - restDays_full - halfRestDays_full + halfRestDays_full*0.5;
    workDays_days = daysCount - restDays_full;
    const statRestDays_full = restDays_full + halfRestDays_full*0.5;
    const estWorkHours_full = totalWorkHours_full + notFilledDays_full*7.5;
    const statWorkHours = totalWorkHours;

    // 締め日までの残業集計
    let statOT = 0, statOTLaw = 0;
    for(let d=0; d<daysCount; d++){
      const dayDate = new Date(statYear, statMonth, d+1);
      if(dayDate > cutoff) continue;
      const day = member.days[d]; if(day.type === '休') continue;
      const checkin = day.values && day.values[0];
      const checkout = day.values && day.values[1];
      if(checkin && checkout){
        if(checkin === '13:00'){
          const [outH,outM] = checkout.split(':').map(x=>parseInt(x,10));
          let ot = (outH*60+outM - (17*60+30))/60; if(ot < 0) ot = 0;
          let otLaw = (outH*60+outM - (18*60))/60; if(otLaw < 0) otLaw = 0;
          statOT += ot; statOTLaw += otLaw;
        } else if(checkout === '12:00'){
          // no OT
        } else {
          const [inH,inM] = checkin.split(':').map(x=>parseInt(x,10));
          const [outH,outM] = checkout.split(':').map(x=>parseInt(x,10));
          let ot = (outH*60+outM - inH*60-inM)/60 - 1 - 7.5; if(ot < 0) ot = 0;
          let otLaw = (outH*60+outM - inH*60-inM)/60 - 1 - 8; if(otLaw < 0) otLaw = 0;
          statOT += ot; statOTLaw += otLaw;
        }
      }
    }

    // f/g/h
    const diff200 = 200 - estWorkHours_full;
    const remainDays = workDays_days - workDays - (halfRestDays_full - halfRestDays);
    const remainOT = remainDays > 0 ? diff200 / remainDays : 0;

    // 统计日之后到月底的该成员剩余工作日数（用于 140 小时计算的日均分配）
    let remainingWorkDaysAfterStat = 0;
    for(let d = statDaysCount; d < daysCount; d++){
      const day = member.days[d]; if(day && day.type !== '休') remainingWorkDaysAfterStat++;
    }

    return {
      name: member.name,
      workDays_days,
      workDays,
      statWorkHours,
      estWorkHours_full,
      diff200,
      remainOT,
      statRestDays_full,
      statOT,
      statOTLaw,
      remainingWorkDaysAfterStat
    };
  });

  // 各メンバーについて、締め日基準の予想（締め日までの実働 + 締め日以降の残業日数*7.5）を計算
  memberStats.forEach(ms => {
    ms.estCutoff = ms.statWorkHours + (ms.remainingWorkDaysAfterStat || 0) * 7.5;
    ms.diff200_cutoff = 200 - ms.estCutoff;
    ms.remainDays_cutoff = ms.remainingWorkDaysAfterStat || 0;
    ms.perDayCutoff = ms.remainDays_cutoff > 0 ? (ms.diff200_cutoff / ms.remainDays_cutoff) : 0;
  });

  // ヘッダに (140時間) を表示するか（メンバーのうち1人でも予想工時 < 140 の場合に表示）
  // 締め日基準の予想が140時間未満のメンバーがいるか
  const anyBelow140 = memberStats.some(ms => ms.estCutoff < 140 - 1e-6);
  // ヘッダ作成
  let html = '<table class="stat-table" border="1" cellpadding="4" style="border-collapse:collapse;min-width:900px;font-size:13px">';
  const remHeader = anyBelow140
    ? `<th><span style="color:#429af1">200時間</span><span style="color:#f00">(140時間)</span>までの残り業務時間</th>`
    : `<th>200時間までの残り業務時間</th>`;
  const avgHeader = anyBelow140
    ? `<th><span style="color:#429af1">200時間</span><span style="color:#f00">(140時間)</span>になるまで、今日から毎日できる業務外時間平均値</th>`
    : `<th>200時間になるまで、今日から毎日できる業務外時間平均値</th>`;

  html += `<thead><tr><th>名前</th><th>出勤日数(${statDaysCount}日までの出勤日数)</th><th>${statDaysCount}日までの稼働時間</th><th>締め日基準の予想稼働時間</th>${remHeader}${avgHeader}<th>休日数</th><th>業務外時間</th></tr></thead><tbody>`;

  // 各メンバーの行を追加
  memberStats.forEach(ms=>{
    // 締め日基準の表示：HhMm 形式に整形
    const statWorkMins = hoursToMinutes(ms.statWorkHours);
    const estCutoffMins = hoursToMinutes(ms.estCutoff);

    // 残り（200 - 締め日基準予想）および日別で必要な時間を計算して表示
    let remDisplay = `<span style="color:#429af1">${formatMinutesToHM(Math.max(0, hoursToMinutes(ms.diff200)))}</span>`;
    let avgDisplay = `<span style="color:#429af1">${formatMinutesToHM(hoursToMinutes(ms.remainOT))}</span>`;
    if(ms.estCutoff < 140 - 1e-6){
      const shortageH = 140 - ms.estCutoff;
      const shortageM = hoursToMinutes(shortageH);
      remDisplay = `<span style="color:#f00">(${formatMinutesToHM(shortageM)})</span>`;
      const days = ms.remainDays_cutoff || 1;
      const perDayH = shortageH / days;
      avgDisplay = `<span style="color:#f00">(${formatMinutesToHM(hoursToMinutes(perDayH))})</span>`;
    } else {
      // 200 時間到達のための差（締め日基準）
      const diff200_cutoff = ms.diff200_cutoff;
      remDisplay = `<span style="color:#429af1">${formatMinutesToHM(hoursToMinutes(Math.max(0, diff200_cutoff)))}</span>`;
      const perDayH = ms.perDayCutoff;
      avgDisplay = `<span style="color:#429af1">${formatMinutesToHM(hoursToMinutes(Math.max(0, perDayH)))}</span>`;
    }

    // 今月の予想は締め日基準の予想を表示。全月ベースの予想は title に入れる
    const estFullTitle = `全月予想: ${formatMinutesToHM(hoursToMinutes(ms.estWorkHours_full))}`;
    html += `<tr><td>${ms.name}</td><td>${ms.workDays_days}日(${ms.workDays}日)</td><td>${formatMinutesToHM(statWorkMins)}</td><td title="${estFullTitle}">${formatMinutesToHM(estCutoffMins)}</td><td>${remDisplay}</td><td>${avgDisplay}</td><td>${ms.statRestDays_full.toFixed(2)}</td><td>${ms.statOT.toFixed(2)}(${ms.statOTLaw.toFixed(2)})</td></tr>`;
  });

  html += '</tbody></table>';
  document.getElementById('statTableContainer').innerHTML = html;
}
  // 統計ボタンのイベント（日時選択モーダルを表示。デフォルトは今日または当表の最終日）
  const statBtn = document.getElementById('statBtn');
  const statModal = document.getElementById('statModal');
  const statModalDate = document.getElementById('statModalDate');
  const statModalApply = document.getElementById('statModalApply');
  const statModalCancel = document.getElementById('statModalCancel');
  function openStatModal(){
    const st = readStore();
    if(st.current<0) return alert('勤務表を未選択');
    const s = st.sheets[st.current];
    const days = daysInMonth(s.year, s.month);
    const min = `${s.year}-${String(s.month).padStart(2,'0')}-01`;
    const max = `${s.year}-${String(s.month).padStart(2,'0')}-${String(days).padStart(2,'0')}`;
    const today = new Date();
    const lastOfMonth = new Date(s.year, s.month-1, days);
    const defaultDate = today <= lastOfMonth ? today : lastOfMonth;
    statModalDate.min = min; statModalDate.max = max;
    statModalDate.value = `${defaultDate.getFullYear()}-${String(defaultDate.getMonth()+1).padStart(2,'0')}-${String(defaultDate.getDate()).padStart(2,'0')}`;
    statModal.style.display = 'block'; statModal.setAttribute('aria-hidden','false');
    statModalDate.focus();
  }
  function closeStatModal(){ statModal.style.display='none'; statModal.setAttribute('aria-hidden','true'); }

  if(statBtn && statModal && statModalDate && statModalApply && statModalCancel){
    statBtn.addEventListener('click', ()=>{ openStatModal(); });
    statModalApply.addEventListener('click', ()=>{
      const v = statModalDate.value; if(!v) return alert('请选择日期');
      renderStatTable(v);
      closeStatModal();
    });
    statModalCancel.addEventListener('click', ()=>{ closeStatModal(); });
  }
// 完全な JSON データをエクスポートする
function exportAllData(){
  const store = readStore();
  const json = JSON.stringify(store, null, 2);
  const blob = new Blob([json], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'timesheet_data.json';
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
}

// 完全な JSON データをインポートする
function importAllData(file){
  const reader = new FileReader();
  reader.onload = function(e){
    try {
      const data = JSON.parse(e.target.result);
      if(!data || typeof data !== 'object' || !Array.isArray(data.sheets)){
        alert('フォーマットが正しくない');
        return;
      }
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      alert('インポート成功！');
      renderAll();
    } catch(err){
      alert('インポート失敗：' + err.message);
    }
  };
  reader.readAsText(file);
}
  // エクスポート/インポート用ボタンのイベント（オプションで年-月範囲をサポート）
  function _parseMonthKey(v){
  if(!v) return null; // 期待される形式: 'YYYY-MM'
    const parts = v.split('-'); if(parts.length<2) return null;
    const y = parseInt(parts[0],10); const m = parseInt(parts[1],10);
    if(isNaN(y) || isNaN(m)) return null; return y*12 + (m-1);
  }

  function exportDataByMonthRange(fromMonthStr, toMonthStr){
    const store = readStore();
    const fromKey = _parseMonthKey(fromMonthStr);
    const toKey = _parseMonthKey(toMonthStr);
    if(!fromKey && !toKey){
      exportAllData(); return;
    }
    const start = fromKey || toKey; const end = toKey || fromKey;
    const s = Math.min(start, end); const e = Math.max(start, end);
    const filtered = (store.sheets || []).filter(sh => {
      const k = sh.year*12 + (sh.month-1);
      return k >= s && k <= e;
    });
    if(!filtered.length){ alert('指定した期間のデータが見つかりません'); return; }
    const out = { sheets: filtered, current: -1, showAllDays: store.showAllDays };
    const json = JSON.stringify(out, null, 2);
    const blob = new Blob([json], {type:'application/json'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const filenameFrom = fromMonthStr ? fromMonthStr.replace('-','') : 'start';
    const filenameTo = toMonthStr ? toMonthStr.replace('-','') : 'end';
    a.download = `sheet_${filenameFrom}-${filenameTo}.json`;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  }

  function importDataByMonthRange(file, fromMonthStr, toMonthStr){
    const reader = new FileReader();
    reader.onload = function(e){
      try{
        const data = JSON.parse(e.target.result);
        if(!data || typeof data !== 'object' || !Array.isArray(data.sheets)){
          alert('フォーマットが正しくない'); return;
        }
        const fromKey = _parseMonthKey(fromMonthStr);
        const toKey = _parseMonthKey(toMonthStr);
        let filtered = data.sheets;
        if(fromKey || toKey){
          const start = fromKey || toKey; const end = toKey || fromKey;
          const s = Math.min(start, end); const e = Math.max(start, end);
          filtered = (data.sheets || []).filter(sh=>{
            const k = sh.year*12 + (sh.month-1);
            return k >= s && k <= e;
          });
        }
        if(!filtered.length){ alert('インポートファイル内に指定範囲のデータが見つかりませんでした'); return; }
        const store = readStore();
        let added = 0, replaced = 0;
        filtered.forEach(sheet =>{
          const idx = store.sheets.findIndex(s => s.year === sheet.year && s.month === sheet.month);
          if(idx >= 0){ store.sheets[idx] = sheet; replaced++; }
          else { store.sheets.push(sheet); added++; }
        });
        writeStore(store);
        alert(`インポート完了。追加: ${added}、更新: ${replaced}`);
        renderAll();
      }catch(err){ alert('インポート失敗：' + err.message); }
    };
    reader.readAsText(file);
  }

  const exportDataBtn = document.getElementById('exportDataBtn');
  const importDataBtn = document.getElementById('importDataBtn');
  const importDataInput = document.getElementById('importDataInput');
  // 範囲モーダルの要素（エクスポート/インポート用）。ボタン押下でモーダルを開き、適用で処理を実行する
  const rangeModal = document.getElementById('rangeModal');
  const rangeModalFrom = document.getElementById('rangeModalFrom');
  const rangeModalTo = document.getElementById('rangeModalTo');
  const rangeModalApply = document.getElementById('rangeModalApply');
  const rangeModalCancel = document.getElementById('rangeModalCancel');
  let _pendingRangeAction = null; // 'export' or 'import'
  let _pendingImportRange = null; // ファイル選択時にインポートする範囲を一時保存する

  function openRangeModal(action){
  // action: 'export' | 'import'（エクスポートまたはインポート）
    const st = readStore();
    if(st.current<0 && action === 'export') return alert('勤務表を未選択');
    const today = new Date();
    const defaultYM = st.current>=0 ? `${st.sheets[st.current].year}-${String(st.sheets[st.current].month).padStart(2,'0')}` : `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}`;
    rangeModalFrom.value = defaultYM; rangeModalTo.value = defaultYM;
    rangeModal.style.display = 'block'; rangeModal.setAttribute('aria-hidden','false');
    _pendingRangeAction = action;
  }
  function closeRangeModal(){ rangeModal.style.display = 'none'; rangeModal.setAttribute('aria-hidden','true'); _pendingRangeAction = null; }

  if(exportDataBtn && rangeModal && rangeModalApply && rangeModalCancel && rangeModalFrom && rangeModalTo){
    exportDataBtn.addEventListener('click', ()=>{ openRangeModal('export'); });
    rangeModalApply.addEventListener('click', ()=>{
      const fromVal = rangeModalFrom.value || '';
      const toVal = rangeModalTo.value || '';
      if(_pendingRangeAction === 'export'){
        if(fromVal || toVal) exportDataByMonthRange(fromVal, toVal);
        else exportAllData();
      }else if(_pendingRangeAction === 'import'){
  // 選択した範囲を保存してファイル選択ダイアログを開く
        _pendingImportRange = { from: fromVal, to: toVal };
  // ファイル選択ダイアログを開く
        if(importDataInput){ importDataInput.value = ''; importDataInput.click(); }
      }
      closeRangeModal();
    });
    rangeModalCancel.addEventListener('click', ()=>{ closeRangeModal(); });
  }

  if(importDataBtn && importDataInput){
    importDataBtn.addEventListener('click', ()=>{ openRangeModal('import'); });
    importDataInput.addEventListener('change', (e)=>{
      if(e.target.files && e.target.files[0]){
        const file = e.target.files[0];
        const fromVal = _pendingImportRange ? _pendingImportRange.from : '';
        const toVal = _pendingImportRange ? _pendingImportRange.to : '';
  // 使用後に保留範囲をリセットする
        _pendingImportRange = null;
        if(fromVal || toVal) importDataByMonthRange(file, fromVal, toVal);
        else importAllData(file);
      }
    });
  }
// 複数の勤務表を管理する（v1）
const STORAGE_KEY = 'Sheets_v1';

function generateTimeOptions(startMinute, endMinute, stepMin){
  const out = [];
  for(let m = startMinute; m <= endMinute; m += stepMin){
    const hh = Math.floor(m/60);
    const mm = m % 60;
    const s = String(hh).padStart(2,'0') + ':' + String(mm).padStart(2,'0');
    out.push(s);
  }
  return out;
}

function readStore(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if(!raw) return { sheets: [], current: -1, showAllDays: true };
    const store = JSON.parse(raw);
    if(typeof store.showAllDays === 'undefined') store.showAllDays = true;
    let changed = false;
    if(store && Array.isArray(store.sheets)){
      store.sheets.forEach(s=>{
        if(!s.components || s.components < 2){ s.components = 2; changed = true; }
        if(Array.isArray(s.members)){
          s.members.forEach(m=>{
            if(Array.isArray(m.days)){
              m.days.forEach(d=>{
                if(d.type === 'work'){
                  if(!Array.isArray(d.values)){
                    d.values = Array.from({length: s.components}, ()=>''); changed = true;
                  } else if(d.values.length < s.components){
                    while(d.values.length < s.components) d.values.push(''); changed = true;
                  }
                }
              });
            }
          });
        }
      });
    }
    if(changed) writeStore(store);
    return store;
  }catch(e){ console.error('load store', e); return { sheets: [], current: -1, showAllDays: true }; }
}

function writeStore(store){ localStorage.setItem(STORAGE_KEY, JSON.stringify(store)); }

function daysInMonth(y,m){ return new Date(y, m, 0).getDate(); }
function weekdayOf(y,m,d){ return new Date(y, m-1, d).getDay(); }

// ヘルパ: 小数時間（例: 7.5）→分に変換
function hoursToMinutes(h){ return Math.round((h || 0) * 60); }
// ヘルパ: 分を "7h30m" 形式に整形。0分は '0h' を返す
function formatMinutesToHM(mins){ if(mins == null || isNaN(mins)) return '0h'; mins = Math.round(mins); if(mins <= 0) return '0h'; const h = Math.floor(mins/60); const m = mins % 60; if(h>0) return m? (h + 'h' + String(m).padStart(2,'0') + 'm') : (h + 'h'); return (m + 'm'); }

function openNewModal(){
  const modal = document.getElementById('newModal');
  const now = new Date();
  document.getElementById('newYear').value = now.getFullYear();
  document.getElementById('newMonth').value = now.getMonth()+1;
  document.getElementById('newMembers').value = 2;
  modal.setAttribute('aria-hidden','false');
}
function closeNewModal(){ document.getElementById('newModal').setAttribute('aria-hidden','true'); }

function createSheet(){
  const y = parseInt(document.getElementById('newYear').value,10);
  const m = parseInt(document.getElementById('newMonth').value,10);
  const comps = 2;
  const membersCount = Math.max(1, Math.min(200, parseInt(document.getElementById('newMembers').value,10)||5));
  const days = daysInMonth(y,m);
  const members = [];
  for(let i=0;i<membersCount;i++){
    const daysArr = [];
    for(let d=1; d<=days; d++){
      const wk = weekdayOf(y,m,d);
      if(wk===0 || wk===6) daysArr.push({ type: '休', values: null });
      else daysArr.push({ type: 'work', values: Array.from({length:comps}, ()=>'') });
    }
    members.push({ id: Date.now()+i, name: `メンバー ${i+1}`, days: daysArr });
  }
  const sheet = { id: 's_' + Date.now(), year: y, month: m, components: comps, members, createdAt: new Date().toISOString() };
  const store = readStore(); store.sheets.push(sheet); store.current = store.sheets.length - 1; writeStore(store);
  closeNewModal(); renderAll();
}

function renderSheetsList(){
  const list = document.getElementById('sheetsList'); list.innerHTML = '';
  const store = readStore();
  store.sheets.forEach((s, idx)=>{
    const div = document.createElement('div'); div.className = 'sheet-item' + (idx===store.current? ' active':'');
    div.textContent = `${s.year}年 ${s.month}月 (${s.members.length}人)`;
    div.onclick = ()=>{ store.current = idx; writeStore(store); renderAll(); };
    list.appendChild(div);
  });
}

function renderTableForCurrent(){
  const container = document.getElementById('tableContainer'); container.innerHTML = '';
  const store = readStore();
  if(store.current < 0 || !store.sheets[store.current]){ document.getElementById('sheetTitle').textContent = '（勤務表を未選択）'; return; }
  const sheet = store.sheets[store.current];
  const sheetTitleEl = document.getElementById('sheetTitle');
  // ヘルパー: 日付文字列 'HH:MM' を分に変換
  function parseTimeToMinutes(t){ if(!t) return null; const parts = String(t).split(':'); if(parts.length<2) return null; const h = parseInt(parts[0],10); const m = parseInt(parts[1],10); if(isNaN(h)||isNaN(m)) return null; return h*60 + m; }
  // ヘルパー: day オブジェクトからその日の実働分数を計算する（概算）
  function computeMinutesFromDay(dayObj){
    if(!dayObj || dayObj.type === '休') return 0;
    const vals = Array.isArray(dayObj.values) ? dayObj.values : [];
    const checkin = vals[0] || '';
    const checkout = vals[1] || '';
    // 半休判定
    if(checkin === '13:00' || checkout === '12:00'){
      if(checkout === '12:00' && checkin){ const inMin = parseTimeToMinutes(checkin); if(inMin==null) return 0; return Math.max(0, 12*60 - inMin); }
      if(checkin === '13:00' && checkout){ const outMin = parseTimeToMinutes(checkout); if(outMin==null) return 0; return Math.max(0, outMin - 13*60); }
    }
    // 通常の出退勤
    if(checkin && checkout){ const inMin = parseTimeToMinutes(checkin); const outMin = parseTimeToMinutes(checkout); if(inMin==null||outMin==null) return 0; const mins = outMin - inMin - 60; return mins>0? mins:0; }
    // 複数コンポーネントの合算（ペアで計算）
    let total = 0;
    for(let ci=0; ci+1<vals.length; ci+=2){ const a = vals[ci], b = vals[ci+1]; if(a && b){ const am = parseTimeToMinutes(a), bm = parseTimeToMinutes(b); if(am!=null && bm!=null){ const wh = bm - am; if(!isNaN(wh)) total += Math.max(0, wh); } } }
    return total;
  }
  // ヘルパー: 分を「7h30m」形式に整形（0 は空文字列として返す）
  function formatMinutesToHM(mins){ if(!mins || mins <= 0) return ''; const h = Math.floor(mins/60); const m = mins % 60; if(h>0){ return m? (h + 'h' + String(m).padStart(2,'0') + 'm') : (h + 'h'); } return (m + 'm'); }
  // sheetTitleEl.textContent = `${sheet.year}年 ${sheet.month}月 — ${sheet.components} 項目/日`;
  sheetTitleEl.textContent = `${sheet.year}年 ${sheet.month}月`;
  // トグルボタン
  if(!document.getElementById('toggleRestBtn')){
    const btn = document.createElement('button'); btn.id = 'toggleRestBtn'; btn.style.marginBottom = '10px';
    btn.addEventListener('click', ()=>{ const st = readStore(); st.showAllDays = !st.showAllDays; writeStore(st); renderAll(); });
    sheetTitleEl.parentNode.insertBefore(btn, sheetTitleEl.nextSibling);
  }
  document.getElementById('toggleRestBtn').textContent = store.showAllDays ? '休日を隠す' : '休日を表示';

  // メンバー追加ボタン（未作成なら作成）
  if(!document.getElementById('addMemberBtn')){
    const addBtn = document.createElement('button'); addBtn.id = 'addMemberBtn'; addBtn.style.marginLeft = '8px';
    addBtn.textContent = 'メンバー追加';
    addBtn.addEventListener('click', ()=>{
      const st = readStore(); if(st.current<0) return alert('勤務表を未選択');
      const s = st.sheets[st.current];
      const days = daysInMonth(s.year, s.month);
      const comps = s.components || 2;
      const daysArr = [];
      for(let d=1; d<=days; d++){
        const wk = weekdayOf(s.year, s.month, d);
        if(wk===0 || wk===6) daysArr.push({ type: '休', values: null });
        else daysArr.push({ type: 'work', values: Array.from({length: comps}, ()=>'') });
      }
      const newMember = { id: Date.now(), name: `メンバー ${s.members.length+1}`, days: daysArr };
      s.members.push(newMember);
      writeStore(st); renderAll();
    });
    sheetTitleEl.parentNode.insertBefore(addBtn, document.getElementById('toggleRestBtn').nextSibling);
  }

  const daysCount = daysInMonth(sheet.year, sheet.month);
  const visibleDays = [];
  // 今日の列をハイライトするための現在日付
  const now = new Date(); const nowY = now.getFullYear(); const nowM = now.getMonth()+1; const nowD = now.getDate();
  for(let d=1; d<=daysCount; d++){
    const idx = d-1;
    const allRest = sheet.members.every(m => m.days[idx] && m.days[idx].type === '休');
    if(!store.showAllDays && allRest) continue;
    visibleDays.push(d);
  }

  const tbl = document.createElement('table');
  const thead = document.createElement('thead'); const hr = document.createElement('tr');
  const thName = document.createElement('th'); thName.className='member-name'; thName.textContent = '名前'; hr.appendChild(thName);
  const wnames = ['日','月','火','水','木','金','土'];
  visibleDays.forEach((d, visIndex)=>{
    const th = document.createElement('th'); const wk = weekdayOf(sheet.year, sheet.month, d);
  const dayLabel = store.showAllDays ? `${d}` : `${visIndex+1}`;
  // 表示：序号 / 曜日 / 月日
  const monthDay = `${String(sheet.month).padStart(2,'0')}.${String(d).padStart(2,'0')}`;
  th.innerHTML = `<div>${dayLabel}</div><div class="weekday" style="color:#0538ff">${wnames[wk]}</div><div class="month-day" style="font-size:0.8em;font-weight:bold;color:#f00">${monthDay}</div>`;
  if(wk===0||wk===6) th.className = 'weekend';
  // このシートが今日の年月と一致する場合に当日の列をハイライト
  if(sheet.year === nowY && sheet.month === nowM && d === nowD){ th.classList.add('col-today'); }
    th.style.cursor = 'pointer';
    th.addEventListener('click', ()=>{
      const st = readStore(); if(st.current < 0) return; const s = st.sheets[st.current]; const idx = d-1;
      const allRest = s.members.every(m=>m.days[idx] && m.days[idx].type === '休');
      let msg = allRest
        ? `${sheet.month}月${d}日“出勤日”にしますか？`
        : `${sheet.month}月${d}日“休日”にしますか？`;
      if(!confirm(msg)) return;
      if(allRest){ s.members.forEach(m=>{ m.days[idx].type = 'work'; m.days[idx].values = Array.from({length: s.components}, ()=>''); }); }
      else { s.members.forEach(m=>{ m.days[idx].type = '休'; m.days[idx].values = null; }); }
      writeStore(st); renderAll();
    });
    hr.appendChild(th);
  });
  // 操作列ヘッダ
  const thOp = document.createElement('th'); thOp.textContent = '操作'; thOp.className = 'col-ops';
  hr.appendChild(thOp);
  thead.appendChild(hr); tbl.appendChild(thead);

  const tbody = document.createElement('tbody');
  sheet.members.forEach((member, mi)=>{
    const tr = document.createElement('tr');
    const nameTd = document.createElement('td'); nameTd.className='member-name';
  const nameInput = document.createElement('input');
  nameInput.value = member.name;
  nameInput.style.maxWidth = '90px';
  nameInput.style.width = '90px';
  nameInput.style.boxSizing = 'border-box';
  nameInput.style.overflow = 'hidden';
  nameInput.style.textOverflow = 'ellipsis';
  nameInput.style.verticalAlign = 'middle';
    nameInput.addEventListener('change', (e)=>{ const store2 = readStore(); store2.sheets[store2.current].members[mi].name = e.target.value; writeStore(store2); renderSheetsList(); });
      // 姓名とA/B/Cのドロップダウンを並べて表示し、両方とも可視化
      const nameWrap = document.createElement('span');
      nameWrap.style.display = 'inline-block';
      nameWrap.style.verticalAlign = 'middle';
      nameWrap.appendChild(nameInput);

  // nameWrap は現在名前入力のみを含む。操作列には後でタイプ選択と削除ボタンを配置します
      nameTd.appendChild(nameWrap);
    tr.appendChild(nameTd);

    visibleDays.forEach(d=>{
      const di = d-1; const dayObj = member.days[di]; const td = document.createElement('td');
      const isTodayCol = (sheet.year === nowY && sheet.month === nowM && d === nowD);
  // 時間マーカーに基づいて、この日が半休（午前/午後半休）か判定
      const checkinVal = dayObj.values && dayObj.values[0];
      const checkoutVal = dayObj.values && dayObj.values[1];
      const isHalf = (checkinVal === '13:00') || (checkoutVal === '12:00');
      if(dayObj.type === '休'){
        td.className = 'cell-weekend'; td.textContent = '休';
        if(isTodayCol) td.classList.add('col-today');
        td.addEventListener('click', ()=>{ const st = readStore(); if(st.current<0) return; const s = st.sheets[st.current]; const target = s.members[mi].days[di]; target.type = 'work'; target.values = Array.from({length: s.components}, ()=>''); writeStore(st); renderAll(); });
      } else {
        if(isHalf) td.className = 'cell-halfrest';
        if(isTodayCol) td.classList.add('col-today');
        const wrapper = document.createElement('div'); wrapper.className = 'day-inputs';
        for(let c=0;c<sheet.components;c++){
          const sel = document.createElement('select'); const placeholder = document.createElement('option'); placeholder.value=''; placeholder.textContent='—'; sel.appendChild(placeholder);
          if(c===0){ const restOpt = document.createElement('option'); restOpt.value='休'; restOpt.textContent='休'; sel.appendChild(restOpt); const opts = generateTimeOptions(6*60,13*60,60); opts.forEach(t=>{ const o=document.createElement('option'); o.value=t; o.textContent=t; sel.appendChild(o); }); }
          else { const opts = generateTimeOptions(12*60,24*60,15); opts.forEach(t=>{ const o=document.createElement('option'); o.value=t; o.textContent=t; sel.appendChild(o); }); }
          sel.value = dayObj.values && dayObj.values[c] ? dayObj.values[c] : '';
          sel.addEventListener('change', (e)=>{
            const st = readStore(); if(st.current<0) return; const val = e.target.value;
            if(c===0 && val === '休'){ st.sheets[st.current].members[mi].days[di].type = '休'; st.sheets[st.current].members[mi].days[di].values = null; writeStore(st); renderAll(); return; }
            const targetDay = st.sheets[st.current].members[mi].days[di]; if(targetDay.type === '休'){ targetDay.type='work'; targetDay.values = Array.from({length: st.sheets[st.current].components}, ()=>''); }
            st.sheets[st.current].members[mi].days[di].values[c] = val; writeStore(st); // 保存後に即座に再描画して日別実働時間を更新
            renderAll();
          });
          wrapper.appendChild(sel);
        }
        td.appendChild(wrapper);
        td.addEventListener('click', (e)=>{ if(e.target && e.target.tagName){ const t = e.target.tagName.toUpperCase(); if(t === 'INPUT' || t === 'SELECT' || t === 'OPTION') return; } const st = readStore(); if(st.current<0) return; const s = st.sheets[st.current]; const target = s.members[mi].days[di]; target.type='休'; target.values=null; writeStore(st); renderAll(); });
      }
      tr.appendChild(td);
    });
    // 操作列（最後の列）: 包含 タイプ ドロップダウン と 削除 ボタン
    const opsTd = document.createElement('td'); opsTd.className = 'cell-ops';
  
  // プリセット選択（タイプ）
    const presetSel2 = document.createElement('select');
    const ph2 = document.createElement('option'); ph2.value=''; ph2.textContent='タイプ'; presetSel2.appendChild(ph2);
    ['A','B','C'].forEach(k=>{ const o = document.createElement('option'); o.value = k; o.textContent = k; presetSel2.appendChild(o); });
    presetSel2.style.marginRight = '6px';
    presetSel2.addEventListener('mousedown', e=>{ e.stopPropagation(); });
    presetSel2.addEventListener('click', e=>{ e.stopPropagation(); });
    presetSel2.addEventListener('change', (e)=>{
      const v = e.target.value; if(!v) return;
      const st = readStore(); if(st.current < 0) return;
      const s = st.sheets[st.current];
  // このメンバーに既に入力済みデータがあるかを検出
      const memberDays = s.members[mi].days || [];
      let hasFilled = false;
      for(let di=0; di<memberDays.length; di++){
        const dd = memberDays[di];
        if(dd && dd.type === 'work' && Array.isArray(dd.values)){
          for(let vi=0; vi<dd.values.length; vi++){
            const vv = dd.values[vi]; if(vv && String(vv).trim() !== ''){ hasFilled = true; break; }
          }
        }
        if(hasFilled) break;
      }
  // データがある場合は事前に確認を行う
      if(hasFilled){
        const ok = confirm('このメンバーには既に入力があります。タイプを変更すると既存の時刻が上書きされます。続行しますか？');
        if(!ok){ e.target.value = ''; return; }
      }
      const map = { A: '08:00', B: '09:00', C: '10:00' };
      const mapEnd = { A: '16:30', B: '17:30', C: '18:30' };
      const setTime = map[v];
      const setTimeEnd = mapEnd[v];
      const comps = s.components || 2;
      s.members[mi].days.forEach(d=>{
        if(d.type === 'work'){
          if(!Array.isArray(d.values)) d.values = Array.from({length: comps}, ()=>'');
          while(d.values.length < comps) d.values.push('');
          d.values[0] = setTime;
          d.values[1] = setTimeEnd;
        }
      });
      writeStore(st);
      renderAll();
    });
    opsTd.appendChild(presetSel2);
    const delBtn2 = document.createElement('button');
    delBtn2.id = 'delMemberBtn'; delBtn2.type = 'button'; delBtn2.textContent = '削除'; delBtn2.style.marginLeft = '6px'; delBtn2.style.fontSize = '0.85em';
    delBtn2.addEventListener('click', (ev)=>{ ev.stopPropagation(); const st = readStore(); if(st.current<0) return; const s = st.sheets[st.current]; if(s.members.length <= 1){ alert('最低1名必要です'); return; } if(!confirm('このメンバーを削除しますか？')) return; s.members.splice(mi,1); writeStore(st); renderAll(); });
    opsTd.appendChild(delBtn2);
    tr.appendChild(opsTd);
    tbody.appendChild(tr);
    // --- 追加行: 各日の実働時間を表示する行を作成 ---
    const hoursTr = document.createElement('tr');
    hoursTr.className = 'member-daily-hours-row';
    const hoursNameTd = document.createElement('td');
    hoursNameTd.className = 'member-name';
    hoursNameTd.textContent = '日毎(時間)';
    hoursTr.appendChild(hoursNameTd);

    // 各日について集計して小セルを作る
    visibleDays.forEach(d=>{
      const di = d-1; const dayObj = member.days[di]; const td = document.createElement('td');
      if(!dayObj || dayObj.type === '休'){
        td.textContent = '0'; 
        td.className = 'cell-weekend';
      } else {
        // 分単位で計算し、HHhMMm 形式で表示
        const mins = computeMinutesFromDay(dayObj);
        td.textContent = formatMinutesToHM(Math.round(mins));
        td.className = 'member-daily-hours';
      }
      hoursTr.appendChild(td);
    });
    
    // メンバーの月合計時間を計算して表示（分 -> HhMm 表示）
    const hoursOpTd = document.createElement('td'); 
    hoursOpTd.className = 'cell-ops'; 
    let memberTotalMins = 0;
    for(let dd = 0; dd < sheet.members[mi].days.length; dd++){
      const dm = computeMinutesFromDay(sheet.members[mi].days[dd]);
      memberTotalMins += dm;
    }
    const totalSpan = document.createElement('span'); totalSpan.className = 'member-total-hours';
    totalSpan.style.marginRight = '8px'; totalSpan.style.fontWeight = '600';
    totalSpan.textContent = '合計: ' + (formatMinutesToHM(memberTotalMins) || '0h');
    hoursOpTd.appendChild(totalSpan);
    hoursTr.appendChild(hoursOpTd);

    tbody.appendChild(hoursTr);
  });
  tbl.appendChild(tbody); container.appendChild(tbl);
}

function goPrev(){ const store = readStore(); if(store.sheets.length===0) return; store.current = Math.max(0, (store.current||0) - 1); writeStore(store); renderAll(); }
function goNext(){ const store = readStore(); if(store.sheets.length===0) return; store.current = Math.min(store.sheets.length-1, (store.current||0) + 1); writeStore(store); renderAll(); }

function deleteCurrentSheet(){ const store = readStore(); if(store.current < 0 || !store.sheets[store.current]) return alert('勤務表を未選択'); const s = store.sheets[store.current]; const ok = confirm(`${s.year}年 ${s.month}月の勤務表を削除しますか？元に戻らない。`); if(!ok) return; store.sheets.splice(store.current, 1); if(store.sheets.length===0) store.current = -1; else store.current = Math.min(store.current, store.sheets.length-1); writeStore(store); renderAll(); }

function deleteAllSheets(){
  const store = readStore();
  if(!store.sheets || store.sheets.length === 0) return alert('削除する勤務表がありません');
  const ok = confirm('すべての勤務表を削除しますか？元に戻せません。');
  if(!ok) return;
  // ストレージをクリア
  writeStore({ sheets: [], current: -1, showAllDays: true });
  renderAll();
}

function exportCSVCurrent(){ const store = readStore(); if(store.current<0) return alert('勤務表を未選択'); const s = store.sheets[store.current]; const headers = ['名前']; for(let d=1; d<=daysInMonth(s.year,s.month); d++) headers.push(`${d}`); const rows = s.members.map(m=>{ const arr = [m.name]; m.days.forEach(dobj=>{ if(dobj.type==='休') arr.push('休'); else arr.push(dobj.values.map(v=>v||'').join(';')); }); return arr; }); const csv = [headers, ...rows].map(r=>r.map(cell=>`"${String(cell).replace(/"/g,'""')}"`).join(',')).join('\n'); const blob = new Blob([csv],{type:'text/csv;charset=utf-8;'}); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `sheet_${s.year}-${String(s.month).padStart(2,'0')}.csv`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url); }

function renderAll(){ renderSheetsList(); renderTableForCurrent(); }

function init(){
  document.getElementById('newBtn').addEventListener('click', openNewModal);
  document.getElementById('cancelCreate').addEventListener('click', closeNewModal);
  document.getElementById('createSheet').addEventListener('click', createSheet);
  document.getElementById('prevSheet').addEventListener('click', goPrev);
  document.getElementById('nextSheet').addEventListener('click', goNext);
  document.getElementById('exportBtn').addEventListener('click', exportCSVCurrent);
  const delBtn = document.getElementById('deleteBtn'); if(delBtn) delBtn.addEventListener('click', deleteCurrentSheet);
  const delAllBtn = document.getElementById('deleteAllBtn'); if(delAllBtn) delAllBtn.addEventListener('click', deleteAllSheets);
  const dateInput = document.getElementById('dateFilter'); const today = new Date();
  // 月選択を YYYY-MM に初期化
  dateInput.value = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}`;
  // シートが存在しない月を選択した場合に戻せるよう、前の値を保持
  let _prevDateFilterVal = dateInput.value;
  // ユーザーが月を選択したら該当する勤務表に移動。無ければ警告して元に戻す
  dateInput.addEventListener('change', ()=>{
    const v = dateInput.value; if(!v) return;
    const parts = v.split('-'); const yy = parts[0]; const mm = parts[1];
    const store = readStore();
    const idx = store.sheets.findIndex(s=>String(s.year)===yy && String(s.month).padStart(2,'0')===mm);
    if(idx>=0){ store.current = idx; writeStore(store); renderAll(); _prevDateFilterVal = v; }
    else{
      alert('該当する勤務表が見つかりません: ' + v);
  // 入力を前の値に戻す
      dateInput.value = _prevDateFilterVal;
    }
  });
  renderAll();
}

if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
