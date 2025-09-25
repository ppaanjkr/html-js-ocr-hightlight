(() => {
  'use strict';

  /* ===================== GLOBAL ERROR GUARD ===================== */
  function unlockUI(msg){
    try{
      document.getElementById('runAll')?.removeAttribute('disabled');
      document.getElementById('downloadAll')?.removeAttribute('disabled');
      document.getElementById('clearAll')?.removeAttribute('disabled');
      document.getElementById('progressWrap')?.classList.add('hidden');
      if (msg) alert(msg);
    }catch{}
  }
  window.addEventListener('error', e => { console.error(e.error||e); unlockUI('เกิดข้อผิดพลาด (ดู Console)'); });
  window.addEventListener('unhandledrejection', e => { console.error(e.reason||e); unlockUI('เกิดข้อผิดพลาด (ดู Console)'); });

  window.addEventListener('DOMContentLoaded', () => {

    /* ===================== CONFIG ===================== */
    const DEBUG = false;

    // บีบอัด/ย่อภาพก่อน OCR เสมอ
    const MAX_SIDE = 1600;
    function autoJpegQuality(w, h){
      const mp = (w*h)/1e6;
      if (mp <= 1) return 0.92;
      if (mp <= 2) return 0.88;
      if (mp <= 3.5) return 0.85;
      if (mp <= 6) return 0.82;
      return 0.80;
    }

    // OCR: ลองหลายภาษาและหลาย PSM
    const LANGS = ['eng+tha', 'tha', 'eng'];
    const PSM_CANDIDATES = [6, 4, 3, 11]; // 6=single block, 4=single column, 3=auto, 11=sparse
    const OCR_COMMON = { preserve_interword_spaces: '1', user_defined_dpi: '300' };

    // หา “วันที่” เผื่อหน้าต่างก่อน/หลัง
    const LOOK_BEHIND = 1, LOOK_AHEAD = 2;

    /* ===================== DOM ===================== */
    const fileInput     = document.getElementById('file');
    const targetInput   = document.getElementById('target');
    const runAllBtn     = document.getElementById('runAll');
    const downloadAllBtn= document.getElementById('downloadAll');
    const clearAllBtn   = document.getElementById('clearAll');
    const listEl        = document.getElementById('list');
    const summaryWrap   = document.getElementById('summaryWrap');
    const globalCountEl = document.getElementById('globalCount');
    const fileDoneEl    = document.getElementById('fileDone');
    const pWrap         = document.getElementById('progressWrap');
    const pbar          = document.getElementById('pbar');
    const ptext         = document.getElementById('ptext');
    const startInput    = document.getElementById('start');
    const endInput      = document.getElementById('end');
    const startDisp     = document.getElementById('startDisplay');
    const endDisp       = document.getElementById('endDisplay');

    if (!fileInput || !runAllBtn || !downloadAllBtn || !clearAllBtn || !startInput || !endInput) {
      console.error('ไม่พบ element id บางตัวบนหน้า HTML');
      return;
    }

    /* ===== Default date: เดือนปัจจุบัน → start=1, end=วันนี้ (local) ===== */
    const pad2 = n => String(n).padStart(2,'0');
    const fmtLocalYYYYMMDD = d => `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
    const fmtTH_DDMMYYYY = d => `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()+543}`;

    (function setDefaultCurrentMonth(){
      const now = new Date();
      const start = new Date(now.getFullYear(), now.getMonth(), 1);
      startInput.value = fmtLocalYYYYMMDD(start);
      endInput.value   = fmtLocalYYYYMMDD(now);
      startDisp.textContent = fmtTH_DDMMYYYY(start);
      endDisp.textContent   = fmtTH_DDMMYYYY(now);
    })();

    startInput.addEventListener('change', ()=>{
      const d = new Date(startInput.value);
      if (!isNaN(d)) startDisp.textContent = fmtTH_DDMMYYYY(d);
    });
    endInput.addEventListener('change', ()=>{
      const d = new Date(endInput.value);
      if (!isNaN(d)) endDisp.textContent = fmtTH_DDMMYYYY(d);
    });

    /* ===================== Helpers (clean & date) ===================== */
    const thDigitMap = {'๐':'0','๑':'1','๒':'2','๓':'3','๔':'4','๕':'5','๖':'6','๗':'7','๘':'8','๙':'9'};
    const thNumToArabic = s => (s||'').replace(/[\u0E50-\u0E59]/g, d => thDigitMap[d] ?? d);
    const isThai = ch => /[\u0E00-\u0E7F]/.test(ch);

    // ใช้ทำความสะอาด "คำค้น" (ไม่สนพิมพ์เล็ก/ใหญ่)
    function cleanForSearch(s){
      return (s||"")
        .replace(/[^a-z0-9ก-๙]+/gi, " ")                        // สัญลักษณ์ → ช่องว่าง
        .replace(/([\u0E00-\u0E7F])\s+([\u0E00-\u0E7F])/g,"$1$2")// ตัดช่องว่างไทย-ไทย
        .replace(/\s{2,}/g, " ")
        .toLowerCase()
        .trim();
    }

    function parseThaiDate(text){
      const t = thNumToArabic((text||''))
                  .replace(/\bBE\b/i,'')
                  .replace(/[,\u00B7]/g,' ')
                  .replace(/\s+/g,' ')
                  .trim();
      const m = t.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:.*?(\d{1,2}):(\d{2}))?/);
      if(!m) return null;
      let d=+m[1], mo=+m[2], y=+m[3], hh=m[4]?+m[4]:0, mm=m[5]?+m[5]:0;
      if (y>=2400 && y<=2600) y-=543;
      const dt = new Date(y, mo-1, d, hh, mm, 0);
      return isNaN(dt) ? null : dt;
    }
    // ดึง “วันที่ทั้งหมด” จากบรรทัด
    function extractDates(lineText){
      const t = thNumToArabic(lineText||'');
      const re = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[^\d]{0,8}(\d{1,2}):(\d{2}))?/g;
      const out = []; let m;
      while ((m = re.exec(t)) !== null) {
        let d=parseInt(m[1],10), mo=parseInt(m[2],10)-1, y=parseInt(m[3],10);
        const hh=m[4]?parseInt(m[4],10):0, mm=m[5]?parseInt(m[5],10):0;
        if (y>=2400 && y<=2600) y-=543;
        const dt = new Date(y, mo, d, hh, mm, 0);
        if (!isNaN(dt.getTime())) out.push(dt);
      }
      return out;
    }
    function getSelectedRangeCE(){
      const s=startInput.value, e=endInput.value;
      if(!s||!e) return {start:null,end:null};
      let [sy,sm,sd]=s.split('-').map(Number);
      let [ey,em,ed]=e.split('-').map(Number);
      let start=new Date(sy,sm-1,sd,0,0,0,0);
      let end  =new Date(ey,em-1,ed,23,59,59,999);
      if (start>end) [start,end]=[end,start];
      return {start,end};
    }
    const inRange = (dt,s,e)=> dt && s && e && dt>=s && dt<=e;

    /* ===================== Geometry / BBox ===================== */
    const vOverlap = (a,b)=> Math.max(0, Math.min(a.y1,b.y1) - Math.max(a.y0,b.y0));
    const vOverlapRatio = (a,b)=> {
      const ov=vOverlap(a,b); const denom=Math.min(a.y1-a.y0, b.y1-b.y0);
      return denom<=0?0:ov/denom;
    };
    const unionBox = (a,b)=> ({ x:a.x0, y:Math.min(a.y0,b.y0), w:b.x1-a.x0, h:Math.max(a.y1,b.y1)-Math.min(a.y0,b.y0) });

    /* ===================== Line building (จาก words เสมอ) ===================== */
    function buildLinesByClustering(words){
      if (!Array.isArray(words)) return [];
      const ws = words.filter(w => (w.text||'').trim() && w.bbox)
                      .sort((a,b)=>a.bbox.y0 - b.bbox.y0);
      const lines = [];
      const THRESH = 0.25; // ลดเกณฑ์รวมบรรทัด

      for (const w of ws){
        if (!lines.length){
          lines.push({ y0:w.bbox.y0, y1:w.bbox.y1, words:[w] });
          continue;
        }
        const cur = lines[lines.length-1];
        const r = vOverlapRatio(w.bbox, { y0:cur.y0, y1:cur.y1 });
        if (r >= THRESH){
          cur.words.push(w);
          cur.y0 = Math.min(cur.y0, w.bbox.y0);
          cur.y1 = Math.max(cur.y1, w.bbox.y1);
        } else {
          lines.push({ y0:w.bbox.y0, y1:w.bbox.y1, words:[w] });
        }
      }
      for (const L of lines) L.words.sort((a,b)=>a.bbox.x0 - b.bbox.x0);
      return lines.map(L => L.words);
    }
    function buildLines(ocrData){
      const words = (ocrData.words||[]).map(w => ({ text:w.text, bbox:w.bbox })).filter(w => w.text && w.bbox);
      const lines = buildLinesByClustering(words);
      if (DEBUG) {
        console.log('[buildLines] words=', words.length, 'lines=', lines.length);
        const dump = lines.map((L,i)=> `${i}: "${L.map(w=>w.text).join(' ')}"`).join('\n');
        console.log('LINES (clustered from words):\n' + dump);
      }
      return lines;
    }

    /* ===================== Index / Search mapping ===================== */
    function buildLineIndex(wordsInLine){
      let raw=""; const rawIdxToWord=[];
      for(let wi=0; wi<wordsInLine.length; wi++){
        const w=wordsInLine[wi];
        if (wi>0){ raw+=" "; rawIdxToWord.push(null); }
        const t=w.text||"";
        for(let k=0;k<t.length;k++){ raw+=t[k]; rawIdxToWord.push(wi); }
      }
      let clean="", cleanNoSp=""; const mapC2R=[], mapC2RNoSp=[]; let lastSpace=false;
      for(let i=0;i<raw.length;i++){
        const ch=raw[i], prev=raw[i-1]||"", next=raw[i+1]||"";
        if (ch===" " && isThai(prev) && isThai(next)) continue; // ตัดช่องว่างไทย-ไทย
        const lower=ch.toLowerCase(); const isWord=/[a-z0-9ก-๙]/i.test(ch);
        if (!isWord || lower===" "){
          if(!lastSpace){ clean+=" "; mapC2R.push(i); lastSpace=true; }
          continue;
        }
        clean+=lower; mapC2R.push(i); lastSpace=false;
        cleanNoSp+=lower; mapC2RNoSp.push(i);
      }
      return { raw, clean, mapC2R, cleanNoSp, mapC2RNoSp, rawIdxToWord, words:wordsInLine };
    }
    function searchInIndexedLine(lineIdx, targetPhrase){
      const qClean = cleanForSearch(targetPhrase);
      if (!qClean) return null;
      let at = lineIdx.clean.indexOf(qClean);
      let qLen = qClean.length;
      let useNoSp = false;
      if (at < 0){
        const qNoSp = qClean.replace(/\s+/g,'');
        at = lineIdx.cleanNoSp.indexOf(qNoSp);
        if (at >= 0){ useNoSp=true; qLen=qNoSp.length; }
      }
      if (at < 0) return null;

      const map = useNoSp ? lineIdx.mapC2RNoSp : lineIdx.mapC2R;
      const rawStart = map[at];
      const rawEnd   = map[at + qLen - 1];

      function nearestWordIndex(rawI, dir){
        let i=rawI;
        while(i>=0 && i<lineIdx.rawIdxToWord.length){
          const w=lineIdx.rawIdxToWord[i];
          if(w!==null && w!==undefined) return w;
          i+=dir;
        }
        return null;
      }
      const s = nearestWordIndex(rawStart,-1) ?? nearestWordIndex(rawStart,+1);
      const e = nearestWordIndex(rawEnd,+1)   ?? nearestWordIndex(rawEnd,-1);
      if (s==null || e==null) return null;
      return { startWord:Math.min(s,e), endWord:Math.max(s,e) };
    }

    /* ===================== Image preprocess ===================== */
    async function prepareImageForOCR(file){
      let bmp, w, h;
      if (window.createImageBitmap){
        bmp = await createImageBitmap(file, { imageOrientation:'from-image' });
        w=bmp.width; h=bmp.height;
      }else{
        const url = URL.createObjectURL(file);
        const img = await new Promise((res,rej)=>{ const i=new Image(); i.onload=()=>res(i); i.onerror=rej; i.src=url;});
        URL.revokeObjectURL(url);
        w = img.naturalWidth||img.width; h = img.naturalHeight||img.height; bmp = img;
      }
      const longSide = Math.max(w,h);
      const scale = longSide>MAX_SIDE ? (MAX_SIDE/longSide) : 1;
      const outW = Math.max(1, Math.round(w*scale));
      const outH = Math.max(1, Math.round(h*scale));

      const c=document.createElement('canvas');
      c.width=outW; c.height=outH;
      const cx=c.getContext('2d');
      cx.imageSmoothingEnabled=true;
      cx.imageSmoothingQuality='high';
      cx.drawImage(bmp,0,0,outW,outH);

      const q = autoJpegQuality(outW,outH);
      const blob = await new Promise(res=>c.toBlob(res,'image/jpeg',q));
      const url  = URL.createObjectURL(blob);
      return { blob, url, width: outW, height: outH };
    }

    /* ===================== OCR helper (ลองหลาย PSM) ===================== */
    async function recognizeWithFallback(imageOrBlob, logger){
      let lastErr;
      for (const lang of LANGS){
        for (const psm of PSM_CANDIDATES){
          try{
            if (DEBUG) console.log(`[OCR] try lang=${lang} psm=${psm}`);
            const { data } = await Tesseract.recognize(
              imageOrBlob,
              lang,
              { logger, ...OCR_COMMON, tessedit_pageseg_mode: String(psm) }
            );
            if (DEBUG) console.log(`[OCR] success lang=${lang} psm=${psm}, textLen=${(data.text||'').length}`);
            return { data };
          }catch(e){
            lastErr = e;
            if (DEBUG) console.warn(`[OCR] fail lang=${lang} psm=${psm}`, e);
          }
        }
      }
      throw lastErr || new Error('OCR failed');
    }

    /* ===================== UI helpers ===================== */
    const items=[]; // {file, card, canvas, ctx, processed, matchCount, imgW, imgH}

    function createCard(it){
      const card=document.createElement('div');
      card.className='bg-white rounded-2xl shadow overflow-hidden';
      card.innerHTML=`
        <div class="p-3 border-b">
          <div class="text-sm font-medium truncate">${it.file.name}</div>
          <div class="text-xs text-gray-500">${Math.round(it.file.size/1024)} KB</div>
        </div>
        <div class="p-3">
          <canvas class="canvas w-full"></canvas>
          <div class="mt-3">
            <div class="flex justify-between text-xs text-gray-600 mb-1">
              <span>ความคืบหน้าไฟล์นี้</span><span class="progress-text">0%</span>
            </div>
            <div class="w-full h-2 bg-gray-200 rounded"><div class="progress-bar h-2 bg-pink-500 rounded" style="width:0%"></div></div>
          </div>
          <div class="mt-3 text-sm">พบในช่วงวันที่: <span class="percount font-bold">0</span> รายการ</div>
          <div class="mt-3 flex flex-wrap gap-2">
            <button class="run px-3 py-1.5 rounded bg-pink-600 text-white hover:bg-pink-700">รันไฟล์นี้</button>
            <button class="save px-3 py-1.5 rounded bg-emerald-600 text-white hover:bg-emerald-700" disabled>ดาวน์โหลดไฟล์นี้</button>
          </div>
          <div class="mt-2 text-xs text-gray-500 status">สถานะ: รอประมวลผล</div>
        </div>`;
      it.card=card;
      it.canvas=card.querySelector('canvas'); it.ctx=it.canvas.getContext('2d');
      it.pbar=card.querySelector('.progress-bar'); it.ptext=card.querySelector('.progress-text');
      it.percount=card.querySelector('.percount'); it.statusEl=card.querySelector('.status');
      it.runBtn=card.querySelector('.run'); it.saveBtn=card.querySelector('.save');
      it.runBtn.addEventListener('click',()=>processOne(it));
      it.saveBtn.addEventListener('click',()=>downloadOne(it));
      listEl.appendChild(card);
    }
    function setFileProgress(it,p,status='กำลังประมวลผล'){
      const pct=Math.round(Math.max(0,Math.min(1,p))*100);
      it.pbar.style.width=pct+'%'; it.ptext.textContent=pct+'%';
      it.statusEl.textContent=`สถานะ: ${status} ${pct}%`;
    }
    function setGlobalProgress(ratio){
      pWrap.classList.remove('hidden');
      const pct=Math.round(Math.max(0,Math.min(1,ratio))*100);
      pbar.style.width=pct+'%'; ptext.textContent=pct+'%';
    }
    function loadPreviewToCanvas(it, url, w, h){
      return new Promise((res,rej)=>{
        const img=new Image();
        img.onload=()=>{ it.canvas.width=w; it.canvas.height=h; it.ctx.drawImage(img,0,0,w,h); res(); };
        img.onerror=rej; img.src=url;
      });
    }
    function renderGlobalSummary(){
      let total=0, files=0;
      for(const it of items) if(it.processed){ total+=(it.matchCount||0); files++; }
      globalCountEl.textContent=total;
      fileDoneEl.textContent=files;
      summaryWrap.classList.toggle('hidden', files===0);
      downloadAllBtn.disabled = files===0;
    }

    /* ===================== Core OCR per file ===================== */
    async function processOne(it, idx=0, total=1){
      it.runBtn.disabled=true; it.saveBtn.disabled=true;
      setFileProgress(it,0,'เริ่ม'); it.statusEl.textContent='สถานะ: เตรียมเริ่ม OCR...';

      const phrase = targetInput.value || "";
      const { start, end } = getSelectedRangeCE();
      if (!phrase || !start || !end){
        alert('กรุณากรอก "คำค้นหา" และเลือก "ช่วงวันที่"');
        it.runBtn.disabled=false; return;
      }

      // 1) บีบอัด/ย่อ + แสดง preview
      const prep = await prepareImageForOCR(it.file);
      await loadPreviewToCanvas(it, prep.url, prep.width, prep.height);

      // progress รวม
      const onProgress = (m) => {
        if (m?.progress != null) {
          const ratio = Math.min(0.99, (idx + (m.progress||0)) / total);
          setGlobalProgress(ratio);
          setFileProgress(it, m.progress||0, m.status||'กำลังประมวลผล');
        }
        if (DEBUG) console.log('[tesseract]', m);
      };

      try{
        const { data } = await recognizeWithFallback(prep.blob, onProgress);

        if (DEBUG){
          console.group(`OCR: ${it.file.name}`);
          console.log('RAW OCR TEXT:\n' + (data.text||'').trim());
        }

        const lines = buildLines(data);
        if (!lines || !lines.length){
          it.statusEl.textContent='สถานะ: ไม่พบบรรทัดจาก OCR'; setFileProgress(it,1,'เสร็จสิ้น');
          if (DEBUG) console.groupEnd();
          it.runBtn.disabled=false; return;
        }

        // สร้าง index รายบรรทัด
        const rawLines = lines.map(L => L.map(w=>w.text).join(' '));
        const lineIdxs = lines.map(L => buildLineIndex(L));

        if (DEBUG){
          const cleanDump = lineIdxs.map((idx,i)=> `${i}: "${idx.clean}"`).join('\n');
          console.log('CLEANED LINES:\n' + cleanDump);
        }

        let perCount = 0;

        for (let i=0;i<lines.length;i++){
          const span = searchInIndexedLine(lineIdxs[i], phrase);
          if (!span) continue;

          // หา date ในแถวนี้/รอบ ๆ
          let dates = extractDates(rawLines[i]);
          if (!dates.length){
            const from=Math.max(0,i-LOOK_BEHIND), to=Math.min(lines.length-1,i+LOOK_AHEAD);
            for (let k=from;k<=to;k++){ if (k===i) continue;
              const more = extractDates(rawLines[k]); if (more.length){ dates.push(...more); break; }
            }
          }
          const ok = dates.some(d=>inRange(d,start,end));
          if (!ok) continue;

          // ไฮไลท์ชื่อ
          const sBox=lines[i][span.startWord].bbox, eBox=lines[i][span.endWord].bbox;
          const nameBox=unionBox(sBox,eBox);
          it.ctx.fillStyle='rgba(255,255,0,0.35)'; it.ctx.fillRect(nameBox.x, nameBox.y, nameBox.w, nameBox.h);
          it.ctx.lineWidth=3; it.ctx.strokeStyle='red'; it.ctx.strokeRect(nameBox.x, nameBox.y, nameBox.w, nameBox.h);

          // (ถ้าอยากไฮไลท์วันที่ด้วย เปิดบรรทัดด้านล่าง)
          // let dateBox=null; const from=Math.max(0,i-LOOK_BEHIND), to=Math.min(lines.length-1,i+LOOK_AHEAD);
          // for (let k=from;k<=to;k++){
          //   const ds=extractDates(rawLines[k]);
          //   if (ds.length && ds.some(d=>inRange(d,start,end))){
          //     const first=lines[k][0].bbox, last=lines[k].at(-1).bbox;
          //     dateBox=unionBox(first,last); break;
          //   }
          // }
          // if (dateBox){
          //   it.ctx.fillStyle='rgba(0,128,255,0.18)'; it.ctx.fillRect(dateBox.x, dateBox.y, dateBox.w, dateBox.h);
          //   it.ctx.lineWidth=3; it.ctx.strokeStyle='#0b66c3'; it.ctx.strokeRect(dateBox.x, dateBox.y, dateBox.w, dateBox.h);
          // }

          perCount++;
        }

        it.percount.textContent=perCount;
        it.statusEl.textContent='สถานะ: เสร็จแล้ว ✔';
        setFileProgress(it,1,'เสร็จสิ้น');
        it.processed=true; it.matchCount=perCount; it.saveBtn.disabled=false;

        renderGlobalSummary();
        if (DEBUG) console.groupEnd();
      }catch(e){
        console.error(e); it.statusEl.textContent='สถานะ: เกิดข้อผิดพลาด (ดู Console)';
      }finally{
        it.runBtn.disabled=false;
      }
    }

    /* ===================== Download (มือถือรองรับ) ===================== */
    const isIOS = () =>
      /iPad|iPhone|iPod/i.test(navigator.userAgent) ||
      (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

    function triggerBlobDownload(blob, filename){
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      if ('download' in a) {
        a.href=url; a.download=filename; document.body.appendChild(a);
        a.click(); a.remove();
        setTimeout(()=>URL.revokeObjectURL(url), 4000);
      } else {
        window.open(url, '_blank');
      }
    }

    function downloadCanvasMobileAware(canvas, filename='annotated.jpg'){
      if (isIOS()){
        const dataUrl = canvas.toDataURL('image/jpeg', 0.92);
        const win = window.open('', '_blank');
        if (win) {
          win.document.write(`<html><head><title>${filename}</title></head>
            <body style="margin:0"><img src="${dataUrl}" style="width:100%;height:auto"/></body></html>`);
          win.document.close();
        } else {
          alert('โปรดอนุญาตป๊อปอัป/แท็บใหม่ เพื่อบันทึกรูป');
        }
      } else {
        canvas.toBlob(b => triggerBlobDownload(b, filename), 'image/jpeg', 0.92);
      }
    }

    function downloadOne(it){
      const name = it.file.name.replace(/\.[^.]+$/, '') + '_annotated.jpg';
      downloadCanvasMobileAware(it.canvas, name);
    }
    function downloadAll(){
      const done = items.filter(i=>i.processed);
      if(!done.length) return;
      if (isIOS() && done.length > 1){
        alert('บน iPhone/iPad การดาวน์โหลดหลายไฟล์อัตโนมัติอาจถูกบล็อกป๊อปอัป\nระบบจะเปิดทีละรูปในแท็บใหม่ ให้บันทึกด้วยตนเอง');
      }
      done.forEach((it,idx)=> setTimeout(()=>downloadOne(it), idx*500));
    }

    /* ===================== Events ===================== */
    fileInput.addEventListener('change', async e=>{
      const files = Array.from(e.target.files||[]);
      if(!files.length) return;
      for(const f of files){
        const it={file:f, processed:false, matchCount:0};
        items.push(it); createCard(it);
      }
      window.scrollTo({top:document.body.scrollHeight, behavior:'smooth'});
    });

    runAllBtn.addEventListener('click', async ()=>{
      if(!items.length) return alert('กรุณาเลือกรูปก่อน');
      runAllBtn.disabled=true; downloadAllBtn.disabled=true; pWrap.classList.remove('hidden'); setGlobalProgress(0);
      try{
        for (let i=0;i<items.length;i++){
          await processOne(items[i], i, items.length);
        }
      } finally {
        runAllBtn.disabled=false;
        downloadAllBtn.disabled = items.filter(i=>i.processed).length===0;
        setGlobalProgress(1);
      }
    });

    downloadAllBtn.addEventListener('click', downloadAll);

    clearAllBtn.addEventListener('click', ()=>{
      items.length=0; listEl.innerHTML='';
      globalCountEl.textContent='0'; fileDoneEl.textContent='0'; summaryWrap.classList.add('hidden');
      pWrap.classList.add('hidden'); pbar.style.width='0%'; ptext.textContent='0%';
      fileInput.value='';
    });
  });
})();