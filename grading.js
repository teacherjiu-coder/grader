// 컴활 자동채점 엔진 (브라우저 & Node 공용)
// 의존성은 인자로 주입: XLSX, JSZip, DOMParser
async function gradeWorkbook(buf, deps, KEY){
  const {XLSX, JSZip, DOMParser} = deps;
  const wb = XLSX.read(buf, {type:'array', cellFormula:true, cellDates:true});
  const zip = await JSZip.loadAsync(buf);
  const files = {};
  for(const nm of Object.keys(zip.files)){
    if(/\.(xml|rels)$/.test(nm)) files[nm] = await zip.files[nm].async('string');
  }
  const parse = s => new DOMParser().parseFromString(s, 'application/xml');
  const tag = (el,t)=>[...el.getElementsByTagName(t)];

  // ---- 시트 이름 -> sheetN.xml 경로 매핑 ----
  const wbx = parse(files['xl/workbook.xml']);
  const rels = parse(files['xl/_rels/workbook.xml.rels']);
  const ridToTarget = {};
  tag(rels,'Relationship').forEach(r=>ridToTarget[r.getAttribute('Id')]=r.getAttribute('Target'));
  const sheetPath = {};
  tag(wbx,'sheet').forEach(s=>{
    const nm=s.getAttribute('name');
    const rid=s.getAttribute('r:id')||s.getAttribute('relationship:id')||
      [...s.attributes].find(a=>a.name.endsWith(':id'))?.value;
    let tgt=ridToTarget[rid]; if(tgt && !tgt.startsWith('xl/')) tgt='xl/'+tgt;
    sheetPath[nm]=tgt;
  });
  const sheetDoc = nm => { const p=sheetPath[nm]; return p&&files[p]?parse(files[p]):null; };

  // ---- styles.xml ----
  const styles = files['xl/styles.xml'] ? parse(files['xl/styles.xml']) : null;
  let fonts=[],fills=[],borders=[],cellXfs=[];
  if(styles){
    const fEl=tag(styles,'fonts')[0]; if(fEl) fonts=tag(fEl,'font');
    const flEl=tag(styles,'fills')[0]; if(flEl) fills=tag(flEl,'fill');
    const bEl=tag(styles,'borders')[0]; if(bEl) borders=tag(bEl,'border');
    const xEl=tag(styles,'cellXfs')[0]; if(xEl) cellXfs=tag(xEl,'xf');
  }
  function cellStyleIdx(sheetName,coord){
    const d=sheetDoc(sheetName); if(!d) return null;
    const c=tag(d,'c').find(x=>x.getAttribute('r')===coord);
    if(!c) return {missing:true};
    return {s:parseInt(c.getAttribute('s')||'0'), cEl:c, doc:d};
  }
  function styleOf(sheetName,coord){
    const info=cellStyleIdx(sheetName,coord);
    if(!info||info.missing) return null;
    const xf=cellXfs[info.s]; if(!xf) return {};
    const out={};
    // fill
    const fillId=parseInt(xf.getAttribute('fillId')||'0');
    if(fills[fillId]){
      const fg=tag(fills[fillId],'fgColor')[0];
      if(fg&&fg.getAttribute('rgb')) out.fill=fg.getAttribute('rgb').slice(-6).toUpperCase();
    }
    // font
    const fontId=parseInt(xf.getAttribute('fontId')||'0');
    if(fonts[fontId]){
      const f=fonts[fontId];
      const nm=tag(f,'name')[0]||tag(f,'rFont')[0];
      out.font=nm?nm.getAttribute('val'):null;
      const sz=tag(f,'sz')[0]; out.size=sz?parseFloat(sz.getAttribute('val')):null;
      out.bold=tag(f,'b').length>0;
    }
    if(fonts[fontId]){
      const f=fonts[fontId];
      out.italic=tag(f,'i').length>0;
      out.underline=tag(f,'u').length>0;
      const col=tag(f,'color')[0];
      out.fontColor=col?(col.getAttribute('rgb')||'').slice(-6).toUpperCase()||null:null;
    }
    // numFmt (표시형식)
    const numFmtId=xf.getAttribute('numFmtId');
    out.numFmtId=numFmtId;
    out.numFmtCode=null;
    if(styles){
      const nf=tag(styles,'numFmt').find(n=>n.getAttribute('numFmtId')===numFmtId);
      if(nf) out.numFmtCode=nf.getAttribute('formatCode');
    }
    // border
    const borderId=parseInt(xf.getAttribute('borderId')||'0');
    if(borders[borderId]){
      const b=borders[borderId];const get=t=>{const e=tag(b,t)[0];return e?e.getAttribute('style'):null;};
      out.border={left:get('left'),right:get('right'),top:get('top'),bottom:get('bottom')};
    }
    // alignment
    const al=tag(xf,'alignment')[0];
    out.align=al?al.getAttribute('horizontal'):null;
    return out;
  }

  // ---- comments ----
  function relsOf(sheetName){
    const p=sheetPath[sheetName]; if(!p) return null;
    const base=p.split('/').pop();
    const relPath='xl/worksheets/_rels/'+base+'.rels';
    return files[relPath]?parse(files[relPath]):null;
  }
  function commentAt(sheetName,coord){
    // 해당 시트에 연결된 comments 파일만 확인 (전역 검색 금지)
    const rd=relsOf(sheetName); if(!rd) return null;
    let cpath=null;
    tag(rd,'Relationship').forEach(r=>{if(/comment/i.test(r.getAttribute('Target'))){cpath=r.getAttribute('Target').replace('../','xl/');}});
    if(!cpath||!files[cpath]) return null;
    const cd=parse(files[cpath]);
    const cm=tag(cd,'comment').find(x=>x.getAttribute('ref')===coord);
    return cm?tag(cm,'t').map(t=>t.textContent).join(''):null;
  }
  function chartsForSheet(sheetName){
    const rd=relsOf(sheetName); if(!rd) return [];
    const drawings=[];
    tag(rd,'Relationship').forEach(r=>{const t=r.getAttribute('Target');if(/drawing\d+\.xml$/.test(t)) drawings.push(t.replace('../','xl/'));});
    const want=new Set();
    drawings.forEach(dp=>{
      const drel='xl/drawings/_rels/'+dp.split('/').pop()+'.rels';
      if(files[drel]){const dd=parse(files[drel]);tag(dd,'Relationship').forEach(r=>{const t=r.getAttribute('Target');if(/chart\d+\.xml$/.test(t)) want.add(t.replace('../','xl/'));});}
    });
    return charts.filter(c=>want.has(c.file));
  }
  function rangeCells(rng){
    const m=rng.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/); if(!m) return [rng];
    const col=s=>{let n=0;for(const ch of s)n=n*26+(ch.charCodeAt(0)-64);return n;};
    const colL=n=>{let s='';while(n>0){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26);}return s;};
    const c1=col(m[1]),r1=+m[2],c2=col(m[3]),r2=+m[4];const out=[];
    for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)out.push(colL(c)+r);
    return out;
  }

  // ---- defined names ----
  function definedName(name){
    const dn=tag(wbx,'definedName').find(d=>d.getAttribute('name')===name);
    return dn?dn.textContent:null;
  }

  // ---- conditional formatting ----
  function cfRules(sheetName){
    const d=sheetDoc(sheetName); if(!d) return [];
    return tag(d,'conditionalFormatting').map(cf=>({
      sqref:cf.getAttribute('sqref'),
      formulas:tag(cf,'formula').map(f=>f.textContent),
      type:(tag(cf,'cfRule')[0]||{}).getAttribute?tag(cf,'cfRule')[0].getAttribute('type'):null
    }));
  }

  // ---- charts ----
  const C='http://schemas.openxmlformats.org/drawingml/2006/chart';
  const A='http://schemas.openxmlformats.org/drawingml/2006/main';
  function serName(s){
    const txEl=s.getElementsByTagNameNS(C,'tx')[0];
    return txEl?([...txEl.getElementsByTagNameNS(C,'v')].map(x=>x.textContent).join('')||
              [...txEl.getElementsByTagNameNS(A,'t')].map(x=>x.textContent).join('')):'';
  }
  function parseCharts(){
    const out=[];
    for(const nm of Object.keys(files)){
      if(!/xl\/charts\/chart\d+\.xml$/.test(nm)) continue;
      const d=parse(files[nm]);
      const gt=(p,t)=>[...d.getElementsByTagNameNS(p,t)];
      const groups=[]; let barDir=null;
      for(const g of ['barChart','lineChart','pieChart','areaChart','scatterChart']){
        for(const el of gt(C,g)){
          const bd=el.getElementsByTagNameNS(C,'barDir')[0];
          if(bd&&!barDir) barDir=bd.getAttribute('val');
          const sers=[...el.getElementsByTagNameNS(C,'ser')].map(s=>({name:serName(s),dLbls:s.getElementsByTagNameNS(C,'dLbls').length>0,type:g}));
          groups.push({type:g,barDir:bd?bd.getAttribute('val'):null,series:sers});
        }
      }
      const allSeries=groups.flatMap(g=>g.series);
      const title=gt(C,'title').length?[...gt(C,'title')[0].getElementsByTagNameNS(A,'t')].map(x=>x.textContent).join(''):'';
      const lp=gt(C,'legendPos')[0];
      const valAx=gt(C,'valAx').map(ax=>{
        const mx=ax.getElementsByTagNameNS(C,'max')[0],mn=ax.getElementsByTagNameNS(C,'min')[0],mu=ax.getElementsByTagNameNS(C,'majorUnit')[0];
        return {max:mx?mx.getAttribute('val'):null,min:mn?mn.getAttribute('val'):null,unit:mu?mu.getAttribute('val'):null};
      });
      const axisTitles=[...gt(C,'valAx'),...gt(C,'catAx')].map(ax=>
        [...ax.getElementsByTagNameNS(A,'t')].map(x=>x.textContent).join('')).filter(Boolean);
      const rounded=gt(C,'roundedCorners')[0];
      out.push({file:nm,groups,barDir,title,hasTitle:gt(C,'title').length>0,series:allSeries,axisTitles,
        legend:gt(C,'legend').length>0, legendPos:lp?lp.getAttribute('val'):null,
        nValAx:gt(C,'valAx').length, valAx,
        dTable:gt(C,'dTable').length>0,
        rounded:rounded?rounded.getAttribute('val')!=='0':false});
    }
    return out;
  }
  const charts=parseCharts();

  // ===== helpers =====
  const normS=v=>String(v).trim().replace(/\s+/g,' ');
  const isEmpty=x=>x==null||normS(x)==='';
  function cellVal(sheetName,coord){const sh=wb.Sheets[sheetName];const c=sh?sh[coord]:null;return c?c.v:null;}
  function cellFormula(sheetName,coord){const sh=wb.Sheets[sheetName];const c=sh?sh[coord]:null;return c&&c.f?c.f:'';}
  function valEq(s,e,isDate){
    if(isEmpty(e))return isEmpty(s);
    if(s==null)return false;
    if(isDate){let x=s instanceof Date?s.toISOString().slice(0,10):normS(s);return x===e||x.slice(0,10)===e;}
    if(typeof e==='number'){const n=typeof s==='number'?s:parseFloat(String(s).replace(/,/g,''));return !isNaN(n)&&Math.abs(n-e)<0.5;}
    return normS(s)===normS(e);
  }

  // ===== grade each problem =====
  const results=[];
  for(const p of KEY.problems){
    let ok=true, msg='', manual=false;
    if(p.type==='input'||p.type==='sort'||p.type==='macro_value'){
      const wrong=p.cells.filter(c=>!valEq(cellVal(p.sheet,c.coord),c.value,c.is_date)).map(c=>c.coord);
      if(wrong.length){ok=false;msg=`셀 ${wrong.slice(0,6).join(', ')}${wrong.length>6?' 외':''} 불일치`;}
    }
    else if(p.type==='formula_group'){
      const bad=[];
      let groupFuncOk=(p.req_funcs||[]).length===0, groupTextOk=!p.req_text;
      for(const c of p.cells){
        const f=cellFormula(p.sheet,c.coord).toUpperCase().replace(/_XLFN\./g,'');
        if(f&&(p.req_funcs||[]).every(fn=>f.includes(fn))) groupFuncOk=true;
        if(p.req_text&&cellFormula(p.sheet,c.coord).includes(p.req_text)) groupTextOk=true;
      }
      for(const c of p.cells){ if(!valEq(cellVal(p.sheet,c.coord),c.value,c.is_date)) bad.push(c.coord); }
      if(bad.length){ok=false;msg=`${bad.slice(0,4).join(', ')} 값 불일치`;}
      else if(!groupFuncOk||!groupTextOk){ok=false;msg=`지정 함수/식(${(p.req_funcs||[]).join(',')}${p.req_text?' '+p.req_text:''}) 미사용`;}
    }
    else if(p.type==='fmt_align'){
      const bad=p.checks.filter(c=>{const st=styleOf(p.sheet,c.coord);return !st||st.align!==c.align;}).map(c=>c.coord);
      if(bad.length){ok=false;msg=`맞춤 미적용: ${bad.join(', ')}`;}
    }
    else if(p.type==='fmt_fillfont'){
      const st=styleOf(p.sheet,p.range.split(':')[0]);
      if(!st){ok=false;msg='서식 없음';}
      else{
        const errs=[];
        if((st.fill||'')!==p.fill) errs.push(`음영색(${st.fill||'없음'})`);
        if(st.font!==p.font) errs.push(`글꼴(${st.font||'없음'})`);
        if(st.size!==p.size) errs.push(`크기(${st.size||'없음'})`);
        if(!st.bold) errs.push('굵게');
        if(errs.length){ok=false;msg=errs.join(', ')+' 불일치';}
      }
    }
    else if(p.type==='fmt_border'){
      const [c1,c2]=p.range.split(':');
      const tl=styleOf(p.sheet,c1), br=styleOf(p.sheet,c2);
      const thick=s=>s==='medium'||s==='thick';
      const inner=styleOf(p.sheet,'D6');
      if(!tl||!br||!thick(tl.border?.left)||!thick(tl.border?.top)||!thick(br.border?.right)||!thick(br.border?.bottom)){
        ok=false;msg='굵은 바깥쪽 테두리 미적용';
      } else if(!inner||!inner.border?.left){ok=false;msg='모든 테두리(안쪽) 미적용';}
    }
    else if(p.type==='defined_name'){
      const ref=definedName(p.name);
      const norm=ref?ref.replace(/\$/g,'').replace(/'/g,''):'';
      if(!ref||!norm.includes(p.ref_contains)){ok=false;msg=`이름 '${p.name}' 미정의 또는 범위 불일치`;}
    }
    else if(p.type==='comment'){
      const t=commentAt(p.sheet,p.coord);
      if(!t){ok=false;msg=`[${p.coord}] 메모 없음`;}
      else if(!t.includes('매출액')){ok=false;msg=`메모 내용 불일치(${t.slice(0,15)})`;}
    }
    else if(p.type==='fmt_cells'){
      const bad=[];
      for(const chk of p.checks){
        const st=styleOf(p.sheet,chk.coord);
        if(!st){bad.push(chk.coord+'(서식없음)');continue;}
        for(const k of Object.keys(chk)){
          if(k==='coord')continue;
          let val=st[k];
          if(k==='fill'||k==='fontColor') val=(val||'');
          if(k==='numFmtCode'){ if(!val||!String(val).includes(chk[k])){bad.push(chk.coord+'(표시형식)');break;} continue; }
          if(String(val)!==String(chk[k])){bad.push(chk.coord+'('+k+')');break;}
        }
      }
      if(bad.length){ok=false;msg=bad.slice(0,5).join(', ')+' 불일치';}
    }
    else if(p.type==='merge'){
      const d=sheetDoc(p.sheet); let merged=[];
      if(d) merged=tag(d,'mergeCell').map(m=>m.getAttribute('ref'));
      if(!merged.includes(p.range)){ok=false;msg='['+p.range+'] 병합 안됨';}
    }
    else if(p.type==='macro_fill'){
      const bad=rangeCells(p.range).filter(cd=>{const st=styleOf(p.sheet,cd);return !st||(st.fill||'')!==p.fill;});
      if(bad.length){ok=false;msg=`[${p.range}] 음영색(${p.fill}) 미적용`;}
    }
    else if(p.type==='cond_format'){
      const rules=cfRules(p.sheet);
      let hit;
      if(p.rule_type){ hit=rules.find(r=>r.type===p.rule_type); if(!hit){ok=false;msg=`조건부서식(${p.rule_type}) 규칙 없음`;} }
      else { hit=rules.find(r=>r.formulas.some(f=>(p.req_funcs||[]).every(fn=>f.toUpperCase().includes(fn)))); if(!hit){ok=false;msg=`조건부서식 규칙 없음 또는 함수(${(p.req_funcs||[]).join(',')}) 미사용`;} }
    }
    else if(p.type==='sheet_protect'){
      const d=sheetDoc(p.sheet);
      if(!d||tag(d,'sheetProtection').length===0){ok=false;msg='시트 보호 미설정';}
    }
    else if(p.type==='data_validation'){
      const d=sheetDoc(p.sheet);
      const dvs=d?tag(d,'dataValidation'):[];
      if(dvs.length===0){ok=false;msg='데이터 유효성 검사 미설정';}
      else if(p.sqref_contains && !dvs.some(dv=>(dv.getAttribute('sqref')||'').includes(p.sqref_contains))){ok=false;msg=`[${p.sqref_contains}] 유효성 검사 없음`;}
    }
    else if(p.type==='manual'){
      manual=true; msg=p.note||'수동 채점 필요';
    }
    else if(p.type==='scenario'){
      const all=new Set();
      wb.SheetNames.forEach(n=>{const s=wb.Sheets[n];Object.keys(s).forEach(k=>{if(k[0]!=='!'&&typeof s[k].v==='number')all.add(Math.round(s[k].v));});});
      const found=p.result_values.filter(v=>all.has(Math.round(v))).length;
      const hasSummary=wb.SheetNames.some(n=>n.includes('시나리오'));
      if(!hasSummary||found<p.result_values.length){ok=false;msg=`시나리오 요약 시트/결과값 불일치 (결과 ${found}/${p.result_values.length})`;}
    }
    else if(p.type==='chart_series'){
      const cs=chartsForSheet(p.sheet);
      const good=cs.some(c=>{const names=c.series.map(s=>s.name);
        return p.include.every(i=>names.includes(i)) && !p.exclude.some(e=>names.includes(e));});
      if(!good){ok=false;msg=`'${p.exclude.join(',')}' 계열이 제거되지 않음`;}
    }
    else if(p.type==='chart_type'){
      if(!chartsForSheet(p.sheet).some(c=>c.barDir===p.barDir)){ok=false;msg=`'묶은 가로 막대형'(barDir=bar) 아님`;}
    }
    else if(p.type==='chart_titles'){
      const cs=chartsForSheet(p.sheet);const c=cs.find(c=>c.hasTitle)||cs[0];
      const needAxis=p.axis||0;
      if(!c||!c.hasTitle){ok=false;msg='차트 제목 없음';}
      else if(needAxis>0 && c.axisTitles.length<needAxis){ok=false;msg=`축 제목 부족 (${c.axisTitles.length}/${needAxis})`;}
    }
    else if(p.type==='chart_dlbl'){
      const c=chartsForSheet(p.sheet).find(c=>c.series.some(s=>s.name===p.series&&s.dLbls));
      if(!c){ok=false;msg=`'${p.series}' 계열 데이터 레이블 없음`;}
    }
    else if(p.type==='chart_combo'){
      const good=chartsForSheet(p.sheet).some(c=>c.series.some(s=>s.name===p.series&&s.type==='lineChart'));
      if(!good){ok=false;msg=`'${p.series}' 계열이 꺾은선형 아님`;}
    }
    else if(p.type==='chart_secondary'){
      if(!chartsForSheet(p.sheet).some(c=>c.nValAx>=2)){ok=false;msg='보조 축 없음';}
    }
    else if(p.type==='chart_legend'){
      if(!chartsForSheet(p.sheet).some(c=>c.legendPos===p.pos)){ok=false;msg=`범례 위치(${p.pos}) 아님`;}
    }
    else if(p.type==='chart_axis'){
      const good=chartsForSheet(p.sheet).some(c=>c.valAx.some(a=>
        (!p.max||a.max===p.max)&&(!p.min||a.min===p.min)&&(!p.unit||a.unit===p.unit)));
      if(!good){ok=false;msg=`축 설정(${[p.max&&'최대'+p.max,p.min&&'최소'+p.min,p.unit&&'단위'+p.unit].filter(Boolean).join(',')}) 불일치`;}
    }
    else if(p.type==='chart_datatable'){
      if(!chartsForSheet(p.sheet).some(c=>c.dTable)){ok=false;msg='데이터 표 없음';}
    }
    else if(p.type==='chart_rounded'){
      if(!chartsForSheet(p.sheet).some(c=>c.rounded)){ok=false;msg='둥근 모서리 미적용';}
    }
    else if(p.type==='chart_style'){
      manual=true; msg='차트 영역 윤곽선/둥근모서리/그림자 — 수동 확인 필요';
    }
    results.push({id:p.id,area:p.area,points:p.points,ok:(manual?true:ok),earned:(manual?p.points:(ok?p.points:0)),msg:(manual?(p.note||msg||'자동 확인 불가 → 만점 처리'):msg),answer:p.answer,type:p.type,manual});
  }
  const total=results.reduce((s,r)=>s+r.earned,0);
  const manualPts=results.filter(r=>r.manual).reduce((s,r)=>s+r.points,0);
  const autoMax=100-manualPts;
  return {total,autoMax,manualPts,pass:total>=KEY.pass,results,charts};
}
if(typeof module!=='undefined') module.exports={gradeWorkbook};
if(typeof window!=='undefined') window.gradeWorkbook=gradeWorkbook;
