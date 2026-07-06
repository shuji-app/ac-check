import { useState, useRef, useEffect, memo } from "react";

const INDOOR_FIELDS = [
  { code:"b1", label:"吸込温度",       unit:"°C",    step:0.1,  group:"indoor" },
  { code:"b2", label:"吹出温度",       unit:"°C",    step:0.1,  group:"indoor" },
  { code:"b3", label:"室内熱交・入口", unit:"°C",    step:0.1,  group:"indoor" },
  { code:"b4", label:"室内熱交・中間", unit:"°C",    step:0.1,  group:"indoor" },
  { code:"b5", label:"室内熱交・出口", unit:"°C",    step:0.1,  group:"indoor" },
  { code:"b6", label:"膨張弁開度",     unit:"pulse", step:1,    group:"indoor" },
  { code:"b8", label:"リモコン感知",   unit:"°C",    step:0.1,  group:"indoor" },
];
const OUTDOOR_FIELDS = [
  { code:"E1", label:"吐出ガス温度",   unit:"°C",    step:0.1,  group:"outdoor" },
  { code:"E2", label:"吸入ガス温度",   unit:"°C",    step:0.1,  group:"outdoor" },
  { code:"E3", label:"室外熱交・蒸発", unit:"°C",    step:0.1,  group:"outdoor" },
  { code:"E4", label:"圧縮機頂部",     unit:"°C",    step:0.1,  group:"outdoor" },
  { code:"F1", label:"外気温度",       unit:"°C",    step:0.1,  group:"outdoor" },
  { code:"F2", label:"過冷却液温度",   unit:"°C",    step:0.1,  group:"outdoor" },
  { code:"F3", label:"室外熱交・出口", unit:"°C",    step:0.1,  group:"outdoor" },
  { code:"H1", label:"高圧圧力",       unit:"MPa",   step:0.01, group:"outdoor" },
  { code:"H2", label:"低圧圧力",       unit:"MPa",   step:0.01, group:"outdoor" },
  { code:"H3", label:"運転電流",       unit:"A",     step:0.1,  group:"outdoor" },
  { code:"H4", label:"運転周波数",     unit:"Hz",    step:1,    group:"outdoor" },
];
const ALL_FIELDS = [...INDOOR_FIELDS, ...OUTDOOR_FIELDS];
const defVis = () => { const v={}; ALL_FIELDS.forEach(f=>v[f.code]=true); return v; };
const defLim = () => { const v={}; ALL_FIELDS.forEach(f=>v[f.code]={enabled:false,min:"",max:""}); return v; };
const emptyVal = () => { const v={}; ALL_FIELDS.forEach(f=>v[f.code]=""); return v; };

const emptyForm = (inspector="",date="") => ({
  floor:"", room:"", managementNo:"", unitNo:"",
  inspectionDate: date || new Date().toISOString().slice(0,10),
  inspector, preOperation:"", preMode:"", preWind:"", preSetTemp:"",
  values: emptyVal(), checks:{}, remarks:"",
});

function isAbn(code,val,limits) {
  const v=parseFloat(val); if(val===""||isNaN(v)) return false;
  const l=limits[code]; if(!l||!l.enabled) return false;
  if(l.min!==""&&!isNaN(parseFloat(l.min))&&v<parseFloat(l.min)) return true;
  if(l.max!==""&&!isNaN(parseFloat(l.max))&&v>parseFloat(l.max)) return true;
  return false;
}
// 列名候補からキーを探す（大文字小文字・全半角を無視）
function findColKey(keys, ...candidates) {
  const norm = s => s.trim().replace(/\s/g,"").toLowerCase();
  for(const cand of candidates){
    const found = keys.find(k=>norm(k)===norm(cand));
    if(found) return found;
  }
  return null;
}
// JSON行配列（xlsxのsheet_to_json結果）をdevListに変換
function parseDevRows(rows, origKeys) {
  if(!rows||rows.length===0) return [];
  const keys = origKeys || Object.keys(rows[0]);
  const fk = findColKey(keys,"階","floor","フロア","エリア","area");
  const rk = findColKey(keys,"部屋名","room","部屋","室名");
  const mk = findColKey(keys,"管理番号","managementno","管理No","管理no","管理");
  const uk = findColKey(keys,"機器番号","unitno","機器No","機器no","機器","unit");
  // 見つからない場合は列順で対応
  const fk2=fk||keys[0]||null, rk2=rk||keys[1]||null;
  const mk2=mk||keys[2]||null, uk2=uk||keys[3]||null;
  return rows.map(row=>{
    const raw={};
    keys.forEach(k=>{raw[k]=String(row[k]||"");});
    return {
      floor: String(row[fk2]||"").trim(),
      room:  String(row[rk2]||"").trim(),
      managementNo: String(row[mk2]||"").trim(),
      unitNo: String(row[uk2]||"").trim(),
      _raw: raw,
    };
  }).filter(r=>r.managementNo||r.unitNo);
}
// CSV文字列からdevListに変換
function parseDevCSV(text) {
  const lines = text.trim().split(/\r?\n/).filter(l=>l.trim());
  if(lines.length<2) return [];
  const origKeys = lines[0].split(",").map(s=>s.trim().replace(/^"|"$/g,""));
  const rows = lines.slice(1).map(line=>{
    const c=line.split(",").map(s=>s.trim().replace(/^"|"$/g,""));
    const obj={};
    origKeys.forEach((k,i)=>{obj[k]=c[i]||"";});
    return obj;
  });
  return parseDevRows(rows, origKeys);
}
function parseInspCSV(text) {
  return text.trim().split(/\r?\n/).map(l=>l.trim().replace(/^"|"$/g,"")).filter(Boolean);
}
function doExport(records,visibility,checkFields=[]) {
  const vf=ALL_FIELDS.filter(f=>visibility[f.code]);
  const rows=[
    ["点検日","点検者","階","部屋名","管理番号","機器番号","運転","モード","風量","設定温度",...vf.map(f=>f.code+"("+f.unit+")"),...checkFields.map(f=>f.label),"備考"],
    ...records.map(r=>[r.inspectionDate,r.inspector,r.floor,r.room,r.managementNo,r.unitNo,
      r.preOperation||"",r.preMode||"",r.preWind||"",r.preSetTemp||"",...vf.map(f=>r.values[f.code]),...checkFields.map(f=>r.checks?.[f.code]||""),r.remarks])
  ];
  const tsv=rows.map(r=>r.map(c=>'"'+String(c??"").replace(/"/g,'""')+'"').join("\t")).join("\n");
  const blob=new Blob(["\uFEFF"+tsv],{type:"text/tab-separated-values;charset=utf-8;"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
  a.download="ac_check_"+new Date().toISOString().slice(0,10)+".tsv"; a.click();
}

const PS="@media print{body>*{display:none!important;}#print-area{display:block!important;position:static!important;}@page{size:A4 landscape;margin:10mm;}}";
// ─── ブラウザ保存ヘルパー（機器リスト等の「取込データ」永続化用）───
// アプリのコードを更新（再読込）しても、設定画面で読み込んだデータが消えないようにする。
// 対象：機器リスト・表示列設定・点検項目・点検者リスト・表示項目設定・正常値範囲
// 対象外：入力中の点検記録（indoorRecords/outdoorRecords）・セッション情報は従来通り
function lsGet(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw !== null ? JSON.parse(raw) : fallback;
  } catch(e) { return fallback; }
}
function lsSet(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch(e) {}
}
const LS_KEYS = {
  devList:"acDevList", devColumns:"acDevColumns", devVisibleCols:"acDevVisibleCols",
  inspList:"acInspList", checkFields:"acCheckFields", limits:"acLimits", vis:"acVis",
};
const C={
  navy:"#1B3A6B",blue:"#2563B0",teal:"#0D7A6B",green:"#059669",
  red:"#DC2626",purple:"#7C3AED",
  g50:"#F8FAFC",g100:"#F1F5F9",g200:"#E2E8F0",
  g300:"#CBD5E1",g400:"#94A3B8",g500:"#64748B",g600:"#475569",g800:"#1E293B",
  white:"#FFFFFF",inp:"#F7F9FC",
};

const FieldRow = memo(function FR({f,isIn,idx,active,val,abn,dimmed,onClick,fRef}) {
  const fill=val!=="";
  return (
    <tr ref={fRef} onClick={onClick} style={{
      background:active?"linear-gradient(90deg,#2563B022,#2563B00A)":abn?"#FEF2F2":dimmed?"#F5F7FA":idx%2===0?C.white:C.g50,
      cursor:"pointer",outline:active?"2px solid "+C.blue:"none",outlineOffset:-1,opacity:dimmed?0.45:1,
    }}>
      <td style={{width:5,padding:0,background:active?C.blue:abn?C.red:"transparent",borderBottom:active?"2px solid "+C.blue:"1px solid "+C.g100}}/>
      <td style={{padding:active?"18px 10px":"9px 8px",fontFamily:"monospace",fontWeight:700,fontSize:active?20:13,color:active?C.blue:isIn?C.blue:C.teal,borderBottom:active?"2px solid "+C.blue:"1px solid "+C.g100,whiteSpace:"nowrap"}}>{f.code}</td>
      <td style={{padding:active?"18px 10px":"9px 8px",fontSize:active?18:13,color:active?C.navy:abn?C.red:C.g600,fontWeight:active?700:400,borderBottom:active?"2px solid "+C.blue:"1px solid "+C.g100}}>
        {f.label}{abn&&<span style={{marginLeft:4,fontSize:9,background:C.red,color:C.white,padding:"1px 4px",borderRadius:3,fontWeight:700}}>⚠️</span>}
      </td>
      <td style={{padding:active?"18px 6px":"9px 6px",fontSize:active?14:11,color:C.g400,borderBottom:active?"2px solid "+C.blue:"1px solid "+C.g100,textAlign:"center",whiteSpace:"nowrap"}}>{f.unit}</td>
      <td style={{padding:active?"10px 8px":"4px 7px",borderBottom:active?"2px solid "+C.blue:"1px solid "+C.g100,width:110}}>
        <div style={{padding:active?"12px 14px":"6px 10px",borderRadius:8,fontSize:active?26:14,fontFamily:"monospace",textAlign:"right",fontWeight:800,
          border:"2px solid "+(active?C.blue:fill?C.green:C.g200),
          background:active?"#EFF6FF":fill?"#F0FDF4":C.white,
          color:abn?C.red:fill?C.g800:C.g300,minHeight:active?52:32,display:"flex",alignItems:"center",justifyContent:"flex-end"}}>
          {fill ? val : active ? <span style={{color:C.blue+"80",fontSize:18}}>—</span> : "—"}
        </div>
      </td>
    </tr>
  );
});

function Numpad({mode="numeric",display,onPress,onConfirm,canConfirm,checkLabel,checkCategory,checkValue,onCheckPress,onPrev,onNext,canPrev,canNext,onSave,saveComplete,saveMissing}) {
  const isCheck = mode==="check";
  const KEYS=[[7,8,9],[4,5,6],[1,2,3],[0,"."]];
  const kb=en=>({flex:1,padding:0,height:56,borderRadius:12,border:"none",cursor:en?"pointer":"not-allowed",fontWeight:800,fontFamily:"monospace",fontSize:22,background:en?C.white:C.g100,color:en?C.g800:C.g300,boxShadow:en?"0 2px 6px rgba(0,0,0,0.10)":"none",display:"flex",alignItems:"center",justifyContent:"center"});
  const nb=en=>({flex:1,height:52,borderRadius:12,border:"2px solid "+(en?C.blue:C.g200),cursor:en?"pointer":"not-allowed",fontSize:22,fontWeight:800,background:en?C.blue+"15":C.g50,color:en?C.blue:C.g300,display:"flex",alignItems:"center",justifyContent:"center"});
  return (
    <div style={{width:208,flexShrink:0,display:"flex",flexDirection:"column",background:C.g50,borderLeft:"2px solid "+C.g200,padding:"8px 8px 10px",gap:6}}>
      {/* 表示エリア：数値入力中は入力値、チェック項目選択中は項目ラベル（枠の大きさは画面によらず統一） */}
      <div style={{background:C.white,borderRadius:12,padding:"8px 12px",border:"2px solid "+(isCheck?C.green:C.blue),height:58,boxSizing:"border-box",display:"flex",flexDirection:"column",justifyContent:"center"}}>
        {isCheck ? (
          <>
            <div style={{fontSize:10,fontWeight:700,color:"#059669",letterSpacing:"0.04em",minHeight:14}}>{checkCategory||""}</div>
            <div style={{fontSize:16,fontWeight:800,color:C.navy,lineHeight:1.3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{checkLabel||"—"}</div>
          </>
        ) : (
          <>
            <div style={{fontSize:10,fontWeight:700,color:C.g400,letterSpacing:"0.04em"}}>入力値</div>
            <div style={{fontFamily:"monospace",fontSize:32,fontWeight:800,color:display?C.navy:C.g300,textAlign:"right",lineHeight:1.1}}>
              {display || <span style={{color:C.g200}}>—</span>}
            </div>
          </>
        )}
      </div>
      {/* 数値キーパッド：常時表示。チェック項目選択中は無効化（グレーアウト）してテンキーと○×を同じ配置で常設する */}
      {KEYS.map((row,ri) => (
        <div key={ri} style={{display:"flex",gap:6}}>
          {row.map(k => (
            <button key={k} onClick={()=>!isCheck&&onPress(k)} disabled={isCheck} style={{...kb(!isCheck),flex:(ri===3&&k===0)?2:1}}>{k}</button>
          ))}
        </div>
      ))}
      {/* ○×大ボタン：常時表示。数値項目選択中は無効化（グレーアウト） */}
      <div style={{display:"flex",gap:6}}>
        {["○","×"].map(v=>(
          <button key={v} onClick={()=>isCheck&&onCheckPress&&onCheckPress(v)} disabled={!isCheck}
            style={{flex:1,height:56,borderRadius:12,border:"2px solid "+(isCheck&&checkValue===v?(v==="○"?C.green:C.red):C.g200),
              background:isCheck?(checkValue===v?(v==="○"?C.green:C.red):C.white):C.g100,
              color:isCheck?(checkValue===v?C.white:C.g800):C.g300,fontWeight:800,fontSize:28,cursor:isCheck?"pointer":"not-allowed",transition:"all 0.1s",
              display:"flex",alignItems:"center",justifyContent:"center",boxShadow:isCheck?"0 2px 6px rgba(0,0,0,0.10)":"none"}}>
            {v}
          </button>
        ))}
      </div>
      {/* ▲▼：モードによらず常に同じ位置・同じ構造で表示（切替時のちらつき防止）。▼を押すと入力値を確定して次項目へ進む（Enter機能を統合） */}
      <div style={{display:"flex",gap:6}}>
        <button onClick={onPrev} disabled={!canPrev} style={{...nb(canPrev),flex:1}}>▲</button>
        <button onClick={onNext} disabled={!canNext} style={{...nb(canNext),flex:1}}>▼</button>
      </div>
      {/* 保存ボタン（右サイドパネルに統合。備考欄の隣ではなくここに常時表示） */}
      {onSave && (
        <button onClick={onSave} disabled={!saveComplete}
          style={{marginTop:2,padding:"12px 8px",borderRadius:12,border:"none",cursor:saveComplete?"pointer":"not-allowed",fontWeight:800,fontSize:saveComplete?16:12,
            background:saveComplete?"linear-gradient(135deg,"+C.green+",#047857)":C.g200,
            color:saveComplete?C.white:C.g400,
            boxShadow:saveComplete?"0 4px 14px rgba(5,150,105,0.35)":"none",lineHeight:1.4,textAlign:"center"}}>
          {saveComplete?"💾 保存":"⏳ あと"+(saveMissing||0)+"項目"}
        </button>
      )}
    </div>
  );
}

function S1Head({num,label,done,active,onClick}) {
  return (
    <div onClick={onClick}
      style={{display:"flex",alignItems:"center",gap:8,padding:"8px 14px",cursor:"pointer",
        background:done?C.green:active?C.blue+"10":C.g50,
        borderBottom:"1px solid "+C.g100,transition:"background 0.15s"}}>
      <div style={{width:22,height:22,borderRadius:11,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:11,
        background:done?C.white:active?C.blue+"30":C.g200,
        color:done?C.green:active?C.blue:C.g500}}>
        {done?"✓":num}
      </div>
      <span style={{fontSize:13,fontWeight:700,color:done?C.white:active?C.blue:C.g600,flex:1}}>{label}</span>
    </div>
  );
}

function Step2View({form,setInfo,handleSave,setStep,visIn,visOut,visFields,activeCode,numDisp,limits,onPress,onConfirm,onRowClick,moveActive,rowRefs,listRef,complete,missing,editIdx,ALL_FIELDS,vis,isAbn,setCheck,checkFields,inspectionMode,focusSeq,isCheckCode,setCheckAndAdvance,hideHeader,hideNumpad}) {
  const isOutdoor = inspectionMode==="outdoor";
  const chkFields = checkFields||[];
  const outChkFields = chkFields.filter(f=>f.group==="check_out");
  const ciFields = chkFields.filter(f=>f.group==="check_in");
  // 室外機モード：チェック項目のみで完了判定
  const filled = isOutdoor
    ? outChkFields.filter(f=>(form.checks?.[f.code]||"")!=="").length
    : visFields.filter(f=>form.values[f.code]!=="").length;
  const total = isOutdoor ? outChkFields.length : visFields.length;
  const pct=total>0?Math.round(filled/total*100):0;
  const seq = focusSeq||visFields;
  const ai=seq.findIndex(f=>f.code===activeCode);
  const activeIsCheck = !isOutdoor && !!activeCode && !!isCheckCode && isCheckCode(activeCode);
  const activeCheckField = activeIsCheck ? ciFields.find(f=>f.code===activeCode) : null;
  return (
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
      {!hideHeader && (
      <div style={{flexShrink:0,background:"linear-gradient(135deg,"+C.navy+","+C.blue+")",padding:"8px 14px",boxShadow:"0 2px 8px rgba(0,0,0,0.2)"}}>
        <div style={{display:"flex",gap:14,flexWrap:"wrap",alignItems:"center"}}>
          {[["点検日",form.inspectionDate],["点検者",form.inspector],["階",form.floor],["部屋名",form.room],["管理番号",form.managementNo],["機器番号",form.unitNo]].map(([k,v])=>(
            <div key={k} style={{display:"flex",gap:4,alignItems:"baseline"}}>
              <span style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,0.55)",textTransform:"uppercase",letterSpacing:"0.05em"}}>{k}</span>
              <span style={{fontSize:13,fontWeight:700,color:C.white}}>{v||"—"}</span>
            </div>
          ))}
          <button onClick={()=>setStep(1)} style={{marginLeft:"auto",padding:"4px 12px",borderRadius:6,border:"1.5px solid rgba(255,255,255,0.4)",background:"rgba(255,255,255,0.15)",color:C.white,cursor:"pointer",fontSize:12,fontWeight:700}}>← 修正</button>
        </div>
      </div>
      )}
      <div style={{flex:1,display:"flex",overflow:"hidden"}}>
        {isOutdoor && (
          <div style={{flex:1,overflowY:"auto",scrollbarGutter:"stable",background:C.white,padding:"8px"}}>
            <div style={{fontSize:12,fontWeight:700,color:C.teal,marginBottom:8,padding:"6px 10px",background:C.teal+"10",borderRadius:8}}>🏭 室外機点検チェック（{form.floor} {form.room} {form.managementNo}）</div>
            {(()=>{
              const cats=[...new Set(outChkFields.map(f=>f.category))];
              return cats.map(cat=>(
                <div key={cat} style={{marginBottom:12}}>
                  <div style={{fontSize:11,fontWeight:700,color:C.teal,marginBottom:4,padding:"3px 8px",background:C.teal+"15",borderRadius:5,display:"inline-block"}}>{cat}</div>
                  {outChkFields.filter(f=>f.category===cat).map((f,i)=>{
                    const val=form.checks?.[f.code]||"";
                    return (
                      <div key={f.code} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",borderBottom:"1px solid "+C.g100,background:i%2===0?C.white:C.g50}}>
                        <span style={{flex:1,fontSize:14,color:C.g700}}>{f.label}</span>
                        <div style={{display:"flex",gap:6}}>
                          {["○","×"].map(v=>(
                            <button key={v} onClick={()=>setCheck(f.code,v)}
                              style={{width:52,height:44,borderRadius:10,border:"2px solid "+(val===v?(v==="○"?C.green:C.red):C.g200),
                                background:val===v?(v==="○"?C.green:C.red):C.white,
                                color:val===v?C.white:C.g400,fontWeight:800,fontSize:22,cursor:"pointer",transition:"all 0.1s"}}>
                              {v}
                            </button>
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </div>
              ));
            })()}
          </div>
        )}
        {!isOutdoor && <div ref={listRef} style={{flex:1,overflowY:"auto",scrollbarGutter:"stable",background:C.white,position:"relative"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead>
              <tr style={{background:C.g100,position:"sticky",top:0,zIndex:2}}>
                <th style={{width:5,padding:0,borderBottom:"2px solid "+C.g200}}/>
                {[["コード",62],["項目名",null],["単位",48],["測定値",110]].map(([h,w])=>(
                  <th key={h} style={{padding:"8px 8px",fontSize:11,fontWeight:700,color:C.g500,textAlign:h==="測定値"?"right":"center",borderBottom:"2px solid "+C.g200,width:w||undefined,whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {visIn.length>0 && (
                <>
                  <tr><td colSpan={5} style={{padding:"6px 10px",background:C.blue+"12",fontSize:11,fontWeight:700,color:C.blue,borderBottom:"1px solid "+C.blue+"20"}}>室内機（インドア）</td></tr>
                  {visIn.map((f,i)=>{
                    const act=activeCode===f.code;
                    const v=act?numDisp:form.values[f.code];
                    return <FieldRow key={f.code} f={f} isIn idx={i} active={act} val={v} abn={isAbn(f.code,v,limits)} dimmed={!!activeCode&&!act} onClick={()=>onRowClick(f)} fRef={el=>{rowRefs.current[f.code]=el;}}/>;
                  })}
                </>
              )}
              {visOut.length>0 && (
                <>
                  <tr><td colSpan={5} style={{padding:"6px 10px",background:C.teal+"12",fontSize:11,fontWeight:700,color:C.teal,borderBottom:"1px solid "+C.teal+"20"}}>室外機（アウトドア）</td></tr>
                  {visOut.map((f,i)=>{
                    const act=activeCode===f.code;
                    const v=act?numDisp:form.values[f.code];
                    return <FieldRow key={f.code} f={f} isIn={false} idx={i} active={act} val={v} abn={isAbn(f.code,v,limits)} dimmed={!!activeCode&&!act} onClick={()=>onRowClick(f)} fRef={el=>{rowRefs.current[f.code]=el;}}/>;
                  })}
                </>
              )}
            </tbody>
          </table>
          {/* 室内機チェック項目（室外機チェックと同じ3列表：項目／点検内容／○×） */}
          {ciFields.length>0 && (
            <div style={{padding:"14px 12px 20px",background:C.g50,borderTop:"2px solid "+C.g200}}>
              <div style={{fontSize:12,fontWeight:800,color:"#059669",marginBottom:8}}>✅ 室内機チェック項目</div>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead>
                  <tr>
                    <th style={{width:90,padding:"6px 8px",textAlign:"left",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>項目</th>
                    <th style={{padding:"6px 8px",textAlign:"left",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>点検内容</th>
                    <th style={{width:110,padding:"6px 8px",textAlign:"center",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>○×</th>
                  </tr>
                </thead>
                <tbody>
                  {ciFields.map((f,i)=>{
                    const val=form.checks?.[f.code]||"";
                    const act=activeCode===f.code;
                    return (
                      <tr key={f.code} ref={el=>{if(rowRefs)rowRefs.current[f.code]=el;}}
                        onClick={()=>onRowClick&&onRowClick(f)}
                        style={{background:act?C.blue+"0C":i%2===0?C.white:C.g50,boxShadow:act?"inset 0 0 0 2px "+C.blue:"none",cursor:"pointer",transition:"box-shadow 0.1s"}}>
                        <td style={{padding:"9px 8px",fontSize:11,fontWeight:700,color:C.teal,borderBottom:"1px solid "+C.g100,verticalAlign:"middle"}}>{f.category}</td>
                        <td style={{padding:"9px 8px",fontSize:14,color:C.g700,borderBottom:"1px solid "+C.g100,verticalAlign:"middle"}}>{f.label}</td>
                        <td style={{padding:"6px 8px",borderBottom:"1px solid "+C.g100,verticalAlign:"middle"}}>
                          <div style={{display:"flex",gap:6,justifyContent:"center"}} onClick={e=>e.stopPropagation()}>
                            {["○","×"].map(v=>(
                              <button key={v} onClick={()=>setCheckAndAdvance?setCheckAndAdvance(f.code,v):setCheck(f.code,v)}
                                style={{width:48,height:40,borderRadius:9,border:"2px solid "+(val===v?(v==="○"?C.green:C.red):C.g200),
                                  background:val===v?(v==="○"?C.green:C.red):C.white,
                                  color:val===v?C.white:C.g400,fontWeight:800,fontSize:19,cursor:"pointer",transition:"all 0.1s"}}>{v}</button>
                            ))}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>}
        {!isOutdoor && !hideNumpad && (
          <Numpad
            mode={activeIsCheck?"check":"numeric"}
            display={numDisp} onPress={onPress} onConfirm={onConfirm} canConfirm={!!activeCode&&!activeIsCheck&&numDisp!==""}
            checkLabel={activeCheckField?.label} checkCategory={activeCheckField?.category}
            checkValue={form.checks?.[activeCode]||""}
            onCheckPress={v=>setCheckAndAdvance&&setCheckAndAdvance(activeCode,v)}
            onPrev={()=>moveActive(-1)} onNext={()=>moveActive(1)}
            canPrev={ai>0} canNext={ai>=0}
          />
        )}
      </div>
      <div style={{flexShrink:0,background:C.white,borderTop:"2px solid "+C.g200,padding:"10px 12px",display:"flex",gap:10,alignItems:"stretch"}}>
        <div style={{flex:1,display:"flex",flexDirection:"column",gap:4}}>
          <div style={{fontSize:12,fontWeight:700,color:C.g500}}>📝 備考・特記事項</div>
          <textarea value={form.remarks} onChange={e=>setInfo("remarks",e.target.value)} placeholder="異常箇所、特記事項など..."
            style={{flex:1,width:"100%",padding:"10px 12px",borderRadius:9,fontSize:14,border:"1.5px solid "+C.g200,background:C.inp,outline:"none",boxSizing:"border-box",fontFamily:"inherit",resize:"none",minHeight:100,lineHeight:1.5}}/>
        </div>
      </div>
    </div>
  );
}

// ─── セッション開始画面 ────────────────────────────────────────────────────
function SessionView({ devList, inspList, sessionInfo, setSessionInfo, onStart, onSelectOutdoor, records, undoneOnly, setUndoneOnly, devColumns, devVisibleCols }) {
  const today = new Date().toISOString().slice(0,10);
  const [date, setDate] = useState(sessionInfo?.date || today);
  const [inspector, setInspector] = useState(sessionInfo?.inspector || "");
  const [inspector2, setInspector2] = useState(sessionInfo?.inspector2 || "");
  const [showInspector2, setShowInspector2] = useState(!!sessionInfo?.inspector2);
  const [access, setAccess] = useState(sessionInfo?.roomAccess || {});
  const [memoOpen, setMemoOpen] = useState({});
  const [floorSortAsc, setFloorSortAsc] = useState(false); // デフォルト降順（10F→1F）
  const [selectedBuildings, setSelectedBuildings] = useState(()=> sessionInfo?.selectedBuildings || []); // 選択中の建物
  // 機器種別（機器リストの「分類」列）
  const categoryKey = devColumns.find(k=>/^(分類|category|class)$/i.test(k.trim()))||null;
  const allCategories = categoryKey
    ? [...new Set(devList.map(d=>d._raw?.[categoryKey]).map(v=>String(v||"").trim()).filter(Boolean))]
    : [];
  const hasCategoryPanel = allCategories.length>0;
  const [selectedType, setSelectedType] = useState(()=>
    sessionInfo?.selectedType || allCategories.find(c=>/室内機/.test(c)) || null
  );
  // 「室内機」が選択されている場合のみ点検エリア確認を表示（分類列が無い場合は従来通り常時表示）
  const showAccessCheck = !hasCategoryPanel || (!!selectedType && /室内機/.test(selectedType));
  // 点検エリア確認のスキップ（ONにすると部屋一覧を非表示にする）
  const [skipAccessCheck, setSkipAccessCheck] = useState(!!sessionInfo?.skipAccessCheck);
  const isOutdoorSelected = !!selectedType && /室外機/.test(selectedType);
  // 点検エリア確認の機器リストは「室内機」のみを対象にする（分類列が無い場合は従来通り全件）
  const getCategory = d => categoryKey && d._raw ? String(d._raw[categoryKey]||"").trim() : "";
  const indoorDevList = categoryKey ? devList.filter(d=>/室内機/.test(getCategory(d))) : devList;
  // 対象階（前回引き継ぎ、なければ全階選択）
  const allFloors = [...new Set(devList.map(d=>String(d.floor||"").trim()).filter(s=>s.length>0))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:"base"}));
  const visibleFloors = allFloors;
  const [targetFloors, setTargetFloors] = useState(()=>
    sessionInfo?.targetFloors || []
  );

  // 建物列を特定してallBuildings・建物→階マップを生成
  const buildingKey = devColumns.find(k=>/建物|building|棟|ビル/i.test(k))||null;
  // _rawから階を取得するヘルパー（d.floorが空の場合に備える）
  const floorKey4Room = devColumns.find(k=>/^(階|floor|フロア|階数|階層)$/i.test(k.trim()))||null;
  const getFloor = d => {
    if(d.floor) return d.floor;
    if(floorKey4Room&&d._raw) return String(d._raw[floorKey4Room]||"").trim();
    return "";
  };
  const allBuildings = buildingKey
    ? [...new Set(devList.map(d=>d._raw?.[buildingKey]).filter(Boolean))]
    : [];
  // 建物ごとの階リスト（_rawから取得）
  const buildingFloorMap = buildingKey ? allBuildings.reduce((acc,b)=>{
    acc[b]=[...new Set(devList
      .filter(d=>d._raw?.[buildingKey]===b)
      .map(d=>getFloor(d))
      .filter(Boolean))];
    return acc;
  },{}) : {};
  // 建物選択に応じて表示する階を絞り込む
  const floorsInSelectedBuildings = selectedBuildings.length>0
    ? [...new Set(selectedBuildings.flatMap(b=>buildingFloorMap[b]||[]))]
    : visibleFloors;
  const sortedFloors = floorSortAsc ? [...floorsInSelectedBuildings] : [...floorsInSelectedBuildings].reverse();
  // 部屋リストも建物選択で絞り込む（室内機のみのindoorDevListが元）
  const filteredDevList = selectedBuildings.length>0&&buildingKey
    ? indoorDevList.filter(d=>selectedBuildings.includes(d._raw?.[buildingKey]))
    : indoorDevList;
  const rooms = [...new Map(filteredDevList.map(d=>{
    const fl=getFloor(d);
    return [fl+"__"+d.room,{floor:fl,room:d.room}];
  }).filter(([k,v])=>v.floor||v.room)).values()]
    .sort((a,b)=>a.floor.localeCompare(b.floor,undefined,{numeric:true})||a.room.localeCompare(b.room));

  const key = (floor, room) => floor+"__"+room;
  const setAcc = (k, val) => setAccess(p=>({...p,[k]:{...p[k],...val}}));
  const getAcc = (k) => access[k]?.status || "OK";
  const getMemo = (k) => access[k]?.memo || "";

  const toggleFloor = fl => setTargetFloors(p=>
    p.includes(fl) ? p.filter(f=>f!==fl) : [...p, fl]
  );
  // 対象階の部屋のみでNG件数カウント
  const targetRooms = targetFloors.length>0 ? rooms.filter(r=>targetFloors.includes(r.floor)) : rooms;
  const ngCount = targetRooms.filter(r=>getAcc(key(r.floor,r.room))==="NG").length;
  const undoneCount = indoorDevList.filter(d=>{
    if(targetFloors&&targetFloors.length>0&&!targetFloors.includes(d.floor)) return false;
    return !records.some(r=>r.managementNo===d.managementNo&&r.unitNo===d.unitNo&&Object.values(r.values).some(v=>v!==""));
  }).length;
  const canStart = !!date && !!inspector && (visibleFloors.length===0 || targetFloors.length>0);

  return (
    <div style={{flex:1,overflowY:"auto",scrollbarGutter:"stable",padding:"10px 12px",display:"flex",flexDirection:"column",gap:10,maxWidth:680,margin:"0 auto",width:"100%"}}>

      {/* ヘッダー */}
      <div style={{background:"linear-gradient(135deg,"+C.navy+","+C.blue+")",borderRadius:14,padding:"10px 14px",color:C.white,boxShadow:"0 2px 10px rgba(27,58,107,0.25)"}}>
        <div style={{fontSize:15,fontWeight:800,marginBottom:2}}>🚪 点検エリア確認</div>
        <div style={{fontSize:11,color:"rgba(255,255,255,0.75)"}}>点検日・点検者・対象階・点検エリアを設定してから点検を開始してください</div>
      </div>

      {/* 点検日 */}
      <div style={{background:C.white,borderRadius:14,padding:"10px 12px",boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
        <div style={{fontSize:12,fontWeight:800,color:C.navy,marginBottom:6}}>📅 点検日</div>
        <input type="date" value={date} onChange={e=>setDate(e.target.value)}
          onClick={e=>{ try{ e.target.showPicker&&e.target.showPicker(); }catch(err){} }}
          onFocus={e=>{ try{ e.target.showPicker&&e.target.showPicker(); }catch(err){} }}
          style={{width:"100%",fontSize:14,fontWeight:700,padding:"7px 10px",border:"2px solid "+(date?C.blue:C.g200),borderRadius:8,outline:"none",color:C.g800,background:C.g50,cursor:"pointer"}}/>
      </div>

      {/* 点検者 */}
      <div style={{background:C.white,borderRadius:14,padding:"10px 12px",boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
        <div style={{fontSize:12,fontWeight:800,color:C.navy,marginBottom:6}}>👤 点検者</div>
        {inspList.length>0 ? (
          <select value={inspector} onChange={e=>setInspector(e.target.value)}
            style={{width:"100%",fontSize:14,fontWeight:700,padding:"7px 10px",
              border:"2px solid "+(inspector?C.blue:C.g200),borderRadius:8,outline:"none",
              color:inspector?C.g800:C.g400,background:C.g50,appearance:"auto",cursor:"pointer"}}>
            <option value="">— 点検者を選択 —</option>
            {inspList.map(name=>(<option key={name} value={name}>{name}</option>))}
          </select>
        ) : (
          <input type="text" value={inspector} onChange={e=>setInspector(e.target.value)}
            placeholder="点検者名を入力"
            style={{width:"100%",fontSize:13,padding:"7px 10px",border:"2px solid "+(inspector?C.blue:C.g200),borderRadius:8,outline:"none",background:C.g50}}/>
        )}
        {showInspector2 ? (
          <div style={{marginTop:8,display:"flex",gap:6,alignItems:"center"}}>
            {inspList.length>0 ? (
              <select value={inspector2} onChange={e=>setInspector2(e.target.value)}
                style={{flex:1,fontSize:14,fontWeight:700,padding:"7px 10px",
                  border:"2px solid "+(inspector2?C.blue:C.g200),borderRadius:8,outline:"none",
                  color:inspector2?C.g800:C.g400,background:C.g50,appearance:"auto",cursor:"pointer"}}>
                <option value="">— 二人目の点検者を選択 —</option>
                {inspList.filter(n=>n!==inspector).map(name=>(<option key={name} value={name}>{name}</option>))}
              </select>
            ) : (
              <input type="text" value={inspector2} onChange={e=>setInspector2(e.target.value)}
                placeholder="二人目の点検者名を入力"
                style={{flex:1,fontSize:13,padding:"7px 10px",border:"2px solid "+(inspector2?C.blue:C.g200),borderRadius:8,outline:"none",background:C.g50}}/>
            )}
            <button onClick={()=>{setInspector2("");setShowInspector2(false);}}
              style={{padding:"7px 10px",borderRadius:8,border:"1.5px solid "+C.g200,background:C.g50,color:C.g500,cursor:"pointer",fontSize:12,fontWeight:700}}>✕</button>
          </div>
        ) : (
          <button onClick={()=>setShowInspector2(true)}
            style={{marginTop:8,padding:"6px 12px",borderRadius:8,border:"1.5px dashed "+C.g300,background:"none",color:C.blue,cursor:"pointer",fontSize:12,fontWeight:700}}>
            ＋ 二人目の点検者を追加
          </button>
        )}
      </div>

      {/* 対象階の選択 */}
      {allFloors.length===0 ? (
        <div style={{background:"#FFF7ED",borderRadius:14,padding:"10px 12px",border:"1.5px solid #F59E0B",boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
          <div style={{fontSize:12,fontWeight:800,color:"#92400E",marginBottom:4}}>🏢 本日の対象エリア</div>
          <div style={{fontSize:11,color:"#92400E"}}>⚠️ 機器リストCSVが読み込まれていません。<br/>設定画面から機器リストCSVを読み込むと、エリアの選択・点検エリア確認が利用できます。</div>
        </div>
      ) : (
        <div style={{background:C.white,borderRadius:14,padding:"10px 12px",boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
            <div style={{fontSize:12,fontWeight:800,color:C.navy,flex:1}}>🏢 本日の対象エリア</div>
          </div>
          {/* 建物パネル（建物列がある場合のみ表示） */}
          {allBuildings.length>0&&(
            <div style={{marginBottom:8}}>
              <div style={{fontSize:11,fontWeight:700,color:C.g500,marginBottom:5}}>🏗️ 建物</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                {allBuildings.map(b=>{
                  const sel=selectedBuildings.includes(b);
                  return (
                    <button key={b}
                      onClick={()=>{
                        setSelectedBuildings(sel?[]:[b]);
                        setTargetFloors([]); // 建物変更時に階選択をリセット
                      }}
                      style={{padding:"7px 14px",borderRadius:8,border:"2px solid "+(sel?C.navy:C.g200),
                        background:sel?"linear-gradient(135deg,"+C.navy+",#374151)":C.white,
                        color:sel?C.white:C.g600,fontWeight:700,fontSize:13,cursor:"pointer",transition:"all 0.15s"}}>
                      {b}{sel?" ✓":""}
                    </button>
                  );
                })}
              </div>
              <div style={{marginTop:5,height:"1px",background:C.g200}}/>
            </div>
          )}
          {/* 階ボタン */}
          <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
            {/* すべての階ボタン */}
            {(()=>{
              const allSel = floorsInSelectedBuildings.length>0 && floorsInSelectedBuildings.every(f=>targetFloors.includes(f));
              return (
                <button onClick={()=>setTargetFloors(allSel?[]:floorsInSelectedBuildings)}
                  style={{padding:"7px 14px",borderRadius:8,border:"2px solid "+(allSel?C.teal:C.g200),
                    background:allSel?"linear-gradient(135deg,"+C.teal+",#0D9488)":C.white,
                    color:allSel?C.white:C.g600,fontWeight:700,fontSize:13,cursor:"pointer",
                    transition:"all 0.15s",minWidth:62,textAlign:"center"}}>
                  すべて{allSel?" ✓":""}
                </button>
              );
            })()}
            {sortedFloors.map(fl=>{
              const sel=targetFloors.includes(fl);
              return (
                <button key={fl} onClick={()=>toggleFloor(fl)}
                  style={{padding:"7px 14px",borderRadius:8,border:"2px solid "+(sel?C.blue:C.g200),
                    background:sel?"linear-gradient(135deg,"+C.navy+","+C.blue+")":C.white,
                    color:sel?C.white:C.g600,fontWeight:700,fontSize:13,cursor:"pointer",
                    transition:"all 0.15s",minWidth:50,textAlign:"center"}}>
                  {fl}{sel?" ✓":""}
                </button>
              );
            })}
          </div>
          {targetFloors.length===0&&(
            <div style={{marginTop:6,fontSize:11,color:C.red,fontWeight:700}}>⚠️ 対象エリアを1つ以上選択してください</div>
          )}
          {/* 未入力分ボタン */}
          {targetFloors.length>0&&indoorDevList.length>0&&(
            <div style={{marginTop:8,display:"flex",gap:6,alignItems:"center"}}>
              <button
                onClick={()=>setUndoneOnly(false)}
                style={{flex:1,padding:"7px",borderRadius:8,border:"2px solid "+(undoneOnly?C.g200:C.blue),
                  background:undoneOnly?C.white:"linear-gradient(135deg,"+C.navy+","+C.blue+")",
                  color:undoneOnly?C.g500:C.white,fontWeight:700,fontSize:12,cursor:"pointer",transition:"all 0.15s"}}>
                📋 全機器
              </button>
              <button
                onClick={()=>setUndoneOnly(true)}
                style={{flex:1,padding:"7px",borderRadius:8,border:"2px solid "+(undoneOnly?C.red:C.g200),
                  background:undoneOnly?"#FEF2F2":C.white,
                  color:undoneOnly?C.red:C.g500,fontWeight:700,fontSize:12,cursor:"pointer",transition:"all 0.15s"}}>
                ⏳ 未入力分 {undoneCount>0?"("+undoneCount+")":"(なし)"}
              </button>
            </div>
          )}
          {/* 階ソート */}
          {allFloors.length>1&&(
            <div style={{marginTop:8,display:"flex",justifyContent:"flex-end"}}>
              <button onClick={()=>setFloorSortAsc(p=>!p)}
                style={{padding:"3px 9px",borderRadius:6,border:"1.5px solid "+C.teal,background:C.teal+"18",color:C.teal,fontSize:10,fontWeight:700,cursor:"pointer"}}>
                表示順 {floorSortAsc?"▲ 昇順":"▼ 降順"}
              </button>
            </div>
          )}
        </div>
      )}

      {/* 点検エリア確認（対象階のみ） */}
      {targetRooms.length>0 && showAccessCheck && (
        <div style={{background:C.white,borderRadius:14,padding:"10px 12px",boxShadow:"0 1px 8px rgba(0,0,0,0.07)"}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
            <div style={{fontSize:12,fontWeight:800,color:C.navy,flex:1}}>🚪 入室可否チェック</div>
            <button onClick={()=>setSkipAccessCheck(p=>!p)}
              style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid "+(skipAccessCheck?C.teal:C.g200),
                background:skipAccessCheck?C.teal:C.white,color:skipAccessCheck?C.white:C.g500,
                fontSize:11,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>
              ⏭️ スキップ{skipAccessCheck?" ✓":""}
            </button>
          </div>

          {skipAccessCheck ? (
            <div style={{padding:"8px 10px",background:C.teal+"10",border:"1.5px solid "+C.teal+"40",borderRadius:8,fontSize:11,color:C.teal,fontWeight:700}}>
              ⏭️ 入室可否チェックをスキップしました（部屋一覧は表示されません）
            </div>
          ) : (<>

          {ngCount>0&&(
            <div style={{marginBottom:6,padding:"5px 10px",background:"#FEF2F2",border:"1.5px solid "+C.red,borderRadius:7,fontSize:11,fontWeight:700,color:C.red}}>
              ⚠️ NG {ngCount}部屋 — 点検フォームでグレーアウトされます
            </div>
          )}

          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {sortedFloors.filter(fl=>targetFloors.includes(fl)).map(fl=>{
              const flRooms=rooms.filter(r=>r.floor===fl);
              const flNg=flRooms.filter(r=>getAcc(key(r.floor,r.room))==="NG").length;
              return (
                <div key={fl}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                    <div style={{fontSize:11,fontWeight:800,color:C.g600,flex:1,letterSpacing:"0.05em"}}>{fl}</div>
                    {flNg>0&&<span style={{fontSize:9,fontWeight:700,color:C.red}}>NG {flNg}</span>}
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:3}}>
                    {flRooms.map(r=>{
                      const k=key(r.floor,r.room);
                      const acc=getAcc(k);
                      const isNG=acc==="NG";
                      const memo=getMemo(k);
                      const open=memoOpen[k];
                      const ngType=access[k]?.ngType||"today";
                      return (
                        <div key={k} style={{borderRadius:8,border:"1.5px solid "+(isNG?C.red:C.g200),overflow:"hidden",transition:"border-color 0.15s"}}>
                          {/* メイン行：部屋名 ＋ NG理由（1行） ＋ OK/NGボタン */}
                          <div style={{display:"flex",alignItems:"center",gap:7,padding:"7px 10px",background:isNG?"#FFF0F0":C.g50,cursor:isNG&&!open?"pointer":"default"}}
                            onClick={isNG&&!open?()=>setMemoOpen(p=>({...p,[k]:true})):undefined}>
                            {/* 表示列に従って情報を表示 */}
                            <div style={{flex:1,display:"flex",gap:8,alignItems:"center",overflow:"hidden",minWidth:0,flexWrap:"wrap"}}>
                              {(()=>{
                                const cols=devVisibleCols&&devVisibleCols.length>0?devVisibleCols:(devColumns&&devColumns.length>0?devColumns:null);
                                // devListからこの部屋の機器データを取得（_rawを持つ）
                                const devItem=indoorDevList.find(d=>d.floor===r.floor&&d.room===r.room);
                                if(cols){
                                  return cols.map(col=>{
                                    // _rawがあればそこから、なければfloor/room/managementNo/unitNoにフォールバック
                                    const val=devItem?._raw?.[col]
                                      ||(col==="階"||col.toLowerCase()==="floor"?r.floor
                                        :col==="部屋名"||col.toLowerCase()==="room"?r.room
                                        :col==="管理番号"||col.toLowerCase()==="managementno"?devItem?.managementNo
                                        :col==="機器番号"||col.toLowerCase()==="unitno"?devItem?.unitNo:"");
                                    return val?<span key={col} style={{fontSize:12,fontWeight:col==="部屋名"||col.toLowerCase()==="room"?700:400,color:isNG?C.red:C.g800,whiteSpace:"nowrap"}}>{val}</span>:null;
                                  });
                                }
                                return <span style={{fontSize:12,fontWeight:700,color:isNG?C.red:C.g800}}>{r.room}</span>;
                              })()}
                              {isNG&&!open&&(
                                <span style={{fontSize:10,color:C.red,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flexShrink:0}}>
                                  🚫 {[["today","本日のみNG"],["am","午前のみNG"],["pm","午後のみNG"],["discharge","退院までNG"]].find(([v])=>v===ngType)?.[1]||"本日のみNG"}
                                  {memo&&"　"+memo}
                                </span>
                              )}
                            </div>
                            {/* OK/NG トグル */}
                            <div style={{display:"flex",gap:3}} onClick={e=>e.stopPropagation()}>
                              {["OK","NG"].map(s=>(
                                <button key={s}
                                  onClick={()=>{
                                    setAcc(k,{status:s});
                                    setMemoOpen(p=>({...p,[k]:s==="NG"}));
                                  }}
                                  style={{padding:"5px 12px",borderRadius:6,
                                    border:"2px solid "+(acc===s?(s==="OK"?C.green:C.red):C.g200),
                                    background:acc===s?(s==="OK"?C.green:C.red):C.white,
                                    color:acc===s?C.white:C.g500,
                                    fontWeight:800,fontSize:11,cursor:"pointer",transition:"all 0.12s"}}>
                                  {s}
                                </button>
                              ))}
                            </div>
                          </div>
                          {/* NG詳細パネル（open時のみ展開） */}
                          {isNG&&open&&(
                            <div style={{padding:"8px 10px",borderTop:"1.5px solid #FECACA",background:"#FEF2F2",display:"flex",flexDirection:"column",gap:6}}>
                              {/* NGタイプ選択：タップで選択＋パネルを閉じる */}
                              <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                                {[["today","本日のみNG"],["am","午前のみNG"],["pm","午後のみNG"],["discharge","退院までNG"]].map(([v,label])=>(
                                  <button key={v} onClick={()=>{
                                    setAcc(k,{ngType:v});
                                    setMemoOpen(p=>({...p,[k]:false}));
                                  }}
                                    style={{padding:"4px 10px",borderRadius:6,
                                      border:"1.5px solid "+(ngType===v?C.red:C.g300),
                                      background:ngType===v?C.red:C.white,
                                      color:ngType===v?C.white:C.g600,
                                      fontSize:11,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap",transition:"all 0.1s"}}>
                                    {label}
                                  </button>
                                ))}
                              </div>
                              {/* 備考（3倍高さのtextarea） */}
                              <textarea value={memo} onChange={e=>setAcc(k,{memo:e.target.value})}
                                placeholder="備考（入室不可理由など）"
                                inputMode="text"
                                style={{width:"100%",fontSize:11,padding:"6px 9px",border:"1.5px solid #FECACA",borderRadius:7,outline:"none",background:C.white,boxSizing:"border-box",resize:"none",minHeight:66,lineHeight:1.5,fontFamily:"inherit"}}/>
                              <button onClick={()=>setMemoOpen(p=>({...p,[k]:false}))}
                                style={{alignSelf:"flex-end",padding:"4px 12px",borderRadius:6,border:"1.5px solid "+C.g300,background:C.white,color:C.g500,fontSize:10,fontWeight:700,cursor:"pointer"}}>
                                閉じる
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
          </>)}
        </div>
      )}

      {/* 開始ボタン */}
      <button
        disabled={!canStart}
        onClick={()=>{
          const combinedInspector = inspector2 ? inspector+"・"+inspector2 : inspector;
          const info={date, inspector:combinedInspector, targetFloors, roomAccess:access, selectedBuildings, selectedType, skipAccessCheck};
          try{ localStorage.setItem("acSessionInfo", JSON.stringify(info)); }catch(e){}
          if(isOutdoorSelected){ onSelectOutdoor(info); } else { onStart(info); }
        }}
        style={{padding:"13px",borderRadius:12,border:"none",cursor:canStart?"pointer":"not-allowed",
          background:canStart?"linear-gradient(135deg,"+(isOutdoorSelected?C.teal:C.navy)+","+(isOutdoorSelected?"#0D9488":C.blue)+")":"#CBD5E1",
          color:C.white,fontWeight:800,fontSize:14,
          boxShadow:canStart?"0 3px 12px rgba(37,99,176,0.35)":"none",
          opacity:canStart?1:0.7,transition:"all 0.2s",flexShrink:0}}>
        {!date||!inspector?"点検日・点検者を入力してください":
         visibleFloors.length>0&&targetFloors.length===0?"対象エリアを選択してください":
         visibleFloors.length>0?(isOutdoorSelected?"🏭 室外機点検を開始する（":"✅ 点検を開始する（")+targetFloors.slice().sort().join("・")+")":
         (isOutdoorSelected?"🏭 室外機点検を開始する":"✅ 点検を開始する")}
      </button>

      <div style={{height:4}}/>
    </div>
  );
}


function Step1View({
  form, setInfo, inspList, devList, devSearch, setDevSearch,
  s1DateDone, s1InspDone, s1DevDone, step1Valid, records,
  s1Focus, setS1Focus, goToStep2, handleStep1TmpSave, setForm,
  editIdx, lastInsp, lastDate, setStep, setView,
  sessionInfo, undoneOnly, setUndoneOnly,
  devColumns, devVisibleCols, inspectionMode,
  checkFields, setCheck, handleSave, complete, missing,
  visIn, visOut, visFields, activeCode, numDisp, limits,
  onPress, onConfirm, onRowClick, moveActive, rowRefs, listRef,
  ALL_FIELDS, vis, isAbn, focusSeq, isCheckCode, setCheckAndAdvance, onSwitchMode,
}) {
  const allFloors = [...new Set(devList.map(d=>d.floor).filter(Boolean))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:"base"}));
  // allFloorsが空の場合：_rawから階候補列を探して補完
  const allFloorsResolved = allFloors.length>0 ? allFloors : (()=>{
    const floorKey = devColumns.find(k=>/^(階|floor|フロア|階数|F|階層)$/i.test(k.trim()));
    if(!floorKey) return [];
    return [...new Set(devList.map(d=>String(d._raw?.[floorKey]||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:"base"}));
  })();
  const roomAccess = sessionInfo?.roomAccess || {};
  const targetFloors = sessionInfo?.targetFloors || null;
  const displayFloors = (targetFloors&&targetFloors.length>0) ? targetFloors : allFloorsResolved;
  // 非対象階 or 入室NGならグレーアウト
  const isNG = (floor, room) => {
    if(targetFloors && targetFloors.length>0 && !targetFloors.includes(floor)) return true;
    return (roomAccess[floor+"__"+room]?.status || "OK") === "NG";
  };
  const isNGReason = (floor, room) => {
    if(targetFloors && targetFloors.length>0 && !targetFloors.includes(floor)) return "対象外";
    return (roomAccess[floor+"__"+room]?.status || "OK")==="NG" ? "入室NG" : null;
  };
  // 機器選択リストを、選択中のモード（室内機／室外機）かつ点検エリア確認で選択した建物・階のみに絞り込む
  const categoryKey = devColumns.find(k=>/^(分類|category|class)$/i.test(k.trim()))||null;
  const buildingKey = devColumns.find(k=>/建物|building|棟|ビル/i.test(k))||null;
  const remarksKey = devColumns.find(k=>/^(備考|remarks?|memo|note)$/i.test(k.trim()))||null;
  const selectedBuildings = sessionInfo?.selectedBuildings || [];
  const categoryPattern = inspectionMode==="outdoor" ? /室外機/ : /室内機/;
  const baseDevList = devList.filter(d=>{
    if(categoryKey && !categoryPattern.test(String(d._raw?.[categoryKey]||"").trim())) return false;
    if(buildingKey && selectedBuildings.length>0 && !selectedBuildings.includes(d._raw?.[buildingKey])) return false;
    if(targetFloors && targetFloors.length>0 && !targetFloors.includes(d.floor)) return false;
    return true;
  });
  // 室外機チェック項目（画面下部リスト＋右側テンキーパネルの両方で使う）
  const outFields = (checkFields||[]).filter(f=>f.group==="check_out");
  const outTotal = outFields.length;
  const outFilled = outFields.filter(f=>(form.checks?.[f.code]||"")!=="").length;
  const outPct = outTotal>0 ? Math.round(outFilled/outTotal*100) : 0;
  const isCheckFocus = outFields.some(f=>f.code===s1Focus);
  const focusedCheckField = outFields.find(f=>f.code===s1Focus);
  const outListRef = useRef();
  const outRowRefs = useRef({});
  // 入力欄が常に画面内に収まるよう、フォーカス項目が変わるたびにスクロール位置を調整する
  const scrollToCheckRow = (code,smooth=true) => {
    requestAnimationFrame(()=>requestAnimationFrame(()=>{
      const el=outRowRefs.current[code]; const ct=outListRef.current;
      if(!el||!ct) return;
      let top=0,node=el;
      while(node&&node!==ct){top+=node.offsetTop;node=node.offsetParent;}
      ct.scrollTo({top:Math.max(0,top-ct.clientHeight/2+el.offsetHeight/2),behavior:smooth?"smooth":"instant"});
    }));
  };
  const focusCheckRow = code => { setS1Focus(code); if(code)scrollToCheckRow(code); };
  // 点検日・点検者・建物・階・機器選択が揃ったら、Step2Viewと同じ1行バーに折りたたむ（修正ボタンで再展開）
  const [summaryExpanded, setSummaryExpanded] = useState(false);
  // 機器選択の検索窓：この画面（機器未選択の検索・一覧状態）になったら自動的にフォーカスする（室内機・室外機とも）
  const devSearchRef = useRef(null);
  useEffect(()=>{
    if(!s1DevDone && devSearchRef.current){
      devSearchRef.current.focus();
    }
  },[s1DevDone, inspectionMode]);
  // 点検前状態：一次保存/スキップを押すと折りたたむ（測定データ入力はブロックしない。機器を切り替えたら自動的に展開状態に戻る）
  const [preStateConfirmedFor, setPreStateConfirmedFor] = useState(null);
  const deviceKey = (form.managementNo||"")+"|"+(form.unitNo||"");
  const preStateOpen = preStateConfirmedFor !== deviceKey;
  // 機器リストのソート：点検前状態・データ入力の済/未でスライドボタン切替（ONのとき未入力を先頭に表示）
  const [preStateSortDir, setPreStateSortDir] = useState(null); // null|"done"|"undone"
  const [dataSortDir, setDataSortDir] = useState(null); // null|"done"|"undone"
  const devPreStateDone = d => records.some(r=>r.managementNo===d.managementNo&&r.unitNo===d.unitNo&&(r.preOperation||r.preMode||r.preWind||r.preSetTemp));
  const devDataDone = d => records.some(r=>r.managementNo===d.managementNo&&r.unitNo===d.unitNo&&(inspectionMode==="outdoor"?Object.values(r.checks||{}).some(v=>v!==""):Object.values(r.values).some(v=>v!=="")));
  // 機器選択時の共通処理：機器を選ぶとすぐ下にデータ入力（測定値・チェック）が表示されるため、最初の項目に自動フォーカスする
  const selectDevice = dev => {
    setForm(p=>({...p,floor:dev.floor,room:dev.room,managementNo:dev.managementNo,unitNo:dev.unitNo,
      preOperation:"",preMode:"",preWind:"",preSetTemp:""}));
    setDevSearch("");
    setSummaryExpanded(false);
    if(inspectionMode==="outdoor" && outFields.length>0){ focusCheckRow(outFields[0].code); }
    else if(inspectionMode==="indoor" && focusSeq && focusSeq.length>0){ onRowClick && onRowClick(focusSeq[0]); }
    else { setS1Focus(null); }
  };
  // 建物・階などの絞り込みで候補が1件だけになっている場合は、タップしなくても自動的に選択する
  // （画面を開き直すたびに同じ機器を選び直す手間を無くすため）
  useEffect(()=>{
    if(!s1DevDone && !devSearch && s1InspDone && baseDevList.length===1){
      selectDevice(baseDevList[0]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[baseDevList.length, s1DevDone, devSearch, s1InspDone]);
  // チェックの値を設定し、続けて次のチェック項目へフォーカスを進める（テンキーパネル・一覧の○×ボタン共通）
  // ※反転（同じ値を押すとクリア）はしない。何度押しても選んだ値のまま（クリアはCLRボタンで行う）
  const advanceCheck = (code,v) => {
    setCheck && setCheck(code, v);
    const idx = outFields.findIndex(f=>f.code===code);
    const next = outFields[idx+1];
    if(next){ setS1Focus(next.code); scrollToCheckRow(next.code); }
    else { setS1Focus(null); }
  };
  const topSummary = (!step1Valid || summaryExpanded) ? (<>
          {/* ── 1行目：点検日・点検者／2行目：建物・階（それぞれ修正ボタン付き） ── */}
          {(()=>{
            const bKey=devColumns.find(k=>/建物|building|棟|ビル/i.test(k));
            const sessionBuildings=sessionInfo?.selectedBuildings||[];
            const showBuildings=bKey?(sessionBuildings.length>0?sessionBuildings:[...new Set(devList.map(d=>d._raw?.[bKey]).filter(Boolean))]):[];
            const floorsSorted=[...displayFloors].sort((a,b)=>b.localeCompare(a,undefined,{numeric:true}));
            return (
              <>
              <div style={{background:C.white,borderRadius:12,padding:"10px 12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                <div style={{display:"flex",gap:16,rowGap:8,flexWrap:"wrap",alignItems:"center"}}>
                  <div style={{display:"flex",alignItems:"baseline",gap:6}}>
                    <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ 点検日</span>
                    <span style={{fontSize:14,fontWeight:700,color:C.navy,fontFamily:"monospace"}}>{form.inspectionDate||"—"}</span>
                  </div>
                  <div style={{display:"flex",alignItems:"baseline",gap:6}}>
                    <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ 点検者</span>
                    <span style={{fontSize:14,fontWeight:700,color:C.navy}}>{form.inspector||"—"}</span>
                  </div>
                  <button onMouseDown={e=>e.preventDefault()} onClick={()=>setStep(0)}
                    style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"1.5px solid "+C.g300,background:C.g50,color:C.g600,cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>修正</button>
                </div>
              </div>
              {(showBuildings.length>0||allFloorsResolved.length>0)&&(
                <div style={{background:C.white,borderRadius:12,padding:"10px 12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginTop:6}}>
                  <div style={{display:"flex",gap:16,rowGap:8,flexWrap:"wrap",alignItems:"center"}}>
                    {showBuildings.length>0&&(
                      <div style={{display:"flex",alignItems:"baseline",gap:6,flexWrap:"wrap"}}>
                        <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ 建物</span>
                        {showBuildings.map(b=>(
                          <span key={b} style={{fontSize:13,fontWeight:700,color:C.navy,background:C.navy+"10",padding:"2px 9px",borderRadius:6}}>{b}</span>
                        ))}
                      </div>
                    )}
                    {allFloorsResolved.length>0&&(
                      <div style={{display:"flex",alignItems:"baseline",gap:6,flexWrap:"wrap"}}>
                        <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ 階</span>
                        {floorsSorted.map(fl=>(
                          <span key={fl} style={{fontSize:13,fontWeight:700,color:C.blue,background:C.blue+"15",padding:"2px 9px",borderRadius:6}}>{fl}</span>
                        ))}
                      </div>
                    )}
                    <button onMouseDown={e=>e.preventDefault()} onClick={()=>setStep(0)}
                      style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"1.5px solid "+C.g300,background:C.g50,color:C.g600,cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>修正</button>
                  </div>
                </div>
              )}
              </>
            );
          })()}

          {/* ── 機器種別（室内機／室外機。第二ヘッダーから移設）── */}
          <div style={{display:"flex",background:C.g200,gap:"1px"}}>
            <div style={{width:110,flexShrink:0,background:C.green+"18",display:"flex",alignItems:"center",padding:"8px 12px"}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:18,height:18,borderRadius:9,background:C.green,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:C.white}}>✓</div>
                <span style={{fontSize:12,fontWeight:700,color:C.green}}>機器種別</span>
              </div>
            </div>
            <div style={{flex:1,background:C.white,padding:"7px 10px",display:"flex",alignItems:"center",gap:8}}>
              {[["indoor","🏠","室内機",C.blue],["outdoor","🏭","室外機",C.teal]].map(([mode,icon,label,color])=>{
                const sel=inspectionMode===mode;
                return (
                  <button key={mode} onMouseDown={e=>e.preventDefault()}
                    onClick={()=>onSwitchMode&&onSwitchMode(mode)}
                    style={{padding:"7px 16px",borderRadius:9,border:"2px solid "+(sel?color:C.g200),cursor:"pointer",fontWeight:700,fontSize:13,
                      background:sel?"linear-gradient(135deg,"+color+","+color+"CC)":C.white,
                      color:sel?C.white:C.g600,transition:"all 0.12s",whiteSpace:"nowrap"}}>
                    {icon} {label}{sel?" ✓":""}
                  </button>
                );
              })}
            </div>
          </div>

          {/* ── 機器選択 ── */}
          <div style={{display:"flex",background:C.g200,gap:"1px",opacity:s1InspDone?1:0.45}}>
            <div style={{width:110,flexShrink:0,background:s1DevDone?C.green+"18":C.g50,padding:"10px 12px"}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:18,height:18,borderRadius:9,background:s1DevDone?C.green:C.g300,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:C.white}}>{s1DevDone?"✓":"3"}</div>
                <span style={{fontSize:12,fontWeight:700,color:s1DevDone?C.green:C.g600}}>機器選択</span>
              </div>
            </div>
            <div style={{flex:1,background:C.white,padding:"8px 10px",display:"flex",flexDirection:"column",gap:5,pointerEvents:s1InspDone?"auto":"none"}}>
              {/* 対象エリアバッジは上の行に移動 */}
              {false&&null}
              {devList.length>0 ? (
                <>
                  <div style={{display:"flex",gap:8,alignItems:"center"}}>
                    <div style={{position:"relative",flex:1}}>
                      <input value={devSearch} ref={devSearchRef}
                        onChange={e=>{setDevSearch(e.target.value);}}
                        onFocus={()=>setS1Focus("devSearch")}
                        placeholder="管理番号・部屋名で検索…"
                        style={{width:"100%",padding:"7px 10px 7px 28px",borderRadius:7,fontSize:13,
                          border:"1.5px solid "+(s1Focus==="devSearch"?C.blue:C.g200),
                          background:C.inp,outline:"none",boxSizing:"border-box",fontFamily:"inherit"}}/>
                      <span style={{position:"absolute",left:8,top:"50%",transform:"translateY(-50%)",fontSize:12,color:C.blue,pointerEvents:"none"}}>🔍</span>
                      {devSearch&&<button onMouseDown={e=>e.preventDefault()} onClick={()=>setDevSearch("")} style={{position:"absolute",right:6,top:"50%",transform:"translateY(-50%)",border:"none",background:"none",cursor:"pointer",fontSize:14,color:C.g400,lineHeight:1}}>✕</button>}
                    </div>
                    {s1DevDone&&(
                      <button onMouseDown={e=>e.preventDefault()} onClick={()=>{setForm(p=>({...p,floor:"",room:"",managementNo:"",unitNo:""}));setDevSearch("");}}
                        style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid "+C.g200,background:C.white,color:C.g500,cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>変更</button>
                    )}
                  </div>
                  {s1DevDone&&!devSearch ? (
                    <div style={{padding:"8px 12px",borderRadius:9,background:C.g50,border:"1.5px solid "+C.g200,display:"flex",flexWrap:"wrap",rowGap:6,gap:14,alignItems:"center"}}>
                      {[["階",form.floor],["部屋名",form.room],["管理番号",form.managementNo],["機器番号",form.unitNo]].map(([k,v])=>(
                        <div key={k} style={{display:"flex",alignItems:"baseline",gap:6}}>
                          <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ {k}</span>
                          <span style={{fontSize:13,fontWeight:700,color:C.navy}}>{v||"—"}</span>
                        </div>
                      ))}
                      {(()=>{
                        const selDev=devList.find(d=>d.managementNo===form.managementNo&&d.unitNo===form.unitNo);
                        const remarksVal=remarksKey?selDev?._raw?.[remarksKey]:"";
                        return remarksVal ? (
                          <span style={{flexBasis:"100%",fontSize:11,color:C.g500,whiteSpace:"pre-wrap",wordBreak:"break-word"}}>📝 {remarksVal}</span>
                        ) : null;
                      })()}
                    </div>
                  ) : (
                    <>
                    {/* 並び替えボタン：点検前状態・データ入力それぞれ「入力済」「未入力」を選ぶと、その状態を先頭に表示 */}
                    <div style={{display:"flex",gap:10,flexWrap:"wrap",flexShrink:0,alignItems:"center"}}>
                      {[
                        {label:"点検前状態",dir:preStateSortDir,set:setPreStateSortDir},
                        {label:"データ入力",dir:dataSortDir,set:setDataSortDir},
                      ].map(({label,dir,set})=>(
                        <div key={label} style={{display:"flex",alignItems:"center",gap:4}}>
                          <span style={{fontSize:11,fontWeight:700,color:C.g500,whiteSpace:"nowrap"}}>{label}：</span>
                          {[["done","入力済"],["undone","未入力"]].map(([v,txt])=>(
                            <button key={v} onMouseDown={e=>e.preventDefault()} onClick={()=>set(p=>p===v?null:v)}
                              style={{padding:"4px 10px",borderRadius:16,border:"1.5px solid "+(dir===v?C.blue:C.g200),
                                background:dir===v?C.blue:C.white,color:dir===v?C.white:C.g500,cursor:"pointer",fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                              {txt}
                            </button>
                          ))}
                        </div>
                      ))}
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:3,overflowY:"auto",scrollbarGutter:"stable",maxHeight:420,minHeight:80,flexShrink:0}}>
                      {(()=>{
                        // 対象階の表示順（sortedFloors=降順 or 昇順）を尊重
                        const floorOrder = targetFloors && targetFloors.length>0
                          ? targetFloors  // セッション画面の選択順（sortedFloorsと同じ降順）
                          : null;
                        const sorted = [...baseDevList].sort((a,b)=>{
                          if(dataSortDir){
                            const da=devDataDone(a)?1:0, db=devDataDone(b)?1:0;
                            const cmp = dataSortDir==="undone" ? da-db : db-da;
                            if(da!==db) return cmp;
                          }
                          if(preStateSortDir){
                            const pa=devPreStateDone(a)?1:0, pb=devPreStateDone(b)?1:0;
                            const cmp = preStateSortDir==="undone" ? pa-pb : pb-pa;
                            if(pa!==pb) return cmp;
                          }
                          if(!floorOrder) return a.floor.localeCompare(b.floor)||a.room.localeCompare(b.room);
                          const ai=floorOrder.indexOf(a.floor);
                          const bi=floorOrder.indexOf(b.floor);
                          // 対象階内は降順（sortedFloorsと同じ並び）
                          const ar=ai<0?999:ai;
                          const br=bi<0?999:bi;
                          if(ar!==br) return ar-br;
                          // 同じ階内は部屋名順
                          return a.room.localeCompare(b.room);
                        });
                        const filtered = sorted.filter(d=>{
                          const n=s=>s.replace(/-/g,"").toLowerCase();
                          if(devSearch){
                            const q=n(devSearch);
                            // devVisibleColsの全列＋基本4フィールドでヒット判定
                            const cols=devVisibleCols&&devVisibleCols.length>0?devVisibleCols:[];
                            const rawHit=cols.some(col=>{
                              const v=d._raw?d._raw[col]:"";
                              return n(String(v)).includes(q);
                            });
                            const basicHit=n(d.managementNo).includes(q)||n(d.unitNo).includes(q)||n(d.room).includes(q)||n(d.floor).includes(q);
                            if(!rawHit&&!basicHit) return false;
                          }
                          if(undoneOnly){
                            const hasMeas=records.some(r=>r.managementNo===d.managementNo&&r.unitNo===d.unitNo&&(inspectionMode==="outdoor"?Object.values(r.checks||{}).some(v=>v!==""):Object.values(r.values).some(v=>v!=="")));
                            if(hasMeas) return false;
                          }
                          return true;
                        });
                        if(filtered.length===0) return <div style={{fontSize:12,color:C.g400,padding:"6px 4px"}}>該当する機器がありません</div>;
                        return filtered.map((dev,i)=>{
                          const sel=form.managementNo===dev.managementNo&&form.unitNo===dev.unitNo;
                          const hasMeas=records.some(r=>r.managementNo===dev.managementNo&&r.unitNo===dev.unitNo&&(inspectionMode==="outdoor"?Object.values(r.checks||{}).some(v=>v!==""):Object.values(r.values).some(v=>v!=="")));
                          const ngRoom=isNG(dev.floor, dev.room);
                          const ngReason=isNGReason(dev.floor, dev.room);
                          if(ngRoom) return (
                            <div key={i} style={{padding:"7px 10px",borderRadius:7,border:"1.5px solid "+C.g200,background:C.g100,
                              display:"flex",gap:6,alignItems:"center",flexShrink:0,opacity:0.45,userSelect:"none",flexWrap:"wrap"}}>
                              {(()=>{
                                const cols=devVisibleCols&&devVisibleCols.length>0?devVisibleCols:(devColumns&&devColumns.length>0?devColumns:null);
                                const remarksVal=remarksKey?dev._raw?.[remarksKey]:"";
                                const remarksStyle={flexBasis:"100%",fontSize:10,color:C.g500,whiteSpace:"pre-wrap",wordBreak:"break-word"};
                                if(cols){
                                  return <>
                                    {cols.map(col=>{
                                      const val=dev._raw?dev._raw[col]:(col==="階"?dev.floor:col==="部屋名"?dev.room:col==="管理番号"?dev.managementNo:col==="機器番号"?dev.unitNo:"");
                                      return val?<span key={col} style={{fontSize:13,fontWeight:700,color:C.g400,textDecoration:"line-through"}}>{val}</span>:null;
                                    })}
                                    <span style={{fontSize:10,fontWeight:700,color:ngReason==="対象外"?C.g400:C.red,background:ngReason==="対象外"?C.g200:"#FEF2F2",padding:"1px 6px",borderRadius:4,whiteSpace:"nowrap"}}>{ngReason}</span>
                                    {remarksVal&&<span style={remarksStyle}>📝 {remarksVal}</span>}
                                  </>;
                                }
                                return <>
                                  <span style={{fontSize:13,fontWeight:700,flex:1,color:C.g400,textDecoration:"line-through"}}>{dev.floor}　{dev.room}</span>
                                  <span style={{fontSize:12,fontFamily:"monospace",opacity:0.6,color:C.g400}}>{dev.managementNo} / {dev.unitNo}</span>
                                  <span style={{fontSize:10,fontWeight:700,color:ngReason==="対象外"?C.g400:C.red,background:ngReason==="対象外"?C.g200:"#FEF2F2",padding:"1px 6px",borderRadius:4,whiteSpace:"nowrap"}}>{ngReason}</span>
                                  {remarksVal&&<span style={remarksStyle}>📝 {remarksVal}</span>}
                                </>;
                              })()}
                            </div>
                          );
                          if(hasMeas && !sel) return (
                            <button key={i}
                              onMouseDown={e=>e.preventDefault()}
                              onClick={()=>selectDevice(dev)}
                              style={{padding:"7px 10px",borderRadius:7,border:"2px solid "+C.green,cursor:"pointer",textAlign:"left",
                                background:C.green+"10",color:C.g800,display:"flex",gap:6,alignItems:"center",flexShrink:0,transition:"all 0.1s",flexWrap:"wrap"}}>
                              {(()=>{
                                const cols=devVisibleCols&&devVisibleCols.length>0?devVisibleCols:(devColumns&&devColumns.length>0?devColumns:null);
                                const remarksVal=remarksKey?dev._raw?.[remarksKey]:"";
                                const showExtraRemarks=remarksVal&&!(cols&&cols.includes(remarksKey));
                                const remarksStyle={flexBasis:"100%",fontSize:10,fontWeight:400,color:C.g500,textAlign:"left",whiteSpace:"pre-wrap",wordBreak:"break-word"};
                                if(cols){
                                  return <>
                                    {cols.map(col=>{
                                      const val=dev._raw?dev._raw[col]:(col==="階"?dev.floor:col==="部屋名"?dev.room:col==="管理番号"?dev.managementNo:col==="機器番号"?dev.unitNo:"");
                                      return val?<span key={col} style={{fontSize:13,fontWeight:700}}>{val}</span>:null;
                                    })}
                                    <span style={{fontSize:10,fontWeight:700,color:C.green,background:C.green+"20",padding:"1px 6px",borderRadius:4,whiteSpace:"nowrap"}}>入力済</span>
                                    {showExtraRemarks&&(
                                      <span style={remarksStyle}>📝 {remarksVal}</span>
                                    )}
                                  </>;
                                }
                                return <>
                                  <span style={{fontSize:13,fontWeight:700,flex:1}}>{dev.floor}　{dev.room}</span>
                                  <span style={{fontSize:12,fontFamily:"monospace",opacity:0.8}}>{dev.managementNo} / {dev.unitNo}</span>
                                  <span style={{fontSize:10,fontWeight:700,color:C.green,background:C.green+"20",padding:"1px 6px",borderRadius:4,whiteSpace:"nowrap"}}>入力済</span>
                                  {showExtraRemarks&&(
                                    <span style={remarksStyle}>📝 {remarksVal}</span>
                                  )}
                                </>;
                              })()}
                            </button>
                          );
                          return (
                            <button key={i}
                              onMouseDown={e=>e.preventDefault()}
                              onClick={()=>selectDevice(dev)}
                              style={{padding:"7px 10px",borderRadius:7,border:"2px solid "+(sel?C.blue:C.g200),cursor:"pointer",textAlign:"left",
                                background:sel?"linear-gradient(135deg,"+C.navy+","+C.blue+")":C.white,
                                color:sel?C.white:C.g800,display:"flex",gap:6,alignItems:"center",flexShrink:0,transition:"all 0.1s",flexWrap:"wrap"}}>
                              {(()=>{
                                const cols=devVisibleCols&&devVisibleCols.length>0?devVisibleCols:(devColumns&&devColumns.length>0?devColumns:null);
                                const remarksVal=remarksKey?dev._raw?.[remarksKey]:"";
                                const showExtraRemarks=remarksVal&&!(cols&&cols.includes(remarksKey));
                                const remarksStyle={flexBasis:"100%",fontSize:10,fontWeight:400,color:sel?"rgba(255,255,255,0.85)":C.g500,textAlign:"left",whiteSpace:"pre-wrap",wordBreak:"break-word"};
                                if(cols){
                                  return <>
                                    {cols.map(col=>{
                                      const val=dev._raw?dev._raw[col]:(col==="階"?dev.floor:col==="部屋名"?dev.room:col==="管理番号"?dev.managementNo:col==="機器番号"?dev.unitNo:"");
                                      return val?<span key={col} style={{fontSize:13,fontWeight:700}}>{val}</span>:null;
                                    })}
                                    {showExtraRemarks&&(
                                      <span style={remarksStyle}>📝 {remarksVal}</span>
                                    )}
                                  </>;
                                }
                                return <>
                                  <span style={{fontSize:13,fontWeight:700,flex:1}}>{dev.floor}　{dev.room}</span>
                                  <span style={{fontSize:12,fontFamily:"monospace",opacity:0.8}}>{dev.managementNo} / {dev.unitNo}</span>
                                  {showExtraRemarks&&(
                                    <span style={remarksStyle}>📝 {remarksVal}</span>
                                  )}
                                </>;
                              })()}
                            </button>
                          );
                        });
                      })()}
                    </div>
                    </>
                  )}
                </>
              ) : (
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 0.5fr 1.4fr",gap:5}}>
                  {[{k:"managementNo",l:"管理番号",p:"A-001"},{k:"unitNo",l:"機器番号",p:"IDU-001"},{k:"floor",l:"階",p:"1F"},{k:"room",l:"部屋名",p:"事務室"}].map(f=>(
                    <div key={f.k}>
                      <div style={{fontSize:10,fontWeight:700,color:C.g500,marginBottom:3}}>{f.l}</div>
                      <input value={form[f.k]} placeholder={f.p} onChange={e=>setInfo(f.k,e.target.value)}
                        style={{width:"100%",padding:"7px 9px",borderRadius:7,fontSize:13,
                          border:"1.5px solid "+(form[f.k]?C.green:C.g200),
                          background:C.inp,outline:"none",boxSizing:"border-box",fontFamily:"inherit"}}/>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
          </>) : (
            /* ── 折りたたみ済み：1行目=点検日・点検者／2行目=建物・階／3行目=機器選択（部屋名・管理番号・機器番号）。それぞれに修正ボタン ── */
            <>
            {(()=>{
              const bKey=devColumns.find(k=>/建物|building|棟|ビル/i.test(k));
              const sessionBuildings=sessionInfo?.selectedBuildings||[];
              const showBuildings=bKey?(sessionBuildings.length>0?sessionBuildings:[...new Set(devList.map(d=>d._raw?.[bKey]).filter(Boolean))]):[];
              return (
                <>
                <div style={{background:C.white,borderRadius:12,padding:"10px 12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <div style={{display:"flex",gap:16,rowGap:8,flexWrap:"wrap",alignItems:"center"}}>
                    {[["点検日",form.inspectionDate],["点検者",form.inspector]].map(([k,v])=>(
                      <div key={k} style={{display:"flex",alignItems:"baseline",gap:6}}>
                        <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ {k}</span>
                        <span style={{fontSize:14,fontWeight:700,color:C.navy}}>{v||"—"}</span>
                      </div>
                    ))}
                    <button onMouseDown={e=>e.preventDefault()} onClick={()=>setStep(0)}
                      style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"1.5px solid "+C.g300,background:C.g50,color:C.g600,cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>修正</button>
                  </div>
                </div>
                <div style={{background:C.white,borderRadius:12,padding:"10px 12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginTop:6}}>
                  <div style={{display:"flex",gap:16,rowGap:8,flexWrap:"wrap",alignItems:"center"}}>
                    {showBuildings.length>0&&(
                      <div style={{display:"flex",alignItems:"baseline",gap:6}}>
                        <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ 建物</span>
                        <span style={{fontSize:14,fontWeight:700,color:C.navy}}>{showBuildings.join("・")}</span>
                      </div>
                    )}
                    <div style={{display:"flex",alignItems:"baseline",gap:6}}>
                      <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ 階</span>
                      <span style={{fontSize:14,fontWeight:700,color:C.navy}}>{form.floor||"—"}</span>
                    </div>
                    <button onMouseDown={e=>e.preventDefault()} onClick={()=>setStep(0)}
                      style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"1.5px solid "+C.g300,background:C.g50,color:C.g600,cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>修正</button>
                  </div>
                </div>
                </>
              );
            })()}
            <div style={{background:C.white,borderRadius:12,padding:"10px 12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginTop:6}}>
              <div style={{display:"flex",gap:16,rowGap:8,flexWrap:"wrap",alignItems:"center"}}>
                {[["部屋名",form.room],["管理番号",form.managementNo],["機器番号",form.unitNo]].map(([k,v])=>(
                  <div key={k} style={{display:"flex",alignItems:"baseline",gap:6}}>
                    <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ {k}</span>
                    <span style={{fontSize:14,fontWeight:700,color:C.navy}}>{v||"—"}</span>
                  </div>
                ))}
                <button onMouseDown={e=>e.preventDefault()}
                  onClick={()=>{setSummaryExpanded(true);setForm(p=>({...p,floor:"",room:"",managementNo:"",unitNo:""}));setDevSearch("");}}
                  style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"1.5px solid "+C.g300,background:C.g50,color:C.g600,cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>修正</button>
              </div>
            </div>
            </>
          );
  const seqIndoor = focusSeq||[];
  const aiIndoor = seqIndoor.findIndex(f=>f.code===activeCode);
  const activeIsCheckIndoor = !!activeCode && !!isCheckCode && isCheckCode(activeCode);
  const ciFieldsIndoor = (checkFields||[]).filter(f=>f.group==="check_in");
  const activeCheckFieldIndoor = activeIsCheckIndoor ? ciFieldsIndoor.find(f=>f.code===activeCode) : null;
  return (
    <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"row",gap:0}}>

      {/* ── 左：フォームエリア（常時） ── */}
      <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column",padding:"10px 12px 12px",gap:0}}>
        <div ref={outListRef} style={{flex:1,display:"flex",flexDirection:"column",gap:"1px",background:C.g200,borderRadius:14,overflow:"auto",scrollbarGutter:"stable",boxShadow:"0 1px 8px rgba(0,0,0,0.08)",position:"relative"}}>

          {topSummary}

            <>
          {/* ── 点検前状態（室内機のみ）／ 室外機チェックリスト（室外機点検の統合画面）── */}
          {inspectionMode==="outdoor" ? (
            <div style={{opacity:(s1DevDone&&!devSearch)?1:0.45,pointerEvents:(s1DevDone&&!devSearch)?"auto":"none"}}>
              <div style={{padding:"12px 12px 4px"}}>
                <div style={{fontSize:12,fontWeight:800,color:C.teal,marginBottom:8}}>🏭 室外機点検チェック</div>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead>
                    <tr>
                      <th style={{width:90,padding:"6px 8px",textAlign:"left",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>項目</th>
                      <th style={{padding:"6px 8px",textAlign:"left",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>点検内容</th>
                      <th style={{width:110,padding:"6px 8px",textAlign:"center",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>○×</th>
                    </tr>
                  </thead>
                  <tbody>
                    {outFields.map((f,i)=>{
                      const val=form.checks?.[f.code]||"";
                      const act=s1Focus===f.code;
                      return (
                        <tr key={f.code} ref={el=>{outRowRefs.current[f.code]=el;}}
                          onClick={()=>focusCheckRow(f.code)}
                          style={{background:act?C.blue+"0C":i%2===0?C.white:C.g50,boxShadow:act?"inset 0 0 0 2px "+C.blue:"none",cursor:"pointer",transition:"box-shadow 0.1s"}}>
                          <td style={{padding:"9px 8px",fontSize:11,fontWeight:700,color:C.teal,borderBottom:"1px solid "+C.g100,verticalAlign:"middle"}}>{f.category}</td>
                          <td style={{padding:"9px 8px",fontSize:14,color:C.g700,borderBottom:"1px solid "+C.g100,verticalAlign:"middle"}}>{f.label}</td>
                          <td style={{padding:"6px 8px",borderBottom:"1px solid "+C.g100,verticalAlign:"middle"}}>
                            <div style={{display:"flex",gap:6,justifyContent:"center"}} onClick={e=>e.stopPropagation()}>
                              {["○","×"].map(v=>(
                                <button key={v} onClick={()=>advanceCheck(f.code,v)}
                                  style={{width:48,height:40,borderRadius:9,border:"2px solid "+(val===v?(v==="○"?C.green:C.red):C.g200),
                                    background:val===v?(v==="○"?C.green:C.red):C.white,
                                    color:val===v?C.white:C.g400,fontWeight:800,fontSize:19,cursor:"pointer",transition:"all 0.1s"}}>{v}</button>
                              ))}
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          ) : (
          (s1DevDone && !devSearch) ? (
          preStateOpen ? (<>
          {(()=>{
            const prevRec = records.find(r=>r.managementNo===form.managementNo&&r.unitNo===form.unitNo&&(r.preOperation||r.preMode||r.preWind||r.preSetTemp));
            if(!prevRec) return null;
            const prevSummary=[prevRec.preOperation,prevRec.preMode,prevRec.preWind,prevRec.preSetTemp?prevRec.preSetTemp+"°C":null].filter(Boolean).join(" ・ ");
            return (
              <div style={{padding:"6px 10px",marginBottom:4,background:"#FFF7ED",border:"1.5px solid #F59E0B",borderRadius:8,fontSize:11,color:"#92400E",fontWeight:700}}>
                📋 一次保存済み：前回の点検前状態「{prevSummary}」
              </div>
            );
          })()}
          <div style={{display:"flex",background:C.g200,gap:"1px"}}>
            <div style={{width:110,flexShrink:0,background:C.teal+"10",padding:"10px 12px"}}>
              <div style={{display:"flex",alignItems:"center",gap:7}}>
                <div style={{width:18,height:18,borderRadius:9,background:C.teal+"50",display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:C.teal}}>4</div>
                <span style={{fontSize:12,fontWeight:700,color:C.teal}}>点検前状態</span>
              </div>
              <div style={{fontSize:10,color:C.g400,marginTop:2,paddingLeft:25}}>任意</div>
            </div>
            <div style={{flex:1,background:C.teal+"05",padding:"12px 12px",pointerEvents:s1DevDone?"auto":"none",display:"flex",gap:8,alignItems:"flex-start",borderTop:"1px solid "+C.teal+"20",flexWrap:"nowrap",overflowX:"auto"}}>
              {/* 共通ボタンスタイル関数 */}
              {[
                {label:"運転",   items:[["ON","🟢 ON",C.blue],["OFF","⭕ OFF",C.blue]], key:"preOperation"},
                {label:"運転モード", items:[["冷房","❄️ 冷房",C.teal],["暖房","🔥 暖房",C.teal],["送風","💨 送風",C.teal],["除湿","💧 除湿",C.teal]], key:"preMode"},
                {label:"風量",   items:[["自動","🔄 自動","#7C3AED"],["弱","💨 弱","#7C3AED"],["強","💨 強","#7C3AED"],["急風","🌪️ 急風","#7C3AED"]], key:"preWind"},
              ].map(({label,items,key})=>(
                <div key={key} style={{display:"flex",flexDirection:"column",gap:4,flexShrink:0}}>
                  <div style={{fontSize:10,fontWeight:700,color:C.g500,marginBottom:2}}>{label}</div>
                  {items.map(([v,txt,col])=>{
                    const sel=form[key]===v;
                    return (
                      <button key={v} onClick={()=>setInfo(key,sel?"":v)}
                        style={{padding:"10px 16px",borderRadius:10,border:"2px solid "+(sel?col:C.g200),cursor:"pointer",fontWeight:700,fontSize:15,
                          background:sel?"linear-gradient(135deg,"+col+","+col+"BB)":C.white,
                          color:sel?C.white:C.g600,transition:"all 0.12s",textAlign:"center",whiteSpace:"nowrap"}}>
                        {txt}
                      </button>
                    );
                  })}
                </div>
              ))}
              {/* 設定温度（上下スピン） */}
              <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"center"}}>
                <div style={{fontSize:10,fontWeight:700,color:C.g500,marginBottom:2}}>設定温度</div>
                <button onClick={()=>setInfo("preSetTemp",Math.min(30,parseFloat(form.preSetTemp||20)+0.5).toFixed(1))}
                  style={{width:48,height:40,borderRadius:10,border:"1.5px solid "+C.g200,cursor:"pointer",fontSize:20,fontWeight:700,background:C.white,color:C.g600}}>＋</button>
                <div onClick={()=>setS1Focus(s1Focus==="preSetTemp"?null:"preSetTemp")}
                  style={{width:80,padding:"8px 6px",borderRadius:10,cursor:"pointer",
                    border:"2px solid "+(s1Focus==="preSetTemp"?C.blue:form.preSetTemp?C.green:C.g200),
                    background:s1Focus==="preSetTemp"?"#EFF6FF":C.inp,
                    textAlign:"center",fontFamily:"monospace",fontSize:22,fontWeight:800,
                    color:form.preSetTemp?C.navy:C.g300,transition:"all 0.12s"}}>
                  {form.preSetTemp ? parseFloat(form.preSetTemp).toFixed(1) : "—"}<br/><span style={{fontSize:9,fontWeight:400,color:C.g400}}>°C</span>
                </div>
                <button onClick={()=>setInfo("preSetTemp",Math.max(16,parseFloat(form.preSetTemp||20)-0.5).toFixed(1))}
                  style={{width:48,height:40,borderRadius:10,border:"1.5px solid "+C.g200,cursor:"pointer",fontSize:20,fontWeight:700,background:C.white,color:C.g600}}>－</button>
              </div>
            </div>
          </div>
          {/* 一次保存/スキップ：点検前状態を折りたたむ（測定データはブロックしない・既に表示・操作可能） */}
          {(()=>{
            const preStateSet = !!(form.preOperation||form.preMode||form.preWind||form.preSetTemp);
            return (
              <button
                onClick={()=>{ setPreStateConfirmedFor(deviceKey); if(focusSeq&&focusSeq.length>0){ onRowClick&&onRowClick(focusSeq[0]); } }}
                style={{marginTop:6,display:"block",width:"100%",boxSizing:"border-box",padding:"10px",borderRadius:10,border:"none",cursor:"pointer",fontWeight:800,fontSize:14,
                  background:preStateSet?"linear-gradient(135deg,"+C.green+",#047857)":C.g200,
                  color:preStateSet?C.white:C.g600}}>
                {preStateSet?"💾 一次保存":"⏭️ スキップ"}
              </button>
            );
          })()}
          </>) : (
            /* ── 折りたたみ済み：点検日等のバーと同じ書式（白背景＋✓アイコン）で1行表示（タップで再展開） ── */
            <div style={{width:"100%",boxSizing:"border-box",background:C.white,borderRadius:12,padding:"10px 12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
              <div style={{display:"flex",gap:16,rowGap:8,flexWrap:"wrap",alignItems:"center"}}>
                <div style={{display:"flex",alignItems:"baseline",gap:6}}>
                  <span style={{fontSize:11,fontWeight:700,color:C.green,whiteSpace:"nowrap"}}>✓ 点検前状態</span>
                  <span style={{fontSize:13,fontWeight:700,color:C.navy}}>
                    {[form.preOperation,form.preMode,form.preWind,form.preSetTemp?form.preSetTemp+"°C":null].filter(Boolean).join(" ・ ")||"未入力（スキップ済み）"}
                  </span>
                </div>
                <button onMouseDown={e=>e.preventDefault()} onClick={()=>setPreStateConfirmedFor(null)}
                  style={{marginLeft:"auto",padding:"5px 12px",borderRadius:7,border:"1.5px solid "+C.g300,background:C.g50,color:C.g600,cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>修正</button>
              </div>
            </div>
          )
          ) : null
          )}
          {inspectionMode==="indoor" && s1DevDone && !devSearch && (
            <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",opacity:preStateOpen?0.4:1,pointerEvents:preStateOpen?"none":"auto",transition:"opacity 0.15s"}}>
              <Step2View hideHeader hideNumpad
                form={form} setInfo={setInfo} handleSave={handleSave} setStep={()=>setSummaryExpanded(true)}
                visIn={visIn} visOut={visOut} visFields={visFields} activeCode={activeCode} numDisp={numDisp} limits={limits}
                onPress={onPress} onConfirm={onConfirm} onRowClick={onRowClick} moveActive={moveActive}
                rowRefs={rowRefs} listRef={listRef} complete={complete} missing={missing} editIdx={editIdx}
                ALL_FIELDS={ALL_FIELDS} vis={vis} isAbn={isAbn} setCheck={setCheck} checkFields={checkFields} inspectionMode={inspectionMode}
                focusSeq={focusSeq} isCheckCode={isCheckCode} setCheckAndAdvance={setCheckAndAdvance}
              />
            </div>
          )}
            </>

        </div>

        {/* フッター：室外機のみ（備考欄。保存ボタンは右サイドパネルのNumpadに統合済み）。室内機は備考欄も保存ボタンもStep2View側に既にあるためここでは何も表示しない */}
        {inspectionMode==="outdoor" && (
          <div style={{marginTop:10,background:C.white,borderRadius:12,padding:"10px 12px",display:"flex",gap:10,alignItems:"stretch",boxShadow:"0 1px 8px rgba(0,0,0,0.08)"}}>
            <div style={{flex:1,display:"flex",flexDirection:"column",gap:4}}>
              <div style={{fontSize:12,fontWeight:700,color:C.g500}}>📝 備考・特記事項</div>
              <textarea value={form.remarks} onChange={e=>setInfo("remarks",e.target.value)} placeholder="異常箇所、特記事項など..."
                style={{flex:1,width:"100%",padding:"10px 12px",borderRadius:9,fontSize:14,border:"1.5px solid "+C.g200,background:C.inp,outline:"none",boxSizing:"border-box",fontFamily:"inherit",resize:"none",minHeight:100,lineHeight:1.5}}/>
            </div>
          </div>
        )}
      </div>

      {(inspectionMode==="indoor" && s1DevDone && !devSearch) ? (
        <div style={{opacity:preStateOpen?0.4:1,pointerEvents:preStateOpen?"none":"auto",transition:"opacity 0.15s",display:"flex"}}>
        <Numpad
          mode={activeIsCheckIndoor?"check":"numeric"}
          display={numDisp} onPress={onPress} onConfirm={onConfirm} canConfirm={!!activeCode&&!activeIsCheckIndoor&&numDisp!==""}
          checkLabel={activeCheckFieldIndoor?.label} checkCategory={activeCheckFieldIndoor?.category}
          checkValue={form.checks?.[activeCode]||""}
          onCheckPress={v=>setCheckAndAdvance&&setCheckAndAdvance(activeCode,v)}
          onPrev={()=>moveActive(-1)} onNext={()=>moveActive(1)}
          canPrev={aiIndoor>0} canNext={aiIndoor>=0}
          onSave={()=>handleSave&&handleSave("next")} saveComplete={complete} saveMissing={missing}
        />
        </div>
      ) : (
        <>
      {/* ── 右：常時テンキー（室外機は室内機と同じNumpadを共有。数字は使わないためグレーアウトし○×のみ機能） ── */}
      {inspectionMode==="outdoor" ? (
        <div style={{opacity:(s1DevDone&&!devSearch)?1:0.45,pointerEvents:(s1DevDone&&!devSearch)?"auto":"none",transition:"opacity 0.15s",display:"flex"}}>
        <Numpad
          mode="check"
          checkLabel={focusedCheckField?.label} checkCategory={focusedCheckField?.category}
          checkValue={form.checks?.[s1Focus]||""}
          onCheckPress={v=>advanceCheck(s1Focus,v)}
          onPrev={()=>{
            const idx=outFields.findIndex(f=>f.code===s1Focus);
            if(idx>0) focusCheckRow(outFields[idx-1].code);
          }}
          onNext={()=>{
            const idx=outFields.findIndex(f=>f.code===s1Focus);
            const next=outFields[idx+1];
            if(next) focusCheckRow(next.code);
          }}
          canPrev={outFields.findIndex(f=>f.code===s1Focus)>0}
          canNext={(()=>{const idx=outFields.findIndex(f=>f.code===s1Focus);return idx>=0&&idx<outFields.length-1;})()}
          onSave={()=>handleSave&&handleSave("next")} saveComplete={complete} saveMissing={missing}
        />
        </div>
      ) : (
      <div style={{width:208,flexShrink:0,display:"flex",flexDirection:"column",background:C.g50,borderLeft:"2px solid "+C.g200,padding:"10px 8px",gap:6}}>
        <div style={{background:C.white,borderRadius:10,padding:"8px 10px",border:"2px solid "+(isCheckFocus?C.green:s1Focus?C.blue:C.g200),minHeight:52,display:"flex",flexDirection:"column",justifyContent:"center",transition:"border-color 0.15s"}}>
          {isCheckFocus ? (
            <>
              <div style={{fontSize:10,fontWeight:700,color:"#059669",letterSpacing:"0.04em"}}>{focusedCheckField?.category||""}</div>
              <div style={{fontSize:14,fontWeight:800,color:C.navy,lineHeight:1.3}}>{focusedCheckField?.label||"—"}</div>
            </>
          ) : (
            <>
              <div style={{fontSize:10,fontWeight:700,color:C.g400,letterSpacing:"0.04em"}}>
                {s1Focus==="preSetTemp"?"設定温度を入力中":s1Focus==="devSearch"?"検索ワードを入力":"検索窓または設定温度をタップ"}
              </div>
              <div style={{fontFamily:"monospace",fontSize:22,fontWeight:800,textAlign:"right",lineHeight:1.2,color:s1Focus?C.navy:C.g300,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                {s1Focus==="preSetTemp"
                  ? (form.preSetTemp||<span style={{color:C.g200}}>—</span>)
                  : s1Focus==="devSearch"
                  ? (devSearch||<span style={{fontSize:13,color:C.g300}}>入力してください</span>)
                  : <span style={{fontSize:13}}>—</span>}
              </div>
            </>
          )}
        </div>
        {isCheckFocus ? (
          <div style={{display:"flex",gap:8,height:226}}>
            {["○","×"].map(v=>{
              const val=form.checks?.[s1Focus]||"";
              return (
                <button key={v} onClick={()=>advanceCheck(s1Focus,v)}
                  style={{flex:1,borderRadius:14,border:"2px solid "+(val===v?(v==="○"?C.green:C.red):C.g200),
                    background:val===v?(v==="○"?C.green:C.red):C.white,
                    color:val===v?C.white:C.g400,fontWeight:800,fontSize:44,cursor:"pointer",transition:"all 0.1s",
                    display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 2px 6px rgba(0,0,0,0.10)"}}>
                  {v}
                </button>
              );
            })}
          </div>
        ) : (
          [[7,8,9],[4,5,6],[1,2,3],[0,"."]].map((row,ri)=>(
            <div key={ri} style={{display:"flex",gap:5}}>
              {row.map(k=>(
                <button key={k}
                  disabled={!s1Focus}
                  onMouseDown={e=>e.preventDefault()}
                  onClick={()=>{
                    if(s1Focus==="preSetTemp"){
                      const prev=String(form.preSetTemp||"");
                      let next;
                      if(k==="."){next=prev.includes(".")?prev:(prev||"0")+".";}
                      else if(prev==="0"){next=String(k);}
                      else{next=prev.length<4?prev+k:prev;}
                      setInfo("preSetTemp",next);
                    } else if(s1Focus==="devSearch"){
                      setDevSearch(p=>p+String(k));
                    }
                  }}
                  style={{flex:(ri===3&&k===0)?2:1,height:52,borderRadius:10,border:"none",
                    cursor:s1Focus?"pointer":"default",fontWeight:800,fontFamily:"monospace",fontSize:22,
                    background:s1Focus?C.white:C.g100,color:s1Focus?C.g800:C.g300,
                    boxShadow:s1Focus?"0 2px 5px rgba(0,0,0,0.09)":"none",
                    display:"flex",alignItems:"center",justifyContent:"center",transition:"background 0.1s"}}>
                  {k}
                </button>
              ))}
            </div>
          ))
        )}
        <div style={{display:"flex",gap:5}}>
          <button disabled={!s1Focus||isCheckFocus}
            onMouseDown={e=>e.preventDefault()}
            onClick={()=>{
              if(s1Focus==="preSetTemp") setInfo("preSetTemp",String(form.preSetTemp||"").slice(0,-1));
              else if(s1Focus==="devSearch") setDevSearch(p=>p.slice(0,-1));
            }}
            style={{flex:1,height:48,borderRadius:10,border:"none",cursor:(s1Focus&&!isCheckFocus)?"pointer":"default",
              fontSize:18,fontWeight:800,background:(s1Focus&&!isCheckFocus)?C.g200:C.g100,color:(s1Focus&&!isCheckFocus)?C.g600:C.g300,
              display:"flex",alignItems:"center",justifyContent:"center"}}>⌫</button>
          <button disabled={!s1Focus}
            onMouseDown={e=>e.preventDefault()}
            onClick={()=>{
              if(s1Focus==="preSetTemp"){setInfo("preSetTemp","");setS1Focus(null);}
              else if(s1Focus==="devSearch"){setDevSearch("");setS1Focus(null);}
              else if(isCheckFocus){setCheck&&setCheck(s1Focus,"");}
            }}
            style={{flex:1,height:48,borderRadius:10,border:"none",cursor:s1Focus?"pointer":"default",
              fontSize:14,fontWeight:800,background:s1Focus?"#FEF2F2":C.g100,color:s1Focus?C.red:C.g300,
              display:"flex",alignItems:"center",justifyContent:"center"}}>CLR</button>
        </div>
        <div style={{marginTop:"auto",padding:"4px",fontSize:11,color:C.g400,textAlign:"center",lineHeight:1.6}}>
          {isCheckFocus ? <>チェック項目をタップ<br/>すると連続入力できます</> : <>検索窓または<br/>設定温度をタップ</>}
        </div>
      </div>
      )}
        </>
      )}
    </div>
  );
}

function ACInspectionApp() {
  const [view,setView]   = useState("form");
  const [step,setStep]   = useState(0); // 0=点検エリア確認 1=機器選択 2=測定データ
  const [undoneOnly, setUndoneOnly] = useState(false); // 未入力分のみ表示
  const [sessionInfo, setSessionInfo] = useState(()=>{
    try {
      const saved = localStorage.getItem("acSessionInfo");
      return saved ? JSON.parse(saved) : null;
    } catch(e){ return null; }
  }); // {date, inspector, targetFloors:[], roomAccess:{}}
  const [form,setForm]   = useState(emptyForm());
  const [inspectionMode, setInspectionMode] = useState("indoor"); // "indoor" | "outdoor"
  const [indoorRecords, setIndoorRecords] = useState([]);
  const [outdoorRecords, setOutdoorRecords] = useState([]);
  // 現在のモードに応じたrecords/setRecords
  const records = inspectionMode==="indoor" ? indoorRecords : outdoorRecords;
  const setRecords = inspectionMode==="indoor" ? setIndoorRecords : setOutdoorRecords;
  const [editIdx,setEditIdx] = useState(null);
  const [flash,setFlash] = useState("");

  const [limits,setLimits]         = useState(()=>lsGet(LS_KEYS.limits, defLim()));
  const [tmpLim,setTmpLim]         = useState(defLim());
  const [vis,setVis]               = useState(()=>lsGet(LS_KEYS.vis, defVis()));
  const [tmpVis,setTmpVis]         = useState(defVis());
  const [devList,setDevList]       = useState(()=>lsGet(LS_KEYS.devList, [])); // [{floor,room,managementNo,unitNo,_raw:{}}]
  const [devColumns,setDevColumns]   = useState(()=>lsGet(LS_KEYS.devColumns, [])); // CSVの全列名
  const [devVisibleCols,setDevVisibleCols] = useState(()=>lsGet(LS_KEYS.devVisibleCols, [])); // 表示する列名（空=全列）
  const [inspList,setInspList]     = useState(()=>lsGet(LS_KEYS.inspList, []));
  const [checkFields,setCheckFields] = useState(()=>lsGet(LS_KEYS.checkFields, [
    // 室内機チェック項目
    {code:"ci1", label:"配管類支持異常の有無", category:"据付状態",       group:"check_in"},
    {code:"ci2", label:"異音・異常振動の有無", category:"フィルター点検・清掃", group:"check_in"},
    {code:"ci3", label:"異音・異常振動の有無", category:"運転確認",       group:"check_in"},
    // 室外機チェック項目
    {code:"co1", label:"防振装置異常の有無",   category:"据付状態",       group:"check_out"},
    {code:"co2", label:"配管類支持異常の有無", category:"据付状態",       group:"check_out"},
    {code:"co3", label:"ガスリークテスト",     category:"冷媒系統",       group:"check_out"},
    {code:"co4", label:"配管系統外観点検",     category:"冷媒系統",       group:"check_out"},
    {code:"co5", label:"異音・異常振動の有無", category:"送排風機系統",   group:"check_out"},
    {code:"co6", label:"ドレン配管異常の有無", category:"排水系統",       group:"check_out"},
    {code:"co7", label:"フィン汚れの有無",     category:"熱交換器系統",   group:"check_out"},
    {code:"co8", label:"異音・異常振動の有無", category:"熱交換器系統",   group:"check_out"},
    {code:"co9", label:"外面清掃",             category:"作業終了時",     group:"check_out"},
  ])); // [{code,label,category,group}]
  const [lastInsp,setLastInsp]     = useState("");
  const [lastDate,setLastDate]     = useState(new Date().toISOString().slice(0,10));
  const [devSearch,setDevSearch]   = useState("");
  const [openSec,setOpenSec]       = useState({device:false,inspector:false,cols:false,checkitems:false,vis:false,lim:false});
  const [modalSec,setModalSec]       = useState(null); // 設定画面で開いているモーダルのid

  const [numDisp,setNumDisp]   = useState("");
  const isOvr = useRef(false);
  const [activeCode,setActiveCode] = useState(null);
  const rowRefs   = useRef({});
  const listRef   = useRef();
  const devRef    = useRef();
  const inspRef   = useRef();

  useEffect(()=>{
    if(view==="settings"){ setTmpLim(JSON.parse(JSON.stringify(limits))); setTmpVis(JSON.parse(JSON.stringify(vis))); }
  },[view]);

  // ─── 取込データの自動保存（ブラウザのlocalStorageへ）───
  // 機器リスト・表示列設定・点検項目・点検者リスト・表示項目設定・正常値範囲を
  // アプリ更新後も引き継げるようにする（入力中の点検記録は対象外）
  useEffect(()=>{ lsSet(LS_KEYS.devList, devList); },[devList]);
  useEffect(()=>{ lsSet(LS_KEYS.devColumns, devColumns); },[devColumns]);
  useEffect(()=>{ lsSet(LS_KEYS.devVisibleCols, devVisibleCols); },[devVisibleCols]);
  useEffect(()=>{ lsSet(LS_KEYS.inspList, inspList); },[inspList]);
  useEffect(()=>{ lsSet(LS_KEYS.checkFields, checkFields); },[checkFields]);
  useEffect(()=>{ lsSet(LS_KEYS.limits, limits); },[limits]);
  useEffect(()=>{ lsSet(LS_KEYS.vis, vis); },[vis]);

  const showFlash = msg => { setFlash(msg); setTimeout(()=>setFlash(""),2400); };
  const visIn  = INDOOR_FIELDS.filter(f=>vis[f.code]);
  const visOut = OUTDOOR_FIELDS.filter(f=>vis[f.code]);
  const visFields = ALL_FIELDS.filter(f=>vis[f.code]);
  const ciFields = checkFields.filter(f=>f.group==="check_in");
  // 測定値項目＋室内機チェック項目を1本の連続シーケンスにして、同じ右側パネル・同じ手の位置でENTERまたは○×を続けて入力できるようにする
  const focusSeq = [...visFields, ...ciFields];
  const isCheckCode = code => ciFields.some(f=>f.code===code);
  const outChkFieldsAll = checkFields.filter(f=>f.group==="check_out");
  const complete = inspectionMode==="outdoor"
    ? (outChkFieldsAll.length>0 && outChkFieldsAll.every(f=>(form.checks?.[f.code]||"")!==""))
    : (visFields.every(f=>form.values[f.code]!=="") && ciFields.every(f=>(form.checks?.[f.code]||"")!==""));
  const missing = inspectionMode==="outdoor"
    ? outChkFieldsAll.filter(f=>(form.checks?.[f.code]||"")==="").length
    : visFields.filter(f=>form.values[f.code]==="").length + ciFields.filter(f=>(form.checks?.[f.code]||"")==="").length;

  const setInfo  = (k,v) => setForm(p=>({...p,[k]:v}));
  const setVal   = (code,v) => setForm(p=>({...p,values:{...p.values,[code]:v}}));
  const setCheck = (code,v) => setForm(p=>({...p,checks:{...p.checks,[code]:v}}));

  // Step1 state
  const s1DateDone   = !!form.inspectionDate;
  const s1InspDone   = !!form.inspector;
  const s1DevDone    = !!(form.managementNo&&form.unitNo&&form.floor&&form.room);
  const step1Valid   = s1DateDone && s1InspDone && s1DevDone;

  const scrollToCenter = (code,smooth=true) => {
    requestAnimationFrame(()=>requestAnimationFrame(()=>{
      const el=rowRefs.current[code]; const ct=listRef.current;
      if(!el||!ct) return;
      let top=0,node=el;
      while(node&&node!==ct){top+=node.offsetTop;node=node.offsetParent;}
      ct.scrollTo({top:Math.max(0,top-ct.clientHeight/2+el.offsetHeight/2),behavior:smooth?"smooth":"instant"});
    }));
  };

  const onPress = key => {
    if(!activeCode||isCheckCode(activeCode)) return;
    setNumDisp(prev=>{
      if(key==="C"){isOvr.current=false;return"";}
      if(key==="⌫"){isOvr.current=false;return prev.slice(0,-1);}
      if(isOvr.current){isOvr.current=false;return key==="."?"0.":String(key);}
      if(key==="."){if(prev.includes("."))return prev;return(prev||"0")+".";}
      if(prev==="0")return String(key);
      return prev+key;
    });
  };
  const focusField = (code,smooth=true) => {
    setActiveCode(code);
    setNumDisp(code&&!isCheckCode(code)?(form.values[code]||""):"");
    isOvr.current=true;
    if(code)scrollToCenter(code,smooth);
  };
  const onConfirm = () => {
    if(!activeCode||isCheckCode(activeCode)||numDisp==="") return;
    setVal(activeCode,numDisp);
    const idx=focusSeq.findIndex(f=>f.code===activeCode);
    const next=focusSeq[idx+1];
    if(next){focusField(next.code);}
    else{setActiveCode(null);setNumDisp("");isOvr.current=false;}
  };
  const moveActive = dir => {
    if(!activeCode){
      if(focusSeq.length>0){focusField(focusSeq[0].code,true);}
      return;
    }
    if(!isCheckCode(activeCode)&&numDisp!=="")setVal(activeCode,numDisp);
    const idx=focusSeq.findIndex(f=>f.code===activeCode);
    const next=focusSeq[idx+dir];
    if(next){ focusField(next.code,false); }
    else { setActiveCode(null); setNumDisp(""); isOvr.current=false; }
  };
  const onRowClick = f => {
    if(activeCode&&!isCheckCode(activeCode)&&numDisp!=="")setVal(activeCode,numDisp);
    focusField(f.code);
  };
  // ○×チェック項目：値を設定し、続けて次の項目（測定値/チェックの連続シーケンス上の次）へ自動的にフォーカスを移す
  // ※反転（同じ値を押すとクリア）はしない。何度押しても選んだ値のまま
  const setCheckAndAdvance = (code,v) => {
    setCheck(code, v);
    const idx=focusSeq.findIndex(f=>f.code===code);
    const next=focusSeq[idx+1];
    if(next){focusField(next.code);}
    else{setActiveCode(null);setNumDisp("");}
  };
  const goToStep2 = () => {
    if(!step1Valid) return;
    const first=focusSeq[0]; setStep(2);
    setActiveCode(first?.code||null);setNumDisp(first&&!isCheckCode(first.code)?(form.values[first.code]||""):"");isOvr.current=true;
    if(first)scrollToCenter(first.code);
  };
  // 機器選択画面内で室内機／室外機を切り替える（点検日・点検者・建物・階はそのまま、機器選択のみクリア）
  const onSwitchMode = mode => {
    if(mode===inspectionMode) return;
    setInspectionMode(mode);
    setForm(p=>({...p, floor:"", room:"", managementNo:"", unitNo:"", values:emptyVal(), checks:{}}));
    setEditIdx(null);
    setActiveCode(null); setNumDisp("");
  };
  const [saveModal, setSaveModal] = useState(null); // 保存完了モーダル用データ
  const [measZoom, setMeasZoom] = useState(false); // 測定データ拡大表示

  const handleSave = (mode="next") => {
    if(!complete) return;
    setLastInsp(form.inspector);setLastDate(form.inspectionDate);
    const saved = {...form};
    if(editIdx!==null){setRecords(p=>p.map((r,i)=>i===editIdx?saved:r));setEditIdx(null);}
    else{setRecords(p=>[...p,saved]);}
    setSaveModal({...saved, _mode:mode}); // モーダル表示
  };
  // ① 次へ：測定データ入力 → 保存後そのまま測定データ入力画面へ
  const closeSaveNext = () => {
    const insp=saveModal?.inspector||form.inspector;
    const date=saveModal?.inspectionDate||form.inspectionDate;
    setSaveModal(null);
    setForm(emptyForm(insp,date));
    setStep(inspectionMode==="outdoor"?1:1);setActiveCode(null);setNumDisp("");isOvr.current=false;
  };
  // ② 一次保存 → 点検日・点検者はそのまま、機器情報リセットして基本情報へ
  const closeSaveTmp = () => {
    const insp=saveModal?.inspector||form.inspector;
    const date=saveModal?.inspectionDate||form.inspectionDate;
    setSaveModal(null);
    setForm(p=>({...emptyForm(insp,date),inspector:insp,inspectionDate:date}));
    setStep(1);setActiveCode(null);setNumDisp("");isOvr.current=false;
  };
  const handleEdit = i => { setForm({...records[i]});setEditIdx(i);setStep(1);setView("form"); };
  const handleDel  = i => {
    if(!window.confirm("削除しますか？")) return;
    setRecords(p=>p.filter((_,j)=>j!==i));
    if(editIdx===i){setEditIdx(null);setForm(emptyForm(lastInsp,lastDate));}
  };
  // 一覧（室内機/室外機まとめ表示）用：対象の記録がどちらのモードかを指定して編集・削除する
  const handleEditFor = (rec, mode) => {
    const arr = mode==="indoor" ? indoorRecords : outdoorRecords;
    const idx = arr.indexOf(rec);
    if(idx<0) return;
    setInspectionMode(mode);
    setForm({...rec});
    setEditIdx(idx);
    setStep(1);
    setView("form");
  };
  const handleDelFor = (rec, mode) => {
    if(!window.confirm("削除しますか？")) return;
    const setRec = mode==="indoor" ? setIndoorRecords : setOutdoorRecords;
    setRec(p=>p.filter(r=>r!==rec));
  };
  const handleInputFor = (row, mode) => {
    setInspectionMode(mode);
    setForm(p=>({...emptyForm(lastInsp,lastDate),floor:row.floor,room:row.room,managementNo:row.managementNo,unitNo:row.unitNo,inspectionDate:p.inspectionDate,inspector:p.inspector}));
    const bKey=devColumns.find(k=>/建物|building|棟|ビル/i.test(k));
    const rowBuilding=bKey&&row._raw?row._raw[bKey]:null;
    setSessionInfo(p=>({...p,
      targetFloors:row.floor?[row.floor]:(p?.targetFloors||[]),
      selectedBuildings:rowBuilding?[rowBuilding]:(p?.selectedBuildings||[]),
    }));
    setEditIdx(null);setStep(1);setView("form");
  };
  // 一覧の行タップ用：既存レコードがあれば編集、無ければ新規入力へ（室内機/室外機いずれか該当する方）
  const handleEditOrInputFor = (row, mode) => {
    const rec = mode==="indoor" ? row.indoorRecord : row.outdoorRecord;
    if(rec) handleEditFor(rec, mode);
    else handleInputFor(row, mode);
  };
  const loadXLSX = () => new Promise((resolve,reject) => {
    if(window.XLSX){resolve(window.XLSX);return;}
    const s=document.createElement("script");
    s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload=()=>resolve(window.XLSX); s.onerror=reject;
    document.head.appendChild(s);
  });
  const readFile = (file, onParsed) => {
    const isXlsx = /\.xlsx?$/i.test(file.name);
    const reader = new FileReader();
    if(isXlsx) {
      reader.onload = async ev => {
        const XLSX = await loadXLSX();
        const wb = XLSX.read(ev.target.result, {type:"array"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        // sheet_to_json でヘッダー行をキーとして取得（列名の揺れに対応）
        const json = XLSX.utils.sheet_to_json(ws, {defval:"", raw:false});
        if(json.length===0){ onParsed(null,[]); return; }
        // __EMPTYなどSheetJS内部列を除外
        const keys = Object.keys(json[0]).filter(k=>!k.startsWith("__"));
        onParsed(json, keys); // JSON配列をそのまま渡す
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = ev => {
        const bytes = new Uint8Array(ev.target.result);
        const hasUtf8Bom = bytes[0]===0xEF && bytes[1]===0xBB && bytes[2]===0xBF;
        const enc = hasUtf8Bom ? "UTF-8" : "Shift_JIS";
        const text = new TextDecoder(enc).decode(bytes).replace(/^\uFEFF/,"");
        const cols = text.trim().split(/\r?\n/)[0]?.split(",").map(s=>s.trim().replace(/^"|"$/g,"")).filter(Boolean)||[];
        onParsed(text, cols);
      };
      reader.readAsArrayBuffer(file);
    }
  };
  const handleDevCSV  = e=>{const f=e.target.files[0];if(!f)return;readFile(f,(data,cols)=>{
    let l;
    if(Array.isArray(data)){
      // xlsx: JSON配列
      l=parseDevRows(data, cols);
    } else {
      // csv: 文字列
      l=parseDevCSV(data||"");
    }
    setDevList(l);
    if(cols&&cols.length>0){
      const filteredCols=cols.filter(k=>!k.startsWith("__"));
      setDevColumns(filteredCols);
      setDevVisibleCols(filteredCols);
    }
    showFlash("✅ 機器リスト "+l.length+"件 読込");
  });e.target.value="";};
  const handleInspCSV = e=>{const f=e.target.files[0];if(!f)return;readFile(f,(data,cols)=>{
    let text;
    if(Array.isArray(data)){
      // xlsx: 1列目の値を名前リストとして取得
      const nameKey=cols&&cols.length>0?cols[0]:Object.keys(data[0]||{})[0]||"";
      text=data.map(row=>String(row[nameKey]||"").trim()).filter(Boolean).join("\n");
    } else {
      text=data||"";
    }
    const l=parseInspCSV(text);setInspList(l);showFlash("✅ 点検者 "+l.length+"名 読込");
  });e.target.value="";};
  const toggleSec = k => setOpenSec(p=>({...p,[k]:!p[k]}));

  const vf=ALL_FIELDS.filter(f=>vis[f.code]);
  const [listFilter, setListFilter] = useState("all"); // "all"|"done"|"undone"
  const [sortCol, setSortCol]   = useState(null); // ソート列
  const [hoverRow, setHoverRow] = useState(null); // 一覧テーブルのマウスオーバー中の行
  const [sortDir, setSortDir]   = useState("asc");
  const [showStats, setShowStats] = useState(false); // 集計モーダル
  const [floorFilter, setFloorFilter] = useState(null); // 階フィルター（null=全階）

  // 一覧・集計・印刷では室内機/室外機の記録を両方まとめて1行にする
  const tRows=(()=>{
    if(devList.length>0){
      return devList.map(d=>({
        ...d,
        indoorRecord: indoorRecords.find(r=>r.managementNo===d.managementNo&&r.unitNo===d.unitNo)||null,
        outdoorRecord: outdoorRecords.find(r=>r.managementNo===d.managementNo&&r.unitNo===d.unitNo)||null,
      }));
    }
    const map=new Map();
    indoorRecords.forEach(r=>{
      map.set(r.managementNo+"|"+r.unitNo,{floor:r.floor,room:r.room,managementNo:r.managementNo,unitNo:r.unitNo,indoorRecord:r,outdoorRecord:null});
    });
    outdoorRecords.forEach(r=>{
      const key=r.managementNo+"|"+r.unitNo;
      const ex=map.get(key);
      if(ex) ex.outdoorRecord=r;
      else map.set(key,{floor:r.floor,room:r.room,managementNo:r.managementNo,unitNo:r.unitNo,indoorRecord:null,outdoorRecord:r});
    });
    return [...map.values()];
  })();
  // 行から値を取り出す共通ヘルパー（室内機優先→室外機フォールバック）
  const rowFieldVal = (row,code) => {
    const iv=row.indoorRecord?.values?.[code]; if(iv) return iv;
    const ov=row.outdoorRecord?.values?.[code]; if(ov) return ov;
    return "";
  };
  const rowMeta = (row,key) => row.indoorRecord?.[key] || row.outdoorRecord?.[key] || "";
  const rowRemarks = row => {
    const a=row.indoorRecord?.remarks||"", b=row.outdoorRecord?.remarks||"";
    if(a&&b&&a!==b) return a+" / "+b;
    return a||b||"";
  };
  const rowIndoorDone = row => !!row.indoorRecord && Object.values(row.indoorRecord.values).some(v=>v!=="");
  const rowOutdoorDone = row => !!row.outdoorRecord && Object.values(row.outdoorRecord.checks||{}).some(v=>v!=="");
  const rowDone = row => rowIndoorDone(row) || rowOutdoorDone(row);
  // 室内機の行は室内機チェック（check_in）、室外機の行は室外機チェック（check_out）の結果を表示する
  const rowChecksList = row => {
    if(row.indoorRecord) return checkFields.filter(f=>f.group==="check_in").map(f=>({code:f.code,label:f.label,val:row.indoorRecord.checks?.[f.code]||""}));
    if(row.outdoorRecord) return checkFields.filter(f=>f.group==="check_out").map(f=>({code:f.code,label:f.label,val:row.outdoorRecord.checks?.[f.code]||""}));
    return [];
  };

  const handleStep1TmpSave = () => {
    if(!step1Valid) return;
    const saved = {...form, values:emptyVal()};
    setLastInsp(form.inspector); setLastDate(form.inspectionDate);
    if(editIdx!==null) {
      setRecords(p=>p.map((r,i)=>i===editIdx?saved:r)); setEditIdx(null);
    } else {
      const ex=records.findIndex(r=>r.managementNo===form.managementNo&&r.unitNo===form.unitNo);
      if(ex>=0) setRecords(p=>p.map((r,i)=>i===ex?saved:r));
      else setRecords(p=>[...p,saved]);
    }
    setForm(p=>({...emptyForm(p.inspector,p.inspectionDate),inspector:p.inspector,inspectionDate:p.inspectionDate}));
    showFlash("💾 一次保存しました");
  };

  // ─── STEP1 ────────────────────────────────────────────
  // STEP1 テンキー対象フィールド
  const [s1Focus, setS1Focus] = useState(null); // 'date'|'inspector'|'device'|'preSetTemp'

  // Step1View → top-level component

  // フィルター適用
  const listBuildingKey = devColumns.find(k=>/建物|building|棟|ビル/i.test(k))||null;
  // 列の値を取得する共通ヘルパー（オートフィルタ・ソート両方で使用）
  const colValue = (row,col) => {
    if(col==="建物") return listBuildingKey?String(row._raw?.[listBuildingKey]||"（未設定）"):"（未設定）";
    if(col==="階") return row.floor||"（未設定）";
    if(col==="部屋名") return row.room||"（未設定）";
    if(col==="管理番号") return row.managementNo||"（未設定）";
    if(col==="機器番号") return row.unitNo||"（未設定）";
    if(col==="室内機") return rowIndoorDone(row)?"入力済":"未入力";
    if(col==="室外機") return rowOutdoorDone(row)?"入力済":"未入力";
    if(col==="点検日") return rowMeta(row,"inspectionDate")||"（未設定）";
    if(col==="点検者") return rowMeta(row,"inspector")||"（未設定）";
    if(row._raw&&row._raw[col]!==undefined) return String(row._raw[col]||"（未設定）");
    return "";
  };
  // 絞り込み：完了状況・機器種別・建物・階（パネル式、それぞれ単一選択）
  const statsCategoryKey = devColumns.find(k=>/^(分類|category|class)$/i.test(k.trim()))||null;
  const rowStatsType = row => {
    if(statsCategoryKey && row._raw){
      const c = String(row._raw[statsCategoryKey]||"");
      if(/室内機/.test(c)) return "indoor";
      if(/室外機/.test(c)) return "outdoor";
    }
    if(row.indoorRecord) return "indoor";
    if(row.outdoorRecord) return "outdoor";
    return "unknown";
  };
  const [typeFilter, setTypeFilter] = useState(null); // null|"indoor"|"outdoor"
  const [buildingFilter, setBuildingFilter] = useState(null); // null=すべての建物
  const filteredRows = (()=>{
    const rows = tRows.filter(row => {
      if(floorFilter && row.floor!==floorFilter) return false;
      if(buildingFilter && colValue(row,"建物")!==buildingFilter) return false;
      if(typeFilter && rowStatsType(row)!==typeFilter) return false;
      if(listFilter==="done")   return rowDone(row);
      if(listFilter==="undone") return !rowDone(row);
      return true;
    });
    if(!sortCol) return rows;
    return [...rows].sort((a,b)=>{
      const va=colValue(a,sortCol), vb=colValue(b,sortCol);
      const cmp=va.localeCompare(vb,undefined,{numeric:true,sensitivity:"base"});
      return sortDir==="asc"?cmp:-cmp;
    });
  })();

  // データ一覧の絞り込み状態をまとめて説明文にする（CSV出力前の確認ポップアップ用）
  const activeFilterDescriptions = () => {
    const list = [];
    if(listFilter!=="all") list.push("完了状況："+(listFilter==="done"?"入力済":"未入力"));
    if(typeFilter) list.push("機器種別："+(typeFilter==="indoor"?"🏠 室内機":"🏭 室外機"));
    if(buildingFilter) list.push("建物："+buildingFilter);
    if(floorFilter) list.push("階："+floorFilter);
    if(sortCol) list.push("並び替え："+sortCol+"（"+(sortDir==="asc"?"昇順":"降順")+"）");
    return list;
  };
  // データ一覧に表示されている形（見えている列・行）のままCSV（TSV）出力する
  const exportListAsCSV = () => {
    const cols = devVisibleCols.length>0?devVisibleCols:["階","部屋名","管理番号","機器番号"];
    const header = [...cols,"室内機","室外機","点検日","点検者","運転","モード","風量","設定温度",...vf.map(f=>f.code+"("+f.unit+")"),...checkFields.map(f=>f.label),"備考"];
    const dataRows = filteredRows.map(row=>{
      const ir=row.indoorRecord, or_=row.outdoorRecord;
      const metaSrc = ir||or_;
      const colVals = cols.map(col=>{
        const val=row._raw?row._raw[col]:(col==="階"?row.floor:col==="部屋名"?row.room:col==="管理番号"?row.managementNo:col==="機器番号"?row.unitNo:"");
        return val||"";
      });
      const checksVals = checkFields.map(f=>{
        if(ir && f.group==="check_in") return ir.checks?.[f.code]||"";
        if(or_ && f.group==="check_out") return or_.checks?.[f.code]||"";
        return "";
      });
      return [...colVals, rowIndoorDone(row)?"入力済":"未入力", rowOutdoorDone(row)?"入力済":"未入力",
        metaSrc?.inspectionDate||"", metaSrc?.inspector||"", metaSrc?.preOperation||"", metaSrc?.preMode||"", metaSrc?.preWind||"", metaSrc?.preSetTemp||"",
        ...vf.map(f=>rowFieldVal(row,f.code)), ...checksVals, rowRemarks(row)||""];
    });
    const rows=[header,...dataRows];
    const tsv=rows.map(r=>r.map(c=>'"'+String(c??"").replace(/"/g,'""')+'"').join("\t")).join("\n");
    const blob=new Blob(["\uFEFF"+tsv],{type:"text/tab-separated-values;charset=utf-8;"});
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
    a.download="ac_check_list_"+new Date().toISOString().slice(0,10)+".tsv"; a.click();
  };
  const [showExportConfirm, setShowExportConfirm] = useState(null); // 絞り込み中のCSV出力確認ポップアップ（null=非表示、配列=絞り込み内容一覧）
  const handleCSVExportClick = () => {
    if(filteredRows.length===0){ showFlash("⚠️ データがありません"); return; }
    const active = activeFilterDescriptions();
    if(active.length>0){ setShowExportConfirm(active); }
    else { exportListAsCSV(); }
  };

  // 集計（建物×階、室内機/室外機を分けて集計）
  const _floors = [...new Set(tRows.map(r=>r.floor).filter(Boolean))].sort();
  const groupStats = (()=>{
    const map = new Map();
    tRows.forEach(row=>{
      const building = listBuildingKey ? String(row._raw?.[listBuildingKey]||"—") : "—";
      const key = building+"|"+(row.floor||"—");
      if(!map.has(key)) map.set(key,{building,floor:row.floor||"—",indoorRows:[],outdoorRows:[]});
      const g = map.get(key);
      const t = rowStatsType(row);
      if(t==="indoor") g.indoorRows.push(row);
      else if(t==="outdoor") g.outdoorRows.push(row);
    });
    return [...map.values()].map(g=>({
      building:g.building, floor:g.floor,
      indoorTotal:g.indoorRows.length, indoorDone:g.indoorRows.filter(rowIndoorDone).length,
      outdoorTotal:g.outdoorRows.length, outdoorDone:g.outdoorRows.filter(rowOutdoorDone).length,
    })).sort((a,b)=> a.building.localeCompare(b.building,undefined,{numeric:true}) || a.floor.localeCompare(b.floor,undefined,{numeric:true}));
  })();
  const floorStats = _floors.map(f => {
    const rows = tRows.filter(r=>r.floor===f);
    const done = rows.filter(rowDone).length;
    return {floor:f, total:rows.length, done, undone:rows.length-done};
  });

  const modeLabel = m => ({"冷房":"❄️ 冷房","暖房":"🔥 暖房","送風":"💨 送風","除湿":"💧 除湿"}[m]||m||"—");
  const windLabel = w => ({"自動":"🔄 自動","弱":"💨 弱","強":"💨 強","急風":"🌪️ 急風"}[w]||w||"—");
  const closeNext = saveModal && saveModal._mode==="tmp" ? closeSaveTmp : closeSaveNext;
  const closeTmp  = saveModal && saveModal._mode==="tmp" ? closeSaveNext : closeSaveTmp;
  const nextLabel = saveModal && saveModal._mode==="tmp" ? "→ 基本情報へ" : "→ 次の機器へ";

  // ─── render ───────────────────────────────────────────
  return (
    <div style={{fontFamily:"Hiragino Sans, Meiryo, Arial, sans-serif",height:"100%",display:"flex",flexDirection:"column",background:C.g100,color:C.g800,overflow:"hidden",fontSize:16}}>
      <style>{PS}</style>

      {/* topbar（メインヘッダー） */}
      <div style={{background:"linear-gradient(135deg,"+C.navy+" 0%,"+C.blue+" 100%)",color:C.white,flexShrink:0,boxShadow:"0 2px 10px rgba(0,0,0,0.25)",display:"flex",flexDirection:"column"}}>
        <div style={{padding:"8px 16px 4px",fontWeight:700,fontSize:16,letterSpacing:"0.04em"}}>🌡️ エアコン点検</div>
        <div style={{display:"flex",gap:6,padding:"6px 12px 8px",flexWrap:"wrap",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {[["form","📝 点検データ入力"],["list","📋 データ一覧"+(tRows.length>0?" ("+tRows.length+")":"")]].map(([v,label])=>(
              <button key={v} onClick={()=>{setView(v);if(v==="form"&&editIdx===null)setStep(inspectionMode==="outdoor"?1:0);}}
                style={{padding:"6px 12px",borderRadius:7,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap",
                  background:view===v?C.white:"rgba(255,255,255,0.18)",color:view===v?C.navy:C.white}}>
                {label}
              </button>
            ))}
            <button onClick={()=>setShowStats(true)} disabled={tRows.length===0}
              style={{padding:"6px 10px",borderRadius:7,border:"1.5px solid rgba(255,255,255,0.4)",cursor:tRows.length>0?"pointer":"not-allowed",fontSize:12,fontWeight:700,background:"rgba(255,255,255,0.15)",color:C.white,opacity:tRows.length>0?1:0.5,whiteSpace:"nowrap"}}>
              📊 集計
            </button>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            <button key="settings" onClick={()=>setView("settings")}
              style={{padding:"6px 12px",borderRadius:7,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap",
                background:view==="settings"?C.white:"rgba(255,255,255,0.18)",color:view==="settings"?C.navy:C.white}}>
              ⚙️ 設定
            </button>
          </div>
        </div>
      </div>

      <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>

        {/* FORM */}
        {view==="form" && (
          <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
            {/* step=0: 点検エリア確認（室内機・室外機とも、機器選択画面の「修正」から遷移可能） */}
            {step===0 && (
              <SessionView
                devList={devList} inspList={inspList} records={records}
                sessionInfo={sessionInfo} setSessionInfo={setSessionInfo}
                undoneOnly={undoneOnly} setUndoneOnly={setUndoneOnly}
                devColumns={devColumns} devVisibleCols={devVisibleCols}
                onStart={(info)=>{
                  setSessionInfo(info);
                  setForm(p=>({...p, inspectionDate:info.date, inspector:info.inspector}));
                  setLastInsp(info.inspector); setLastDate(info.date);
                  setStep(1);
                }}
                onSelectOutdoor={(info)=>{
                  setInspectionMode("outdoor");
                  setSessionInfo(p=>({...(p||{}), date:info.date, inspector:info.inspector}));
                  setForm(emptyForm(info.inspector,info.date));
                  setLastInsp(info.inspector); setLastDate(info.date);
                  setEditIdx(null);
                  setStep(1);
                }}
              />
            )}
            {/* step>=1: 機器選択/データ入力（ステップバー表示は廃止） */}
            {step>=1 && (<>
            <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column",paddingTop:8}}>
              <Step1View
                form={form} setInfo={setInfo} inspList={inspList} devList={devList}
                devSearch={devSearch} setDevSearch={setDevSearch}
                s1DateDone={s1DateDone} s1InspDone={s1InspDone} s1DevDone={s1DevDone}
                step1Valid={step1Valid} records={records}
                s1Focus={s1Focus} setS1Focus={setS1Focus}
                goToStep2={goToStep2} handleStep1TmpSave={handleStep1TmpSave}
                setForm={setForm} editIdx={editIdx} lastInsp={lastInsp}
                lastDate={lastDate} setStep={setStep} setView={setView}
                sessionInfo={sessionInfo}
                undoneOnly={undoneOnly} setUndoneOnly={setUndoneOnly}
                devColumns={devColumns} devVisibleCols={devVisibleCols} inspectionMode={inspectionMode}
                checkFields={checkFields} setCheck={setCheck} handleSave={handleSave} complete={complete} missing={missing}
                visIn={visIn} visOut={visOut} visFields={visFields}
                activeCode={activeCode} numDisp={numDisp} limits={limits}
                onPress={onPress} onConfirm={onConfirm} onRowClick={onRowClick} moveActive={moveActive}
                rowRefs={rowRefs} listRef={listRef}
                ALL_FIELDS={ALL_FIELDS} vis={vis} isAbn={isAbn}
                focusSeq={focusSeq} isCheckCode={isCheckCode} setCheckAndAdvance={setCheckAndAdvance}
                onSwitchMode={onSwitchMode}
              />
            </div>
          </>)}
          </div>
        )}

        {/* LIST */}
        {view==="list" && (
          <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",padding:"10px 16px"}}>

            {/* ── データ一覧ヘッダー：タイトル＋CSV出力ボタン ── */}
            <div style={{flexShrink:0,display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
              <div style={{fontSize:15,fontWeight:800,color:C.navy}}>📋 データ一覧</div>
              <button onClick={handleCSVExportClick}
                style={{padding:"7px 14px",borderRadius:8,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,
                  background:filteredRows.length>0?C.teal:C.g200,color:filteredRows.length>0?C.white:C.g400}}>
                💾 CSV出力
              </button>
            </div>

            {/* ── 絞り込み：パネル式 ── */}
            {/* 1行目：件数＋すべて／入力済／未入力／室内機／室外機 */}
            <div style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,marginBottom:6,flexWrap:"wrap"}}>
              <span style={{fontSize:12,fontWeight:700,color:C.g600,marginRight:2}}>
                {devList.length>0?filteredRows.filter(rowDone).length+"/"+tRows.length+"台":tRows.length+"件"}
              </span>
              {["all","done","undone"].map(f=>(
                <button key={f} onClick={()=>setListFilter(f)}
                  style={{padding:"6px 12px",borderRadius:8,border:"1.5px solid "+(listFilter===f?C.blue:C.g200),cursor:"pointer",fontSize:12,fontWeight:700,
                    background:listFilter===f?C.blue:C.white,color:listFilter===f?C.white:C.g500}}>
                  {f==="all"?"すべて":f==="done"?"入力済":"未入力"}
                </button>
              ))}
              <span style={{width:1,height:20,background:C.g200,margin:"0 2px"}}/>
              {[["indoor","🏠 室内機",C.blue],["outdoor","🏭 室外機",C.teal]].map(([v,label,col])=>(
                <button key={v} onClick={()=>setTypeFilter(p=>p===v?null:v)}
                  style={{padding:"6px 12px",borderRadius:8,border:"1.5px solid "+(typeFilter===v?col:C.g200),cursor:"pointer",fontSize:12,fontWeight:700,
                    background:typeFilter===v?col:C.white,color:typeFilter===v?C.white:C.g500,whiteSpace:"nowrap"}}>
                  {label}
                </button>
              ))}
            </div>
            {/* 2行目：建物パネル */}
            {(()=>{
              const buildings = listBuildingKey ? [...new Set(tRows.map(r=>colValue(r,"建物")).filter(b=>b&&b!=="（未設定）"))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true})) : [];
              if(buildings.length===0) return null;
              return (
                <div style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,marginBottom:6,flexWrap:"wrap"}}>
                  <span style={{fontSize:10,fontWeight:700,color:C.g400,minWidth:32}}>建物：</span>
                  <button onClick={()=>{setBuildingFilter(null);setFloorFilter(null);}}
                    style={{padding:"5px 11px",borderRadius:7,border:"1.5px solid "+(!buildingFilter?C.navy:C.g200),cursor:"pointer",fontSize:11,fontWeight:700,
                      background:!buildingFilter?C.navy:C.white,color:!buildingFilter?C.white:C.g600}}>すべて</button>
                  {buildings.map(b=>(
                    <button key={b} onClick={()=>{setBuildingFilter(p=>{const next=p===b?null:b; setFloorFilter(null); return next;});}}
                      style={{padding:"5px 11px",borderRadius:7,border:"1.5px solid "+(buildingFilter===b?C.navy:C.g200),cursor:"pointer",fontSize:11,fontWeight:700,
                        background:buildingFilter===b?C.navy:C.white,color:buildingFilter===b?C.white:C.g600,whiteSpace:"nowrap"}}>
                      {b}
                    </button>
                  ))}
                </div>
              );
            })()}
            {/* 3行目：階パネル（選択中の建物に存在する階のみ表示） */}
            {(()=>{
              const floorsForBuilding = buildingFilter
                ? [...new Set(tRows.filter(r=>colValue(r,"建物")===buildingFilter).map(r=>r.floor).filter(Boolean))].sort()
                : _floors;
              if(floorsForBuilding.length===0) return null;
              return (
                <div style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,marginBottom:8,flexWrap:"wrap"}}>
                  <span style={{fontSize:10,fontWeight:700,color:C.g400,minWidth:32}}>階：</span>
                  <button onClick={()=>setFloorFilter(null)}
                    style={{padding:"5px 11px",borderRadius:7,border:"1.5px solid "+(!floorFilter?C.teal:C.g200),cursor:"pointer",fontSize:11,fontWeight:700,
                      background:!floorFilter?C.teal:C.white,color:!floorFilter?C.white:C.g600}}>すべて</button>
                  {floorsForBuilding.map(fl=>(
                    <button key={fl} onClick={()=>setFloorFilter(floorFilter===fl?null:fl)}
                      style={{padding:"5px 11px",borderRadius:7,border:"1.5px solid "+(floorFilter===fl?C.teal:C.g200),cursor:"pointer",fontSize:11,fontWeight:700,
                        background:floorFilter===fl?C.teal:C.white,color:floorFilter===fl?C.white:C.g600}}>
                      {fl}
                    </button>
                  ))}
                </div>
              );
            })()}

            {/* ── 階フィルター中ポップアップバー ── */}
            {floorFilter&&(
              <div style={{flexShrink:0,display:"flex",alignItems:"center",gap:8,marginBottom:6,padding:"8px 14px",background:"linear-gradient(90deg,"+C.teal+"18,"+C.teal+"08)",border:"1.5px solid "+C.teal,borderRadius:10}}>
                <span style={{fontSize:13,fontWeight:800,color:C.teal,flex:1}}>🏢 {floorFilter} のみ表示中</span>
                <button onClick={()=>setFloorFilter(null)}
                  style={{padding:"5px 12px",borderRadius:7,border:"1.5px solid "+C.teal,background:C.white,color:C.teal,cursor:"pointer",fontSize:11,fontWeight:700}}>
                  ✕ 解除
                </button>
                <button onClick={()=>setShowStats(true)}
                  style={{padding:"5px 12px",borderRadius:7,border:"none",background:C.teal,color:C.white,cursor:"pointer",fontSize:11,fontWeight:700}}>
                  📊 集計へ戻る
                </button>
              </div>
            )}

            {/* ── テーブル ── */}
            <div style={{flex:1,overflowX:"auto",overflowY:"auto",scrollbarGutter:"stable",background:C.white,borderRadius:12,boxShadow:"0 2px 10px rgba(0,0,0,0.07)"}}>
              {filteredRows.length===0 ? (
                <div style={{padding:60,textAlign:"center",color:C.g400}}>
                  <div style={{fontSize:36,marginBottom:8}}>📭</div>
                  <div style={{fontWeight:700,fontSize:14}}>{listFilter==="undone"?"未入力の機器はありません":"データがありません"}</div>
                </div>
              ) : (
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{position:"sticky",top:0,zIndex:2}}>
                      {[...(devVisibleCols.length>0?devVisibleCols:["階","部屋名","管理番号","機器番号"]),"室内機","室外機","点検日","点検者","運転","モード","風量","設定温度"].map(h=>{
                        const isSort = sortCol===h;
                        return (
                          <th key={h}
                            onClick={()=>{ if(sortCol===h){setSortDir(d=>d==="asc"?"desc":"asc");}else{setSortCol(h);setSortDir("asc");} }}
                            style={{background:isSort?C.blue+"18":C.g100,color:isSort?C.blue:C.g600,padding:"6px 6px",textAlign:"center",fontWeight:700,fontSize:10,
                              whiteSpace:"normal",wordBreak:"keep-all",borderBottom:"2px solid "+C.g200,cursor:"pointer",userSelect:"none",
                              maxWidth:64,lineHeight:1.3,verticalAlign:"bottom"}}>
                            {h.length>4?<>{h.slice(0,Math.ceil(h.length/2))}<br/>{h.slice(Math.ceil(h.length/2))}</>:h}{isSort?(sortDir==="asc"?" ▲":" ▼"):""}
                          </th>
                        );
                      })}
                      {vf.map((f,fi)=>(
                        <th key={f.code} style={{background:C.g100,padding:"8px 6px",textAlign:"center",fontWeight:700,fontSize:10,whiteSpace:"nowrap",borderBottom:"2px solid "+C.g200,color:f.group==="indoor"?C.blue:C.teal,borderLeft:fi===0?"2px solid "+C.g200:undefined}}>
                          {f.code}<br/><span style={{fontSize:8,fontWeight:400}}>{f.unit}</span>
                        </th>
                      ))}
                      <th style={{background:C.g100,color:C.g600,padding:"8px 6px",textAlign:"center",fontWeight:700,fontSize:10,whiteSpace:"nowrap",borderBottom:"2px solid "+C.g200,borderLeft:"2px solid "+C.g200}}>チェック結果</th>
                      <th style={{background:C.g100,color:C.g600,padding:"8px 8px",textAlign:"center",fontWeight:700,fontSize:10,borderBottom:"2px solid "+C.g200,borderLeft:"2px solid "+C.g200}}>備考</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.map((row,i)=>{
                      const ir=row.indoorRecord, or_=row.outdoorRecord;
                      const indoorDone=rowIndoorDone(row), outdoorDone=rowOutdoorDone(row);
                      const hasMeasure=indoorDone||outdoorDone;
                      const bg = hoverRow===i ? "#EFF6FF" : (!hasMeasure?(i%2===0?"#FAFAFA":"#F5F5F5"):i%2===0?C.white:C.g50);
                      const metaSrc = ir||or_;
                      const rType = rowStatsType(row);
                      const StatusCell = ({rec,done,mode,label}) => (
                        <td style={{padding:"5px 6px",textAlign:"center",borderBottom:"1px solid "+C.g100,background:bg}}>
                          <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                            {done
                              ? <span style={{background:C.green+"22",color:C.green,fontWeight:700,fontSize:9,padding:"2px 6px",borderRadius:5,whiteSpace:"nowrap"}}>入力済</span>
                              : <span style={{background:"#FEF2F2",color:C.red,fontWeight:700,fontSize:9,padding:"2px 6px",borderRadius:5,whiteSpace:"nowrap"}}>未入力</span>}
                            {rec ? (
                              <div style={{display:"flex",gap:2,justifyContent:"center"}}>
                                <button onClick={e=>{e.stopPropagation();handleEditFor(rec,mode);}} style={{padding:"2px 5px",borderRadius:4,border:"none",cursor:"pointer",background:C.blue,color:C.white,fontSize:9,fontWeight:700}}>編集</button>
                                <button onClick={e=>{e.stopPropagation();handleDelFor(rec,mode);}} style={{padding:"2px 5px",borderRadius:4,border:"none",cursor:"pointer",background:C.red,color:C.white,fontSize:9,fontWeight:700}}>削除</button>
                              </div>
                            ) : (
                              <button onClick={e=>{e.stopPropagation();handleInputFor(row,mode);}} style={{padding:"2px 6px",borderRadius:4,border:"1.5px solid "+C.blue,cursor:"pointer",background:"white",color:C.blue,fontSize:9,fontWeight:700,whiteSpace:"nowrap"}}>{label}入力</button>
                            )}
                          </div>
                        </td>
                      );
                      const NACell = () => (
                        <td style={{padding:"5px 6px",textAlign:"center",borderBottom:"1px solid "+C.g100,background:bg,color:C.g300,fontSize:11}}>—</td>
                      );
                      return (
                        <tr key={i} style={{opacity:hasMeasure?1:0.65,cursor:"pointer"}}
                          onMouseEnter={()=>setHoverRow(i)}
                          onMouseLeave={()=>setHoverRow(p=>p===i?null:p)}
                          onClick={()=>{ if(rType==="outdoor"){ handleEditOrInputFor(row,"outdoor"); } else { handleEditOrInputFor(row,"indoor"); } }}>
                          {(devVisibleCols.length>0?devVisibleCols:["階","部屋名","管理番号","機器番号"]).map((col,j)=>{
                            const val=row._raw?row._raw[col]
                              :(col==="階"?row.floor:col==="部屋名"?row.room:col==="管理番号"?row.managementNo:col==="機器番号"?row.unitNo:"");
                            return <td key={j} style={{padding:"7px 8px",textAlign:"center",borderBottom:"1px solid "+C.g100,background:bg,fontSize:11,whiteSpace:"nowrap",fontWeight:600}}>{val||"—"}</td>;
                          })}
                          {rType==="outdoor" ? <NACell/> : <StatusCell rec={ir} done={indoorDone} mode="indoor" label="🏠"/>}
                          {rType==="indoor" ? <NACell/> : <StatusCell rec={or_} done={outdoorDone} mode="outdoor" label="🏭"/>}
                          {[metaSrc?.inspectionDate,metaSrc?.inspector,metaSrc?.preOperation,metaSrc?.preMode,metaSrc?.preWind,metaSrc?.preSetTemp].map((v,j)=>(
                            <td key={j} style={{padding:"7px 8px",textAlign:"center",borderBottom:"1px solid "+C.g100,background:bg,fontSize:11,color:v?C.g800:C.g300,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:90}}>{v||"—"}</td>
                          ))}
                          {vf.map((f,fi)=>{
                            const val=rowFieldVal(row,f.code); const ab=val?isAbn(f.code,val,limits):false;
                            return <td key={f.code} style={{padding:"7px 7px",textAlign:"right",borderBottom:"1px solid "+C.g100,background:ab?"#FEF2F2":bg,fontFamily:"monospace",fontSize:11,color:ab?C.red:val?C.g800:C.g300,borderLeft:fi===0?"2px solid "+C.g200:undefined}}>{val||"—"}</td>;
                          })}
                          <td style={{padding:"6px 6px",textAlign:"center",borderBottom:"1px solid "+C.g100,background:bg,borderLeft:"2px solid "+C.g200,maxWidth:180}}>
                            <div style={{display:"flex",flexWrap:"nowrap",gap:2,justifyContent:"flex-start",overflowX:"auto",maxWidth:170}}>
                              {rowChecksList(row).map(c=>(
                                <span key={c.code} title={c.label}
                                  style={{display:"inline-flex",flexShrink:0,alignItems:"center",justifyContent:"center",width:16,height:16,borderRadius:3,fontSize:10,fontWeight:800,
                                    background:c.val==="○"?C.green+"22":c.val==="×"?C.red+"22":C.g100,
                                    color:c.val==="○"?C.green:c.val==="×"?C.red:C.g300}}>
                                  {c.val||"—"}
                                </span>
                              ))}
                            </div>
                          </td>
                          <td style={{padding:"7px 8px",textAlign:"center",borderBottom:"1px solid "+C.g100,background:bg,fontSize:11,borderLeft:"2px solid "+C.g200,color:rowRemarks(row)?C.g800:C.g300,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:150}}>{rowRemarks(row)||"—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}

        {/* SETTINGS */}
        {view==="settings" && (
          <div style={{flex:1,overflowY:"auto",scrollbarGutter:"stable",padding:"14px 16px 28px",display:"flex",flexDirection:"column",gap:12,maxWidth:840,margin:"0 auto",width:"100%",boxSizing:"border-box"}}>
            {/* 保存状態バナー */}
            <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",background:C.teal+"10",border:"1.5px solid "+C.teal+"40",borderRadius:12}}>
              <span style={{fontSize:18}}>💾</span>
              <div style={{flex:1,fontSize:11,color:C.g600,lineHeight:1.5}}>
                機器リスト・点検者リスト・点検項目・表示設定・正常値範囲は、このブラウザに自動保存されます。<br/>
                アプリを更新（再読込）しても読み込み直す必要はありません。
              </div>
              <button onClick={()=>{
                if(!window.confirm("機器リスト・点検者リスト・点検項目・表示設定・正常値範囲の保存データを全て削除します。よろしいですか？")) return;
                Object.values(LS_KEYS).forEach(k=>{try{localStorage.removeItem(k);}catch(e){}});
                setDevList([]);setDevColumns([]);setDevVisibleCols([]);setInspList([]);
                setLimits(defLim());setVis(defVis());
                setCheckFields([
                  {code:"ci1", label:"配管類支持異常の有無", category:"据付状態",       group:"check_in"},
                  {code:"ci2", label:"異音・異常振動の有無", category:"フィルター点検・清掃", group:"check_in"},
                  {code:"ci3", label:"異音・異常振動の有無", category:"運転確認",       group:"check_in"},
                  {code:"co1", label:"防振装置異常の有無",   category:"据付状態",       group:"check_out"},
                  {code:"co2", label:"配管類支持異常の有無", category:"据付状態",       group:"check_out"},
                  {code:"co3", label:"ガスリークテスト",     category:"冷媒系統",       group:"check_out"},
                  {code:"co4", label:"配管系統外観点検",     category:"冷媒系統",       group:"check_out"},
                  {code:"co5", label:"異音・異常振動の有無", category:"送排風機系統",   group:"check_out"},
                  {code:"co6", label:"ドレン配管異常の有無", category:"排水系統",       group:"check_out"},
                  {code:"co7", label:"フィン汚れの有無",     category:"熱交換器系統",   group:"check_out"},
                  {code:"co8", label:"異音・異常振動の有無", category:"熱交換器系統",   group:"check_out"},
                  {code:"co9", label:"外面清掃",             category:"作業終了時",     group:"check_out"},
                ]);
                showFlash("🗑️ 保存データを削除しました");
              }} style={{flexShrink:0,padding:"6px 10px",borderRadius:7,border:"1.5px solid "+C.g300,cursor:"pointer",background:C.white,color:C.g500,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>
                🗑️ 保存データを削除
              </button>
            </div>
            {/* グリッド形式のアイコンメニュー */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
              {[
                {id:"device",title:"機器リスト",icon:"📂",color:C.navy,badge:devList.length>0?devList.length+"件":null},
                {id:"cols",title:"表示列設定",icon:"🗂️",color:C.teal,badge:devVisibleCols.length>0?devVisibleCols.length+"列":devColumns.length>0?"全列":null},
                {id:"checkitems",title:"点検項目",icon:"✅",color:"#059669",badge:checkFields.length>0?checkFields.length+"項目":null},
                {id:"inspector",title:"点検者リスト",icon:"👤",color:C.blue,badge:inspList.length>0?inspList.length+"名":null},
                {id:"vis",title:"表示項目",icon:"👁️",color:C.teal,badge:null},
                {id:"lim",title:"正常値範囲",icon:"⚙️",color:C.purple,badge:null},
              ].map(({id,title,icon,color,badge})=>(
                <button key={id} onClick={()=>setModalSec(id)}
                  style={{background:C.white,borderRadius:14,boxShadow:"0 2px 10px rgba(0,0,0,0.07)",border:"none",cursor:"pointer",
                    padding:"18px 10px",display:"flex",flexDirection:"column",alignItems:"center",gap:8,transition:"all 0.15s"}}>
                  <span style={{fontSize:32}}>{icon}</span>
                  <span style={{fontSize:13,fontWeight:700,color:color}}>{title}</span>
                  {badge&&<span style={{fontSize:11,background:color,color:C.white,padding:"2px 8px",borderRadius:8,fontWeight:700}}>{badge}</span>}
                </button>
              ))}
            </div>

            {/* 設定モーダル */}
            {modalSec&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"flex-end",justifyContent:"center"}}
                onClick={()=>setModalSec(null)}>
                <div style={{background:C.white,borderRadius:"20px 20px 0 0",width:"100%",maxWidth:680,maxHeight:"80vh",display:"flex",flexDirection:"column",boxShadow:"0 -4px 30px rgba(0,0,0,0.2)"}}
                  onClick={e=>e.stopPropagation()}>
                  {/* モーダルヘッダー */}
                  {(()=>{
                    const sec={device:{title:"機器リスト CSV読込",icon:"📂",color:C.navy},cols:{title:"表示列設定",icon:"🗂️",color:C.teal},checkitems:{title:"点検項目 Excel読込",icon:"✅",color:"#059669"},inspector:{title:"点検者リスト CSV読込",icon:"👤",color:C.blue},vis:{title:"表示項目設定",icon:"👁️",color:C.teal},lim:{title:"正常値範囲設定",icon:"⚙️",color:C.purple}}[modalSec]||{};
                    return (
                      <div style={{background:"linear-gradient(135deg,"+C.navy+","+C.blue+")",borderRadius:"20px 20px 0 0",padding:"14px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:10}}>
                          <span style={{fontSize:20}}>{sec.icon}</span>
                          <span style={{fontSize:15,fontWeight:800,color:C.white}}>{sec.title}</span>
                        </div>
                        <button onClick={()=>setModalSec(null)} style={{background:"rgba(255,255,255,0.2)",border:"none",color:C.white,borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:18,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
                      </div>
                    );
                  })()}
                  {/* モーダルコンテンツ（スクロール） */}
                  <div style={{overflowY:"auto",scrollbarGutter:"stable",padding:"18px 20px",flex:1}}>
                    {[{id:"device"},{id:"cols"},{id:"checkitems"},{id:"inspector"},{id:"vis"},{id:"lim"}].map(({id})=>modalSec===id&&(
                      <div key={id}>
                      {id==="device" && (
                        <div style={{display:"flex",flexDirection:"column",gap:14}}>
                          <div>
                            <p style={{fontSize:13,color:C.g500,marginBottom:10,lineHeight:1.6}}>1行目を項目名として自動認識します。</p>
                            <input ref={devRef} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={handleDevCSV}/>
                            <div style={{display:"flex",gap:10,alignItems:"center"}}>
                              <button onClick={()=>devRef.current.click()} style={{padding:"10px 20px",borderRadius:9,border:"none",cursor:"pointer",background:C.blue,color:C.white,fontWeight:700,fontSize:13}}>📁 CSVを選択</button>
                              {devList.length>0 && <span style={{fontSize:13,color:C.green,fontWeight:700}}>✅ {devList.length}件読込済</span>}
                            </div>
                          </div>
                          {devList.length>0 && (
                            <div>
                              <div style={{fontSize:11,fontWeight:700,color:C.g500,marginBottom:6}}>プレビュー（先頭5件）</div>
                              <div style={{overflowX:"auto"}}>
                                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                                  <thead><tr>{(devVisibleCols.length>0?devVisibleCols:devColumns).map(h=><th key={h} style={{background:C.g100,padding:"6px 8px",textAlign:"center",fontWeight:700,fontSize:11,color:C.g500,borderBottom:"2px solid "+C.g200,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                                  <tbody>{devList.slice(0,5).map((d,i)=>{
                                    const cols=devVisibleCols.length>0?devVisibleCols:devColumns;
                                    return <tr key={i}>{cols.map((col,j)=>{
                                      const val=d._raw?d._raw[col]:d[{階:"floor",部屋名:"room",管理番号:"managementNo",機器番号:"unitNo"}[col]||col]||"";
                                      return <td key={j} style={{padding:"6px 8px",textAlign:"center",borderBottom:"1px solid "+C.g100,background:i%2===0?C.white:C.g50}}>{val}</td>;
                                    })}</tr>;
                                  })}{devList.length>5&&<tr><td colSpan={(devVisibleCols.length||devColumns.length)||4} style={{padding:"6px",textAlign:"center",color:C.g400,fontSize:11}}>…他 {devList.length-5}件</td></tr>}</tbody>
                                </table>
                              </div>
                            </div>
                          )}
                        </div>
                      )}
                      {id==="cols" && (
                        <div>
                          {devColumns.length===0 ? (
                            <div style={{padding:"12px 16px",background:"#FFF7ED",borderRadius:8,fontSize:13,color:"#92400E",border:"1.5px solid #F59E0B"}}>
                              ⚠️ 先に機器リストCSVを読み込んでください。読み込み後に列が自動抽出されます。
                            </div>
                          ) : (
                            <>
                              <p style={{fontSize:12,color:C.g500,marginBottom:10,lineHeight:1.6}}>
                                チェックした列を「入室可否チェック」と「機器選択リスト」に表示します。<br/>
                                ※ 「本日の対象エリア」の階ボタンには影響しません。
                              </p>
                              <div style={{display:"flex",gap:6,marginBottom:10}}>
                                <button onClick={()=>setDevVisibleCols([...devColumns])}
                                  style={{padding:"5px 12px",borderRadius:6,border:"1.5px solid "+C.teal,background:C.teal+"18",color:C.teal,fontSize:11,fontWeight:700,cursor:"pointer"}}>全選択</button>
                                <button onClick={()=>setDevVisibleCols([])}
                                  style={{padding:"5px 12px",borderRadius:6,border:"1.5px solid "+C.g300,background:C.g100,color:C.g500,fontSize:11,fontWeight:700,cursor:"pointer"}}>全解除</button>
                              </div>
                              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                                {devColumns.map(col=>{
                                  const active=devVisibleCols.includes(col);
                                  return (
                                    <label key={col} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 14px",borderRadius:9,
                                      background:active?C.teal+"0A":C.g50,border:"1.5px solid "+(active?C.teal:C.g200),cursor:"pointer"}}>
                                      <input type="checkbox" checked={active} onChange={e=>{
                                        setDevVisibleCols(p=>e.target.checked?[...p,col]:p.filter(c=>c!==col));
                                      }} style={{width:16,height:16,accentColor:C.teal,cursor:"pointer"}}/>
                                      <span style={{fontSize:13,fontWeight:active?700:400,color:active?C.navy:C.g600,flex:1}}>{col}</span>
                                      <span style={{fontSize:11,color:C.g400}}>{devList.filter(d=>d._raw&&d._raw[col]).length}件に値あり</span>
                                    </label>
                                  );
                                })}
                              </div>
                            </>
                          )}
                        </div>
                      )}
                      {id==="checkitems" && (
                        <div>
                          <p style={{fontSize:13,color:C.g500,marginBottom:10,lineHeight:1.6}}>
                            Excelファイルを読み込みます。<br/>
                            1列目に項目名、2列目にカテゴリ（省略可）を縦に記載してください。<br/>
                            例：<code style={{background:C.g100,padding:"1px 6px",borderRadius:4,fontSize:12}}>配管支持異常の有無 / 室内機</code>
                          </p>
                          <input
                            type="file" accept=".xlsx,.xls,.csv"
                            style={{display:"none"}}
                            id="checkItemsFile"
                            onChange={e=>{
                              const f=e.target.files[0]; if(!f) return;
                              readFile(f,(data,cols)=>{
                                let rows=[];
                                if(Array.isArray(data)){
                                  rows=data;
                                } else {
                                  const lines=(data||"").trim().split(/\r?\n/).filter(Boolean);
                                  rows=lines.map(l=>{const c=l.split(",").map(s=>s.trim().replace(/^"|"$/g,""));return {label:c[0]||"",category:c[1]||""};});
                                }
                                const fields=rows.map((row,i)=>{
                                  const label=String(Array.isArray(data)?Object.values(row)[0]:row.label||"").trim();
                                  const category=String(Array.isArray(data)?(Object.values(row)[1]||""):row.category||"").trim();
                                  const group=/室内|indoor/i.test(category)?"check_in":"check_out";
                                  return label?{code:"ck"+(i+1),label,category:category||"点検",group}:null;
                                }).filter(Boolean);
                                setCheckFields(fields);
                                showFlash("✅ 点検項目 "+fields.length+"件 読込");
                              });
                              e.target.value="";
                            }}
                          />
                          <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12}}>
                            <button onClick={()=>document.getElementById("checkItemsFile").click()}
                              style={{padding:"10px 20px",borderRadius:9,border:"none",cursor:"pointer",background:"#059669",color:C.white,fontWeight:700,fontSize:13}}>
                              📁 Excelを選択
                            </button>
                            {checkFields.length>0 && <button onClick={()=>setCheckFields([])}
                              style={{padding:"8px 14px",borderRadius:8,border:"1.5px solid "+C.red,background:"#FEF2F2",color:C.red,cursor:"pointer",fontSize:12,fontWeight:700}}>
                              🗑️ クリア
                            </button>}
                            {checkFields.length>0 && <span style={{fontSize:13,color:C.green,fontWeight:700}}>✅ {checkFields.length}項目読込済</span>}
                          </div>
                          {checkFields.length>0 && (
                            <div style={{display:"flex",flexDirection:"column",gap:2,maxHeight:300,overflowY:"auto",scrollbarGutter:"stable"}}>
                              {(()=>{
                                const cats=[...new Set(checkFields.map(f=>f.category))];
                                return cats.map(cat=>(
                                  <div key={cat} style={{marginBottom:6}}>
                                    <div style={{fontSize:11,fontWeight:700,color:"#059669",marginBottom:3,padding:"2px 6px",background:"#ECFDF5",borderRadius:4,display:"inline-block"}}>{cat}</div>
                                    {checkFields.filter(f=>f.category===cat).map(f=>(
                                      <div key={f.code} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 10px",borderRadius:7,background:C.g50,marginBottom:2}}>
                                        <span style={{fontSize:13,color:C.g700,flex:1}}>{f.label}</span>
                                        <span style={{fontSize:10,color:C.g400}}>{f.code}</span>
                                      </div>
                                    ))}
                                  </div>
                                ));
                              })()}
                            </div>
                          )}
                        </div>
                      )}
                      {id==="inspector" && (
                        <div>
                          <p style={{fontSize:13,color:C.g500,marginBottom:12}}>フォーマット：1行に1名（ヘッダーなし）</p>
                          <input ref={inspRef} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={handleInspCSV}/>
                          <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                            <button onClick={()=>inspRef.current.click()} style={{padding:"11px 22px",borderRadius:9,border:"none",cursor:"pointer",background:C.blue,color:C.white,fontWeight:700,fontSize:14}}>📁 CSVを選択</button>
                            {inspList.length>0 && <span style={{fontSize:13,color:C.green,fontWeight:700}}>✅ {inspList.join("・")}</span>}
                          </div>
                        </div>
                      )}

                      {id==="vis" && (
                        <div>
                          <p style={{fontSize:13,color:C.g500,marginBottom:14}}>チェックを外した項目は入力フォーム・一覧から非表示になります。</p>
                          {[{fields:INDOOR_FIELDS,label:"室内機（インドア）",color:C.blue},{fields:OUTDOOR_FIELDS,label:"室外機（アウトドア）",color:C.teal}].map(({fields,label,color})=>(
                            <div key={label} style={{marginBottom:18}}>
                              <div style={{fontWeight:700,fontSize:13,color,borderBottom:"2px solid "+color,paddingBottom:5,marginBottom:10}}>{label}</div>
                              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                                {fields.map(f=>(
                                  <label key={f.code} style={{display:"flex",alignItems:"center",gap:14,padding:"10px 14px",borderRadius:9,background:tmpVis[f.code]?color+"0A":C.g50,border:"1.5px solid "+(tmpVis[f.code]?color:C.g200),cursor:"pointer"}}>
                                    <input type="checkbox" checked={tmpVis[f.code]} onChange={e=>setTmpVis(p=>({...p,[f.code]:e.target.checked}))} style={{width:18,height:18,accentColor:color,cursor:"pointer"}}/>
                                    <span style={{fontFamily:"monospace",fontWeight:700,color,fontSize:14,minWidth:34}}>{f.code}</span>
                                    <span style={{fontSize:15,color:C.g600,flex:1}}>{f.label}</span>
                                    <span style={{fontSize:11,color:C.g400,background:C.g100,padding:"2px 8px",borderRadius:5}}>{f.unit}</span>
                                  </label>
                                ))}
                              </div>
                            </div>
                          ))}
                          <div style={{display:"flex",gap:10,marginTop:4}}>
                            <button onClick={()=>{setVis({...tmpVis});showFlash("✅ 表示設定を保存しました");}} style={{padding:"11px 26px",borderRadius:9,border:"none",cursor:"pointer",background:C.teal,color:C.white,fontWeight:700,fontSize:14}}>保存する</button>
                            <button onClick={()=>setTmpVis(defVis())} style={{padding:"11px 16px",borderRadius:9,border:"1.5px solid "+C.g200,cursor:"pointer",background:C.white,color:C.g500,fontSize:13}}>全選択</button>
                          </div>
                        </div>
                      )}
                      {id==="lim" && (
                        <div>
                          <p style={{fontSize:13,color:C.g500,marginBottom:14}}>✅ をチェックすると最小・最大の入力欄が表示されます。</p>
                          {[{fields:INDOOR_FIELDS,label:"室内機（インドア）",color:C.blue},{fields:OUTDOOR_FIELDS,label:"室外機（アウトドア）",color:C.teal}].map(({fields,label,color})=>(
                            <div key={label} style={{marginBottom:18}}>
                              <div style={{fontWeight:700,fontSize:13,color,borderBottom:"2px solid "+color,paddingBottom:5,marginBottom:10}}>{label}</div>
                              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                                {fields.map(f=>(
                                  <div key={f.code} style={{borderRadius:10,overflow:"hidden",border:"1.5px solid "+(tmpLim[f.code]?.enabled?C.purple:C.g200)}}>
                                    <label style={{display:"flex",alignItems:"center",gap:14,padding:"10px 14px",background:tmpLim[f.code]?.enabled?C.purple+"0A":C.g50,cursor:"pointer"}}>
                                      <input type="checkbox" checked={!!tmpLim[f.code]?.enabled} onChange={e=>setTmpLim(p=>({...p,[f.code]:{...p[f.code],enabled:e.target.checked}}))} style={{width:18,height:18,accentColor:C.purple,cursor:"pointer"}}/>
                                      <span style={{fontFamily:"monospace",fontWeight:700,color,fontSize:14,minWidth:34}}>{f.code}</span>
                                      <span style={{fontSize:15,color:C.g600,flex:1}}>{f.label}</span>
                                      <span style={{fontSize:11,color:C.g400,background:C.g100,padding:"2px 8px",borderRadius:5}}>{f.unit}</span>
                                    </label>
                                    {tmpLim[f.code]?.enabled && (
                                      <div style={{padding:"12px 14px 14px 46px",background:C.white,display:"flex",gap:12,alignItems:"center",borderTop:"1px solid "+C.g200}}>
                                        <span style={{fontSize:13,color:C.g500,minWidth:36}}>最小</span>
                                        <input type="number" step={f.step} placeholder="—" value={tmpLim[f.code]?.min||""} onChange={e=>setTmpLim(p=>({...p,[f.code]:{...p[f.code],min:e.target.value}}))} style={{width:90,padding:"8px 10px",borderRadius:7,border:"1.5px solid "+C.g200,fontSize:15,fontFamily:"monospace",textAlign:"right",outline:"none"}}/>
                                        <span style={{fontSize:13,color:C.g300}}>〜</span>
                                        <span style={{fontSize:13,color:C.g500,minWidth:36}}>最大</span>
                                        <input type="number" step={f.step} placeholder="—" value={tmpLim[f.code]?.max||""} onChange={e=>setTmpLim(p=>({...p,[f.code]:{...p[f.code],max:e.target.value}}))} style={{width:90,padding:"8px 10px",borderRadius:7,border:"1.5px solid "+C.g200,fontSize:15,fontFamily:"monospace",textAlign:"right",outline:"none"}}/>
                                      </div>
                                    )}
                                  </div>
                                ))}
                              </div>
                            </div>
                          ))}
                          <div style={{display:"flex",gap:10,marginTop:4}}>
                            <button onClick={()=>{setLimits(JSON.parse(JSON.stringify(tmpLim)));showFlash("✅ 正常値範囲を保存しました");}} style={{padding:"11px 26px",borderRadius:9,border:"none",cursor:"pointer",background:C.purple,color:C.white,fontWeight:700,fontSize:14}}>保存する</button>
                            <button onClick={()=>setTmpLim(defLim())} style={{padding:"11px 16px",borderRadius:9,border:"1.5px solid "+C.g200,cursor:"pointer",background:C.white,color:C.g500,fontSize:13}}>リセット</button>
                          </div>
                        </div>
                      )}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* print */}
      <div id="print-area" style={{display:"none"}}>
        <style>{PS}</style>
        <div style={{fontFamily:"Hiragino Sans, Meiryo, sans-serif"}}>
          <h2 style={{textAlign:"center",fontSize:14,marginBottom:10}}>エアコン点検データ一覧</h2>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:8}}>
            <thead>
              <tr style={{background:"#1B3A6B",color:"white"}}>
                {["階","部屋名","管理番号","機器番号","室内機","室外機","点検日","点検者","運転","モード","風量","設定温度",...vf.map(f=>f.code+"("+f.unit+")"),"備考"].map((h,i)=>(
                  <th key={i} style={{border:"1px solid #ccc",padding:"3px 4px",textAlign:"center",fontSize:7}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {tRows.map((row,i)=>(
                <tr key={i} style={{background:i%2===0?"white":"#f8fafc"}}>
                  {[row.floor,row.room,row.managementNo,row.unitNo,rowIndoorDone(row)?"入力済":"未入力",rowOutdoorDone(row)?"入力済":"未入力",rowMeta(row,"inspectionDate"),rowMeta(row,"inspector"),rowMeta(row,"preOperation"),rowMeta(row,"preMode"),rowMeta(row,"preWind"),rowMeta(row,"preSetTemp")].map((v,j)=>(
                    <td key={j} style={{border:"1px solid #ddd",padding:"3px 4px",textAlign:"center",fontSize:7}}>{v}</td>
                  ))}
                  {vf.map(f=>{const val=rowFieldVal(row,f.code);const ab=val?isAbn(f.code,val,limits):false;return(
                    <td key={f.code} style={{border:"1px solid #ddd",padding:"3px 4px",textAlign:"right",fontSize:7,color:ab?"red":undefined,fontWeight:ab?"bold":undefined}}>{val||""}</td>
                  );})}
                  <td style={{border:"1px solid #ddd",padding:"3px 4px",fontSize:7}}>{rowRemarks(row)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{marginTop:8,fontSize:8,color:"#666"}}>出力：{new Date().toLocaleString("ja-JP")}</div>
        </div>
      </div>

      {/* ── 集計モーダル ── */}
      {showStats && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={()=>setShowStats(false)}>
          <div style={{background:C.white,borderRadius:20,width:"100%",maxWidth:560,boxShadow:"0 8px 40px rgba(0,0,0,0.3)",overflow:"hidden",maxHeight:"85vh",overflowY:"auto",scrollbarGutter:"stable"}} onClick={e=>e.stopPropagation()}>
            <div style={{background:"linear-gradient(135deg,"+C.navy+","+C.blue+")",padding:"16px 22px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:22}}>📊</span>
                <div style={{fontSize:16,fontWeight:800,color:C.white}}>集計（建物・階・室内機／室外機）</div>
              </div>
              <button onClick={()=>setShowStats(false)} style={{background:"rgba(255,255,255,0.2)",border:"none",color:C.white,borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
            </div>
            <div style={{padding:"16px 20px",display:"flex",flexDirection:"column",gap:10}}>
              {(()=>{
                const total=tRows.length;
                const totalDone=tRows.filter(rowDone).length;
                const indoorTotal=groupStats.reduce((s,g)=>s+g.indoorTotal,0);
                const indoorDone=groupStats.reduce((s,g)=>s+g.indoorDone,0);
                const outdoorTotal=groupStats.reduce((s,g)=>s+g.outdoorTotal,0);
                const outdoorDone=groupStats.reduce((s,g)=>s+g.outdoorDone,0);
                return (
                  <>
                    <div style={{background:C.g50,borderRadius:12,padding:"12px 16px",display:"flex",flexDirection:"column",gap:6,border:"2px solid "+C.g200}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <span style={{fontWeight:700,fontSize:14,color:C.navy}}>合計</span>
                        <div style={{display:"flex",gap:16,alignItems:"center"}}>
                          <span style={{fontSize:13,color:C.g500}}>{total}台</span>
                          <span style={{fontSize:15,fontWeight:800,color:C.green}}>{totalDone}✓</span>
                          {total-totalDone>0&&<span style={{fontSize:13,fontWeight:700,color:C.red}}>未{total-totalDone}</span>}
                          <div style={{width:80,height:8,background:C.g200,borderRadius:4,overflow:"hidden"}}>
                            <div style={{width:(total>0?Math.round(totalDone/total*100):0)+"%",height:"100%",background:C.green,borderRadius:4}}/>
                          </div>
                          <span style={{fontSize:12,color:C.g500,minWidth:34,textAlign:"right"}}>{total>0?Math.round(totalDone/total*100):0}%</span>
                        </div>
                      </div>
                      <div style={{display:"flex",gap:16,fontSize:12,color:C.g600,paddingLeft:2}}>
                        <span>🏠 室内機：{indoorDone}/{indoorTotal}</span>
                        <span>🏭 室外機：{outdoorDone}/{outdoorTotal}</span>
                      </div>
                    </div>
                    {groupStats.length>0 && (
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead>
                            <tr>
                              <th style={{textAlign:"left",padding:"6px 8px",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>建物</th>
                              <th style={{textAlign:"left",padding:"6px 8px",fontSize:10,fontWeight:700,color:C.g500,borderBottom:"2px solid "+C.g200}}>階</th>
                              <th style={{textAlign:"center",padding:"6px 8px",fontSize:10,fontWeight:700,color:C.blue,borderBottom:"2px solid "+C.g200}}>🏠 室内機</th>
                              <th style={{textAlign:"center",padding:"6px 8px",fontSize:10,fontWeight:700,color:C.teal,borderBottom:"2px solid "+C.g200}}>🏭 室外機</th>
                            </tr>
                          </thead>
                          <tbody>
                            {groupStats.map((g,i)=>(
                              <tr key={i}
                                onClick={()=>{setFloorFilter(g.floor==="—"?null:g.floor);setListFilter("all");setView("list");setShowStats(false);}}
                                style={{cursor:"pointer",background:i%2===0?C.white:C.g50}}>
                                <td style={{padding:"7px 8px",borderBottom:"1px solid "+C.g100,fontWeight:700,color:C.navy}}>{g.building}</td>
                                <td style={{padding:"7px 8px",borderBottom:"1px solid "+C.g100,fontWeight:700,color:C.navy}}>{g.floor}</td>
                                <td style={{padding:"7px 8px",borderBottom:"1px solid "+C.g100,textAlign:"center"}}>
                                  {g.indoorTotal>0 ? (
                                    <span style={{color:g.indoorDone===g.indoorTotal?C.green:C.g600,fontWeight:700}}>{g.indoorDone}/{g.indoorTotal}</span>
                                  ) : <span style={{color:C.g300}}>—</span>}
                                </td>
                                <td style={{padding:"7px 8px",borderBottom:"1px solid "+C.g100,textAlign:"center"}}>
                                  {g.outdoorTotal>0 ? (
                                    <span style={{color:g.outdoorDone===g.outdoorTotal?C.green:C.g600,fontWeight:700}}>{g.outdoorDone}/{g.outdoorTotal}</span>
                                  ) : <span style={{color:C.g300}}>—</span>}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                    {groupStats.length===0&&<div style={{textAlign:"center",padding:"24px 0",color:C.g400,fontSize:13}}>データがありません</div>}
                  </>
                );
              })()}
            </div>
          </div>
        </div>
      )}

      {/* ── CSV出力前の確認ポップアップ（絞り込み中のときだけ表示） ── */}
      {showExportConfirm && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={()=>setShowExportConfirm(null)}>
          <div style={{background:C.white,borderRadius:16,width:"100%",maxWidth:380,boxShadow:"0 8px 40px rgba(0,0,0,0.3)",overflow:"hidden"}} onClick={e=>e.stopPropagation()}>
            <div style={{background:"linear-gradient(135deg,#F59E0B,#D97706)",padding:"14px 18px"}}>
              <div style={{fontSize:14,fontWeight:800,color:C.white}}>⚠️ 絞り込みが選択されています</div>
            </div>
            <div style={{padding:"16px 18px"}}>
              <div style={{fontSize:13,color:C.g700,marginBottom:10,lineHeight:1.6}}>
                データ一覧で以下の絞り込み・並び替えが選択された状態です。<br/>この内容のままCSV出力してよろしいですか？
              </div>
              <div style={{background:C.g50,borderRadius:10,padding:"10px 12px",marginBottom:14,display:"flex",flexDirection:"column",gap:5}}>
                {showExportConfirm.map((desc,i)=>(
                  <div key={i} style={{fontSize:12,fontWeight:700,color:C.navy}}>・{desc}</div>
                ))}
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setShowExportConfirm(null)}
                  style={{flex:1,padding:"10px",borderRadius:9,border:"1.5px solid "+C.g300,background:C.g50,color:C.g600,fontWeight:700,fontSize:13,cursor:"pointer"}}>いいえ</button>
                <button onClick={()=>{exportListAsCSV();setShowExportConfirm(null);}}
                  style={{flex:1,padding:"10px",borderRadius:9,border:"none",background:C.teal,color:C.white,fontWeight:700,fontSize:13,cursor:"pointer"}}>はい・出力する</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {flash && (
        <div style={{position:"fixed",bottom:24,right:24,background:C.green,color:C.white,padding:"12px 24px",borderRadius:10,fontWeight:700,fontSize:15,boxShadow:"0 4px 16px rgba(0,0,0,0.2)",zIndex:9999}}>
          {flash}
        </div>
      )}

      {/* ── 保存完了モーダル ── */}
      {saveModal && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div style={{background:C.white,borderRadius:20,width:"100%",maxWidth:560,boxShadow:"0 8px 40px rgba(0,0,0,0.3)",overflow:"hidden",maxHeight:"90vh",overflowY:"auto",scrollbarGutter:"stable"}}>

            {/* ヘッダー */}
            <div style={{background:"linear-gradient(135deg,"+C.green+",#047857)",padding:"14px 18px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:24}}>✅</span>
                <div>
                  <div style={{fontSize:16,fontWeight:800,color:C.white}}>保存しました</div>
                  <div style={{fontSize:11,color:"rgba(255,255,255,0.8)",marginTop:1}}>{saveModal.floor}　{saveModal.room}　{saveModal.managementNo} / {saveModal.unitNo}</div>
                </div>
              </div>
              <button onClick={closeNext}
                style={{background:"rgba(255,255,255,0.2)",border:"none",color:C.white,borderRadius:8,width:34,height:34,cursor:"pointer",fontSize:18,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>
                ✕
              </button>
            </div>

            <div style={{padding:"14px 18px",display:"flex",flexDirection:"column",gap:12}}>

              {/* 測定データ（タップで拡大） */}
              <div onClick={()=>setMeasZoom(true)} style={{background:C.g50,borderRadius:10,padding:"10px 14px",cursor:"pointer",border:"2px solid "+C.g200,transition:"border-color 0.15s",userSelect:"none"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                  <div style={{fontSize:11,fontWeight:700,color:C.g500}}>📊 測定データ</div>
                  <div style={{fontSize:10,color:C.blue,fontWeight:700}}>タップで拡大 ▶</div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:"4px 12px"}}>
                  {ALL_FIELDS.filter(f=>vis[f.code]&&saveModal.values[f.code]!=="").map(f=>{
                    const abn=isAbn(f.code,saveModal.values[f.code],limits);
                    return (
                      <div key={f.code} style={{display:"flex",alignItems:"baseline",gap:5,padding:"3px 0",borderBottom:"1px solid "+C.g200}}>
                        <span style={{fontSize:11,fontFamily:"monospace",fontWeight:700,color:f.group==="indoor"?C.blue:C.teal,minWidth:24}}>{f.code}</span>
                        <span style={{fontSize:10,color:C.g500,flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{f.label}</span>
                        <span style={{fontSize:13,fontWeight:700,fontFamily:"monospace",color:abn?C.red:C.navy,whiteSpace:"nowrap"}}>
                          {saveModal.values[f.code]}<span style={{fontSize:9,color:C.g400}}> {f.unit}</span>
                          {abn&&<span style={{fontSize:9,color:C.red}}> ⚠️</span>}
                        </span>
                      </div>
                    );
                  })}
                </div>
                {saveModal.remarks&&<div style={{marginTop:8,fontSize:12,color:C.g600,borderTop:"1px solid "+C.g200,paddingTop:6}}>備考: {saveModal.remarks}</div>}
              </div>

              {/* 点検チェック項目（○×） */}
              {(()=>{
                const chkGroup = inspectionMode==="outdoor" ? "check_out" : "check_in";
                const chkList = checkFields.filter(f=>f.group===chkGroup);
                if(chkList.length===0) return null;
                return (
                  <div style={{background:C.g50,borderRadius:10,padding:"10px 14px",border:"2px solid "+C.g200}}>
                    <div style={{fontSize:11,fontWeight:700,color:C.g500,marginBottom:8}}>✅ 点検チェック項目</div>
                    <div style={{display:"flex",flexDirection:"column",gap:4}}>
                      {chkList.map(f=>{
                        const v=saveModal.checks?.[f.code]||"";
                        return (
                          <div key={f.code} style={{display:"flex",alignItems:"center",gap:8,padding:"3px 0",borderBottom:"1px solid "+C.g200}}>
                            <span style={{fontSize:10,fontWeight:700,color:C.teal,minWidth:70,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{f.category}</span>
                            <span style={{fontSize:12,color:C.g700,flex:1}}>{f.label}</span>
                            <span style={{fontSize:16,fontWeight:800,color:v==="○"?C.green:v==="×"?C.red:C.g300,minWidth:20,textAlign:"center"}}>{v||"—"}</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}

              {/* リモコン復元 or 次へボタン */}
              {(saveModal.preOperation||saveModal.preMode||saveModal.preWind||saveModal.preSetTemp) ? (
                <div style={{background:"#FFF7ED",border:"2px solid #F59E0B",borderRadius:12,padding:"14px 16px"}}>
                  <div style={{fontSize:13,fontWeight:800,color:"#92400E",marginBottom:10}}>⚠️ リモコンを点検前の状態に戻してください</div>
                  <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
                    {[
                      {show:!!saveModal.preOperation,label:"運転",val:saveModal.preOperation==="ON"?"🟢 ON":"⭕ OFF"},
                      {show:!!saveModal.preMode,label:"モード",val:modeLabel(saveModal.preMode)},
                      {show:!!saveModal.preWind,label:"風量",val:windLabel(saveModal.preWind)},
                      {show:!!saveModal.preSetTemp,label:"設定温度",val:saveModal.preSetTemp+"°C"},
                    ].filter(p=>p.show).map(p=>(
                      <div key={p.label} style={{flex:1,minWidth:70,background:C.white,border:"2px solid #F59E0B",borderRadius:10,padding:"8px 10px",textAlign:"center"}}>
                        <div style={{fontSize:10,fontWeight:700,color:"#92400E",marginBottom:4}}>{p.label}</div>
                        <div style={{fontSize:15,fontWeight:800,color:"#78350F",fontFamily:p.label==="設定温度"?"monospace":"inherit"}}>{p.val}</div>
                      </div>
                    ))}
                  </div>
                  <button onClick={closeNext}
                    style={{width:"100%",padding:"16px",borderRadius:10,border:"none",cursor:"pointer",
                      background:"linear-gradient(135deg,"+C.green+",#047857)",color:C.white,fontWeight:800,fontSize:16,
                      boxShadow:"0 3px 10px rgba(5,150,105,0.3)"}}>
                    ✅ 戻しました　→　{nextLabel}
                  </button>
                </div>
              ) : (
                <button onClick={closeNext}
                  style={{width:"100%",padding:"16px",borderRadius:10,border:"none",cursor:"pointer",
                    background:"linear-gradient(135deg,"+C.navy+","+C.blue+")",color:C.white,fontWeight:800,fontSize:16,
                    boxShadow:"0 3px 10px rgba(37,99,176,0.3)"}}>
                  ✅ {nextLabel}
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ── 測定データ拡大モーダル ── */}
      {measZoom && saveModal && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",zIndex:10001,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={()=>setMeasZoom(false)}>
          <div style={{background:C.white,borderRadius:20,width:"100%",maxWidth:600,boxShadow:"0 8px 40px rgba(0,0,0,0.4)",overflow:"hidden",maxHeight:"92vh",display:"flex",flexDirection:"column"}} onClick={e=>e.stopPropagation()}>
            <div style={{background:"linear-gradient(135deg,"+C.navy+","+C.blue+")",padding:"12px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
              <div style={{fontSize:15,fontWeight:800,color:C.white}}>📊 測定データ　{saveModal.floor} {saveModal.room}</div>
              <button onClick={()=>setMeasZoom(false)}
                style={{background:"rgba(255,255,255,0.2)",border:"none",color:C.white,borderRadius:8,width:34,height:34,cursor:"pointer",fontSize:18,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center"}}>
                ✕
              </button>
            </div>
            <div style={{overflowY:"auto",scrollbarGutter:"stable",padding:"12px 16px",display:"flex",flexDirection:"column",gap:3}}>
              {ALL_FIELDS.filter(f=>vis[f.code]&&saveModal.values[f.code]!=="").map(f=>{
                const abn=isAbn(f.code,saveModal.values[f.code],limits);
                return (
                  <div key={f.code} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,background:abn?"#FEF2F2":f.group==="indoor"?C.blue+"08":C.teal+"08",border:"1.5px solid "+(abn?C.red:f.group==="indoor"?C.blue+"30":C.teal+"30")}}>
                    <span style={{fontFamily:"monospace",fontWeight:800,fontSize:18,color:f.group==="indoor"?C.blue:C.teal,minWidth:34}}>{f.code}</span>
                    <span style={{fontSize:15,color:C.g600,flex:1}}>{f.label}</span>
                    <span style={{fontFamily:"monospace",fontWeight:800,fontSize:26,color:abn?C.red:C.navy}}>
                      {saveModal.values[f.code]}
                    </span>
                    <span style={{fontSize:13,color:C.g400,minWidth:38}}>{f.unit}{abn&&" ⚠️"}</span>
                  </div>
                );
              })}
              {saveModal.remarks&&<div style={{marginTop:6,padding:"10px 14px",borderRadius:10,background:C.g50,fontSize:13,color:C.g600}}>備考: {saveModal.remarks}</div>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── iPad実機サイズプレビュー枠 ─────────────────────────────────
// iPad 10.2インチ(第9世代) 縦向き論理解像度: 810 x 1080pt を基準
export default function IPadFrame() {
  const IPAD_W = 810;
  const IPAD_H = 1080;
  return (
    <div style={{
      width: "100%", minHeight: "100vh",
      background: "#0F172A",
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: "24px 12px",
      fontFamily: "'Hiragino Sans','Hiragino Kaku Gothic ProN',Meiryo,sans-serif",
    }}>
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 10 }}>
        {/* ラベル */}
        <div style={{ color: "#94A3B8", fontSize: 12, fontWeight: 700, letterSpacing: "0.05em" }}>
          iPad 10.2インチ相当　{IPAD_W} × {IPAD_H}px（縦向き）
        </div>
        {/* iPad筐体風フレーム */}
        <div style={{
          width: IPAD_W + 28, height: IPAD_H + 28,
          background: "#1E293B",
          borderRadius: 40,
          padding: 14,
          boxShadow: "0 20px 60px rgba(0,0,0,0.5), inset 0 0 0 2px #334155",
          display: "flex", alignItems: "center", justifyContent: "center",
        }}>
          {/* 画面領域（実サイズ固定・スクロールはアプリ内部のみ） */}
          <div style={{
            width: IPAD_W, height: IPAD_H,
            background: "#fff",
            borderRadius: 22,
            overflow: "hidden",
            position: "relative",
            boxShadow: "inset 0 0 0 1px rgba(0,0,0,0.1)",
          }}>
            <div style={{ width: "100%", height: "100%", display: "flex", flexDirection: "column" }}>
              <ACInspectionApp />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
