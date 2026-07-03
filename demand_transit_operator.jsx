import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";

/* ============================================================
   オンデマンド交通 予約受付オペレータ・コンソール
   ------------------------------------------------------------
   エンジン仕様（提供Excelに準拠）
   - 所要時間マトリクス TT(x,y)：原典の値を保持。空欄ペアのみ
     既知区間の最短経路（経由）で補完。
   - DRTn: 直行乗車時間 / STn: 寄り道許容 = 三段式（短距離一律→線形→長距離一律）でDRTnから算出
   - MRTn = DRTn + STn（最大乗車許容時間）
   - TW: 発時刻の約束幅（約束発〜約束発+TW の間に乗車）
   - Dwell: 乗降滞留時間
   - 約束発 IPT / 約束着 IDT = IPT + Dwell + MRT は確定後不変
   - 「乗車のまま待たない」：同乗者がいる間の待機は span（車両
     拘束時間）のコスト増として挿入選択で抑制。約束時刻前の早着
     待機は物理的に不可避のため許容し、同乗者への害は約束着
     （IDT）判定で制約する
   - 挿入法：既存ルートの全位置に O/D を挿入試行し、全確定予約の
     約束を守れる実行可能解のうち追加所要が最小のものを採用
   ============================================================ */

// ---- 所要時間マトリクス（秒）BS001〜BS030 ----
// 運行エリア：10km四方（マップ横900px＝10km、1px≒11.1m）。
// 平均速度25km/h（市街地・信号停車込み）で、座標ベースの道路ネットワーク
// （近傍4接続・道路係数1.10〜1.35）上の全ペア最短時間として生成。
// 対称・三角不等式を全数検証済み（矛盾なし）。6秒単位。
// 所要時間レンジ：2.3〜28.8分（中央値13.0分）。実測値が得られた区間から
// 順次差し替え可能（差し替え後は verifyNetworkTT で三角不等式を再検証する）。
let STOPS = Array.from({length:30},(_,i)=>`BS${String(i+1).padStart(3,"0")}`);

// 停留所座標（運行盤上の論理配置。900×640のビューボックス上のpx座標）
let POS=[[255,150],[810,255],[395,90],[665,355],[140,245],[800,140],[650,450],[480,580],[95,400],[400,470],[560,68],[228,308],[76,154],[568,327],[694,58],[128,103],[354,331],[557,237],[347,160],[266,523],[190,420],[837,375],[495,397],[236,71],[747,212],[498,186],[517,504],[681,266],[822,485],[266,375]];

/* ---- 座標ベースの道路ネットワーク生成 ----
   停留所のpx座標を実距離（km）に換算し、各点を近傍K点へ接続した無向グラフを作る。
   区間ごとに道路係数（直線距離に対する迂回率）をランダムに割り当て、
   区間時間＝距離÷速度×3600×係数 を6秒単位に丸めてから全対最短経路（Floyd-Warshall）を計算。
   丸めた区間重みの和で最短路を求めるため、三角不等式は構成上自動的に満たされる。 */
function buildSyntheticNetwork(pos, opts){
  const {mapWidthKm=10, speedKmh=25, kNeighbors=4, coefMin=1.10, coefMax=1.35, seed}=opts||{};
  const N=pos.length;
  const scale=mapWidthKm/900;               // px→km（マップ横幅900px＝mapWidthKm）
  const km=pos.map(([x,y])=>[x*scale,y*scale]);
  let rngState=seed!=null?seed:Math.floor(Math.random()*1e9);
  const rand=()=>{ if(seed==null) return Math.random();
    rngState=(rngState*1103515245+12345)&0x7fffffff; return rngState/0x7fffffff; };
  const eset=new Set(), edges=[];
  for(let i=0;i<N;i++){
    const order=[...Array(N).keys()].filter(j=>j!==i)
      .sort((a,b)=>Math.hypot(km[i][0]-km[a][0],km[i][1]-km[a][1])-Math.hypot(km[i][0]-km[b][0],km[i][1]-km[b][1]))
      .slice(0,kNeighbors);
    for(const j of order){
      const key=i<j?`${i}-${j}`:`${j}-${i}`;
      if(!eset.has(key)){eset.add(key); edges.push([i<j?i:j, i<j?j:i]);}
    }
  }
  const big=1e9;
  const T=Array.from({length:N},()=>Array(N).fill(big));
  for(let i=0;i<N;i++)T[i][i]=0;
  for(const [a,b] of edges){
    const distKm=Math.hypot(km[a][0]-km[b][0],km[a][1]-km[b][1]);
    const coef=coefMin+rand()*(coefMax-coefMin);
    const sec=Math.round(distKm/speedKmh*3600*coef/6)*6;   // 6秒単位に丸め
    T[a][b]=sec; T[b][a]=sec;
  }
  for(let k=0;k<N;k++)for(let i=0;i<N;i++){ if(T[i][k]>=big)continue;
    for(let j=0;j<N;j++){ const v=T[i][k]+T[k][j]; if(v<T[i][j])T[i][j]=v; } }
  return {T,edges};
}
/* ネットワークの整合性を検証（対称性・三角不等式・到達可能性を全数チェック）。
   CSV等でTTを差し替えた後の再検証にも使える。違反があれば内容を返す（開発確認用）。 */
function verifyNetworkTT(T){
  const N=T.length;
  let asym=0,unreachable=0,viol=0;
  for(let i=0;i<N;i++)for(let j=0;j<N;j++){
    if(T[i][j]!==T[j][i])asym++;
    if(T[i][j]>=1e9)unreachable++;
  }
  for(let x=0;x<N;x++)for(let y=0;y<N;y++)for(let z=0;z<N;z++){
    if(T[x][y] > T[x][z]+T[z][y]+1e-6)viol++;
  }
  return {ok:asym===0&&unreachable===0&&viol===0, asym, unreachable, viol, checked:N*N*N};
}
/* ---- コリドー軸の構築（需要カバーと路線コストの両立） ----
   目的：定時定路線（またはセミデマンドの背骨）として引く価値のある1本の軸を、
   需要（人数重み）・路線延長（実TT）・利用者の乗車時間の3点から決める。
   手順：
   1) 端点選定＝接触人数上位の停留所群から「需要×スパン（実TT）」最大の2点。
   2) 逐次挿入＝「新たにカバーされる人数 ÷ 挿入による延長増（分）」が最大の
      停留所を安価挿入位置へ追加。ただし挿入後のカバー済みODの平均乗車時間
      伸び率（路線上所要÷直行TT）が maxStretch を超える挿入は棄却（利便性制約）。
   3) 停止則＝限界効率が現在の平均効率（人/分）の stopRatio 倍を下回ったら打切り
      （需要規模に依存しない相対基準）。停留所数上限は全体の6割。
   4) 2-opt＝経路順序の入替えで総延長を短縮（ジグザグ・後戻りの除去）。
   カバー＝ODの両端が路線上に載ること（路線は両方向運行の前提で無方向に集計）。 */
function buildCorridor(records, opts){
  const {maxStretch=1.6, stopRatio=0.25}=opts||{};
  const N=STOPS.length;
  const flow=new Map();                       // "a-b"(a<b) → 人数
  const touch=Array(N).fill(0);
  let totalPax=0;
  for(const r of records||[]){
    if(!r||r.err||r.o==null||r.d==null||r.o===r.d)continue;
    if(r.o>=N||r.d>=N)continue;
    const w=r.pax||1;
    const key=Math.min(r.o,r.d)+"-"+Math.max(r.o,r.d);
    flow.set(key,(flow.get(key)||0)+w);
    touch[r.o]+=w; touch[r.d]+=w; totalPax+=w;
  }
  if(totalPax===0)return null;
  const pathLen=p=>{let s=0;for(let i=0;i+1<p.length;i++)s+=TT[p[i]][p[i+1]];return s;};
  const evalPath=p=>{                          // カバー人数・平均伸び率（接頭和でO(経路長+流動数)）
    const pos=new Map(); p.forEach((s,i)=>pos.set(s,i));
    const pre=[0]; for(let i=0;i+1<p.length;i++)pre.push(pre[i]+TT[p[i]][p[i+1]]);
    let cov=0,stSum=0;
    for(const [key,w] of flow){
      const sp=key.split("-"),a=+sp[0],b=+sp[1];
      let ia=pos.get(a),ib=pos.get(b);
      if(ia==null||ib==null)continue;
      if(ia>ib){const t=ia;ia=ib;ib=t;}
      cov+=w;
      const direct=Math.max(Math.min(TT[a][b],TT[b][a]),1);
      stSum+=w*((pre[ib]-pre[ia])/direct);     // 路線上の乗車時間（順方向の区間和）
    }
    return {cov, avgStretch:cov>0?stSum/cov:1};
  };
  // 1) 端点：接触上位から需要×スパン最大の2点
  const ranked=touch.map((c,i)=>[i,c]).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]);
  if(ranked.length<2)return null;
  const cand=ranked.slice(0,Math.min(ranked.length,12)).map(e=>e[0]);
  let A=cand[0],B=cand[1],best=-1;
  for(let i=0;i<cand.length;i++)for(let j=i+1;j<cand.length;j++){
    const s=Math.sqrt(touch[cand[i]]*touch[cand[j]])*(TT[cand[i]][cand[j]]+TT[cand[j]][cand[i]]);
    if(s>best){best=s;A=cand[i];B=cand[j];}
  }
  let path=[A,B];
  const inPath=new Set(path);
  const maxStops=Math.min(N,Math.max(4,Math.ceil(N*0.6)));
  let curEval=evalPath(path), curLen=pathLen(path);
  // 2)〜3) 逐次挿入＋相対停止則
  while(path.length<maxStops){
    let bestS=-1,bestRatio=0,bestPos=-1;
    for(let s=0;s<N;s++){
      if(inPath.has(s)||touch[s]===0)continue;
      let dMin=Infinity,pos=-1;
      for(let i=0;i+1<path.length;i++){        // 中間への安価挿入
        const d=TT[path[i]][s]+TT[s][path[i+1]]-TT[path[i]][path[i+1]];
        if(d<dMin){dMin=d;pos=i+1;}
      }
      const dHead=TT[s][path[0]];              // 端への延長も許す
      const dTail=TT[path[path.length-1]][s];
      if(dHead<dMin){dMin=dHead;pos=0;}
      if(dTail<dMin){dMin=dTail;pos=path.length;}
      if(!(dMin<Infinity))continue;
      const p2=path.slice(); p2.splice(pos,0,s);
      const ev=evalPath(p2);
      if(ev.avgStretch>maxStretch)continue;    // 利便性制約：伸び率上限
      const gain=ev.cov-curEval.cov;
      if(gain<=0)continue;
      const ratio=gain/Math.max(dMin/60,0.5);  // 人/分（0.5分でクリップ）
      if(ratio>bestRatio){bestRatio=ratio;bestS=s;bestPos=pos;}
    }
    if(bestS<0)break;
    const avgEff=curLen>0?curEval.cov/(curLen/60):Infinity;
    if(curLen>0 && bestRatio<stopRatio*avgEff)break;   // 限界効率の相対停止則
    path.splice(bestPos,0,bestS); inPath.add(bestS);
    curLen=pathLen(path); curEval=evalPath(path);
  }
  // 4) 2-opt：区間反転で総延長を短縮（非対称TTでも全長再計算で正しく比較）。
  //    延長が縮んでも、カバー済みODの平均伸び率が上限を超える入替えは棄却する
  //    （距離の最短化が利用者の乗車時間を犠牲にしないための制約付き2-opt）。
  let improved=true,guard=0;
  while(improved&&guard++<60){
    improved=false;
    for(let i=0;i<path.length-1;i++)for(let k=i+1;k<path.length;k++){
      const p2=path.slice(0,i).concat(path.slice(i,k+1).reverse(),path.slice(k+1));
      if(pathLen(p2)<pathLen(path)-1e-9 && evalPath(p2).avgStretch<=maxStretch){
        path=p2;improved=true;
      }
    }
  }
  curLen=pathLen(path); curEval=evalPath(path);
  return {stops:path, lengthSec:curLen, coverage:curEval.cov/totalPax,
    avgStretch:curEval.avgStretch, covPax:curEval.cov, totalPax};
}
/* コリドー停留所列からアンカーにする停留所を間引く。
   目的：全停留所を時刻の杭にすると各区間が短く（1〜4分）、寄り道の原資が
   区間ごとに細切れになって実質使えない。杭を主要停留所（約minGap間隔）に
   限定すれば、区間が長くなり余裕がプールされ、寄り道・待ちの許容が増える。
   杭でない停留所も経路上にあるため乗降自体は可能（予約で拾う）。
   始点・終点は常に保持。 */
function thinCorridorStops(pathStops, minGapSec){
  if(!minGapSec||pathStops.length<=2)return pathStops.slice();
  const out=[pathStops[0]];
  let acc=0;
  for(let i=1;i<pathStops.length-1;i++){
    acc+=TT[pathStops[i-1]][pathStops[i]];
    if(acc>=minGapSec){out.push(pathStops[i]);acc=0;}
  }
  out.push(pathStops[pathStops.length-1]);
  return out;
}
/* コリドー軸 → セミデマンド背骨（アンカー列）への変換。
   アンカー間隔＝直行TT×余裕率（分単位切上げ・最低1分）。余裕率−1が寄り道の原資。
   opts.mode："oneway"（片道1本）／"roundtrip"（往復＝便ごとに向きを反転）／
              "loop"（循環＝毎便同方向、便間に終点→始点の回送を挟む）
   opts.trips：便数（片道1本＝1便と数える。往復1周＝2便）
   opts.turnaround：便間の折返し・待機秒（既定180秒）
   opts.endSec：背骨の終了時刻。指定時はtripsを無視し、終了時刻に収まる最大便数を
                自動算出する（時間帯限定セミ運行用）。1便も入らなければ1便のみ生成。 */
function corridorToBackbone(pathStops, startSec, slackFactor, opts){
  const {mode="oneway", trips=1, turnaround=180, endSec=null}=opts||{};
  const gen=(nTrips)=>{
    const fwd=pathStops.slice(), rev=pathStops.slice().reverse();
    const anchors=[];
    let t=Math.round(startSec/60)*60;
    for(let k=0;k<nTrips;k++){
      const seq=(mode==="roundtrip"&&k%2===1)?rev:fwd;
      for(let i=0;i<seq.length;i++){
        if(anchors.length===0){ anchors.push({stop:seq[0],time:t}); continue; }
        const prevStop=anchors[anchors.length-1].stop;
        let iv;
        if(i===0){
          const move=prevStop===seq[0]?0:TT[prevStop][seq[0]]*slackFactor;
          iv=Math.max(60, Math.ceil((move+turnaround)/60)*60);
        }else{
          iv=Math.max(60, Math.ceil(TT[prevStop][seq[i]]*slackFactor/60)*60);
        }
        t=anchors[anchors.length-1].time+iv;
        anchors.push({stop:seq[i],time:t});
      }
    }
    return anchors;
  };
  if(mode==="oneway")return gen(1);
  if(endSec==null)return gen(Math.max(1,trips));
  // 終了時刻指定：収まる最大便数を探索（1便ずつ増やし、終着が終了時刻を超えたら1つ戻す）
  let n=1, bb=gen(1);
  while(true){
    const next=gen(n+1);
    if(next[next.length-1].time>endSec)break;
    n++; bb=next;
    if(n>=48)break;   // 安全弁
  }
  return bb;
}
/* セミデマンド有効性の比較試算（アプリの状態は変更しない）。
   同一の需要行を「全車フル」（構成A）と「指定1台セミ＋残りフル」（構成B）の
   2構成に、それぞれ空の状態から先着順で流し込み、結果指標を並べて返す。
   指標：成立数・成立率・平均ズレ（|約束発−希望発|）・拘束時間計（全車spanの和）・
   平均乗車時間・セミ車の担当件数（構成Bのみ意味を持つ）。 */
function compareSemiEffect(rows, baseFleet, semiVehId, backbone, P){
  const runCfg=(fleet, initRoutes)=>{
    const semiIds=new Set(fleet.filter(v=>v.mode==="semi").map(v=>v.id));
    const routes={}; for(const v of fleet)routes[v.id]=(initRoutes&&initRoutes[v.id])?[...initRoutes[v.id]]:[];
    const rm={}; let okc=0, devSum=0, semiCnt=0, paxOk=0, idx=0;
    for(const q of rows){
      const drt=TT[q.o][q.d], mrt=mrtFromDRT(drt,P);
      const dpt=q.mode==="dep"?q.t:q.t-(mrt+P.dwell);
      const r={id:"X"+(idx++)+"_"+Math.random().toString(36).slice(2,6),num:0,
        o:q.o,d:q.d,dpt,drt,mrt,sa:q.sa,pax:q.pax,ipt:null,idt:null,vehicle:null};
      const cs=searchInsertions(routes,rm,r,P,fleet);
      if(cs.length){
        const c=cs[0];
        r.ipt=c.apt;r.idt=c.apt+P.dwell+mrt;r.vehicle=c.vehicle;
        routes[c.vehicle]=c.route;rm[r.id]=r;okc++;
        devSum+=Math.abs(c.apt-dpt); paxOk+=(q.pax||1);
        if(semiIds.has(c.vehicle))semiCnt++;
      }
    }
    let spanSum=0, rideSum=0, rideN=0;
    for(const v of fleet){
      const sim=simulate(routes[v.id],rm,P,v);
      if(!sim.ok)continue;
      if(sim.events.length)spanSum+=sim.span;
      const om={},dm={};
      for(const e of sim.events){if(e.type==="O")om[e.resId]=e;else if(e.type==="D")dm[e.resId]=e;}
      for(const id in om){const o=om[id],d=dm[id];if(!d)continue;rideSum+=d.adt-o.etd;rideN++;}
    }
    return {ok:okc,total:rows.length,rate:rows.length?okc/rows.length:0,
      devAvg:okc?devSum/okc:0,spanSum,rideAvg:rideN?rideSum/rideN:0,paxOk,semiCnt};
  };
  const fullFleet=baseFleet.map(v=>({...v,mode:"full",backbone:[]}));
  const A=runCfg(fullFleet,null);
  const semiFleet=baseFleet.map(v=>v.id===semiVehId
    ?{...v,mode:"semi",backbone}
    :{...v,mode:"full",backbone:[]});
  const initRoutes={}; initRoutes[semiVehId]=anchorEvents(backbone);
  const B=runCfg(semiFleet,initRoutes);
  return {A,B};
}
const _netInit=buildSyntheticNetwork(POS,{mapWidthKm:10,speedKmh:25,kNeighbors:4,coefMin:1.10,coefMax:1.35,seed:42});
let TT=_netInit.T;
// 道路ネットワークの隣接区間（マップの薄い参照線＝実際に走行しうる道路区間）
let EDGES=_netInit.edges;

// 車両は最大5台。台数（稼働ON/OFF）・定員・運行時間はヘッダの「車両設定」で変更可能。
// 中心拠点（病院・駅・商業施設を想定）：発着が集中する停留所。indexは0始まり。
let HUBS=[0,17,22,12,16];   // BS001, BS018, BS023, BS013, BS017

/* ---------------- ネットワーク（停留所・OD直行時間）の読込 ----------------
   stops.csv  : stop_id,name,lat,lon,is_hub
   od_time.csv: from_id,to_id,seconds
   読み込むと STOPS / POS / HUBS / TT / EDGES を置き換える（letで再代入）。 */
function splitCSV(text){
  return text.replace(/\r/g,"").split("\n").map(l=>l.trim()).filter(l=>l.length>0).map(l=>l.split(","));
}
function parseStopsCSV(text){
  const rows=splitCSV(text);
  if(rows.length<2) throw new Error("stops.csvに行がない");
  const head=rows[0].map(h=>h.trim().toLowerCase());
  const ci=name=>head.findIndex(h=>h===name);
  const iId=ci("stop_id")>=0?ci("stop_id"):0, iName=ci("name"), iLat=ci("lat"), iLon=ci("lon"), iHub=ci("is_hub");
  if(iLat<0||iLon<0) throw new Error("stops.csvに lat/lon 列がない");
  return rows.slice(1).map(c=>({
    id:(c[iId]||"").trim(),
    name:(iName>=0?c[iName]:c[iId]||"").trim(),
    lat:Number(c[iLat]), lon:Number(c[iLon]),
    isHub:iHub>=0 && String(c[iHub]).trim()==="1",
  })).filter(s=>s.id && Number.isFinite(s.lat) && Number.isFinite(s.lon));
}
function parseODCSV(text){
  const rows=splitCSV(text);
  if(rows.length<2) throw new Error("od_time.csvに行がない");
  const head=rows[0].map(h=>h.trim().toLowerCase());
  const ci=name=>head.findIndex(h=>h===name);
  const iF=ci("from_id")>=0?ci("from_id"):0, iT=ci("to_id")>=0?ci("to_id"):1, iS=ci("seconds")>=0?ci("seconds"):2;
  return rows.slice(1).map(c=>({from:(c[iF]||"").trim(),to:(c[iT]||"").trim(),sec:Number(c[iS])}))
    .filter(r=>r.from && r.to && Number.isFinite(r.sec));
}
// lat/lon を 900×640 のビューボックスへ投影（北を上に）。退化時は中央寄せ。
function projectLatLon(stops){
  const W=900,H=640,pad=70;
  const lats=stops.map(s=>s.lat), lons=stops.map(s=>s.lon);
  const laMin=Math.min(...lats),laMax=Math.max(...lats),loMin=Math.min(...lons),loMax=Math.max(...lons);
  const dLa=laMax-laMin||1, dLo=loMax-loMin||1;
  return stops.map(s=>[
    Math.round(pad+(s.lon-loMin)/dLo*(W-2*pad)),
    Math.round(pad+(laMax-s.lat)/dLa*(H-2*pad)),
  ]);
}
function applyNetwork(stops, odRows, odFactor){
  if(!stops||stops.length<2) throw new Error("停留所が2点以上必要");
  // odFactor: CSVの所要時間に掛ける倍率（例1.5＝信号・混雑等の実勢補正）。
  // 実測値にのみ適用し、欠損セルのプレースホルダー（3時間）には掛けない。
  const f=(typeof odFactor==="number"&&isFinite(odFactor)&&odFactor>0)?odFactor:1;
  const N=stops.length;
  const idx={}; stops.forEach((s,i)=>{idx[s.id]=i;});
  const newSTOPS=stops.map(s=>s.name||s.id);
  const newPOS=projectLatLon(stops);
  const newHUBS=stops.map((s,i)=>s.isHub?i:-1).filter(i=>i>=0);
  // TT初期化：対角0、未定義は大きな値（後でCSVで上書き）
  const big=3*3600;
  const newTT=Array.from({length:N},(_,i)=>Array.from({length:N},(_,j)=>i===j?0:big));
  let filled=0;
  for(const r of odRows){
    const a=idx[r.from], b=idx[r.to];
    if(a==null||b==null) continue;
    newTT[a][b]=Math.max(0,Math.round(r.sec*f)); filled++;
  }
  // 網羅率：N×(N-1)通りの有向ペアのうち、実測値（big未満）で埋まっている割合。
  // 欠損セルはbig（3時間）のプレースホルダーのままで、実距離ではない点に注意。
  const totalPairs=N*(N-1);
  let known=0, asym=0;
  const missing=[];
  for(let i=0;i<N;i++)for(let j=0;j<N;j++){ if(i===j)continue;
    if(newTT[i][j]<big) known++; else if(missing.length<8) missing.push([i,j]);
    if(newTT[i][j]<big && newTT[j][i]<big && newTT[i][j]!==newTT[j][i]) asym++;
  }
  const coverage=totalPairs>0?known/totalPairs:0;
  // EDGES（視覚用の参照線）：各停留所を時間の近い上位3点へ接続（無向・重複排除）
  const eset=new Set(), newEDGES=[];
  for(let i=0;i<N;i++){
    const order=[...Array(N).keys()].filter(j=>j!==i).sort((a,b)=>newTT[i][a]-newTT[i][b]).slice(0,3);
    for(const j of order){
      const key=i<j?`${i}-${j}`:`${j}-${i}`;
      if(!eset.has(key)){eset.add(key); newEDGES.push(i<j?[i,j]:[j,i]);}
    }
  }
  STOPS=newSTOPS; POS=newPOS; HUBS=newHUBS.length?newHUBS:[0]; TT=newTT; EDGES=newEDGES;
  return {n:N, odFilled:filled, coverage, known, totalPairs, asym, factor:f,
    missing:missing.map(([i,j])=>`${newSTOPS[i]}→${newSTOPS[j]}`)};
}
const HUB_NAMES={0:"駅前",17:"総合病院",22:"中心商業",12:"市役所",16:"モール"};

// 需要の時間帯ピーク：中心時刻(h)・分散(sigma)・相対強度(weight)。
// 生成（山の形）と需要プレビューの色分け（ピーク帯判定）が共にこれを参照する。
const DEMAND_PEAKS=[
  {label:"朝",   center:9.0,  sigma:0.55, weight:1.0},
  {label:"昼通院",center:12.0, sigma:0.6,  weight:0.7},
  {label:"夕",   center:16.0, sigma:0.6,  weight:0.95},
];
// 時刻hがいずれかの山の中心±sigma以内ならピーク帯とみなす
const isPeakHour=h=>DEMAND_PEAKS.some(p=>Math.abs(h-p.center)<=p.sigma);

const VEHICLE_COLORS=["#FF9E4D","#4DC6CF","#C792EA","#7BC96F","#E8718D"];
const DEFAULT_VEHICLES=[
  {id:1,name:"1号車",color:VEHICLE_COLORS[0],active:true, cap:8, start:9*3600, end:17*3600},
  {id:2,name:"2号車",color:VEHICLE_COLORS[1],active:true, cap:8, start:9*3600, end:17*3600},
  {id:3,name:"3号車",color:VEHICLE_COLORS[2],active:true, cap:8, start:9*3600, end:17*3600},
  {id:4,name:"4号車",color:VEHICLE_COLORS[3],active:false,cap:4, start:10*3600,end:16*3600},
  {id:5,name:"5号車",color:VEHICLE_COLORS[4],active:false,cap:4, start:10*3600,end:16*3600},
];

const fmt = s => {
  s = Math.round(s);
  const h = Math.floor(s/3600), m = Math.floor((s%3600)/60);
  return `${h}:${String(m).padStart(2,"0")}`;
};
const fmtMin = s => `${Math.round(s/60)}分`;
const fmtDev = s => s===0?"±0分":`${s>0?"+":"−"}${Math.round(Math.abs(s)/60)}分`;
const fmtHMS = s => {
  s = Math.round(s);
  const h=Math.floor(s/3600), m=Math.floor((s%3600)/60), c=s%60;
  return `${h}:${String(m).padStart(2,"0")}:${String(c).padStart(2,"0")}`;
};
const fmtS = s => fmt(s); // 表中の時刻（h:mm）
// <input type="time">境界用の相互変換。内部の時刻は常に秒で保持する。
const secToHM = s => {
  if(typeof s!=="number"||!isFinite(s)) return "";
  const h=Math.floor(s/3600), m=Math.floor((s%3600)/60);
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
};
const hmToSec = (t, fallback) => {
  const m=String(t||"").match(/^(\d{1,2}):(\d{2})$/);
  return m ? (+m[1])*3600+(+m[2])*60 : fallback;   // 不正入力は元の値を維持
};

/* ---------------- サービスレベル（寄り道時間ST／最大乗車MRT） ----------------
   STは三段式：DRT≤d1で一律s1（短距離）→ d1〜d2で線形 → DRT≥d2で一律s2（長距離）。
   MRTはMRT=DRT+STで派生（可変にするのはSTだけ。MRTを独立に固定すると
   長距離でMRT<DRTという不可能値が出るため）。
   単調増加（逆転なし）の十分条件は s2>=s1。線形部の傾きは1+(s2-s1)/(d2-d1)。
   値はすべて秒。P.sl={d1,d2,s1,s2}。 */
function stFromDRT(drt,P){
  const {d1,d2,s1,s2}=P.sl;
  if(drt<=d1) return s1;
  if(drt>=d2) return s2;
  return s1+(s2-s1)*(drt-d1)/(d2-d1);
}
function mrtFromDRT(drt,P){ return Math.round(drt+stFromDRT(drt,P)); }

/* ---------------- エンジン ---------------- */
// route: [{stop, type:'O'|'D', resId}] を時刻シミュレート
// 返値: {ok, events:[{...,eta,apt/adt,etd}], span} / ok=false で理由
function simulate(route, resMap, P, veh){
  if(route.length===0) return {ok:true, events:[], span:0, maxLoad:0};
  const events=[]; const onboard=new Set();
  let t=null, prevStop=null, load=0, maxLoad=0;
  for(let k=0;k<route.length;k++){
    const ev=route[k];
    // アンカー（時刻の杭）：乗降なしの経由点。早着は杭時刻まで待機（区間独立の再同期点）、
    // 遅延許容を超えたら不成立。これで各区間の寄り道スラックが自動的に効く。
    if(ev.type==="ANCHOR"){
      const eta = k===0 ? veh.start : t + TT[prevStop][ev.stop];
      const tol = P.anchorTol ?? 300;
      if(ev.atime!=null && eta > ev.atime + tol) return {ok:false, why:"アンカー時刻に間に合わない"};
      const dep = ev.atime!=null ? Math.max(eta, ev.atime) : eta;
      if(dep > veh.end) return {ok:false, why:"運行終了超過"};
      events.push({...ev, eta, apt:dep, etd:dep, atime:ev.atime, load});
      t=dep; prevStop=ev.stop;
      continue;
    }
    const r=resMap[ev.resId];
    let eta;
    if(prevStop===null){
      if(ev.type!=="O") return {ok:false, why:"先頭が降車"};
      eta = Math.max(r.ipt ?? r.dpt, veh.start); // 始発は約束発(未確定なら発希望)と運行開始の遅い方
    }else{
      eta = t + TT[prevStop][ev.stop];
    }
    if(ev.type==="O"){
      const target = r.ipt ?? r.dpt;               // 約束発（未確定は発希望）
      // 乗車許容窓：確定済=IPT〜IPT+TW（後ろのみ）／未確定=DPT±San（前後）
      const lo = r.ipt!=null ? r.ipt : r.dpt - (r.sa ?? P.tw);
      const hi = r.ipt!=null ? r.ipt + P.tw : r.dpt + (r.sa ?? P.tw);
      let apt;
      if(eta < lo){
        // 早着：利用者は約束時刻（lo）まで来ないため待機は不可避。
        // 同乗者がいる場合の待機の害は、各同乗者の約束着（IDT）判定で制約される。
        apt = onboard.size===0 ? target : lo;       // 空車なら希望時刻まで、同乗ありなら最小限の待機
      }else if(eta <= hi){
        apt = (onboard.size===0 && eta < target) ? target : eta; // 空車で早着なら希望まで待つ／乗客ありなら窓内で早乗せ
      }else{
        return {ok:false, why:r.ipt!=null?"約束発+TW超過":"発希望+San超過"};
      }
      apt = Math.max(apt, veh.start);             // 運行開始前には乗せない
      if(apt > hi) return {ok:false, why:"運行時間と希望が不整合"};
      // 定員判定：乗車人数の合計が車両定員を超えないこと
      load += (r.pax ?? 1);
      if(load > veh.cap) return {ok:false, why:`定員超過（${load}名＞定員${veh.cap}名）`};
      maxLoad = Math.max(maxLoad, load);
      const etd = apt + P.dwell;
      onboard.add(ev.resId);
      events.push({...ev, eta, apt, etd, load});
      t=etd; prevStop=ev.stop;
    }else{
      if(!onboard.has(ev.resId)) return {ok:false, why:"乗車前の降車"};
      const adt = eta;                              // 着＝到着時刻
      const pick = events.find(e=>e.resId===ev.resId && e.type==="O");
      const limit = r.idt ?? (pick.apt + P.dwell + r.mrt); // 約束着（未確定はMRT基準）
      if(adt > limit + 1e-6) return {ok:false, why:"約束着超過"};
      const etd = adt + P.dwell;
      if(etd > veh.end) return {ok:false, why:"運行終了超過"};
      load -= (r.pax ?? 1);
      onboard.delete(ev.resId);
      events.push({...ev, eta, adt, etd, load});
      t=etd; prevStop=ev.stop;
    }
  }
  if(onboard.size>0) return {ok:false, why:"未降車あり"};
  return {ok:true, events, span: events[events.length-1].etd - events[0].eta, maxLoad};
}

// 新規予約 r を全車両・全挿入位置で試行し、実行可能な候補を返す
function searchInsertions(routes, resMap, newRes, P, vehicles){
  const cands=[];
  const tmpMap={...resMap, [newRes.id]:newRes};
  for(const v of vehicles.filter(v=>v.active)){
    const base = routes[v.id] || [];
    const baseSim = simulate(base, resMap, P, v);
    const baseSpan = baseSim.ok ? baseSim.span : 0;
    let best=null;
    for(let i=0;i<=base.length;i++){
      for(let j=i;j<=base.length;j++){
        const r2=[...base];
        r2.splice(i,0,{stop:newRes.o,type:"O",resId:newRes.id});
        r2.splice(j+1,0,{stop:newRes.d,type:"D",resId:newRes.id});
        const sim=simulate(r2,tmpMap,P,v);
        if(!sim.ok) continue;
        // セミ車：各アンカー区間内で後戻りする挿入は却下（往復・循環の復路は別区間なので通る）
        if(v.mode==="semi" && v.backbone && v.backbone.length && !bbSegOrdered(r2)) continue;
        const me=sim.events.filter(e=>e.resId===newRes.id);
        const apt=me[0].apt, adt=me[1].adt;
        const added=sim.span-baseSpan;
        const dev=apt-newRes.dpt;                 // 符号付き（負＝希望より早い）
        const cost=added*1.0+Math.abs(dev)*0.2;
        if(!best||cost<best.cost) best={vehicle:v.id,route:r2,sim,apt,adt,added,dev,cost,maxLoad:sim.maxLoad,cap:v.cap};
      }
    }
    if(best) cands.push(best);
  }
  cands.sort((a,b)=>a.cost-b.cost);
  return cands;
}

/* ===== 全体最適化シミュレーション（割付シャッフル）用ヘルパ =====
   現行の先着順固定（確定後は動かさない）に対し、新規予約が1件入るたびに
   確定済みの割付を作り直し、総走行（span）最小の実行可能解を採る。
   約束（IPT〜IPT+TW窓・約束着IDT）と定員・運行時間はsimulateが厳守するため、
   再最適化しても破れない。変わるのは「どの便が運ぶか」と「窓内の実乗車時刻」。 */
function shuffleArr(a){for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}return a;}
function vehMapOf(routes){const m={};for(const vid in routes)for(const ev of routes[vid])if(ev.resId)m[ev.resId]=Number(vid);return m;}
function totalSpanOf(routes,resMap,P,fleet){let s=0;for(const v of fleet){const sim=simulate(routes[v.id]||[],resMap,P,v);if(sim.ok)s+=sim.span;}return s;}

// 先着順固定（比較用ベースライン）：空から上から順に挿入、確定後は動かさない
function greedyPass(rows,P,fleet){
  const routes={};for(const v of fleet)routes[v.id]=[];
  const rm={};const out=[];let okc=0;
  for(const q of rows){
    const drt=TT[q.o][q.d],mrt=mrtFromDRT(drt,P);
    const dpt=q.mode==="dep"?q.t:q.t-(mrt+P.dwell);
    const id="G"+okc+"_"+Math.random().toString(36).slice(2,6);
    const r={id,o:q.o,d:q.d,dpt,drt,mrt,sa:q.sa,pax:q.pax,ipt:null,idt:null,vehicle:null};
    const cs=searchInsertions(routes,rm,r,P,fleet);
    if(cs.length){const c=cs[0];r.ipt=c.apt;r.idt=c.apt+P.dwell+mrt;r.vehicle=c.vehicle;
      routes[c.vehicle]=c.route;rm[id]=r;okc++;
      out.push({...q,ok:true,num:okc,vehicle:c.vehicle,apt:c.apt,idt:r.idt,dev:c.apt-dpt,reassigned:0,rescued:false});}
    else out.push({...q,ok:false,reason:"挿入位置なし"});
  }
  return {rows:out,ok:okc,total:rows.length,span:totalSpanOf(routes,rm,P,fleet),routes,resMap:rm};
}

// 先着順では入らなかった新規nrを、確定済みの割付シャッフルで「救済」する。
// 既存全員＋新規が全員入る並べ替えのみ採用（確定済みは約束を保ったまま別便へ移ってよい）。
// 救済解が複数あれば総走行（span）最小を採る。救済不能ならnull。
function rescueByReshuffle(routes,resMap,confirmedOrder,nr,P,fleet,tries){
  const conf=confirmedOrder.map(id=>resMap[id]);
  const byIpt=conf.slice().sort((a,b)=>a.ipt-b.ipt);          // 約束発の早い順
  const byTrip=conf.slice().sort((a,b)=>b.drt-a.drt);         // 長距離（制約が厳しい）順
  const seeds=[[nr,...byIpt],[...byIpt,nr],[...byTrip,nr]];    // 決め打ちの並べ替え3種
  let best=null;
  const N=Math.max(tries,seeds.length);
  for(let a=0;a<N;a++){
    const order = a<seeds.length ? seeds[a] : shuffleArr([...conf,nr]);
    const rt={};for(const v of fleet)rt[v.id]=[];
    const rm={};let ok=true,newApt=null;
    for(const r of order){
      const cs=searchInsertions(rt,rm,r,P,fleet);
      if(!cs.length){ok=false;break;}                 // 1人でも入らない並べ替えは棄却
      const c=cs[0];rt[c.vehicle]=c.route;
      if(r.id===nr.id){newApt=c.apt;rm[r.id]={...r,ipt:c.apt,idt:c.apt+P.dwell+r.mrt,vehicle:c.vehicle};}
      else rm[r.id]={...r,vehicle:c.vehicle};          // 確定済みはipt/idt不変・便だけ更新
    }
    if(!ok)continue;
    const sp=totalSpanOf(rt,rm,P,fleet);
    if(!best||sp<best.span)best={routes:rt,resMap:rm,vehOf:vehMapOf(rt),newApt,span:sp};
  }
  return best;
}

// 到着順シミュレーション本体。greedyとの比較も返す。
// 方針：先着順で入る予約はそのまま確定し動かさない（むやみに動かすと将来の成立を潰すため）。
// 先着順で入らない予約だけ、割付シャッフルで救済を試みる。つまり最適化は成立を増やす方向にのみ働く。
function reoptimizeOnArrival(rows,P,fleet,tries){
  const greedy=greedyPass(rows,P,fleet);
  let routes={};for(const v of fleet)routes[v.id]=[];
  let resMap={};const confirmedOrder=[];const out=[];let reassignTotal=0,rescued=0;
  for(const q of rows){
    const drt=TT[q.o][q.d],mrt=mrtFromDRT(drt,P);
    const dpt=q.mode==="dep"?q.t:q.t-(mrt+P.dwell);
    const id="Z"+confirmedOrder.length+"_"+Math.random().toString(36).slice(2,6);
    const nr={id,o:q.o,d:q.d,mode:q.mode,dpt,drt,mrt,sa:q.sa,pax:q.pax,ipt:null,idt:null,vehicle:null};
    // (1) まず先着順（既存を動かさず新規だけ挿入）を試す
    const base=searchInsertions(routes,resMap,nr,P,fleet);
    if(base.length){
      const c=base[0];routes={...routes,[c.vehicle]:c.route};
      nr.ipt=c.apt;nr.idt=c.apt+P.dwell+mrt;nr.vehicle=c.vehicle;resMap={...resMap,[id]:nr};
      confirmedOrder.push(id);
      out.push({...q,ok:true,num:confirmedOrder.length,vehicle:nr.vehicle,
        apt:nr.ipt,idt:nr.idt,dev:nr.ipt-dpt,reassigned:0,rescued:false});
      continue;
    }
    // (2) 入らないので救済（既存の割付をシャッフル）
    const beforeVeh={};for(const rid of confirmedOrder)beforeVeh[rid]=resMap[rid].vehicle;
    const sol=rescueByReshuffle(routes,resMap,confirmedOrder,nr,P,fleet,tries);
    if(!sol){out.push({...q,ok:false,reason:"先着順では満車。割付シャッフルでも約束を守って入らず"});continue;}
    routes=sol.routes;resMap=sol.resMap;
    nr.ipt=sol.newApt;nr.idt=sol.newApt+P.dwell+mrt;nr.vehicle=sol.vehOf[id];
    confirmedOrder.push(id);rescued++;
    let reassigned=0;
    for(const rid of confirmedOrder){if(rid===id)continue;if(sol.vehOf[rid]!==beforeVeh[rid])reassigned++;}
    reassignTotal+=reassigned;
    out.push({...q,ok:true,num:confirmedOrder.length,vehicle:nr.vehicle,
      apt:nr.ipt,idt:nr.idt,dev:nr.ipt-dpt,reassigned,rescued:true});
  }
  const online={rows:out,ok:out.filter(r=>r.ok).length,total:rows.length,
    span:totalSpanOf(routes,resMap,P,fleet),reassign:reassignTotal,rescued};
  // 床：到着ごとの救済が全体で改善しなければ先着順を採る（全体最適化は先着順を下回らない）。
  // オンライン処理では、目先の救済が後続の成立を潰し全体で損になることがあるため。
  const better = online.ok>greedy.ok || (online.ok===greedy.ok && online.span<greedy.span);
  const reopt = better ? online
    : {rows:greedy.rows,ok:greedy.ok,total:rows.length,span:greedy.span,reassign:0,rescued:0,fellBack:true};
  // 確定反映用：採用した解のroutes/resMap（床作動時は先着順の解）
  const plan = better ? {routes,resMap} : {routes:greedy.routes,resMap:greedy.resMap};
  return {greedy,reopt,plan};
}

// 一括最適化（前日予約向け）。前日予約はすでに受付済み＝発着時刻を約束している。
// その約束（IPT〜IPT+TW窓・約束着IDT）を全て守ったまま、全車両の割付を組み直して
// 走行を締め、空いた余地に先着順では入らなかった予約（＝当日相当）を詰めて成立を増やす。
// 先着順の実ルートをベース解に持つため、約束は必ず満たされ成立は先着順を下回らない。
function offlineOptimize(rows,P,fleet,tries){
  const greedy=greedyPass(rows,P,fleet);
  const mk=b=>({id:"B"+b.idx,o:b.o,d:b.d,mode:b.mode,dpt:b.dpt,drt:b.drt,mrt:b.mrt,sa:b.sa,pax:b.pax,ipt:b.ipt,idt:b.idt,vehicle:null});
  // 約束済み（先着順で受付＝IPT/IDT固定）と、当日相当（先着順で入らなかった分）に分ける
  const promised=[],extra=[];
  greedy.rows.forEach((r,idx)=>{
    const q=rows[idx];const drt=TT[q.o][q.d],mrt=mrtFromDRT(drt,P);
    const dpt=q.mode==="dep"?q.t:q.t-(mrt+P.dwell);
    const b={idx,o:q.o,d:q.d,mode:q.mode,dpt,drt,mrt,sa:q.sa,pax:q.pax};
    if(r.ok){b.ipt=r.apt;b.idt=r.idt;promised.push(b);}else{b.ipt=null;b.idt=null;extra.push(b);}
  });
  const extByDpt=extra.slice().sort((a,b)=>a.dpt-b.dpt);

  // ベース解：先着順の実ルートをそのまま使い（約束を必ず満たす）、当日相当を順に詰める
  const baseRoutes={};for(const v of fleet)baseRoutes[v.id]=[...(greedy.routes[v.id]||[])];
  const baseMap={...greedy.resMap};
  const placedBase={};greedy.rows.forEach((r,idx)=>{if(r.ok)placedBase[idx]={vehicle:r.vehicle,apt:r.apt,idt:r.idt,dev:r.dev,extra:false};});
  let okEbase=0;
  for(const b of extByDpt){
    const r=mk(b);
    const cs=searchInsertions(baseRoutes,baseMap,r,P,fleet);
    if(cs.length){const c=cs[0];r.ipt=c.apt;r.idt=c.apt+P.dwell+b.mrt;r.vehicle=c.vehicle;
      baseRoutes[c.vehicle]=c.route;baseMap[r.id]=r;okEbase++;
      placedBase[b.idx]={vehicle:c.vehicle,apt:c.apt,idt:r.idt,dev:c.apt-b.dpt,extra:true};}
  }
  let best={served:greedy.ok+okEbase,span:totalSpanOf(baseRoutes,baseMap,P,fleet),placed:placedBase,okE:okEbase,
    routes:baseRoutes,resMap:baseMap};

  // 全シャッフル再構築（約束は固定）。成立が増えるか、同点で走行が縮むなら採用。
  for(let a=0;a<Math.max(tries,6);a++){
    const order=[...shuffleArr(promised.slice()),...shuffleArr(extra.slice())];
    const rt={};for(const v of fleet)rt[v.id]=[];
    const rm={};const placed={};let okP=0,okE=0,feasible=true;
    for(const b of order){
      const r=mk(b);
      const cs=searchInsertions(rt,rm,r,P,fleet);
      if(cs.length){const c=cs[0];
        r.ipt=b.ipt!=null?b.ipt:c.apt;r.idt=b.idt!=null?b.idt:c.apt+P.dwell+b.mrt;r.vehicle=c.vehicle;
        rt[c.vehicle]=c.route;rm[r.id]=r;
        placed[b.idx]={vehicle:c.vehicle,apt:c.apt,idt:r.idt,dev:c.apt-b.dpt,extra:b.ipt==null};
        if(b.ipt!=null)okP++;else okE++;
      }else if(b.ipt!=null){feasible=false;break;}   // 約束を守れない並べ替えは棄却
    }
    if(!feasible)continue;
    const served=okP+okE,span=totalSpanOf(rt,rm,P,fleet);
    if(served>best.served||(served===best.served&&span<best.span))best={served,span,placed,okE,routes:rt,resMap:rm};
  }
  let nm=0;
  const out=rows.map((q,idx)=>{
    const p=best.placed[idx];
    if(p)return {...q,ok:true,num:++nm,vehicle:p.vehicle,apt:p.apt,idt:p.idt,dev:p.dev,reassigned:0,rescued:!!p.extra};
    return {...q,ok:false,reason:"一括最適化でも約束を守って入らず"};
  });
  return {greedy,reopt:{rows:out,ok:best.served,total:rows.length,span:best.span,reassign:0,rescued:best.okE},
    plan:{routes:best.routes,resMap:best.resMap}};
}

/* ---------------- 初期データ（Excelの要求A/Bを再現） ---------------- */
// 背骨（{stop,time}の列）→ アンカーイベント列（時刻の杭）。時刻順に整列。
function anchorEvents(backbone){
  return [...(backbone||[])].filter(a=>a&&a.stop!=null&&a.time!=null)
    .sort((a,b)=>a.time-b.time)
    .map(a=>({type:"ANCHOR",stop:a.stop,atime:a.time}));
}

// セミ車ルートの区間内秩序チェック。
// ルートをANCHOR（時刻の杭）で区切り、各区間内で
//   (1) 区間始点からの所要が単調非減少
//   (2) 次アンカーへの残所要が単調非増加（＝一歩ごとに次の杭へ近づく）
// を要求する。(2)が本質的な測度で、区間の進行方向と逆向きの停留所を弾く。
// 背骨の時間窓の外（最初のアンカー前・最終アンカー以降）は制約なし＝フルデマンドと
// 同じ自由営業。これにより「ピーク時間帯だけ背骨・それ以外はフル」のハイブリッド運行を
// 1台で表現できる。最初のアンカーへの定時到達はsimulateの時刻杭が保証するため、
// 窓前の方向制約は不要（間に合わない挿入は時間面で棄却される）。
// ※旧版は最終アンカー以降の営業を禁止していた（比較純度のため）が、時間帯限定セミの
//   要求により撤回。窓外をフルとして使うか遊ばせるかは背骨の敷き方（終了時刻）で制御する。
// 限界：発と着が別区間に落ちる逆向きODは順序規則では検出されないが、
// アンカーの時刻杭（区間の持ち時間＝直行TT×余裕率）が大きな逆行寄り道を
// 時間面で棄却するため、通り得るのは余裕内に収まる小さな寄り道に限られる。
function bbSegOrdered(route){
  const anchorIdx=[];
  route.forEach((e,i)=>{if(e.type==="ANCHOR")anchorIdx.push(i);});
  if(anchorIdx.length===0)return true;
  // 各アンカー区間（窓の内側のみ制約）
  for(let k=0;k+1<anchorIdx.length;k++){
    const s=route[anchorIdx[k]].stop, e=route[anchorIdx[k+1]].stop;
    let pf=-Infinity, pr=Infinity;
    for(let i=anchorIdx[k]+1;i<anchorIdx[k+1];i++){
      const f=TT[s][route[i].stop], r=TT[route[i].stop][e];
      if(f<pf-1 || r>pr+1)return false;
      pf=f; pr=r;
    }
  }
  return true;
}

function buildInitial(P, vehicles){
  const resMap={}; const routes={1:[],2:[],3:[],4:[],5:[]};
  const seed=[
    {id:"A", o:0, d:1, dpt:9*3600,       sa:600, pax:1, label:"要求A"},  // BS001→BS002 9:00発希望
    {id:"B", o:2, d:6, dpt:9*3600+30*60, sa:600, pax:1, label:"要求B"},  // BS003→BS007 9:30発希望
  ].filter(s=>s.o<STOPS.length && s.d<STOPS.length);
  let num=1001;
  for(const s of seed){
    const drt=TT[s.o][s.d];
    const r={...s, num:num++, drt, mrt:mrtFromDRT(drt,P), sa:s.sa??600, ipt:null, idt:null, vehicle:null};
    const cands=searchInsertions(routes,resMap,r,P,vehicles);
    if(cands.length){
      const c=cands[0];
      r.ipt=c.apt; r.idt=c.apt+P.dwell+r.mrt; r.vehicle=c.vehicle;
      routes[c.vehicle]=c.route; resMap[r.id]=r;
    }
  }
  return {resMap,routes,nextNum:num};
}

/* ---------------- UI ---------------- */
const S = {
  page:{display:"flex",flexDirection:"column",height:"100vh",background:"#EFEDE7",
    fontFamily:"'Hiragino Kaku Gothic ProN','Hiragino Sans','Yu Gothic',sans-serif",color:"#1E2A38"},
  mono:{fontFamily:"'SF Mono','Consolas',monospace",fontVariantNumeric:"tabular-nums"},
};

function TimeInput({value,onChange}){
  return <input type="time" value={value} onChange={e=>onChange(e.target.value)} step={300}
    style={{...S.mono,padding:"6px 8px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:14,background:"#fff"}}/>;
}

export default function App(){
  const [sl,setSl]=useState({d1:300,d2:1200,s1:480,s2:900}); // 寄り道ST三段式：短≤5分→8分／長≥20分→15分
  const [tw,setTw]=useState(300);
  const [dwell,setDwell]=useState(60);
  const P=useMemo(()=>({sl,tw,dwell}),[sl,tw,dwell]);

  const [vehicles,setVehicles]=useState(DEFAULT_VEHICLES);
  const [showVeh,setShowVeh]=useState(false);
  const [showImport,setShowImport]=useState(false);
  // 運行再生（シミュレーション時計）
  const [simTime,setSimTime]=useState(9*3600);
  const [playing,setPlaying]=useState(false);
  const [playOpen,setPlayOpen]=useState(false);      // 運行再生バーの開閉（既定は畳む＝右上のコリドーパネルと領域を分離）
  const [demandViz,setDemandViz]=useState(null);   // 生成/取込した需要（OD放物線用・モーダル再オープン時の復元にも使用）
  const [demandMeta,setDemandMeta]=useState(null);  // 生成設定/取込元（記録用・モーダルをまたいで保持）
  const [showDemand,setShowDemand]=useState(false); // 需要レイヤの表示トグル
  const [showCorridor,setShowCorridor]=useState(false); // コリドー軸の表示トグル
  const [showMarks,setShowMarks]=useState(true);      // 乗降時刻ラベルの表示トグル（simTime近傍の窓のみ表示）
  const [corVeh,setCorVeh]=useState(1);              // コリドー→背骨の適用先号車
  const [corStart,setCorStart]=useState("9:00");     // 背骨の始発時刻（h:mm）
  const [corEnd,setCorEnd]=useState("");             // 背骨の終了時刻（空＝便数指定。指定時は便数を自動算出＝時間帯限定セミ）
  const [corSlack,setCorSlack]=useState(140);        // 余裕率（%）：アンカー間隔＝直行TT×余裕率
  const [corMode,setCorMode]=useState("roundtrip");  // 背骨の方式：片道／往復／循環
  const [corTrips,setCorTrips]=useState(4);          // 便数（片道1本＝1便）
  const [corThin,setCorThin]=useState(480);          // アンカー密度：杭の最小間隔秒（0=全停留所）
  const [corCompare,setCorCompare]=useState(null);   // セミ有効性比較の試算結果
  const [armClear,setArmClear]=useState(false); // 全消去の2クリック確認
  const [netVer,setNetVer]=useState(0);             // ネットワーク再構築時の再描画トリガ
  const [focusVeh,setFocusVeh]=useState(null);       // マップで表示する号車（null=全車）
  // 車両の運行モード（フル/セミ）と背骨を適用。セミ化時はその号車のルートをアンカー列で初期化し、
  // その号車の既存予約は一旦外す（背骨は予約前に定義する前提の第一版）。
  const applyBackbone=(vehId,mode,backbone)=>{
    let cleared=0;
    setVehicles(vs=>vs.map(v=>v.id===vehId?{...v,mode,backbone:mode==="semi"?backbone:[]}:v));
    setState(s=>{
      const anchors=mode==="semi"?anchorEvents(backbone):[];
      const resMap={}; for(const id in s.resMap){ if(s.resMap[id].vehicle!==vehId) resMap[id]=s.resMap[id]; else cleared++; }
      return {...s,resMap,routes:{...s.routes,[vehId]:anchors}};
    });
    setCands(null); setTimetable(null); setTtSel(null);
    return {cleared};
  };
  // コリドーパネルの入力（始発・余裕率・方式・便数）から背骨を組み立てる共通処理。
  // 適用と比較試算の双方から呼ぶ。入力不正・運行時間超過はerrで返す。
  const buildCorBackbone=()=>{
    if(!corridorData)return {err:"コリドー軸が未構成。"};
    const m=String(corStart).match(/^(\d{1,2}):(\d{2})$/);
    if(!m)return {err:"始発時刻は h:mm 形式で入力（例 9:00）。"};
    const startSec=(+m[1])*3600+(+m[2])*60;
    let endSec=null;
    if(String(corEnd).trim()!==""){
      const me=String(corEnd).match(/^(\d{1,2}):(\d{2})$/);
      if(!me)return {err:"終了時刻は h:mm 形式で入力（例 11:00）。空欄なら便数指定になる。"};
      endSec=(+me[1])*3600+(+me[2])*60;
      if(endSec<=startSec)return {err:"終了時刻が始発以前になっている。"};
    }
    const slack=Math.max(1,corSlack/100);
    // アンカー間引き：杭を主要停留所（約corThin間隔）に限定し、区間の余裕をプールする。
    // 全停留所を杭にすると区間が1〜4分に細切れになり、寄り道・待ちの原資が実質使えない。
    const stops=thinCorridorStops(corridorData.stops,corThin);
    const bb=corridorToBackbone(stops,startSec,slack,
      {mode:corMode,trips:corMode==="oneway"?1:corTrips,turnaround:180,endSec});
    const nTrips=Math.max(1,Math.round(bb.length/stops.length));
    const last=bb[bb.length-1];
    const veh=vehicles.find(v=>v.id===corVeh);
    if(veh&&last.time>veh.end)
      return {err:`終便の終点アンカー ${fmt(last.time)} が${veh.name}の運行終了 ${fmt(veh.end)} を超える。便数を減らすか、始発を早めるか、余裕率を下げること。`};
    if(veh&&startSec<veh.start)
      return {err:`始発 ${fmt(startSec)} が${veh.name}の運行開始 ${fmt(veh.start)} より早い。`};
    return {bb,startSec,veh,nTrips};
  };
  // ルートを白紙化する際、セミ車の背骨アンカーは保持する（予約は消すが運行の骨格は残す）。
  // これをしないと routes だけ空になり vehicles 側は mode:"semi" のまま残るため、
  // 背骨のない「見かけフル」状態になって運行分析でセミ車がフルとして集計される。
  const emptyRoutesKeepingBackbone=()=>{
    const r={1:[],2:[],3:[],4:[],5:[]};
    for(const v of vehicles) if(v.mode==="semi"&&v.backbone&&v.backbone.length) r[v.id]=anchorEvents(v.backbone);
    return r;
  };
  // 停留所・OD表を読み込んでネットワークを差し替え、状態をリセットする
  const loadNetwork=(stops,odRows,odFactor)=>{
    const info=applyNetwork(stops,odRows,odFactor);
    // ネットワーク差し替えでは停留所IDの意味が変わりうるため、背骨も含め完全リセットが安全
    setState({resMap:{},routes:{1:[],2:[],3:[],4:[],5:[]},nextNum:1001,lastDemand:null});
    setVehicles(vs=>vs.map(v=>({...v,mode:"full",backbone:[]})));  // 背骨は停留所前提が崩れるので解除
    setDemandViz(null); setDemandMeta(null); setShowDemand(false); setShowCorridor(false);
    setFo(0); setFd(Math.min(1,STOPS.length-1));
    setSimTime(9*3600); setPlaying(false);
    setNetVer(v=>v+1);
    return info;
  };
  // 確定予約・ルート・直近の流し込み記録をすべて白紙に戻す（2クリック確認）
  const clearReservations=()=>{
    setState(s=>({...s,resMap:{},routes:emptyRoutesKeepingBackbone(),nextNum:1001,lastDemand:null}));
    setDemandViz(null); setDemandMeta(null); setShowDemand(false); setShowCorridor(false);
    setArmClear(false);
  };
  const [speed,setSpeed]=useState(120);          // 実時間1秒＝120秒（2分）
  useEffect(()=>{
    if(!playing)return;
    const h=setInterval(()=>{
      setSimTime(t=>{
        const nt=t+speed*0.1;                    // 100ms刻み
        if(nt>=17.5*3600){setPlaying(false);return 17.5*3600;}
        return nt;
      });
    },100);
    return ()=>clearInterval(h);
  },[playing,speed]);
  const [state,setState]=useState(()=>buildInitial({sl:{d1:300,d2:1200,s1:480,s2:900},tw:300,dwell:60},DEFAULT_VEHICLES));
  const {resMap,routes,lastDemand}=state;
  const activeVehicles=vehicles.filter(v=>v.active);
  // コリドー軸：需要から便益÷費用で構築（需要かネットワークが変われば再計算）
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const corridorData=useMemo(()=>demandViz?buildCorridor(demandViz):null,[demandViz,netVer]);
  const mapVehicles=focusVeh?activeVehicles.filter(v=>v.id===focusVeh):activeVehicles;

  // 受付フォーム
  const [fo,setFo]=useState(3);   // BS004
  const [fd,setFd]=useState(1);   // BS002
  const [mode,setMode]=useState("dep"); // dep:発希望 / arr:着希望
  const [ft,setFt]=useState("09:30");
  const [sa,setSa]=useState(600);       // San:希望とのズレ幅（探索範囲）
  const [pax,setPax]=useState(1);       // 乗車人数
  const [cands,setCands]=useState(null);
  const [sel,setSel]=useState(0);
  const [bookMode,setBookMode]=useState("desired"); // desired:希望時刻型 / timetable:時刻表型
  const [timetable,setTimetable]=useState(null);     // 時刻表型：枠ごとの最良便
  const [ttSel,setTtSel]=useState(null);             // 時刻表型：選択中の枠
  const [msg,setMsg]=useState(null);
  const [showTable,setShowTable]=useState(false);
  const [tableTab,setTableTab]=useState(null);   // 確認表モーダルの初期タブ指定
  const [lastChange,setLastChange]=useState(null); // 直前確定の前後比較スナップショット

  const sims=useMemo(()=>{
    const o={};
    for(const v of vehicles) o[v.id]=simulate(routes[v.id]||[],resMap,P,v);
    return o;
  },[routes,resMap,P,vehicles]);
  // 車両設定の変更で既存計画が守れなくなった車両（警告表示用）
  const brokenVehicles=vehicles.filter(v=>(routes[v.id]||[]).length>0&&(!sims[v.id].ok||!v.active));

  const parseT=s=>{const[h,m]=s.split(":").map(Number);return h*3600+m*60;};

  const doSearch=()=>{
    setMsg(null);
    if(fo===fd){setMsg({t:"err",m:"発地と着地が同一。"});setCands(null);return;}
    const drt=TT[fo][fd];
    const mrt=mrtFromDRT(drt,P);
    let dpt;
    if(mode==="dep") dpt=parseT(ft);
    else dpt=parseT(ft)-(mrt+P.dwell); // 着希望→最大乗車でも間に合う発時刻に変換
    const earliest=Math.min(...activeVehicles.map(v=>v.start));
    if(dpt+sa<earliest){setMsg({t:"err",m:`全稼働車両の運行開始（最早 ${fmt(earliest)}）より前。着希望の場合は寄り道時間を見込んだ発時刻が運行時間内である必要がある。`});setCands(null);return;}
    const id="R"+Date.now();
    const r={id,num:state.nextNum,o:fo,d:fd,mode,dpt,drt,mrt,sa,pax,ipt:null,idt:null,vehicle:null,
      label:`予約${state.nextNum}`};
    const c=searchInsertions(routes,resMap,r,P,vehicles);
    if(c.length===0){setMsg({t:"err",m:"全車両で約束を守れる挿入位置なし。時刻変更を提案。"});setCands(null);return;}
    setCands({res:r,list:c});setSel(0);
  };

  // 時刻表型：終日を15分枠で走査し、各枠でそのODが成立する最良便（効率＝added中心）を1つ求める
  const computeTimetable=()=>{
    setMsg(null); setCands(null); setTtSel(null);
    if(fo===fd){setMsg({t:"err",m:"発地と着地が同一。"});setTimetable(null);return;}
    if(activeVehicles.length===0){setMsg({t:"err",m:"稼働車両がない。"});setTimetable(null);return;}
    const drt=TT[fo][fd], mrt=mrtFromDRT(drt,P);
    const T0=Math.min(...activeVehicles.map(v=>v.start));
    const T1=Math.max(...activeVehicles.map(v=>v.end));
    const step=15*60, stamp=Date.now(), slots=[];
    for(let t=Math.ceil(T0/step)*step; t<T1; t+=step){
      const center=t+step/2;
      const r={id:"TT"+t+"_"+stamp,o:fo,d:fd,mode:"dep",dpt:center,drt,mrt,sa:step/2,
        pax,ipt:null,idt:null,vehicle:null,label:"仮"};
      const c=searchInsertions(routes,resMap,r,P,vehicles);
      if(c.length){ const best=c[0]; slots.push({slot:t,...best,res:{...r,dpt:best.apt}}); }
      else slots.push({slot:t,none:true});
    }
    setTimetable(slots);
    const okN=slots.filter(s=>!s.none).length;
    setMsg({t:okN?"ok":"err",m:okN?`成立する便を ${okN}/${slots.length} 枠で提示。1つ選ぶと予約成立（全便が走るダイヤではない）。`:"終日どの枠でも約束を守れる便がない。"});
  };

  // 候補（希望時刻型・時刻表型 共通）を確定してstateへ反映
  const commitCandidate=(c,resBase)=>{
    const r={...resBase,num:state.nextNum,label:`予約${state.nextNum}`};
    r.ipt=c.apt; r.idt=c.apt+P.dwell+r.mrt; r.vehicle=c.vehicle;
    const beforeSim=sims[c.vehicle];
    setLastChange({vehicle:c.vehicle,newResId:r.id,before:beforeSim.ok?beforeSim.events:[],
      after:c.sim.events,resMap:{...resMap,[r.id]:r},P:{...P},stamp:new Date()});
    setState(s=>({...s,resMap:{...s.resMap,[r.id]:r},routes:{...s.routes,[c.vehicle]:c.route},nextNum:s.nextNum+1}));
    setMsg({t:"ok",m:`予約${r.num}を確定。${vehicles.find(v=>v.id===c.vehicle).name}・約束発 ${fmt(c.apt)}（〜${fmt(c.apt+P.tw)}）・約束着 ${fmt(c.apt+P.dwell+r.mrt)}。`});
  };

  const confirm=()=>{
    if(!cands)return;
    commitCandidate(cands.list[sel],cands.res);
    setCands(null);
  };

  const confirmTimetable=()=>{
    if(ttSel==null||!timetable||!timetable[ttSel]||timetable[ttSel].none)return;
    const c=timetable[ttSel];
    commitCandidate(c,c.res);
    setTimetable(null); setTtSel(null);
  };

  const cancel=(id)=>{
    setState(s=>{
      const r=s.resMap[id];
      const routes2={...s.routes};
      routes2[r.vehicle]=routes2[r.vehicle].filter(e=>e.resId!==id);
      const m2={...s.resMap};delete m2[id];
      return {...s,resMap:m2,routes:routes2};
    });
    setMsg({t:"ok",m:"予約を取消し、運行計画を再構成した。"});
  };

  const resList=Object.values(resMap).sort((a,b)=>a.ipt-b.ipt);

  /* ---- 候補プレビュー（マップ・タイムライン用。両モード対応） ---- */
  const preview = bookMode==="timetable"
    ? (ttSel!=null && timetable && timetable[ttSel] && !timetable[ttSel].none ? timetable[ttSel] : null)
    : (cands?cands.list[sel]:null);
  const previewResObj = bookMode==="timetable"
    ? (ttSel!=null && timetable && timetable[ttSel] ? timetable[ttSel].res : null)
    : cands?.res;

  return (
  <div style={S.page}>
    {/* ヘッダ */}
    <div style={{display:"flex",alignItems:"center",gap:14,padding:"10px 18px",
      background:"#14202F",color:"#E8E4DA"}}>
      <div>
        <div style={{fontSize:11,letterSpacing:3,color:"#7E8CA0"}}>DRT (DEMAND RESPONSIVE TRANSIT) SIMULATOR</div>
        <div style={{fontSize:18,fontWeight:700}}>DRTシミュレーター</div>
      </div>

      {/* 件数バッジ：流し込み／確定を常時明示 */}
      <div style={{display:"flex",alignItems:"stretch",gap:0,marginLeft:6,
        border:"1px solid #38465C",borderRadius:9,overflow:"hidden"}}>
        {lastDemand&&lastDemand.rows&&lastDemand.rows.length>0&&<>
          <div style={{padding:"5px 12px",textAlign:"center",background:"#1B2942"}}>
            <div style={{fontSize:9.5,color:"#8FA0B8",letterSpacing:1}}>流し込み</div>
            <div style={{fontSize:16,fontWeight:800,fontFamily:"'SF Mono','Consolas',monospace",lineHeight:1.1}}>{lastDemand.rows.length}<span style={{fontSize:10,fontWeight:400,color:"#8FA0B8"}}> 件</span></div>
          </div>
          <div style={{padding:"5px 12px",textAlign:"center",background:"#173026"}}>
            <div style={{fontSize:9.5,color:"#7DBd9C",letterSpacing:1}}>確定</div>
            <div style={{fontSize:16,fontWeight:800,fontFamily:"'SF Mono','Consolas',monospace",lineHeight:1.1,color:"#7BE0A8"}}>{lastDemand.rows.filter(r=>r.ok).length}<span style={{fontSize:10,fontWeight:400,color:"#7DBd9C"}}> 件</span></div>
          </div>
          <div style={{padding:"5px 12px",textAlign:"center",background:"#301C1C"}}>
            <div style={{fontSize:9.5,color:"#D89A8E",letterSpacing:1}}>不成立</div>
            <div style={{fontSize:16,fontWeight:800,fontFamily:"'SF Mono','Consolas',monospace",lineHeight:1.1,color:"#E9A090"}}>{lastDemand.rows.length-lastDemand.rows.filter(r=>r.ok).length}<span style={{fontSize:10,fontWeight:400,color:"#D89A8E"}}> 件</span></div>
          </div>
        </>}
        <div style={{padding:"5px 12px",textAlign:"center",background:"#1F2D40"}}>
          <div style={{fontSize:9.5,color:"#8FA0B8",letterSpacing:1}}>現在の予約</div>
          <div style={{fontSize:16,fontWeight:800,fontFamily:"'SF Mono','Consolas',monospace",lineHeight:1.1}}>{Object.keys(resMap).length}<span style={{fontSize:10,fontWeight:400,color:"#8FA0B8"}}> 件</span></div>
        </div>
      </div>

      <div style={{flex:1}}/>
      <button onClick={()=>{ if(armClear){clearReservations();} else {setArmClear(true); setTimeout(()=>setArmClear(false),3000);} }}
        style={{padding:"8px 14px",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:700,
          border:`1px solid ${armClear?"#E8607D":"#5C4048"}`,
          background:armClear?"#7A2A38":"#2A1F24",color:armClear?"#FFE1E6":"#D8A9B2"}}>
        {armClear?"本当に全消去？（もう一度）":"予約を全消去"}
      </button>
      <button onClick={()=>setShowVeh(true)}
        style={{padding:"8px 14px",borderRadius:8,border:"1px solid #4D6485",cursor:"pointer",
          background:"#1F2D40",color:"#E8E4DA",fontSize:13,fontWeight:700}}>
        車両設定（{activeVehicles.length}台）
      </button>
      <button onClick={()=>{setTableTab(null);setShowTable(true);}}
        style={{padding:"8px 14px",borderRadius:8,border:"1px solid #4D6485",cursor:"pointer",
          background:"#1F2D40",color:"#E8E4DA",fontSize:13,fontWeight:700}}>
        運行確認表
      </button>
      <button onClick={()=>{setTableTab("ana");setShowTable(true);}}
        style={{padding:"8px 14px",borderRadius:8,border:"1px solid #4D6485",cursor:"pointer",
          background:"#1F2D40",color:"#E8E4DA",fontSize:13,fontWeight:700}}>
        運行分析
      </button>
    </div>

    <div style={{display:"flex",flex:1,minHeight:0}}>
      {/* 左：受付パネル */}
      <div style={{width:330,padding:14,overflowY:"auto",borderRight:"1px solid #D8D3C6",background:"#F5F3EE"}}>
        <SectionTitle>予約受付</SectionTitle>
        <div style={{display:"flex",gap:6,marginBottom:8}}>
          <Toggle on={bookMode==="desired"} onClick={()=>{setBookMode("desired");setTimetable(null);setTtSel(null);setMsg(null);}}>希望時刻型</Toggle>
          <Toggle on={bookMode==="timetable"} onClick={()=>{setBookMode("timetable");setCands(null);setMsg(null);}}>時刻表型</Toggle>
        </div>
        <div style={{background:"#fff",border:"1px solid #D8D3C6",borderRadius:10,padding:14}}>
          <Row label="発地">
            <StopSelect value={fo} onChange={setFo}/>
          </Row>
          <Row label="着地">
            <StopSelect value={fd} onChange={setFd}/>
          </Row>
          {bookMode==="desired"&&<>
          <Row label="希望">
            <div style={{display:"flex",gap:6}}>
              <Toggle on={mode==="dep"} onClick={()=>setMode("dep")}>発希望</Toggle>
              <Toggle on={mode==="arr"} onClick={()=>setMode("arr")}>着希望</Toggle>
            </div>
          </Row>
          <Row label="時刻">
            <TimeInput value={ft} onChange={setFt}/>
          </Row>
          </>}
          <Row label="人数">
            <select value={pax} onChange={e=>setPax(Number(e.target.value))}
              style={{padding:"6px 8px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:14,background:"#fff"}}>
              {[1,2,3,4,5,6].map(n=><option key={n} value={n}>{n}名</option>)}
            </select>
          </Row>
          {bookMode==="desired"&&<Row label="ズレ幅">
            <select value={sa} onChange={e=>setSa(Number(e.target.value))}
              style={{padding:"6px 8px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:14,background:"#fff"}}>
              <option value={300}>前後5分</option>
              <option value={600}>前後10分</option>
              <option value={900}>前後15分</option>
              <option value={1200}>前後20分</option>
              <option value={1800}>前後30分</option>
              <option value={2700}>前後45分</option>
              <option value={3600}>前後60分</option>
              <option value={5400}>前後90分</option>
              <option value={7200}>前後120分</option>
            </select>
            <span style={{fontSize:10,color:"#8A8474"}}>San:希望±探索</span>
          </Row>}
          <Row label="乗降 Dwell">
            <select value={dwell} onChange={e=>setDwell(Number(e.target.value))}
              style={{padding:"6px 8px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:14,background:"#fff"}}>
              <option value={60}>60秒</option><option value={120}>120秒</option>
            </select>
            <span style={{fontSize:10,color:"#8A8474"}}>各停での乗降時間</span>
          </Row>
          <div style={{fontSize:11,color:"#8A8474",margin:"8px 0"}}>
            直行 {fmtMin(TT[fo][fd])} ／ 寄り道許容 +{fmtMin(stFromDRT(TT[fo][fd],P))}（最大乗車 {fmtMin(mrtFromDRT(TT[fo][fd],P))}）
          </div>
          {bookMode==="desired"
            ?<button onClick={doSearch} style={btnPrimary}>運行計画を検索</button>
            :<button onClick={computeTimetable} style={btnPrimary}>時刻表を作成（終日）</button>}
        </div>
        <button onClick={()=>setShowImport(true)}
          style={{...btnPrimary,marginTop:8,background:"#fff",color:"#14202F",border:"1px solid #14202F"}}>
          予約一覧の流し込み（CSV / Excel）
        </button>

        {brokenVehicles.length>0&&<div style={{marginTop:10,padding:"8px 10px",borderRadius:8,fontSize:12,
          background:"#F8E7E3",color:"#9B3B2B"}}>
          警告：車両設定の変更により {brokenVehicles.map(v=>v.name).join("・")} の既存計画に問題
          （{brokenVehicles.map(v=>v.active?sims[v.id].why:"稼働OFFだが予約あり").join("／")}）。設定を戻すか、該当予約の取消・再受付が必要。
        </div>}
        {msg && <div style={{marginTop:10,padding:"8px 10px",borderRadius:8,fontSize:12,
          background:msg.t==="ok"?"#E4F2E9":"#F8E7E3",color:msg.t==="ok"?"#23694A":"#9B3B2B"}}>{msg.m}</div>}

        {cands && <div style={{marginTop:12}}>
          <SectionTitle>挿入候補（約束を守れる便のみ）</SectionTitle>
          {cands.list.map((c,i)=>{
            const v=vehicles.find(x=>x.id===c.vehicle);
            return (
            <div key={i} onClick={()=>setSel(i)}
              style={{cursor:"pointer",marginBottom:8,padding:12,borderRadius:10,background:"#fff",
                border:i===sel?`2px solid ${v.color}`:"1px solid #D8D3C6"}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <Dot c={v.color}/><b>{v.name}</b>
                {i===0&&<span style={{fontSize:10,background:"#14202F",color:"#fff",borderRadius:4,padding:"1px 6px"}}>推奨</span>}
                <span style={{marginLeft:"auto",fontSize:11,color:"#8A8474"}}>車両拘束 +{fmtMin(c.added)}</span>
              </div>
              <div style={{...S.mono,fontSize:13,marginTop:6}}>
                約束発 <b>{fmt(c.apt)}</b>（〜{fmt(c.apt+P.tw)}）<br/>
                約束着 <b>{fmt(c.apt+P.dwell+cands.res.mrt)}</b>（計画着 {fmt(c.adt)}）
              </div>
              <div style={{fontSize:11,color:"#8A8474",marginTop:4}}>
                希望とのずれ {fmtDev(c.dev)}（前後{fmtMin(cands.res.sa)}以内）／ 計画乗車 {fmtMin(c.adt-c.apt-P.dwell)}（最大 {fmtMin(cands.res.mrt)}）<br/>
                車内最大 {c.maxLoad}名／定員 {c.cap}名
              </div>
            </div>);
          })}
          <button onClick={confirm} style={{...btnPrimary,background:"#2E9E6B"}}>この内容で確定（約束成立）</button>
        </div>}

        {bookMode==="timetable" && timetable && (()=>{
          // 成立便を時（発時刻の時）ごとにまとめる。分セルを押して便を選ぶ。
          const byHour={};
          timetable.forEach((s,i)=>{
            if(s.none)return;
            const h=Math.floor(s.apt/3600);
            (byHour[h]=byHour[h]||[]).push({m:Math.floor((s.apt%3600)/60),idx:i});
          });
          const hours=Object.keys(byHour).map(Number).sort((a,b)=>a-b);
          hours.forEach(h=>byHour[h].sort((a,b)=>a.m-b.m));
          const selSlot=ttSel!=null?timetable[ttSel]:null;
          const selV=selSlot&&!selSlot.none?vehicles.find(x=>x.id===selSlot.vehicle):null;
          return (
          <div style={{marginTop:12}}>
            <SectionTitle>運行時刻表（{STOPS[fd]} 行き・択一）</SectionTitle>
            <div style={{fontSize:11,color:"#8A6D1F",background:"#FBF4E6",border:"1px solid #EBD9A8",
              borderRadius:8,padding:"6px 9px",marginBottom:8,lineHeight:1.6}}>
              これは択一の候補一覧。1つ選んで予約すると成立し、他は消える。全便が同時に走るダイヤではない。
            </div>
            {hours.length===0
              ?<div style={{fontSize:12,color:"#8A8474"}}>成立する便がない。</div>
              :<div style={{maxHeight:320,overflowY:"auto",border:"1px solid #9AA7B5",
                borderRadius:8,background:"#C7D2DB"}}>
                {hours.map(h=>(
                  <div key={h} style={{display:"flex",alignItems:"stretch",borderBottom:"1px solid #9AA7B5"}}>
                    <div style={{width:44,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",
                      fontSize:20,fontWeight:700,color:"#2A323C",borderRight:"1px solid #9AA7B5"}}>{h}</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:6,padding:"6px 8px"}}>
                      {byHour[h].map(({m,idx})=>{
                        const on=ttSel===idx;
                        return (
                        <div key={idx} onClick={()=>setTtSel(idx)}
                          style={{cursor:"pointer",minWidth:38,textAlign:"center",padding:"5px 8px",borderRadius:6,
                            fontSize:17,fontWeight:700,fontFamily:"'SF Mono','Consolas',monospace",
                            background:on?"#2E9E6B":"#fff",color:on?"#fff":"#2A323C",
                            border:on?"2px solid #23694A":"1px solid #B9C2CC"}}>
                          {String(m).padStart(2,"0")}
                        </div>);
                      })}
                    </div>
                  </div>
                ))}
              </div>}
            {selSlot&&!selSlot.none&&<div style={{marginTop:8,fontSize:12,color:"#3A3526",
              background:"#DCEAE2",border:"1px solid #A7CDB5",borderRadius:8,padding:"7px 10px"}}>
              <span style={{...S.mono,fontWeight:700}}>{fmt(selSlot.apt)} 発 → {fmt(selSlot.adt)} 着</span>
              {selV&&<span style={{marginLeft:8}}><Dot c={selV.color}/> {selV.name}</span>}
              <span style={{marginLeft:8,color:"#8A8474"}}>車両拘束+{fmtMin(selSlot.added)}</span>
            </div>}
            <button onClick={confirmTimetable} disabled={ttSel==null}
              style={{...btnPrimary,marginTop:8,background:ttSel==null?"#C9C4B8":"#2E9E6B",
                cursor:ttSel==null?"not-allowed":"pointer"}}>
              {ttSel!=null?`${fmt(timetable[ttSel].apt)}発の便で予約`:"時刻を選択"}
            </button>
          </div>);
        })()}

        <SectionTitle style={{marginTop:16}}>確定予約一覧</SectionTitle>
        {resList.length===0&&<div style={{fontSize:12,color:"#8A8474"}}>確定予約なし。</div>}
        {resList.map(r=>{
          const v=vehicles.find(x=>x.id===r.vehicle);
          const sim=sims[r.vehicle];
          const ev=sim.ok?sim.events.filter(e=>e.resId===r.id):[];
          return (
          <div key={r.id} style={{background:"#fff",border:"1px solid #D8D3C6",borderRadius:10,padding:10,marginBottom:8}}>
            <div style={{display:"flex",alignItems:"center",gap:8,fontSize:13}}>
              <Dot c={v.color}/><b>#{r.num}</b>
              <span>{STOPS[r.o]} → {STOPS[r.d]}</span>
              <span style={{fontSize:11,color:"#8A8474"}}>{r.pax??1}名</span>
              <button onClick={()=>cancel(r.id)} style={{marginLeft:"auto",fontSize:11,border:"1px solid #C9C4B8",
                background:"#fff",borderRadius:6,padding:"2px 8px",cursor:"pointer",color:"#9B3B2B"}}>取消</button>
            </div>
            <div style={{...S.mono,fontSize:12,marginTop:4,color:"#444"}}>
              約束発 {fmt(r.ipt)}〜{fmt(r.ipt+P.tw)} ／ 約束着 {fmt(r.idt)}
            </div>
            {ev.length===2&&<div style={{...S.mono,fontSize:12,color:"#23694A"}}>
              計画：乗車 {fmt(ev[0].apt)} → 降車 {fmt(ev[1].adt)}（乗車 {fmtMin(ev[1].adt-ev[0].etd)}）
            </div>}
          </div>);
        })}
      </div>

      {/* 右：運行盤 */}
      <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0}}>
        <div style={{flex:1,minHeight:0,background:"#14202F",position:"relative"}}>
          <BoardMap vehicles={mapVehicles} routes={routes} resMap={resMap} sims={sims} preview={preview} previewRes={previewResObj} simTime={simTime} demandViz={showDemand?demandViz:null} corridor={showCorridor?corridorData:null} showMarks={showMarks}/>
          {/* 需要可視化トグル */}
          <button onClick={()=>setShowDemand(s=>!s)} disabled={!demandViz}
            style={{position:"absolute",left:12,top:10,display:"flex",alignItems:"center",gap:6,
              background:showDemand?"rgba(217,96,125,0.92)":"rgba(20,32,47,0.88)",
              border:`1px solid ${showDemand?"#E8607D":"#38465C"}`,borderRadius:9,padding:"7px 12px",
              cursor:demandViz?"pointer":"not-allowed",opacity:demandViz?1:0.5,
              color:"#E8E4DA",fontSize:12,fontWeight:700}}>
            需要マップ {showDemand?"ON":"OFF"}
            {demandViz&&<span style={{fontWeight:400,fontSize:10.5,opacity:0.85}}>（{demandViz.filter(r=>!r.err&&r.o!==r.d).length}件）</span>}
          </button>
          {/* コリドー軸トグル：その日の需要が最も乗る走行軸を表示 */}
          <button onClick={()=>setShowCorridor(s=>!s)} disabled={!demandViz}
            style={{position:"absolute",left:170,top:10,display:"flex",alignItems:"center",gap:6,
              background:showCorridor?"rgba(232,162,77,0.92)":"rgba(20,32,47,0.88)",
              border:`1px solid ${showCorridor?"#E8A24D":"#38465C"}`,borderRadius:9,padding:"7px 12px",
              cursor:demandViz?"pointer":"not-allowed",opacity:demandViz?1:0.5,
              color:"#E8E4DA",fontSize:12,fontWeight:700}}>
            コリドー軸 {showCorridor?"ON":"OFF"}
          </button>
          {/* 乗降時刻ラベルの表示トグル：既定は再生カーソル±45分の窓のみ表示。OFFで完全非表示 */}
          <button onClick={()=>setShowMarks(s=>!s)}
            style={{position:"absolute",left:280,top:10,display:"flex",alignItems:"center",gap:6,
              background:showMarks?"rgba(87,124,168,0.92)":"rgba(20,32,47,0.88)",
              border:`1px solid ${showMarks?"#7C9CC4":"#38465C"}`,borderRadius:9,padding:"7px 12px",
              cursor:"pointer",color:"#E8E4DA",fontSize:12,fontWeight:700}}>
            乗降ラベル {showMarks?"ON":"OFF"}
          </button>
          {/* コリドー詳細パネル：指標表示＋セミデマンド背骨への適用 */}
          {showCorridor&&corridorData&&(
            <div style={{position:"absolute",right:12,top:10,width:252,maxHeight:"calc(100% - 20px)",
              overflowY:"auto",background:"rgba(20,32,47,0.94)",border:"1px solid #E8A24D",borderRadius:10,
              padding:"9px 11px",color:"#E8E4DA",fontSize:11}}>
              <div style={{fontWeight:800,fontSize:12,color:"#E8A24D",marginBottom:5}}>
                コリドー軸（{corridorData.stops.length}停留所）
              </div>
              <div style={{lineHeight:1.7,color:"#C9CFD9"}}>
                路線延長（片道）：<b style={{color:"#E8E4DA"}}>{(corridorData.lengthSec/60).toFixed(1)}分</b><br/>
                需要カバー率：<b style={{color:"#E8E4DA"}}>{Math.round(corridorData.coverage*100)}%</b>
                （{corridorData.covPax}/{corridorData.totalPax}人）<br/>
                乗車時間の伸び：<b style={{color:"#E8E4DA"}}>平均{corridorData.avgStretch.toFixed(2)}倍</b>（直行比）
              </div>
              <div style={{fontSize:10,color:"#9AA7BA",margin:"4px 0 7px"}}>
                軸に載らない{corridorData.totalPax-corridorData.covPax}人はデマンド車両が受け持つ想定。
              </div>
              <div style={{borderTop:"1px solid #38465C",paddingTop:7}}>
                <div style={{fontWeight:700,marginBottom:4}}>セミデマンド背骨に適用</div>
                <div style={{display:"flex",gap:5,alignItems:"center",marginBottom:5}}>
                  <select value={corVeh} onChange={e=>setCorVeh(+e.target.value)}
                    style={{flex:1,fontSize:11,background:"#1F2D40",color:"#E8E4DA",
                      border:"1px solid #38465C",borderRadius:5,padding:"3px 4px"}}>
                    {activeVehicles.map(v=><option key={v.id} value={v.id}>{v.name}</option>)}
                  </select>
                  <input value={corStart} onChange={e=>setCorStart(e.target.value)}
                    style={{width:46,fontSize:11,background:"#1F2D40",color:"#E8E4DA",
                      border:"1px solid #38465C",borderRadius:5,padding:"3px 4px",textAlign:"center"}}/>
                  <span style={{color:"#9AA7BA"}}>発</span>
                </div>
                {/* 方式（片道・往復・循環）と便数：背骨は「便」の列として敷く */}
                <div style={{display:"flex",gap:5,alignItems:"center",marginBottom:5}}>
                  <select value={corMode} onChange={e=>setCorMode(e.target.value)}
                    style={{flex:1,fontSize:11,background:"#1F2D40",color:"#E8E4DA",
                      border:"1px solid #38465C",borderRadius:5,padding:"3px 4px"}}>
                    <option value="roundtrip">往復（便ごとに向き反転）</option>
                    <option value="loop">循環（毎便同方向）</option>
                    <option value="oneway">片道1本</option>
                  </select>
                  <input type="number" min="1" max="12" value={corTrips} disabled={corMode==="oneway"}
                    onChange={e=>setCorTrips(Math.max(1,Math.min(12,+e.target.value||1)))}
                    style={{width:36,fontSize:11,background:"#1F2D40",color:"#E8E4DA",
                      border:"1px solid #38465C",borderRadius:5,padding:"3px 4px",textAlign:"center",
                      opacity:corMode==="oneway"?0.4:1}}/>
                  <span style={{color:"#9AA7BA"}}>便</span>
                </div>
                {/* 時間帯限定：終了時刻を入れると便数を自動算出し、時間帯外はフルとして動く */}
                <div style={{display:"flex",gap:5,alignItems:"center",marginBottom:5}}>
                  <span style={{color:"#9AA7BA"}}>終了</span>
                  <input value={corEnd} onChange={e=>setCorEnd(e.target.value)} placeholder="空=便数指定"
                    style={{width:74,fontSize:11,background:"#1F2D40",color:"#E8E4DA",
                      border:"1px solid #38465C",borderRadius:5,padding:"3px 4px",textAlign:"center"}}/>
                  <span style={{color:"#9AA7BA",fontSize:10}}>指定時：時間帯に収まる便数を自動。時間帯外はフル営業</span>
                </div>
                <div style={{display:"flex",gap:5,alignItems:"center",marginBottom:5}}>
                  <span style={{color:"#9AA7BA"}}>杭の間隔</span>
                  <select value={corThin} onChange={e=>setCorThin(+e.target.value)}
                    style={{flex:1,fontSize:11,background:"#1F2D40",color:"#E8E4DA",
                      border:"1px solid #38465C",borderRadius:5,padding:"3px 4px"}}>
                    <option value={0}>全停留所（時刻表が細かい・寄り道余地小）</option>
                    <option value={300}>約5分間隔</option>
                    <option value={480}>約8分間隔（推奨）</option>
                    <option value={720}>約12分間隔（寄り道余地大）</option>
                  </select>
                </div>
                <div style={{display:"flex",gap:5,alignItems:"center",marginBottom:7}}>
                  <span style={{color:"#9AA7BA"}}>余裕率</span>
                  <input type="number" min="100" max="300" step="10" value={corSlack}
                    onChange={e=>setCorSlack(+e.target.value)}
                    style={{width:52,fontSize:11,background:"#1F2D40",color:"#E8E4DA",
                      border:"1px solid #38465C",borderRadius:5,padding:"3px 4px",textAlign:"center"}}/>
                  <span style={{color:"#9AA7BA"}}>%（寄り道の原資）</span>
                </div>
                <button onClick={()=>{
                    const g=buildCorBackbone();
                    if(g.err){setMsg({t:"err",m:g.err});return;}
                    const nb=g.bb.length, last=g.bb[nb-1];
                    const r=applyBackbone(corVeh,"semi",g.bb);
                    const modeLabel=(corMode==="roundtrip"?"往復":corMode==="loop"?"循環":"片道")+`${g.nTrips}便`;
                    setMsg({t:"ok",m:`${g.veh?g.veh.name:""}をセミに設定（${modeLabel}・アンカー${nb}点・${fmt(g.startSec)}発→${fmt(last.time)}終着）。この号車の既存予約${r.cleared}件を外した。動きを見るには「予約一覧の流し込み」から同じ需要を自動確定し直すこと（既定で既存予約をクリアしてから流し込むので二重計上にならない）。`});
                  }}
                  style={{width:"100%",fontSize:11.5,fontWeight:700,cursor:"pointer",
                    background:"#E8A24D",color:"#14202F",border:"none",borderRadius:6,padding:"6px 0"}}>
                  この軸を背骨として適用
                </button>
                {/* 有効性比較：同一需要を「全車フル」と「この号車セミ＋残りフル」に流し込んで並べる（状態は変更しない） */}
                <button onClick={()=>{
                    const g=buildCorBackbone();
                    if(g.err){setMsg({t:"err",m:g.err});return;}
                    const rows=(demandViz||[]).filter(r=>!r.err&&r.o!==r.d);
                    if(!rows.length){setMsg({t:"err",m:"比較する需要がない。先に需要を生成すること。"});return;}
                    const res=compareSemiEffect(rows,activeVehicles,corVeh,g.bb,P);
                    setCorCompare({...res,vehName:g.veh?g.veh.name:`${corVeh}号車`,n:rows.length,
                      modeLabel:(corMode==="roundtrip"?"往復":corMode==="loop"?"循環":"片道")+`${g.nTrips}便`+(corEnd?`・${corStart}〜${corEnd}限定`:"")});
                  }}
                  style={{width:"100%",fontSize:11.5,fontWeight:700,cursor:"pointer",marginTop:5,
                    background:"#1F2D40",color:"#E8E4DA",border:"1px solid #E8A24D",borderRadius:6,padding:"6px 0"}}>
                  有効性を比較（全車フル vs セミ＋フル）
                </button>
                {corCompare&&(()=>{
                  const {A,B}=corCompare;
                  const dOk=B.ok-A.ok;
                  const dSpan=B.spanSum-A.spanSum;
                  const cell={padding:"2px 6px",borderBottom:"1px solid #2A3A50",textAlign:"right",
                    fontFamily:"'SF Mono','Consolas',monospace"};
                  const head={...cell,color:"#9AA7BA",fontFamily:"inherit",textAlign:"left"};
                  let verdict, vColor;
                  if(dOk>=0&&dSpan<=0){verdict="セミ有効：成立を落とさず拘束時間を削減。";vColor="#7BC96F";}
                  else if(dOk<0&&dSpan>=0){verdict="セミ不利：成立・拘束時間とも悪化。全車フルが優位。";vColor="#E8718D";}
                  else if(dOk>0){verdict=`トレードオフ：成立+${dOk}件だが拘束+${Math.round(dSpan/60)}分。成立率優先ならセミ、コスト優先ならフル。`;vColor="#E0A93E";}
                  else{verdict=`トレードオフ：拘束${Math.round(dSpan/60)}分削減だが成立${dOk}件。コスト優先ならセミ、成立率優先ならフル。`;vColor="#E0A93E";}
                  return (
                  <div style={{marginTop:7,borderTop:"1px solid #38465C",paddingTop:6}}>
                    <div style={{fontWeight:700,marginBottom:3}}>
                      比較結果（需要{corCompare.n}件・{corCompare.vehName}を{corCompare.modeLabel}のセミに）
                    </div>
                    <table style={{width:"100%",fontSize:10.5,borderCollapse:"collapse",color:"#E8E4DA"}}>
                      <thead><tr>
                        <th style={head}></th>
                        <th style={{...cell,color:"#9AA7BA"}}>全車フル</th>
                        <th style={{...cell,color:"#E8A24D"}}>セミ＋フル</th>
                      </tr></thead>
                      <tbody>
                        <tr><td style={head}>成立</td>
                          <td style={cell}>{A.ok}/{A.total}</td><td style={cell}>{B.ok}/{B.total}</td></tr>
                        <tr><td style={head}>成立率</td>
                          <td style={cell}>{Math.round(A.rate*100)}%</td><td style={cell}>{Math.round(B.rate*100)}%</td></tr>
                        <tr><td style={head}>拘束時間計</td>
                          <td style={cell}>{Math.round(A.spanSum/60)}分</td><td style={cell}>{Math.round(B.spanSum/60)}分</td></tr>
                        <tr><td style={head}>平均ズレ</td>
                          <td style={cell}>{(A.devAvg/60).toFixed(1)}分</td><td style={cell}>{(B.devAvg/60).toFixed(1)}分</td></tr>
                        <tr><td style={head}>平均乗車</td>
                          <td style={cell}>{(A.rideAvg/60).toFixed(1)}分</td><td style={cell}>{(B.rideAvg/60).toFixed(1)}分</td></tr>
                        <tr><td style={head}>拘束/成立</td>
                          <td style={cell}>{A.ok?(A.spanSum/60/A.ok).toFixed(1):"—"}分/件</td>
                          <td style={cell}>{B.ok?(B.spanSum/60/B.ok).toFixed(1):"—"}分/件</td></tr>
                        <tr><td style={head}>セミ車担当</td>
                          <td style={cell}>—</td><td style={cell}>{B.semiCnt}件</td></tr>
                      </tbody>
                    </table>
                    <div style={{marginTop:5,fontSize:10.5,color:vColor,lineHeight:1.5}}>{verdict}</div>
                    <div style={{marginTop:3,fontSize:9.5,color:"#9AA7BA",lineHeight:1.5}}>
                      ※両構成とも空の状態から先着順で流し込んだ試算。現在の予約状態は変更していない。
                      計算は需要100件あたり10秒前後。セミ車は背骨の時間帯のみ路線制約を受け、
                      時間帯の外はフルとして自由に動く。有効性は「成立の減少を最小限に抑えつつ、
                      拘束/成立（1件を運ぶコスト）をどれだけ下げるか」で読む。予約なし乗車・
                      時刻表による需要誘発・運行管理の単純さは試算の外。
                    </div>
                  </div>);
                })()}
              </div>
            </div>
          )}
          {showCorridor&&demandViz&&!corridorData&&(
            <div style={{position:"absolute",right:12,top:10,background:"rgba(20,32,47,0.94)",
              border:"1px solid #38465C",borderRadius:10,padding:"8px 11px",color:"#9AA7BA",fontSize:11}}>
              有効な需要がなくコリドー軸を構成できない。
            </div>
          )}
          {/* 号車フィルタ：選んだ号車だけ地図に表示 */}
          <div style={{position:"absolute",left:12,top:52,display:"flex",gap:5,alignItems:"center",
            background:"rgba(20,32,47,0.88)",border:"1px solid #38465C",borderRadius:9,padding:"5px 8px"}}>
            <span style={{fontSize:10.5,color:"#9AA7BA",marginRight:2}}>表示</span>
            <button onClick={()=>setFocusVeh(null)}
              style={{fontSize:11,fontWeight:700,borderRadius:6,padding:"3px 9px",cursor:"pointer",
                border:`1px solid ${focusVeh===null?"#E8E4DA":"#38465C"}`,
                background:focusVeh===null?"#33445C":"#1F2D40",color:"#E8E4DA"}}>全車</button>
            {activeVehicles.map(v=>(
              <button key={v.id} onClick={()=>setFocusVeh(focusVeh===v.id?null:v.id)}
                style={{fontSize:11,fontWeight:700,borderRadius:6,padding:"3px 9px",cursor:"pointer",
                  display:"flex",alignItems:"center",gap:4,
                  border:`1px solid ${focusVeh===v.id?v.color:"#38465C"}`,
                  background:focusVeh===v.id?"#33445C":"#1F2D40",
                  color:focusVeh===v.id?"#fff":"#9AA7BA"}}>
                <span style={{width:8,height:8,borderRadius:"50%",background:v.color,display:"inline-block"}}/>
                {v.id}
              </button>
            ))}
          </div>
          {/* 運行再生コントロール：右下に配置し既定は畳む（右上はコリドーパネル専用にして衝突を避ける） */}
          <div style={{position:"absolute",right:12,bottom:10,display:"flex",flexDirection:"column",
            alignItems:"flex-end",gap:6}}>
            {playOpen&&(
              <div style={{display:"flex",alignItems:"center",gap:8,
                background:"rgba(20,32,47,0.92)",border:"1px solid #38465C",borderRadius:10,padding:"7px 12px"}}>
                <button onClick={()=>setPlaying(p=>!p)}
                  style={{width:34,height:30,borderRadius:7,border:"none",cursor:"pointer",fontSize:14,
                    background:playing?"#E0A93E":"#2E9E6B",color:"#fff",fontWeight:700}}>
                  {playing?"❚❚":"▶"}
                </button>
                <button onClick={()=>{setPlaying(false);setSimTime(9*3600);}}
                  style={{width:30,height:30,borderRadius:7,border:"1px solid #38465C",cursor:"pointer",
                    background:"#1F2D40",color:"#E8E4DA",fontSize:12}}>⟲</button>
                <span style={{fontFamily:"'SF Mono','Consolas',monospace",fontSize:16,fontWeight:700,
                  color:"#E8E4DA",minWidth:52,textAlign:"center"}}>{fmt(simTime)}</span>
                <input type="range" min={8.5*3600} max={17.5*3600} step={60} value={simTime}
                  onChange={e=>{setPlaying(false);setSimTime(Number(e.target.value));}}
                  style={{width:200}}/>
                <select value={speed} onChange={e=>setSpeed(Number(e.target.value))}
                  style={{background:"#1F2D40",color:"#E8E4DA",border:"1px solid #38465C",
                    borderRadius:6,padding:"4px 6px",fontSize:12}}>
                  <option value={60}>×60</option>
                  <option value={120}>×120</option>
                  <option value={300}>×300</option>
                  <option value={600}>×600</option>
                </select>
                <button onClick={()=>setPlayOpen(false)} title="畳む"
                  style={{width:24,height:24,borderRadius:6,border:"1px solid #38465C",cursor:"pointer",
                    background:"#1F2D40",color:"#9AA7BA",fontSize:11}}>▾</button>
              </div>
            )}
            {!playOpen&&(
              <button onClick={()=>setPlayOpen(true)}
                style={{display:"flex",alignItems:"center",gap:6,
                  background:"rgba(20,32,47,0.88)",border:"1px solid #38465C",borderRadius:9,
                  padding:"6px 11px",cursor:"pointer",color:"#E8E4DA",fontSize:12,fontWeight:700}}>
                {playing?"❚❚":"▶"}
                <span style={{fontFamily:"'SF Mono','Consolas',monospace"}}>{fmt(simTime)}</span>
                <span style={{color:"#9AA7BA",fontWeight:400}}>運行再生 ▴</span>
              </button>
            )}
          </div>
          <div style={{position:"absolute",left:12,bottom:10,display:"flex",gap:14,fontSize:11,color:"#9AA7BA",
            maxWidth:"calc(100% - 190px)",overflow:"hidden",whiteSpace:"nowrap",textOverflow:"ellipsis"}}>
            {activeVehicles.map(v=><span key={v.id}><Dot c={v.color}/> {v.name} 定員{v.cap}・{fmt(v.start)}〜{fmt(v.end)}</span>)}
            <span style={{color:"#E0A93E"}}>- - - 候補で増える区間／細い実線＝その号車の既存ルート</span>
            <span>走行履歴：太さ＝乗車人数（点線＝空車）、30分で消える</span>
            {showDemand&&<span style={{color:"#E8607D"}}>━ 需要（放物線・太さ＝同一ODの件数／円の大きさ＝発着の需要量）</span>}
          </div>
        </div>
        <div style={{height:64+activeVehicles.length*66,background:"#F5F3EE",borderTop:"1px solid #D8D3C6",padding:"8px 14px"}}>
          <Timeline vehicles={activeVehicles} sims={sims} resMap={resMap} P={P} simTime={simTime} preview={preview} previewRes={previewResObj}/>
        </div>
      </div>
    </div>

    {showTable&&<VerifyModal onClose={()=>setShowTable(false)} vehicles={vehicles}
      sims={sims} resMap={resMap} P={P} lastChange={lastChange} initialTab={tableTab} lastDemand={state.lastDemand}/>}
    {showVeh&&<VehicleModal onClose={()=>setShowVeh(false)} vehicles={vehicles} setVehicles={setVehicles}
      sl={sl} setSl={setSl} tw={tw} setTw={setTw} onLoadNetwork={loadNetwork} netVer={netVer}
      onApplyBackbone={applyBackbone}/>}
    {showImport&&<ImportModal onClose={()=>setShowImport(false)} vehicles={vehicles} P={P}
      state={state} setState={setState} setDemandViz={setDemandViz}
      demandViz={demandViz} demandMeta={demandMeta} setDemandMeta={setDemandMeta}
      onShowAnalysis={()=>{setShowImport(false);setTableTab("ana");setShowTable(true);}}/>}
  </div>);
}

/* ---------- 小物 ---------- */
const pSel={background:"#1F2D40",color:"#E8E4DA",border:"1px solid #38465C",borderRadius:6,padding:"4px 6px",fontSize:13};
const btnPrimary={width:"100%",padding:"9px 0",border:"none",borderRadius:8,background:"#14202F",
  color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer"};
const Param=({label,children})=>(
  <label style={{display:"flex",flexDirection:"column",gap:2,fontSize:10,color:"#7E8CA0"}}>{label}{children}</label>);
// 寄り道ST三段式の1値（分単位入力／内部は秒）。0〜120分にクランプ。
const SLNum=({v,on,t})=>(
  <label style={{display:"flex",flexDirection:"column",alignItems:"center",gap:1,fontSize:9,color:"#7E8CA0"}}>{t}
    <input type="number" min={0} max={120} value={Math.round(v/60)}
      onChange={e=>on(Math.max(0,Math.min(120,Number(e.target.value)||0))*60)}
      style={{width:40,background:"#1F2D40",color:"#E8E4DA",border:"1px solid #38465C",
        borderRadius:5,padding:"3px 4px",fontSize:12,textAlign:"center",
        fontFamily:"'SF Mono','Consolas',monospace"}}/>
  </label>);
const SectionTitle=({children,style})=>(
  <div style={{fontSize:12,fontWeight:700,letterSpacing:1,color:"#6B6453",margin:"4px 0 8px",...style}}>{children}</div>);
const Row=({label,children})=>(
  <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
    <div style={{width:42,fontSize:12,color:"#6B6453"}}>{label}</div>{children}</div>);
const Dot=({c})=><span style={{display:"inline-block",width:9,height:9,borderRadius:9,background:c,marginRight:2}}/>;
const Toggle=({on,onClick,children})=>(
  <button onClick={onClick} style={{padding:"5px 12px",borderRadius:6,fontSize:12,cursor:"pointer",
    border:on?"1px solid #14202F":"1px solid #C9C4B8",background:on?"#14202F":"#fff",color:on?"#fff":"#444"}}>{children}</button>);
const StopSelect=({value,onChange})=>(
  <select value={value} onChange={e=>onChange(Number(e.target.value))}
    style={{padding:"6px 8px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:14,background:"#fff"}}>
    {STOPS.map((s,i)=><option key={i} value={i}>{s}</option>)}
  </select>);

/* ---------- 運行盤マップ ---------- */
// 時刻tにおける車両位置（停留所間は直線補間。滞在中は停留所上）
function vehiclePosAt(sim,t){
  if(!sim.ok||sim.events.length===0)return null;
  const ev=sim.events;
  if(t<ev[0].eta-1||t>ev[ev.length-1].etd+1)return null;  // 拘束時間外は表示しない
  for(let i=0;i<ev.length;i++){
    const e=ev[i];
    if(t<=e.etd){
      if(t>=e.eta){
        return {x:POS[e.stop][0],y:POS[e.stop][1],load:e.load,dwell:true,stop:e.stop};
      }
      const prev=ev[i-1];
      const f=Math.max(0,Math.min(1,(t-prev.etd)/(e.eta-prev.etd)));
      const a=POS[prev.stop],b=POS[e.stop];
      return {x:a[0]+(b[0]-a[0])*f,y:a[1]+(b[1]-a[1])*f,load:prev.load,dwell:false};
    }
  }
  return null;
}

const TRAIL_SEC=1800;   // 走行履歴の保持時間（30分で消える）
const TRAIL_STEP=60;    // 履歴の分割幅（秒）。細かいほど滑らかにフェード

function BoardMap({vehicles,routes,resMap,sims,preview,previewRes,simTime,demandViz,corridor,showMarks}){
  // ---- 地図のズーム・パン（viewBox操作） ----
  // ホイール＝カーソル位置を中心に拡縮、ドラッグ＝パン、ボタン＝＋/−/全体。
  // 倍率はviewBox幅900を基準に1〜8倍（最小幅112）。
  const VB_W=900, VB_H=640;
  const svgRef=React.useRef(null);
  const [view,setView]=React.useState({x:0,y:0,w:VB_W,h:VB_H});
  const dragRef=React.useRef(null);
  const clampView=(x,y,w,h)=>({
    x:Math.min(Math.max(x,0),VB_W-w), y:Math.min(Math.max(y,0),VB_H-h), w, h});
  const zoomAt=(px,py,f)=>setView(v=>{
    const w=Math.min(VB_W,Math.max(VB_W/8,v.w*f)), h=w*VB_H/VB_W;
    return clampView(px-(px-v.x)*(w/v.w), py-(py-v.y)*(h/v.h), w, h);
  });
  const clientToSvg=e=>{
    const r=svgRef.current.getBoundingClientRect();
    return {x:view.x+(e.clientX-r.left)/r.width*view.w,
            y:view.y+(e.clientY-r.top)/r.height*view.h};
  };
  React.useEffect(()=>{   // ホイールはネイティブ登録（passive:falseでページスクロールを抑止）
    const el=svgRef.current; if(!el)return;
    const onWheel=e=>{e.preventDefault();const p=clientToSvg(e);zoomAt(p.x,p.y,e.deltaY>0?1.25:0.8);};
    el.addEventListener("wheel",onWheel,{passive:false});
    return ()=>el.removeEventListener("wheel",onWheel);
  });
  const onMouseDown=e=>{
    if(e.button!==0)return;
    dragRef.current={cx:e.clientX,cy:e.clientY,vx:view.x,vy:view.y};
  };
  React.useEffect(()=>{   // ドラッグ中はwindowで追跡（svg外に出ても継続）
    const move=e=>{
      const d=dragRef.current; if(!d||!svgRef.current)return;
      const r=svgRef.current.getBoundingClientRect();
      setView(v=>clampView(d.vx-(e.clientX-d.cx)/r.width*v.w,
                           d.vy-(e.clientY-d.cy)/r.height*v.h, v.w, v.h));
    };
    const up=()=>{dragRef.current=null;};
    window.addEventListener("mousemove",move); window.addEventListener("mouseup",up);
    return ()=>{window.removeEventListener("mousemove",move); window.removeEventListener("mouseup",up);};
  },[]);
  const zoomBtn={width:30,height:28,fontSize:14,fontWeight:700,cursor:"pointer",
    background:"rgba(20,32,47,0.88)",color:"#E8E4DA",border:"1px solid #38465C",borderRadius:7};
  // 走行履歴：simTimeから過去TRAIL_SEC分の移動を、乗車人数を太さ・経過時間を透明度で描く
  const trails=[];
  if(simTime!=null){
    for(const v of vehicles){
      const sim=sims[v.id];
      if(!sim.ok)continue;
      const ev=sim.events;
      for(let i=1;i<ev.length;i++){
        const t0=ev[i-1].etd, t1=ev[i].eta;
        if(t1<=t0)continue;
        const winA=Math.max(t0,simTime-TRAIL_SEC), winB=Math.min(t1,simTime);
        if(winB<=winA)continue;
        const a=POS[ev[i-1].stop], b=POS[ev[i].stop];
        const load=ev[i-1].load;
        for(let s=winA;s<winB;s+=TRAIL_STEP){
          const e2=Math.min(s+TRAIL_STEP,winB);
          const f0=(s-t0)/(t1-t0), f1=(e2-t0)/(t1-t0);
          const age=simTime-(s+e2)/2;
          trails.push({
            x1:a[0]+(b[0]-a[0])*f0, y1:a[1]+(b[1]-a[1])*f0,
            x2:a[0]+(b[0]-a[0])*f1, y2:a[1]+(b[1]-a[1])*f1,
            w:load>0?2.5+load*1.7:1.6,
            empty:load===0,
            color:v.color,
            op:Math.max(0.05,0.92*(1-age/TRAIL_SEC)),
          });
        }
      }
    }
  }
  const lines=[];
  for(const v of vehicles){
    const sim=sims[v.id];
    if(!sim.ok||sim.events.length<2)continue;
    const off=(v.id-2)*6; // 車両ごとに僅かにずらして重なり回避
    for(let i=0;i<sim.events.length-1;i++){
      const a=POS[sim.events[i].stop],b=POS[sim.events[i+1].stop];
      lines.push({a,b,off,color:v.color,key:`${v.id}-${i}`,seq:i+1,vid:v.id});
    }
  }
  // 候補プレビュー：確定後ルートを「新規予約で増える区間」と「元から走る既存区間」に分ける
  const pvNew=[], pvKeep=[];
  if(preview&&preview.sim.ok){
    const newId=previewRes?.id;
    const ev=preview.sim.events;
    for(let i=0;i<ev.length-1;i++){
      const seg=[POS[ev[i].stop],POS[ev[i+1].stop]];
      // 端点が新規予約の乗降ならその区間は「増える区間」
      if(ev[i].resId===newId||ev[i+1].resId===newId) pvNew.push(seg);
      else pvKeep.push(seg);
    }
  }
  // 需要の放物線アーク：同一方向OD（o→d）ごとに件数を集計し、太さ＝件数。
  // 行き帰りで重ならないよう、進行方向の左側に一定割合だけ膨らませる二次ベジェ。
  const demandArcs=[];
  let stopVol=null,maxVol=1;
  if(demandViz&&demandViz.length){
    const agg={};
    stopVol=Array(POS.length).fill(0);
    for(const r of demandViz){
      if(r.err||r.o==null||r.d==null||r.o===r.d)continue;
      const k=`${r.o}-${r.d}`;
      agg[k]=agg[k]||{o:r.o,d:r.d,n:0,pax:0};
      agg[k].n++; agg[k].pax+=(r.pax??1);
      if(r.o<stopVol.length)stopVol[r.o]++;
      if(r.d<stopVol.length)stopVol[r.d]++;
    }
    maxVol=Math.max(1,...stopVol);
    const maxN=Math.max(1,...Object.values(agg).map(a=>a.n));
    for(const a of Object.values(agg)){
      const p0=POS[a.o],p1=POS[a.d];
      const dx=p1[0]-p0[0],dy=p1[1]-p0[1],len=Math.hypot(dx,dy)||1;
      const nx=-dy/len,ny=dx/len;                       // 進行方向左の法線
      const bow=Math.min(70,len*0.22);                  // 膨らみ量
      const cx=(p0[0]+p1[0])/2+nx*bow, cy=(p0[1]+p1[1])/2+ny*bow; // 制御点
      const w=1.5+(a.n/maxN)*9;                          // 太さ1.5〜10.5
      demandArcs.push({d:`M ${p0[0]} ${p0[1]} Q ${cx} ${cy} ${p1[0]} ${p1[1]}`,
        w,n:a.n,mid:[(p0[0]+cx+p1[0])/2/1,(p0[1]+cy+p1[1])/2],
        lx:0.25*p0[0]+0.5*cx+0.25*p1[0],ly:0.25*p0[1]+0.5*cy+0.25*p1[1],end:p1});
    }
  }
  // コリドー軸：buildCorridorで確定済みの順序付き停留所列（prop経由）を描画する。
  // 経路は中点通過の2次曲線でつなぎ、角を落として実路線らしい滑らかな線にする
  // （中間停留所の真上は通らないが、停留所は円マーカーで別途示すため位置は読める）。
  let corridorPath=null, corridorPts=[];
  if(corridor&&corridor.stops&&corridor.stops.length>=2){
    corridorPts=corridor.stops;
    const pts=corridorPts.map(s=>POS[s]);
    if(pts.length===2){
      corridorPath=`M ${pts[0][0]} ${pts[0][1]} L ${pts[1][0]} ${pts[1][1]}`;
    }else{
      let d=`M ${pts[0][0]} ${pts[0][1]} L ${(pts[0][0]+pts[1][0])/2} ${(pts[0][1]+pts[1][1])/2}`;
      for(let i=1;i<pts.length-1;i++){
        const mx=(pts[i][0]+pts[i+1][0])/2, my=(pts[i][1]+pts[i+1][1])/2;
        d+=` Q ${pts[i][0]} ${pts[i][1]} ${mx} ${my}`;
      }
      d+=` L ${pts[pts.length-1][0]} ${pts[pts.length-1][1]}`;
      corridorPath=d;
    }
  }
  // 停留所ごとの乗降マーカー：終日分を常時表示すると密な停留所で重なって読めなくなるため、
  // 再生カーソル（simTime）を中心とした時間窓のみ表示する（運行盤＝今何が起きているかを見る画面という位置づけに合わせる）。
  // 窓内でも同一停留所に複数件が重なる場合があるため、停留所ごとに出現順で縦にスタックし、
  // 上限を超えた分は「+N件」のみ表示して増殖による重なりを構造的に防ぐ。
  const MARK_WINDOW=45*60, MARK_MAX=3;
  const marksRaw=[];
  if(showMarks){
    for(const v of vehicles){
      const sim=sims[v.id]; if(!sim.ok)continue;
      sim.events.forEach((e,i)=>{
        if(e.type==="ANCHOR")return;
        const t=e.type==="O"?e.apt:e.adt;
        if(simTime!=null&&Math.abs(t-simTime)>MARK_WINDOW)return;
        marksRaw.push({stop:e.stop,type:e.type,color:v.color,t,seq:i+1,vid:v.id});
      });
    }
  }
  marksRaw.sort((a,b)=>a.t-b.t);
  const stopCount={}, marks=[], overflow={};
  for(const m of marksRaw){
    const c=stopCount[m.stop]||0;
    if(c<MARK_MAX){marks.push({...m,row:c}); stopCount[m.stop]=c+1;}
    else overflow[m.stop]=(overflow[m.stop]||0)+1;
  }

  return (
  <div style={{position:"relative",width:"100%",height:"100%"}}>
  <svg ref={svgRef} viewBox={`${view.x} ${view.y} ${view.w} ${view.h}`}
    onMouseDown={onMouseDown}
    style={{width:"100%",height:"100%",cursor:view.w<VB_W?"grab":"default",touchAction:"none"}}>
    <defs>
      <pattern id="grid" width="45" height="45" patternUnits="userSpaceOnUse">
        <path d="M45 0H0V45" fill="none" stroke="#1C2A3D" strokeWidth="1"/>
      </pattern>
      <marker id="demArrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="9" markerHeight="9"
        markerUnits="userSpaceOnUse" orient="auto-start-reverse">
        <path d="M1 2 L8 5 L1 8 z" fill="#E8607D"/>
      </marker>
    </defs>
    <rect width="900" height="640" fill="url(#grid)"/>

    {/* コリドー軸：需要カバー×路線コストで構築した走行軸（buildCorridor）。滑らかな曲線＋停留所マーカー */}
    {corridorPath&&<g>
      <path d={corridorPath} fill="none" stroke="#E8A24D" strokeWidth="10" strokeLinecap="round"
        strokeLinejoin="round" opacity="0.42"/>
      <path d={corridorPath} fill="none" stroke="#E8A24D" strokeWidth="2.5" strokeLinecap="round"
        strokeLinejoin="round" opacity="0.85" strokeDasharray="1 9"/>
      {corridorPts.map((s,i)=>(
        <circle key={"cr"+s} cx={POS[s][0]} cy={POS[s][1]} r={i===0||i===corridorPts.length-1?6:4}
          fill="#E8A24D" stroke="#3A2A12" strokeWidth="1" opacity="0.9"/>
      ))}
      {corridorPts.length>=2&&[corridorPts[0],corridorPts[corridorPts.length-1]].map((s,i)=>(
        <text key={"crl"+s} x={POS[s][0]} y={POS[s][1]-9} textAnchor="middle" fontSize="10"
          fill="#F2C079" fontWeight="700">{STOPS[s]}</text>
      ))}
    </g>}

    {/* 車両ルート */}
    {lines.map(l=>{
      const dx=l.b[0]-l.a[0],dy=l.b[1]-l.a[1],len=Math.hypot(dx,dy)||1;
      const nx=-dy/len*l.off,ny=dx/len*l.off;
      const mx=(l.a[0]+l.b[0])/2+nx,my=(l.a[1]+l.b[1])/2+ny;
      return (<g key={l.key}>
        <line x1={l.a[0]+nx} y1={l.a[1]+ny} x2={l.b[0]+nx} y2={l.b[1]+ny}
          stroke={l.color} strokeWidth="1.6" strokeLinecap="round" opacity="0.30"/>
        <circle cx={mx} cy={my} r="7.5" fill="#14202F" stroke={l.color} strokeWidth="1.2" opacity="0.55"/>
        <text x={mx} y={my+3} textAnchor="middle" fontSize="9" fill={l.color} fontWeight="700" opacity="0.75">{l.seq}</text>
      </g>);
    })}

    {/* 走行履歴（残像）：太さ＝乗車人数、古いほど薄い */}
    {trails.map((t,i)=>(
      <line key={"tr"+i} x1={t.x1} y1={t.y1} x2={t.x2} y2={t.y2}
        stroke={t.color} strokeWidth={t.w} strokeLinecap="round"
        strokeDasharray={t.empty?"3 5":undefined} opacity={t.op}/>
    ))}

    {/* 候補プレビュー：既存ルート（淡い実線）＋新規で増える区間（明るい点線） */}
    {pvKeep.map((s,i)=>(
      <line key={`pk${i}`} x1={s[0][0]} y1={s[0][1]} x2={s[1][0]} y2={s[1][1]}
        stroke="#E0A93E" strokeWidth="1.6" opacity="0.30"/>
    ))}
    {pvNew.map((s,i)=>(
      <line key={`pn${i}`} x1={s[0][0]} y1={s[0][1]} x2={s[1][0]} y2={s[1][1]}
        stroke="#E0A93E" strokeWidth="3" strokeDasharray="7 6"/>
    ))}
    {previewRes&&preview&&<>
      <circle cx={POS[previewRes.o][0]} cy={POS[previewRes.o][1]} r="16" fill="none" stroke="#E0A93E" strokeWidth="2" strokeDasharray="4 3"/>
      <circle cx={POS[previewRes.d][0]} cy={POS[previewRes.d][1]} r="16" fill="none" stroke="#E0A93E" strokeWidth="2" strokeDasharray="4 3"/>
    </>}

    {/* 需要のノード量：発着で触れた回数に応じて停留所を膨らませる（拠点集中の可視化） */}
    {stopVol&&stopVol.map((v,i)=> v>0 ? (
      <circle key={"vol"+i} cx={POS[i][0]} cy={POS[i][1]} r={9+(v/maxVol)*24}
        fill="#E8607D" opacity={0.08+0.20*(v/maxVol)}/>
    ):null)}

    {/* 需要の放物線（同一ODの件数で太さ変化） */}
    {demandArcs.map((a,i)=>(
      <g key={"dm"+i}>
        <path d={a.d} fill="none" stroke="#E8607D" strokeWidth={a.w} strokeLinecap="round"
          opacity={0.18+0.32*Math.min(1,a.n/4)} markerEnd="url(#demArrow)"/>
        {a.n>=2&&<text x={a.lx} y={a.ly} textAnchor="middle" fontSize="10" fontWeight="700"
          fill="#FFD0DC" stroke="#7A2438" strokeWidth="0.5" paintOrder="stroke">{a.n}</text>}
      </g>
    ))}

    {/* 停留所 */}
    {STOPS.map((s,i)=>(
      <g key={s}>
        <circle cx={POS[i][0]} cy={POS[i][1]} r="8" fill="#0E1622" stroke="#5E7290" strokeWidth="1.6"/>
        <circle cx={POS[i][0]} cy={POS[i][1]} r="2.6" fill="#9FB2CC"/>
        <text x={POS[i][0]} y={POS[i][1]-13} textAnchor="middle" fontSize="9.5" fill="#C9D4E4"
          fontFamily="'SF Mono','Consolas',monospace" fontWeight="600">{s}</text>
      </g>
    ))}

    {/* 乗降時刻ラベル（simTime±45分の窓内・停留所あたり最大3件、超過分は+N件） */}
    {marks.map((m,i)=>(
      <text key={i} x={POS[m.stop][0]+12} y={POS[m.stop][1]+5+m.row*12} fontSize="9"
        fill={m.color} fontFamily="'SF Mono','Consolas',monospace">
        {m.type==="O"?"▲乗":"▼降"} {fmt(m.t)}
      </text>
    ))}
    {Object.entries(overflow).map(([stop,n])=>(
      <text key={"ov"+stop} x={POS[+stop][0]+12} y={POS[+stop][1]+5+MARK_MAX*12} fontSize="8.5"
        fill="#8A93A0" fontFamily="'SF Mono','Consolas',monospace">+{n}件</text>
    ))}

    {/* 再生：時刻simTimeの車両位置（円内の数字＝現在の乗車人数） */}
    {vehicles.map(v=>{
      const pos=vehiclePosAt(sims[v.id],simTime);
      if(!pos)return null;
      return (<g key={"veh"+v.id}>
        {pos.dwell&&<circle cx={pos.x} cy={pos.y} r="17" fill="none"
          stroke={v.color} strokeWidth="2" opacity="0.45"/>}
        <circle cx={pos.x} cy={pos.y} r="12" fill={v.color} stroke="#0E1622" strokeWidth="2.5"/>
        <text x={pos.x} y={pos.y+4.5} textAnchor="middle" fontSize="12" fill="#0E1622"
          fontWeight="800" fontFamily="'SF Mono','Consolas',monospace">{pos.load}</text>
        <text x={pos.x} y={pos.y-17} textAnchor="middle" fontSize="9" fill={v.color}
          fontWeight="700">{v.name}</text>
      </g>);
    })}
  </svg>
  {/* ズーム操作：＋/−/全体。ホイールでカーソル中心の拡縮、ドラッグでパンも可能 */}
  <div style={{position:"absolute",left:12,bottom:44,display:"flex",flexDirection:"column",gap:4}}>
    <button style={zoomBtn} title="拡大" onClick={()=>zoomAt(view.x+view.w/2,view.y+view.h/2,0.8)}>＋</button>
    <button style={zoomBtn} title="縮小" onClick={()=>zoomAt(view.x+view.w/2,view.y+view.h/2,1.25)}>−</button>
    <button style={{...zoomBtn,fontSize:10}} title="全体表示"
      onClick={()=>setView({x:0,y:0,w:VB_W,h:VB_H})}>全体</button>
  </div>
  {view.w<VB_W&&<div style={{position:"absolute",left:48,bottom:44,fontSize:10,color:"#9AA7BA",
    background:"rgba(20,32,47,0.7)",borderRadius:5,padding:"2px 6px"}}>
    ×{(VB_W/view.w).toFixed(1)}・ドラッグで移動
  </div>}
  </div>);
}

/* ---------- タイムライン（運行ダイヤ＋乗車区間） ----------
   各車両の行に、1運行（最初の乗車〜乗客が途切れる降車）を薄緑ブロック、
   利用者1人の乗車区間を1本の帯として重ねて描く。帯が縦に重なった分だけ
   乗り合いが発生していることが直接読める（占有レーン数＝その瞬間の乗車人数）。
   白丸＝乗車・塗り丸＝降車。定員は点線で示す。空車の移動・待機は基線上の細線。 */
function Timeline({vehicles,sims,resMap,P,simTime,preview,previewRes}){
  const T0=8.5*3600,T1=17.5*3600,W=860;
  const ROW=66, BASE=52;
  const H=24+vehicles.length*ROW+18;
  const x=t=>20+(t-T0)/(T1-T0)*W;
  // 拡大表示：倍率>1でSVG幅を広げ、コンテナ内スクロールで閲覧（縦横等倍＝文字も拡大）
  const [tz,setTz]=React.useState(1);
  const tzBtn=on=>({fontSize:10,padding:"2px 7px",cursor:"pointer",borderRadius:5,
    border:"1px solid #C9C4B8",background:on?"#14202F":"#fff",color:on?"#fff":"#6B6453",fontWeight:700});
  return (
  <div style={{position:"relative",width:"100%",height:"100%"}}>
  <div style={{position:"absolute",right:0,top:-4,display:"flex",gap:4,zIndex:2}}>
    {[1,1.5,2,3].map(z=>(
      <button key={z} style={tzBtn(tz===z)} onClick={()=>setTz(z)}>×{z}</button>
    ))}
  </div>
  <div style={{width:"100%",height:"100%",overflow:tz>1?"auto":"hidden"}}>
  <svg viewBox={`0 0 900 ${H}`}
    style={tz===1?{width:"100%",height:"100%"}:{width:`${tz*100}%`,height:"auto",display:"block"}}>
    {Array.from({length:10},(_,i)=>{
      const t=9*3600+i*3600;
      return (<g key={i}>
        <line x1={x(t)} y1={14} x2={x(t)} y2={H-12} stroke="#DDD8CB" strokeWidth="1"/>
        <text x={x(t)} y={11} textAnchor="middle" fontSize="10" fill="#8A8474"
          fontFamily="'SF Mono','Consolas',monospace">{fmt(t)}</text>
      </g>);
    })}
    {vehicles.map((v,vi)=>{
      const y=BASE+vi*ROW;                    // 基線（人数0）
      const unit=Math.min(3.2,26/v.cap);      // 1人あたりの高さ(px)
      const sim=sims[v.id];
      // 乗車人数の時間区分（イベント時刻→次イベント時刻、その間のload）
      const segs=[];
      let paxSec=0;
      if(sim.ok&&sim.events.length>0){
        const tm=ev=>ev.type==="O"?ev.apt:(ev.type==="D"?ev.adt:ev.apt); // アンカーはapt(=杭時刻)
        for(let i=0;i<sim.events.length;i++){
          const e=sim.events[i];
          const t0=tm(e);
          const t1=i+1<sim.events.length?tm(sim.events[i+1]):e.etd;
          if(t1>t0){segs.push({t0,t1,load:e.load});paxSec+=e.load*(t1-t0);}
        }
      }
      const span=sim.ok&&sim.events.length>0
        ?sim.events[sim.events.length-1].etd-sim.events[0].eta:0;
      const avg=span>0?paxSec/span:0;
      // 利用者個別の乗車区間（1人＝1本の帯。pax人の予約はpax本）。
      // レーン割付＝時刻順に走査し、空いている最下段レーンへ詰める（区間グラフ彩色）。
      // 同時刻に占有されるレーン数＝その瞬間の乗車人数なので、帯の重なりが乗り合いを直接表す。
      const rides=[];
      if(sim.ok){
        const om={},dm={};
        for(const e of sim.events){
          if(e.type==="O")om[e.resId]=e; else if(e.type==="D")dm[e.resId]=e;
        }
        for(const id in om){
          const o=om[id],d=dm[id]; if(!d)continue;
          const px=(resMap[id]&&resMap[id].pax)||1;
          for(let k=0;k<px;k++)rides.push({t0:o.apt,t1:d.adt,id});
        }
        rides.sort((a,b)=>a.t0-b.t0||a.t1-b.t1);
        const laneEnd=[];
        for(const r of rides){
          let ln=laneEnd.findIndex(t=>t<=r.t0+1e-6);
          if(ln<0){ln=laneEnd.length;laneEnd.push(0);}
          laneEnd[ln]=r.t1; r.lane=ln;
        }
      }
      // 1運行＝乗車区間の連結成分（最初の乗車〜乗客が途切れる降車まで）を薄緑ブロックで示す
      const blocks=[];
      {
        const sorted=rides.slice().sort((a,b)=>a.t0-b.t0);
        for(const r of sorted){
          const last=blocks[blocks.length-1];
          if(last&&r.t0<=last.t1+1e-6){last.t1=Math.max(last.t1,r.t1);}
          else blocks.push({t0:r.t0,t1:r.t1,maxLane:0});
        }
        for(const b of blocks)for(const r of rides)
          if(r.t0<b.t1&&r.t1>b.t0)b.maxLane=Math.max(b.maxLane,r.lane);
      }
      return (<g key={v.id}>
        <text x={20} y={y-34} fontSize="11" fill="#444" fontWeight="700">
          {v.name}（定員{v.cap}{v.mode==="semi"?"・セミ":""}）
        </text>
        {/* 運行時間帯（基線） */}
        <line x1={x(v.start)} y1={y} x2={x(v.end)} y2={y} stroke="#E5E1D5" strokeWidth="4" strokeLinecap="round"/>
        <line x1={x(v.start)} y1={y-4} x2={x(v.start)} y2={y+4} stroke="#B9B2A1" strokeWidth="2"/>
        <line x1={x(v.end)} y1={y-4} x2={x(v.end)} y2={y+4} stroke="#B9B2A1" strokeWidth="2"/>
        {/* 定員の参照線 */}
        {span>0&&<line x1={x(v.start)} y1={y-v.cap*unit} x2={x(v.end)} y2={y-v.cap*unit}
          stroke={v.color} strokeWidth="1" strokeDasharray="3 4" opacity="0.55"/>}
        {span>0&&<text x={x(v.start)-4} y={y-v.cap*unit+3} textAnchor="end" fontSize="8"
          fill="#8A8474">定員</text>}
        {/* セミ車：背骨の運行枠（初便始発〜終便終着）を基線上のバンドで示す。
            乗車帯・約束・乗降丸はフル車と完全に同一の導出（O/Dイベント由来）であり、
            セミの違いは「この枠が先に決まっていること」だけ、という構造を表示にも一致させる */}
        {v.mode==="semi"&&sim.ok&&(()=>{
          const as=sim.events.filter(e=>e.type==="ANCHOR");
          if(!as.length)return null;
          const t0=as[0].apt, t1=as[as.length-1].apt;
          return (<g>
            <rect x={x(t0)} y={y-2.5} width={Math.max(x(t1)-x(t0),2)} height={5}
              rx="2" fill="#577CA8" opacity="0.30"/>
            <text x={x(t0)-4} y={y+3} textAnchor="end" fontSize="8" fill="#577CA8"
              fontWeight="700">背骨</text>
          </g>);
        })()}
        {/* 空車で動いている区間（回送・待機）は基線上の細線で示す */}
        {segs.map((s,i)=>s.load===0?(
          <line key={i} x1={x(s.t0)} y1={y} x2={x(s.t1)} y2={y}
            stroke={v.color} strokeWidth="4" strokeLinecap="butt" opacity="0.45"/>
        ):null)}
        {/* 1運行＝薄緑ブロック（最初の乗車〜乗客が途切れる降車まで） */}
        {blocks.map((b,i)=>(
          <rect key={"b"+i} x={x(b.t0)-2} y={y-(b.maxLane+1)*unit-3}
            width={Math.max(x(b.t1)-x(b.t0)+4,3)} height={(b.maxLane+1)*unit+6}
            rx="2" fill="#CFE3A8" opacity="0.8"/>
        ))}
        {/* 利用者1人＝1本の帯（乗車〜降車）。帯の縦の重なり＝乗り合い */}
        {rides.map((r,i)=>(
          <rect key={"rd"+i} x={x(r.t0)} y={y-(r.lane+1)*unit+0.4}
            width={Math.max(x(r.t1)-x(r.t0),1.5)} height={Math.max(unit-0.8,1.2)}
            fill="#2E9E8A" opacity="0.92"/>
        ))}
        {/* 乗降イベント（アンカーは時刻の杭として縦線で表示） */}
        {sim.ok&&sim.events.map((e,i)=>e.type==="ANCHOR"?(
          <line key={"e"+i} x1={x(e.apt)} y1={y-6} x2={x(e.apt)} y2={y+6}
            stroke="#577CA8" strokeWidth="1.5" opacity="0.85"/>
        ):(
          <circle key={"e"+i} cx={x(e.type==="O"?e.apt:e.adt)} cy={y} r="4"
            fill={e.type==="O"?"#fff":v.color} stroke={v.color} strokeWidth="2"/>
        ))}
        {/* 候補プレビュー：この号車に新規予約が入る場合の挿入区間（オレンジ） */}
        {preview&&previewRes&&v.id===preview.vehicle&&(()=>{
          const me=preview.sim.events.filter(e=>e.resId===previewRes.id);
          const o=me.find(e=>e.type==="O"), d=me.find(e=>e.type==="D");
          if(!o||!d)return null;
          const xo=x(o.apt), xd=x(d.adt);
          return (<g>
            <line x1={xo} y1={y} x2={xd} y2={y} stroke="#E0A93E" strokeWidth="4.5"
              strokeDasharray="7 6" strokeLinecap="round"/>
            <circle cx={xo} cy={y} r="6" fill="#fff" stroke="#E0A93E" strokeWidth="2.5" strokeDasharray="3 2"/>
            <circle cx={xd} cy={y} r="6" fill="#E0A93E" stroke="#E0A93E" strokeWidth="2.5"/>
            <text x={xo} y={y+13} textAnchor="middle" fontSize="8.5" fontWeight="700"
              fill="#B8791F" fontFamily="'SF Mono','Consolas',monospace">新規乗</text>
            <text x={xd} y={y-10} textAnchor="middle" fontSize="8.5" fontWeight="700"
              fill="#B8791F" fontFamily="'SF Mono','Consolas',monospace">新規降</text>
          </g>);
        })()}
      </g>);
    })}
    {/* 再生時刻カーソル */}
    {simTime!=null&&simTime>=T0&&simTime<=T1&&<g>
      <line x1={x(simTime)} y1={14} x2={x(simTime)} y2={H-12} stroke="#E0533F" strokeWidth="1.6"/>
      <polygon points={`${x(simTime)-5},14 ${x(simTime)+5},14 ${x(simTime)},21`} fill="#E0533F"/>
    </g>}
    <g fontSize="9" fill="#8A8474">
      <rect x={560} y={H-14} width={18} height={9} rx="2" fill="#CFE3A8"/>
      <text x={582} y={H-6}>1運行</text>
      <rect x={624} y={H-11} width={18} height={3.2} fill="#2E9E8A"/>
      <text x={646} y={H-6}>1人の乗車（縦の重なり＝乗り合い）</text>
      <line x1={838} y1={H-9} x2={856} y2={H-9} stroke="#999" strokeWidth="1" strokeDasharray="3 3"/>
      <text x={860} y={H-6}>定員</text>
      <line x1={20} y1={H-9} x2={40} y2={H-9} stroke="#E0A93E" strokeWidth="3" strokeDasharray="5 4"/>
      <text x={44} y={H-6} fill="#B8791F">新規挿入区間（候補選択時）</text>
    </g>
  </svg>
  </div>
  </div>);
}

/* ---------- 運行確認表（Excel検証表の再現） ----------
   行構成：予約番号 | O/D | 乗降場 | Start | ETA | ETD | End |
           区間DRT 秒/分 | STn 秒/分 | DRTn+STn
   - O行: Start=約束発(赤), End=約束発+TW(赤)
   - D行: Start=計画着, End=約束着上限(赤)
   - 区間列は「この行の停留所→次の停留所」の移動に対する値
   配色：新規予約の行=緑 / 前回から時刻が動いた行=赤字 /
         置換された旧行（before側）=グレー網掛け            */

function buildRows(events, resMap, P, dateStr){
  return events.filter(e=>e.type!=="ANCHOR").map((e,i,arr)=>{
    const r=resMap[e.resId];
    const next=arr[i+1];
    const leg=next?TT[e.stop][next.stop]:null;
    const isO=e.type==="O";
    const dmode=r?.mode??"dep";
    // 利用者が指定した希望時刻。発希望は希望乗車をO行に、着希望は希望降車をD行に表示。
    const desired = !r ? null
      : isO ? (dmode==="dep" ? r.dpt : null)
            : (dmode==="arr" ? r.dpt+r.mrt+P.dwell : null);
    return {
      key:`${e.resId}-${e.type}`,
      date:dateStr, num:r?r.num:"—", od:e.type, stop:STOPS[e.stop],
      desired, dmode,
      // 利用者との約束
      pFrom:isO?(r.ipt??e.apt):null,                  // 約束発（窓の始まり）
      pTo:isO?(r.ipt??e.apt)+P.tw:(r.idt??e.adt),     // 発の最遅／着の上限
      // 最新の運行計画
      eta:e.eta,
      board:isO?e.apt:e.adt,                          // 乗車／降車時刻
      etd:e.etd,
      // 約束に対する余裕（O:最遅発−乗車、D:着上限−降車）
      slack:isO?((r.ipt??e.apt)+P.tw)-e.apt:(r.idt??e.adt)-e.adt,
      legSec:leg, legSt:leg!=null?stFromDRT(leg,P):null, legSum:leg!=null?mrtFromDRT(leg,P):null,
      resId:e.resId,
    };
  });
}

const cTh={padding:"3px 5px",borderBottom:"2px solid #1E2A38",fontSize:10,whiteSpace:"nowrap",textAlign:"center"};
const cTd={padding:"2px 5px",borderBottom:"1px solid #DDD8CB",fontSize:11,whiteSpace:"nowrap",
  fontFamily:"'SF Mono','Consolas',monospace",fontVariantNumeric:"tabular-nums",textAlign:"center"};
const PROMISE_BG="#EAF4EC";  // 約束列の地色（緑系）
const PLAN_BG="#EDF1F7";     // 計画列の地色（青系）

function Tbl({rows,highlight}){
  // highlight: {green:Set(resId), redRows:Set(key), grayRows:Set(key)}
  const gTh=g=>({...cTh,background:g==="p"?PROMISE_BG:g==="l"?PLAN_BG:"#F0EDE4"});
  return (
  <div style={{overflowX:"auto",maxWidth:"100%"}}>
  <table style={{borderCollapse:"collapse",background:"#fff",border:"1px solid #B9B2A1"}}>
    <thead>
      <tr>
        <th style={{...cTh,background:"#F0EDE4",borderBottom:"1px solid #C9C4B8"}} colSpan={3}></th>
        <th style={{...cTh,background:"#FBF4E6",borderBottom:"1px solid #C9C4B8",color:"#8A6D1F"}}>利用者の希望</th>
        <th style={{...cTh,background:PROMISE_BG,borderBottom:"1px solid #C9C4B8",color:"#1E6B40"}} colSpan={2}>利用者との約束（確定後 不変）</th>
        <th style={{...cTh,background:PLAN_BG,borderBottom:"1px solid #C9C4B8",color:"#2C4A77"}} colSpan={3}>最新の運行計画</th>
        <th style={{...cTh,background:"#F0EDE4",borderBottom:"1px solid #C9C4B8"}}></th>
        <th style={{...cTh,background:"#F0EDE4",borderBottom:"1px solid #C9C4B8"}} colSpan={5}>区間（この行→次の行）</th>
      </tr>
      <tr>
        <th style={gTh()}>予約番号</th><th style={gTh()}></th><th style={gTh()}>乗降場</th>
        <th style={{...cTh,background:"#FBF4E6",color:"#8A6D1F"}}>希望乗降</th>
        <th style={gTh("p")}>約束発（窓）</th><th style={gTh("p")}>約束着（上限）</th>
        <th style={gTh("l")}>ETA</th><th style={gTh("l")}>乗降</th><th style={gTh("l")}>ETD</th>
        <th style={gTh()}>約束まで余裕</th>
        <th style={gTh()}>DRTn 秒</th><th style={gTh()}>分</th>
        <th style={gTh()}>STn 秒</th><th style={gTh()}>分</th><th style={gTh()}>DRTn+STn</th>
      </tr>
    </thead>
    <tbody>
      {rows.map(r=>{
        const green=highlight?.green?.has(r.resId);
        const red=highlight?.redRows?.has(r.key);
        const gray=highlight?.grayRows?.has(r.key);
        const timeColor=red?"#C0392B":undefined;
        const nameColor=green?"#1E8449":red?"#C0392B":undefined;
        const slackBad=r.slack<0;
        return (
        <tr key={r.key} style={{background:gray?"#D8D4CA":undefined}}>
          <td style={cTd}>{r.num}</td>
          <td style={cTd}>{r.od}</td>
          <td style={{...cTd,fontWeight:700,color:nameColor}}>{r.stop}</td>
          <td style={{...cTd,background:gray?undefined:"#FBF4E6",fontWeight:700,color:"#8A6D1F"}}>
            {r.desired!=null?fmt(r.desired):"—"}</td>
          <td style={{...cTd,background:gray?undefined:PROMISE_BG,fontWeight:700,color:"#1E6B40"}}>
            {r.pFrom!=null?`${fmt(r.pFrom)}〜${fmt(r.pTo)}`:"—"}</td>
          <td style={{...cTd,background:gray?undefined:PROMISE_BG,fontWeight:700,color:"#1E6B40"}}>
            {r.od==="D"?`〜${fmt(r.pTo)}`:"—"}</td>
          <td style={{...cTd,background:gray?undefined:PLAN_BG,color:timeColor}}>{fmt(r.eta)}</td>
          <td style={{...cTd,background:gray?undefined:PLAN_BG,fontWeight:700,color:timeColor??"#2C4A77"}}>{fmt(r.board)}</td>
          <td style={{...cTd,background:gray?undefined:PLAN_BG,color:timeColor}}>{fmt(r.etd)}</td>
          <td style={{...cTd,fontWeight:700,color:slackBad?"#C0392B":"#666"}}>
            {slackBad?"超過":""}{fmtMin(Math.abs(r.slack))}</td>
          <td style={cTd}>{r.legSec!=null?Math.round(r.legSec):""}</td>
          <td style={cTd}>{r.legSec!=null?fmtHMS(r.legSec):""}</td>
          <td style={cTd}>{r.legSt!=null?(Math.round(r.legSt*10)/10):""}</td>
          <td style={cTd}>{r.legSt!=null?fmtHMS(r.legSt):""}</td>
          <td style={{...cTd,fontWeight:700}}>{r.legSum!=null?fmtHMS(r.legSum):""}</td>
        </tr>);
      })}
      {rows.length===0&&<tr><td style={cTd} colSpan={15}>イベントなし</td></tr>}
    </tbody>
  </table>
  </div>);
}

function VerifyModal({onClose,vehicles,sims,resMap,P,lastChange,initialTab,lastDemand}){
  const today=new Date();
  const dateStr=`${today.getFullYear()}.${String(today.getMonth()+1).padStart(2,"0")}.${String(today.getDate()).padStart(2,"0")}`;
  const [tab,setTab]=React.useState(initialTab??(lastChange?"diff":"now"));

  // 前後比較の差分判定
  let diff=null;
  if(lastChange){
    const beforeRows=buildRows(lastChange.before,lastChange.resMap,lastChange.P,dateStr);
    const afterRows=buildRows(lastChange.after,lastChange.resMap,lastChange.P,dateStr);
    const bMap=Object.fromEntries(beforeRows.map(r=>[r.key,r]));
    const green=new Set([lastChange.newResId]);
    const redRows=new Set(), grayRows=new Set();
    for(const a of afterRows){
      const b=bMap[a.key];
      if(b&&(Math.abs(b.eta-a.eta)>1||Math.abs(b.etd-a.etd)>1||Math.abs(b.board-a.board)>1)){
        redRows.add(a.key);grayRows.add(a.key);
      }
    }
    const vName=vehicles.find(v=>v.id===lastChange.vehicle).name;
    const newRes=lastChange.resMap[lastChange.newResId];
    diff={beforeRows,afterRows,green,redRows,grayRows,vName,newRes};
  }

  return (
  <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(10,16,26,0.55)",
    display:"flex",alignItems:"center",justifyContent:"center",zIndex:50}}>
    <div onClick={e=>e.stopPropagation()} style={{background:"#F5F3EE",borderRadius:12,
      width:"min(960px,94vw)",maxHeight:"88vh",overflowY:"auto",overflowX:"hidden",padding:18,
      boxShadow:"0 18px 60px rgba(0,0,0,0.45)"}}>
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
        <div style={{fontSize:16,fontWeight:700}}>運行確認表</div>
        <span style={{fontSize:11,color:"#8A8474",fontFamily:"'SF Mono','Consolas',monospace"}}>{dateStr}</span>
        <div style={{display:"flex",gap:6}}>
          <Toggle on={tab==="diff"} onClick={()=>setTab("diff")}>直前確定の前後比較</Toggle>
          <Toggle on={tab==="now"} onClick={()=>setTab("now")}>全車両の現況</Toggle>
          <Toggle on={tab==="ana"} onClick={()=>setTab("ana")}>運行分析</Toggle>
          <Toggle on={tab==="tt"} onClick={()=>setTab("tt")}>OD直行時間表</Toggle>
          <Toggle on={tab==="defs"} onClick={()=>setTab("defs")}>変数定義（表1）</Toggle>
        </div>
        <button onClick={onClose} style={{marginLeft:"auto",border:"1px solid #C9C4B8",background:"#fff",
          borderRadius:6,padding:"4px 12px",cursor:"pointer",fontSize:13}}>閉じる</button>
      </div>

      {tab==="diff"&&(diff?(
        <div>
          <div style={{fontSize:12,color:"#6B6453",marginBottom:8}}>
            {diff.vName}：予約 <b style={{color:"#1E8449"}}>#{diff.newRes.num}</b>
            （{STOPS[diff.newRes.o]}→{STOPS[diff.newRes.d]}）確定による変化。
            <span style={{color:"#1E8449",fontWeight:700}}> 緑＝新規予約</span>／
            <span style={{color:"#C0392B",fontWeight:700}}>赤＝計画時刻が変動した行</span>
            （変動行は確定前の表でグレー網掛け）。
            緑地の「約束」列は利用者に伝えた時刻で確定後は動かない。青地の「運行計画」列は乗合の組み替えで動く。
            「約束まで余裕」が0以上であれば約束は守られている。
          </div>
          <div style={{fontSize:11,fontWeight:700,color:"#6B6453",margin:"6px 0 4px"}}>確定前</div>
          <Tbl rows={diff.beforeRows} highlight={{grayRows:diff.grayRows}}/>
          <div style={{textAlign:"center",fontSize:18,color:"#6B6453",padding:"4px 0"}}>↓ ↓ ↓</div>
          <div style={{fontSize:11,fontWeight:700,color:"#6B6453",margin:"0 0 4px"}}>確定後</div>
          <Tbl rows={diff.afterRows} highlight={{green:diff.green,redRows:diff.redRows}}/>
          <div style={{fontSize:11,color:"#8A8474",marginTop:8}}>
            区間列（DRTn/STn/DRTn+STn）は各行の停留所から次の停留所までの移動に対する値。
            計画時刻が動いた行も「約束まで余裕」が負にならないことがエンジンの担保条件。
          </div>
        </div>
      ):(
        <div style={{fontSize:13,color:"#6B6453",padding:"18px 0"}}>
          この画面を開いてからの確定操作がまだない。新規予約を確定すると、確定直前と直後の運行がここで比較できる。
        </div>
      ))}

      {tab==="now"&&(
        <div>
          {vehicles.filter(v=>v.active).map(v=>{
            const sim=sims[v.id];
            const rows=sim.ok?buildRows(sim.events,resMap,P,dateStr):[];
            return (
            <div key={v.id} style={{marginBottom:16}}>
              <div style={{display:"flex",alignItems:"center",gap:8,margin:"4px 0 6px"}}>
                <Dot c={v.color}/><b style={{fontSize:13}}>{v.name}</b>
                <span style={{fontSize:11,color:"#8A8474"}}>
                  定員{v.cap}名・運行 {fmt(v.start)}〜{fmt(v.end)}
                  {sim.ok&&sim.events.length>0&&` ／ 拘束 ${fmt(sim.events[0].eta)}〜${fmt(sim.events[sim.events.length-1].etd)} ／ 車内最大 ${sim.maxLoad}名`}
                </span>
                {!sim.ok&&<span style={{fontSize:11,color:"#C0392B",fontWeight:700}}>実行不可：{sim.why}</span>}
              </div>
              <Tbl rows={rows}/>
            </div>);
          })}
        </div>
      )}

      {tab==="ana"&&<AnalysisTab vehicles={vehicles} sims={sims} resMap={resMap} P={P} lastDemand={lastDemand}/>}

      {tab==="tt"&&<TTMatrix P={P}/>}

      {tab==="defs"&&<VarDefs P={P}/>}
    </div>
  </div>);
}

/* ---------- 表1 問題を定義する変数（原典定義＋本画面での対応） ---------- */
const VAR_DEFS=[
 ["N","全乗客数を示す。","確定予約の件数。"],
 ["R","全乗客からなる運行計画情報を示す。","全車両の運行（マップのルート全体・現況タブの3表）。"],
 ["n","対象としている乗客 n を示す。","表の「予約番号」1件＝乗客n。"],
 ["Rn","乗客 n の運行計画情報を示す。","表中の同一予約番号のO行＋D行の組。"],
 ["DPTn","乗客 n の希望乗車時刻（the Desired pick-up time）を示す。","受付フォームの「発希望」。着希望入力時は DDTn−(MRTn＋Dwell) に換算。"],
 ["EPTn","乗客 n が許容できる乗車時刻の最早時間（the Earliest pick-up time）を示す。","予約時＝DPTn−San（希望より早い乗車も探索）。確定後＝IPTn（O行 Start。約束発より早い乗車は行わない）。"],
 ["LPTn","乗客 n が許容できる乗車時刻の最遅時間（the Latest pick-up time）を示す。","＝IPTn＋TW。O行 End（赤太字）。"],
 ["APTn","乗客 n の実際の乗車時刻（the Actual pick-up time）を示す。","O行 ETA（計画上の実乗車。EPTn≦APTn≦LPTnが成立条件）。"],
 ["DDTn","乗客 n の希望降車時刻（the Desired delivery time）を示す。","受付フォームの「着希望」。"],
 ["EDTn","乗客 n が許容できる降車時刻の最早時間（the Earliest delivery time）を示す。","本実装では下限制約を課していない（早着は許容）。"],
 ["LDTn","乗客 n が許容できる降車時刻の最遅時間（the Latest delivery time）を示す。","＝IDTn。D行 End（赤太字）。これを超える挿入は棄却。"],
 ["ADTn","乗客 n の実際の降車時刻（the Actual delivery time）を示す。","D行 ETA（計画上の実降車。ADTn≦LDTnが成立条件）。"],
 ["ETA","着く時刻（Estimated Time of Arrival）。","表のETA列。車両がその乗降場に着く時刻。O行では乗車処理の開始、D行では降車（＝ADTn）にあたる。"],
 ["ETD","出発時刻（Estimated Time of Departure）。","表のETD列。乗降（Dwell）を終えて車両がその乗降場を発つ時刻。ETD＝乗降開始時刻＋Dwell。"],
 ["+n / −n","乗客 n の乗車（あるいは降車）のイベントを表す。","表のO行（＋n）／D行（−n）。"],
 ["TT(x,y)","地点 x から地点 y までの移動時間（Travel time）を示す。","所要時間マトリクス。表の「DRTn 秒/分」列は当該行→次行の停留所間のTT。"],
 ["p(x)","イベント x が生じる地点を示す。例：p(+n)","表の「乗降場」列。"],
 ["Busn","乗客 n が降車するバスの番号を示す。","予約一覧の号車表示（マップ・ダイヤの色に対応）。"],
 ["DRTn","乗客 n が乗り合いなしで移動し、出発地から目的地まで直行したときの直行乗車時間（the Direct ride time）を示す。","TT(発地,着地)。受付フォームの「直行」表示。"],
 ["MRTn","乗客 n が許容する最大の乗車許容時間（the Maximum ride time）である。","＝DRTn＋STn。表の「DRTn+STn」列。約束着の算定基礎。"],
 ["STn","乗客 n が設定するゆとり時間（Slack Time）を示す。","＝寄り道時間。三段式（短距離は一律最小値→中距離は線形→長距離は一律最大値）でDRTnから算出。MRTn=DRTn+STn。表の「STn 秒/分」列。"],
 ["San","乗客 n が設定する希望とのずれ時間（解の探索範囲：Search Area）を示す。","受付フォームの「ズレ幅」。予約ごとに前後5〜120分で設定し、DPTn−San〜DPTn＋Sanの範囲で乗車時刻を探索。確定後の約束窓はTW（IPTn〜IPTn＋TW）に切り替わる。"],
 ["IPTn","乗客 n に約束した乗車時刻（the Informed pick-up time）を示す。","O行 Start（赤太字）。確定後は不変。"],
 ["IDTn","乗客 n に約束した降車時刻（the Informed delivery time）を示す。※原文は「乗車時刻」だが文脈上、降車の誤記と判断。","D行 End（赤太字）＝IPTn＋Dwell＋MRTn。確定後は不変。"],
 ["V","全車両数を示す。","ヘッダ「車両設定」で最大5台。台数（稼働ON/OFF）・定員・運行時間を車両ごとに設定。"],
];

function VarDefs({P}){
  const th={padding:"5px 10px",borderBottom:"2px solid #1E2A38",fontSize:11,textAlign:"left",whiteSpace:"nowrap"};
  const td={padding:"5px 10px",borderBottom:"1px solid #DDD8CB",fontSize:12.5,verticalAlign:"top",lineHeight:1.5};
  return (
  <div>
    <div style={{fontSize:12,color:"#6B6453",marginBottom:8}}>
      表1 問題を定義する変数（原典）と、この画面の各表示との対応。
      現在の設定：寄り道ST 短≤{Math.round(P.sl.d1/60)}分で{Math.round(P.sl.s1/60)}分・長≥{Math.round(P.sl.d2/60)}分で{Math.round(P.sl.s2/60)}分（間は線形）／TW {P.tw/60}分／Dwell {P.dwell}秒。
    </div>
    <table style={{borderCollapse:"collapse",background:"#fff",border:"1px solid #B9B2A1",width:"100%"}}>
      <thead><tr style={{background:"#F0EDE4"}}>
        <th style={th}>変数</th><th style={th}>説明（原典）</th><th style={th}>本画面での対応</th>
      </tr></thead>
      <tbody>
        {VAR_DEFS.map(([v,def,map])=>(
          <tr key={v}>
            <td style={{...td,fontFamily:"'SF Mono','Consolas',monospace",fontWeight:700,whiteSpace:"nowrap"}}>{v}</td>
            <td style={td}>{def}</td>
            <td style={{...td,color:"#23694A"}}>{map}</td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>);
}

/* ---------- OD直行時間表（TTマトリクス：秒／分の2表） ---------- */
function TTMatrix({P}){
  const [unit,setUnit]=React.useState("min"); // sec | min
  const th={padding:"3px 6px",border:"1px solid #C9C4B8",fontSize:10,background:"#F0EDE4",
    fontFamily:"'SF Mono','Consolas',monospace",whiteSpace:"nowrap",position:"sticky",top:0};
  const thRow={padding:"3px 6px",border:"1px solid #C9C4B8",fontSize:10,background:"#F0EDE4",
    fontFamily:"'SF Mono','Consolas',monospace",whiteSpace:"nowrap",position:"sticky",left:0,textAlign:"left"};
  const td={padding:"3px 6px",border:"1px solid #DDD8CB",fontSize:11,textAlign:"right",
    fontFamily:"'SF Mono','Consolas',monospace",fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap"};
  const val=(i,j)=>{
    if(i===j)return "0";
    const s=TT[i][j];
    return unit==="sec"?String(s):(Math.round(s/60*10)/10).toFixed(1);
  };
  return (
  <div>
    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
      <div style={{fontSize:12,color:"#6B6453"}}>
        OD間の直行所要時間 TT(x,y)・30停留所。DRTn・STn（三段式）・MRTnの算定基礎。
      </div>
      <div style={{display:"flex",gap:6,marginLeft:"auto"}}>
        <Toggle on={unit==="sec"} onClick={()=>setUnit("sec")}>秒</Toggle>
        <Toggle on={unit==="min"} onClick={()=>setUnit("min")}>分</Toggle>
      </div>
    </div>
    <div style={{overflow:"auto",maxHeight:"56vh",border:"1px solid #B9B2A1"}}>
      <table style={{borderCollapse:"collapse",background:"#fff"}}>
        <thead>
          <tr>
            <th style={{...thRow,zIndex:2}}>{unit==="sec"?"秒":"分"}</th>
            {STOPS.map(s=><th key={s} style={th}>{s.replace("BS0","")}</th>)}
          </tr>
        </thead>
        <tbody>
          {STOPS.map((s,i)=>(
            <tr key={s}>
              <th style={thRow}>{s}</th>
              {STOPS.map((_,j)=>(
                <td key={j} style={{...td,background:i===j?"#F0EDE4":undefined}}>{val(i,j)}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
    <div style={{fontSize:11,color:"#8A8474",marginTop:8,lineHeight:1.7}}>
      想定：10km四方の運行エリア（マップ横幅＝10km）、平均速度25km/h（市街地・信号停車込み）。
      停留所座標に基づく道路ネットワーク（各停留所を近傍4箇所と接続、区間ごとに道路係数1.10〜1.35）
      上の全ペア最短時間。対称性と三角不等式（TT(x,y)≦TT(x,z)+TT(z,y)）を全30×30×30通りで検証済み＝矛盾なし（6秒単位）。
      所要時間は2.3〜28.8分・中央値13分。マップの薄い線が隣接区間。<br/>
      参考：寄り道許容 STn＝三段式（短≤{Math.round(P.sl.d1/60)}分で{Math.round(P.sl.s1/60)}分・長≥{Math.round(P.sl.d2/60)}分で{Math.round(P.sl.s2/60)}分・間は線形）、最大乗車 MRTn＝DRTn＋STn。
      実測値が得られた区間から順次差し替え可能（その場合は再度三角不等式の検証を要する）。
    </div>
  </div>);
}

/* ---------- 車両設定モーダル（台数・定員・運行時間） ---------- */
function VehicleModal({onClose,vehicles,setVehicles,sl,setSl,tw,setTw,onLoadNetwork,netVer,onApplyBackbone}){
  const [bbVeh,setBbVeh]=React.useState(null);        // 背骨編集中の号車id
  const [bbList,setBbList]=React.useState([]);        // 編集中の背骨 [{stop,time}]
  const openBackbone=(v)=>{ setBbVeh(v.id);
    // 過去の編集で文字列時刻が混入していても、ここで秒に正規化してから編集に入る
    setBbList((v.backbone&&v.backbone.length)
      ?v.backbone.map(a=>({stop:a.stop,time:typeof a.time==="number"?a.time:hmToSec(a.time,9*3600)}))
      :[]); };
  const addAnchor=()=>setBbList(l=>[...l,{stop:0,time:9*3600}]);
  const setAnchor=(i,k,val)=>setBbList(l=>l.map((a,j)=>j===i?{...a,[k]:val}:a));
  const rmAnchor=(i)=>setBbList(l=>l.filter((_,j)=>j!==i));
  const [netMsg,setNetMsg]=React.useState(null);
  const [stopsFile,setStopsFile]=React.useState(null);
  const [odFile,setOdFile]=React.useState(null);
  const [odFactor,setOdFactor]=React.useState(1.0);   // OD所要時間に掛ける倍率（実勢補正）
  const readFile=f=>new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result));r.onerror=()=>rej(new Error("読込失敗"));r.readAsText(f);});
  const doLoadNetwork=async()=>{
    try{
      if(!stopsFile||!odFile){setNetMsg({err:true,text:"stops.csv と od_time.csv の両方を選んでください。"});return;}
      const [st,od]=await Promise.all([readFile(stopsFile),readFile(odFile)]);
      const stops=parseStopsCSV(st), odRows=parseODCSV(od);
      const f=Number(odFactor);
      if(!isFinite(f)||f<0.1||f>5){setNetMsg({err:true,text:"所要時間の倍率は0.1〜5.0の範囲で指定してください。"});return;}
      const info=onLoadNetwork(stops,odRows,f);
      const covPct=Math.round(info.coverage*100);
      const lines=[`読込成功：停留所 ${info.n}点・OD ${info.odFilled}件${info.factor!==1?`（所要時間×${info.factor}で取込）`:""}。予約はリセットされました。`,
        `網羅率 ${covPct}%（${info.known}/${info.totalPairs}ペア）${info.asym>0?`・非対称 ${info.asym}ペア`:""}`];
      if(covPct<100) lines.push(`欠損ペアは仮値（3時間）で埋まっており実距離ではない。例：${info.missing.slice(0,4).join("／")}${info.missing.length>4?" ほか":""}`);
      setNetMsg({err:covPct<100, text:lines.join("\n")});
    }catch(e){ setNetMsg({err:true,text:"読込失敗："+(e.message||String(e))}); }
  };
  const upd=(id,patch)=>setVehicles(vs=>vs.map(v=>v.id===id?{...v,...patch}:v));
  const toHM=s=>`${String(Math.floor(s/3600)).padStart(2,"0")}:${String(Math.floor(s%3600/60)).padStart(2,"0")}`;
  const toSec=t=>{const[h,m]=t.split(":").map(Number);return h*3600+m*60;};
  const th={padding:"6px 10px",borderBottom:"2px solid #1E2A38",fontSize:11,whiteSpace:"nowrap"};
  const td={padding:"6px 10px",borderBottom:"1px solid #DDD8CB",fontSize:13,whiteSpace:"nowrap",textAlign:"center"};
  const inp={fontFamily:"'SF Mono','Consolas',monospace",padding:"4px 6px",border:"1px solid #C9C4B8",
    borderRadius:6,fontSize:13,background:"#fff"};
  return (
  <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(10,16,26,0.55)",
    display:"flex",alignItems:"center",justifyContent:"center",zIndex:50}}>
    <div onClick={e=>e.stopPropagation()} style={{background:"#F5F3EE",borderRadius:12,
      width:"min(640px,94vw)",maxHeight:"92vh",overflowY:"auto",padding:18,boxShadow:"0 18px 60px rgba(0,0,0,0.45)"}}>
      <div style={{display:"flex",alignItems:"center",marginBottom:10}}>
        <div style={{fontSize:16,fontWeight:700}}>車両設定</div>
        <button onClick={onClose} style={{marginLeft:"auto",border:"1px solid #C9C4B8",background:"#fff",
          borderRadius:6,padding:"4px 12px",cursor:"pointer",fontSize:13}}>閉じる</button>
      </div>

      {/* 基本運行設定（頻繁に変えない設定。寄り道ST・約束幅TW） */}
      <div style={{background:"#fff",border:"1px solid #D8D3C6",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
        <div style={{fontSize:13,fontWeight:700,marginBottom:8,color:"#3A3526"}}>基本運行設定</div>
        <div style={{display:"flex",gap:24,alignItems:"flex-start",flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:11,color:"#6B6453",marginBottom:4}}>
              寄り道 ST（分）{sl.s2<sl.s1&&<span style={{color:"#9B3B2B",marginLeft:4}}>⚠最大{"<"}最小で逆転</span>}
            </div>
            <div style={{display:"flex",gap:6,alignItems:"flex-end"}}>
              <SLNum v={sl.d1} on={x=>setSl(s=>({...s,d1:x}))} t="短≤"/>
              <SLNum v={sl.s1} on={x=>setSl(s=>({...s,s1:x}))} t="で"/>
              <span style={{color:"#9AA0A8",fontSize:12,paddingBottom:5}}>→</span>
              <SLNum v={sl.d2} on={x=>setSl(s=>({...s,d2:x}))} t="長≥"/>
              <SLNum v={sl.s2} on={x=>setSl(s=>({...s,s2:x}))} t="で"/>
            </div>
            <div style={{fontSize:10,color:"#9AA0A8",marginTop:3}}>短≤分→一律・長≥分→一律・間は線形。最大乗車=直行+ST。</div>
          </div>
          <div>
            <div style={{fontSize:11,color:"#6B6453",marginBottom:4}}>約束幅 TW</div>
            <select value={tw} onChange={e=>setTw(Number(e.target.value))} style={inp}>
              <option value={300}>5分</option><option value={600}>10分</option>
            </select>
            <div style={{fontSize:10,color:"#9AA0A8",marginTop:3}}>約束発〜約束発+TWの間に乗車。</div>
          </div>
        </div>
      </div>

      {/* 停留所・OD直行時間表の読込（実ネットワークへ差し替え） */}
      <div style={{background:"#fff",border:"1px solid #D8D3C6",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
        <div style={{fontSize:13,fontWeight:700,marginBottom:2,color:"#3A3526"}}>停留所・OD直行時間表の読込</div>
        <div style={{fontSize:11,color:"#6B6453",marginBottom:8,lineHeight:1.6}}>
          stops.csv（stop_id,name,lat,lon,is_hub）と od_time.csv（from_id,to_id,seconds）を読み込むと、
          地図・停留所・OD直行時間がそのデータに差し替わる。読込時に既存の予約はリセットされる。
          「所要時間×」は経路検索由来の理論値を実勢に補正する倍率（例：1.5＝5割増し）。
        </div>
        <div style={{display:"flex",gap:16,flexWrap:"wrap",alignItems:"center"}}>
          <label style={{fontSize:12,color:"#3A3526"}}>停留所：
            <input type="file" accept=".csv,text/csv" onChange={e=>setStopsFile(e.target.files?.[0]||null)}
              style={{fontSize:11,marginLeft:4}}/>
          </label>
          <label style={{fontSize:12,color:"#3A3526"}}>OD表：
            <input type="file" accept=".csv,text/csv" onChange={e=>setOdFile(e.target.files?.[0]||null)}
              style={{fontSize:11,marginLeft:4}}/>
          </label>
          <label style={{fontSize:12,color:"#3A3526"}} title="CSVの所要時間に掛ける倍率。1.0＝そのまま、1.5＝5割増し（信号待ち・混雑等の実勢補正）。欠損ペアの仮値には掛からない">
            所要時間×
            <input type="number" min="0.1" max="5" step="0.1" value={odFactor}
              onChange={e=>setOdFactor(e.target.value)}
              style={{fontSize:11,width:52,marginLeft:4,padding:"3px 4px",
                border:"1px solid #C9C4B8",borderRadius:5,textAlign:"center"}}/>
          </label>
          <button onClick={doLoadNetwork} disabled={!stopsFile||!odFile}
            style={{border:"1px solid #1E2A38",background:(!stopsFile||!odFile)?"#C9C4B8":"#1E2A38",
              color:"#fff",borderRadius:7,padding:"6px 16px",cursor:(!stopsFile||!odFile)?"not-allowed":"pointer",
              fontSize:12,fontWeight:700}}>読み込む</button>
        </div>
        {netMsg&&<div style={{marginTop:8,fontSize:11.5,fontWeight:600,whiteSpace:"pre-line",lineHeight:1.6,
          color:netMsg.err?"#9B3B2B":"#23694A"}}>{netMsg.text}</div>}
        <div style={{marginTop:6,fontSize:10.5,color:"#9AA0A8"}}>
          現在：停留所 {STOPS.length}点／OD表 {STOPS.length}×{STOPS.length}。
        </div>
      </div>

      {/* 運行モード：フルデマンド／セミデマンド（背骨あり） */}
      <div style={{background:"#fff",border:"1px solid #D8D3C6",borderRadius:10,padding:"12px 14px",marginBottom:14}}>
        <div style={{fontSize:13,fontWeight:700,marginBottom:2,color:"#3A3526"}}>運行モード（フル／セミデマンド）</div>
        <div style={{fontSize:11,color:"#6B6453",marginBottom:8,lineHeight:1.6}}>
          セミは背骨（アンカー＝停留所と時刻）を先に敷く。車両は各アンカー時刻にその停留所へ（早着は待機）、区間の余裕分だけ寄り道して予約を拾う。
          背骨を適用するとその号車の既存予約は一旦外れる。逆方向や背骨外の需要はフルの号車が拾う。
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
          {vehicles.map(v=>(
            <div key={v.id} style={{display:"flex",alignItems:"center",gap:6,border:"1px solid #E3DFD3",
              borderRadius:8,padding:"5px 9px",background:bbVeh===v.id?"#EEF3EF":"#FAF8F2"}}>
              <Dot c={v.color}/><span style={{fontSize:12,fontWeight:700}}>{v.name}</span>
              <span style={{fontSize:10.5,color:v.mode==="semi"?"#23694A":"#8A8474"}}>
                {v.mode==="semi"?`セミ（${(v.backbone||[]).length}点）`:"フル"}
              </span>
              <button onClick={()=>openBackbone(v)} style={{fontSize:10.5,border:"1px solid #B9B2A1",
                background:"#fff",borderRadius:5,padding:"2px 7px",cursor:"pointer"}}>背骨を編集</button>
            </div>
          ))}
        </div>
        {bbVeh!=null && <div style={{border:"1px solid #C9C4B8",borderRadius:8,padding:"10px",background:"#FAF8F2"}}>
          <div style={{fontSize:12,fontWeight:700,marginBottom:6}}>
            {vehicles.find(v=>v.id===bbVeh)?.name} の背骨（{bbList.length}点・上から時刻順・空でフルデマンド）
          </div>
          {bbList.length===0 && <div style={{fontSize:11,color:"#8A8474",marginBottom:6}}>アンカーなし＝この号車はフルデマンド。</div>}
          {/* アンカー一覧：生成背骨（往復×多便）では数十〜百点になるため高さ上限＋スクロール。
              アンカーのtimeは内部的に秒で持つ。TimeInputはHH:MM文字列を扱うため、この境界で相互変換する
              （変換なしで秒を渡すと不正値で --:-- になり、編集すると文字列が保存されて時刻計算を壊す）。 */}
          <div style={{maxHeight:"38vh",overflowY:"auto",paddingRight:4}}>
          {bbList.map((a,i)=>(
            <div key={i} style={{display:"flex",gap:6,alignItems:"center",marginBottom:5}}>
              <span style={{fontSize:11,color:"#8A8474",width:22,textAlign:"right"}}>{i+1}</span>
              <select value={a.stop} onChange={e=>setAnchor(i,"stop",Number(e.target.value))}
                style={{padding:"4px 6px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:12,background:"#fff"}}>
                {STOPS.map((nm,si)=><option key={si} value={si}>{nm}</option>)}
              </select>
              <TimeInput value={secToHM(a.time)} onChange={t=>setAnchor(i,"time",hmToSec(t,a.time))}/>
              <button onClick={()=>rmAnchor(i)} style={{fontSize:11,border:"1px solid #C9A99B",color:"#9B3B2B",
                background:"#fff",borderRadius:5,padding:"2px 8px",cursor:"pointer"}}>削除</button>
            </div>
          ))}
          </div>
          <div style={{display:"flex",gap:8,marginTop:8}}>
            <button onClick={addAnchor} style={{fontSize:11.5,border:"1px solid #1E2A38",background:"#fff",
              borderRadius:6,padding:"4px 12px",cursor:"pointer"}}>＋アンカー追加</button>
            <button onClick={()=>{const r=onApplyBackbone(bbVeh,bbList.length?"semi":"full",bbList);
              const vn=vehicles.find(v=>v.id===bbVeh)?.name||"";
              setNetMsg({err:false,text:bbList.length
                ?`${vn}をセミに設定（アンカー${bbList.length}点）。この号車の既存予約${r.cleared}件を外した。動きを見るには「予約一覧の流し込み」で同じ需要を流し直すこと（既定で既存予約クリア済みの状態から確定する）。`
                :`${vn}をフルに戻した。既存予約${r.cleared}件を外した。`});
              setBbVeh(null);}}
              style={{fontSize:11.5,border:"1px solid #23694A",background:"#2E9E6B",color:"#fff",
                borderRadius:6,padding:"4px 14px",cursor:"pointer",fontWeight:700}}>
              {bbList.length?"セミとして適用":"フルに戻す"}
            </button>
            <button onClick={()=>setBbVeh(null)} style={{fontSize:11.5,border:"1px solid #B9B2A1",
              background:"#fff",borderRadius:6,padding:"4px 12px",cursor:"pointer"}}>閉じる</button>
          </div>
        </div>}
      </div>

      <table style={{borderCollapse:"collapse",background:"#fff",border:"1px solid #B9B2A1",width:"100%"}}>
        <thead><tr style={{background:"#F0EDE4"}}>
          <th style={th}>稼働</th><th style={th}>車両</th><th style={th}>定員</th>
          <th style={th}>運行開始</th><th style={th}>運行終了</th>
        </tr></thead>
        <tbody>
          {vehicles.map(v=>(
            <tr key={v.id} style={{opacity:v.active?1:0.5}}>
              <td style={td}>
                <input type="checkbox" checked={v.active} onChange={e=>upd(v.id,{active:e.target.checked})}/>
              </td>
              <td style={{...td,fontWeight:700}}><Dot c={v.color}/> {v.name}</td>
              <td style={td}>
                <select value={v.cap} onChange={e=>upd(v.id,{cap:Number(e.target.value)})} style={inp}>
                  {[4,6,8,10,12,14].map(n=><option key={n} value={n}>{n}名</option>)}
                </select>
              </td>
              <td style={td}>
                <input type="time" value={toHM(v.start)} step={1800}
                  onChange={e=>upd(v.id,{start:toSec(e.target.value)})} style={inp}/>
              </td>
              <td style={td}>
                <input type="time" value={toHM(v.end)} step={1800}
                  onChange={e=>upd(v.id,{end:toSec(e.target.value)})} style={inp}/>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <div style={{fontSize:11,color:"#8A8474",marginTop:10,lineHeight:1.7}}>
        設定は新規予約の探索から適用される。確定済み予約の約束は変更されないため、
        既に予約が載っている車両の運行時間短縮・定員削減・稼働OFFは既存計画を実行不可にすることがある
        （その場合は受付パネルに警告を表示）。稼働OFFの車両は候補探索の対象外。
      </div>
    </div>
  </div>);
}

/* ---------- 予約一覧の流し込み（CSV / Excel） ----------
   列：発地, 着地, 希望(発/着), 時刻, 人数, ズレ幅(分)
   上から順に「おすすめ（最小コスト候補）」で自動確定。
   結果を 成立件数/総件数 で表示し、不成立行は理由を併記。 */

const IMPORT_SAMPLE=`発地,着地,希望,時刻,人数,ズレ幅
BS001,BS012,発,9:15,1,15
BS007,BS022,発,10:00,2,10
BS015,BS003,着,11:30,1,20
BS028,BS009,発,13:45,3,30`;

function parseStop(v){
  if(v==null)return null;
  const s=String(v).trim().toUpperCase().replace("BS","");
  const n=parseInt(s,10);
  return (n>=1&&n<=STOPS.length)?n-1:null;
}
function parseTimeCell(v){
  if(v==null)return null;
  if(typeof v==="number") return v<1.5 ? Math.round(v*86400) : null; // Excelシリアル時刻
  const m=String(v).trim().match(/^(\d{1,2}):(\d{2})/);
  return m?(+m[1])*3600+(+m[2])*60:null;
}
function parseRows(raw){
  // raw: 配列の配列。先頭行が見出しなら読み飛ばす
  const rows=raw.filter(r=>r&&r.some(c=>c!=null&&String(c).trim()!==""));
  if(rows.length===0)return [];
  const head=rows[0].map(c=>String(c??""));
  const hasHeader=head.some(c=>/発地|着地|希望|時刻|人数|ズレ/.test(c));
  let idx={o:0,d:1,mode:2,time:3,pax:4,sa:5};
  if(hasHeader){
    const find=(re,def)=>{const i=head.findIndex(c=>re.test(c));return i>=0?i:def;};
    idx={o:find(/発地/,0),d:find(/着地/,1),mode:find(/希望/,2),time:find(/時刻/,3),
         pax:find(/人数/,4),sa:find(/ズレ/,5)};
  }
  return rows.slice(hasHeader?1:0).map((r,n)=>{
    const o=parseStop(r[idx.o]), d=parseStop(r[idx.d]);
    const modeStr=String(r[idx.mode]??"発");
    const mode=/着/.test(modeStr)?"arr":"dep";
    const t=parseTimeCell(r[idx.time]);
    const pax=Math.max(1,parseInt(r[idx.pax],10)||1);
    const saMin=parseInt(r[idx.sa],10);
    const sa=(isNaN(saMin)?10:saMin)*60;
    let err=null;
    if(o==null)err="発地が不正";
    else if(d==null)err="着地が不正";
    else if(o===d)err="発地＝着地";
    else if(t==null)err="時刻が不正";
    return {line:n+1,o,d,mode,t,pax,sa,err};
  });
}

// 流し込みデータの書き出し（CSV / Excel）
function rowsToTable(rows,withResult,vehicles){
  const vn=id=>vehicles.find(v=>v.id===id)?.name??id;
  const head=["発地","着地","希望","時刻","人数","ズレ幅"];
  if(withResult)head.push("結果","予約番号","便","約束発","約束着","希望とのずれ");
  const body=rows.map(r=>{
    const base=[
      r.o!=null?STOPS[r.o]:"",
      r.d!=null?STOPS[r.d]:"",
      r.mode==="arr"?"着":"発",
      r.t!=null?fmt(r.t):"",
      r.pax,
      Math.round(r.sa/60),
    ];
    if(withResult){
      base.push(
        r.err?`形式エラー（${r.err}）`:r.ok?"成立":"不成立",
        r.ok?r.num:"",
        r.ok?vn(r.vehicle):"",
        r.ok?fmt(r.apt):"",
        r.ok?fmt(r.idt):"",
        r.ok?fmtDev(r.dev):"",
      );
    }
    return base;
  });
  return [head,...body];
}
function buildCSV(table){
  return table.map(r=>r.map(c=>{
    const s=String(c??"");
    return /[",\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s;
  }).join(",")).join("\n");
}
// ダウンロードを試行（サンドボックス環境では失敗しうるため真偽を返す）
function tryDownload(csv,filename){
  try{
    const url=URL.createObjectURL(new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8"}));
    const a=document.createElement("a");a.href=url;a.download=filename;
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url),5000);
    return true;
  }catch(e){return false;}
}
const stamp=()=>{
  const d=new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}_${String(d.getHours()).padStart(2,"0")}${String(d.getMinutes()).padStart(2,"0")}`;
};

function ImportModal({onClose,vehicles,P,state,setState,onShowAnalysis,setDemandViz,demandViz,demandMeta,setDemandMeta}){
  const [parsed,setParsed]=React.useState(()=>demandViz||null); // 再オープン時は保持中の需要を復元
  const [clearFirst,setClearFirst]=React.useState(true); // 流し込み前に既存予約をクリア（背骨・車両設定は維持）。
                                                          // OFF＝現在の予約に追加（二重計上に注意）
  const [result,setResult]=React.useState(null);
  const [err,setErr]=React.useState(null);
  const [nDemand,setNDemand]=React.useState(100);   // テスト需要の件数
  const [target,setTarget]=React.useState(90);      // 目標成立率(%)
  const [peakInt,setPeakInt]=React.useState(0.6);   // 時間帯ピーク強度 0〜1
  const [hubConc,setHubConc]=React.useState(0.6);   // 拠点集中度 0〜1
  const [genSa,setGenSa]=React.useState(900);       // 生成需要のズレ幅San（秒・全件共通）
  const [genRet,setGenRet]=React.useState(0.5);     // 往復化する割合 0〜1
  const [genStay,setGenStay]=React.useState(3600);  // 目的施設での滞留時間（秒・中央値）
  const [sizing,setSizing]=React.useState(null);    // 規模シミュレーション結果
  const [reopt,setReopt]=React.useState(null);      // 全体最適化シミュレーション結果
  const [pendingCommit,setPendingCommit]=React.useState(null); // 運行確定の確認待ち {plan,label}
  const [exp,setExp]=React.useState(null);           // {name,csv} 出力パネル
  const expRef=React.useRef(null);
  const openExport=(table,name)=>{
    const csv=buildCSV(table);
    setExp({name,csv,tried:tryDownload(csv,name)});
  };
  const copyExp=async()=>{
    try{await navigator.clipboard.writeText(exp.csv);
      setExp(e=>({...e,copied:true}));}
    catch(ex){
      // クリップボードAPIが使えない環境向けの旧来コピー
      if(expRef.current){expRef.current.select();
        try{document.execCommand("copy");setExp(e=>({...e,copied:true}));}
        catch(e2){setExp(e=>({...e,copied:false}));}}
    }
  };

  const onFile=async(e)=>{
    setErr(null);setResult(null);setParsed(null);
    const f=e.target.files?.[0]; if(!f)return;
    try{
      const buf=await f.arrayBuffer();
      let raw;
      if(/\.csv$|\.txt$/i.test(f.name)){
        const text=new TextDecoder("utf-8").decode(buf).replace(/^\uFEFF/,"");
        raw=text.split(/\r?\n/).map(l=>l.split(/[,\t]/));
      }else{
        const wb=XLSX.read(buf,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        raw=XLSX.utils.sheet_to_json(ws,{header:1,raw:true});
      }
      const rows=parseRows(raw);
      if(rows.length===0){setErr("有効な行がない。");return;}
      setDemandMeta({source:"file",name:f.name});
      setParsed(rows);
      setDemandViz?.(rows);
    }catch(ex){setErr("読込失敗："+ex.message);}
  };

  // 実需要モデル：時間帯ピーク（朝夕・通院帯）＋OD方向性（朝=周辺→中心/夕=中心→周辺）。
  // 目的地が拠点（駅・病院など）の行き便は着希望（到着時刻指定）を多めに生成。帰り便は発希望。
  const genTest=(n)=>{
    setSizing(null);
    // 生成窓は稼働車両の運行時間から導出（最早の運行開始〜最遅の運行終了）。
    // 手順として運行時間の設定が先＝車両設定を変えたら生成し直す前提。
    const actives=vehicles.filter(v=>v.active);
    if(!actives.length){setErr("稼働車両が0台。車両設定で車両を有効化し、運行時間を設定してから生成すること。");return;}
    const T_START=Math.min(...actives.map(v=>v.start));
    const T_END=Math.max(...actives.map(v=>v.end));
    if(T_END-T_START<3600){setErr(`運行時間が短すぎる（${fmt(T_START)}〜${fmt(T_END)}）。車両設定で1時間以上を確保すること。`);return;}
    setErr(null);
    const peak=peakInt, hubC=hubConc;   // 0〜1
    const nonHub=STOPS.map((_,i)=>i).filter(i=>!HUBS.includes(i));
    // 時刻の確率分布：朝(8-10)・昼通院(11-13)・夕(15-17)に山。peakで尖り具合を制御。
    // 山の位置は共通定義のまま、生成窓の外は切り落とされる（窓が昼だけなら昼の山だけ効く）
    const slots=[]; // {t, w}
    for(let t=T_START;t<T_END;t+=300){
      const h=t/3600;
      const g=(c,s)=>Math.exp(-((h-c)**2)/(2*s*s));
      const base=0.15;
      const shape=DEMAND_PEAKS.reduce((a,p)=>a+g(p.center,p.sigma)*p.weight,0); // 共通定義の3山
      slots.push({t, w: base + peak*shape*3.2});
    }
    const totW=slots.reduce((a,s)=>a+s.w,0);
    const pickTime=()=>{
      let r=Math.random()*totW;
      for(const s of slots){r-=s.w; if(r<=0)return s.t;}
      return slots[slots.length-1].t;
    };
    const pickHub=()=>HUBS[Math.floor(Math.random()*HUBS.length)];
    const pickNon=()=>nonHub[Math.floor(Math.random()*nonHub.length)];
    const rows=[];
    const ROUND=v=>Math.round(v/300)*300;
    let loopGuard=0;
    while(rows.length<n && loopGuard++<n*30){
      const t=pickTime();
      const h=t/3600;
      const dirToHub = h<11.5 ? true : (h>=14 ? false : Math.random()<0.5);  // 朝＝周辺→中心、夕＝中心→周辺、昼は両方向
      let o,d;
      const useHub=Math.random()<hubC;   // 拠点集中度：この確率で片端を拠点に
      if(useHub){
        if(dirToHub){o=pickNon(); d=pickHub();}
        else {o=pickHub(); d=pickNon();}
      }else{
        o=Math.floor(Math.random()*STOPS.length);
        d=Math.floor(Math.random()*STOPS.length);
      }
      let guard=0;
      while(d===o&&guard++<10)d=Math.floor(Math.random()*STOPS.length);
      if(d===o)continue;
      // 乗車人数：1人80%・2人15%・3人5%（旧配分は1人57%で複数人が過多だった）
      const pr=Math.random();
      const pax=pr<0.80?1:(pr<0.95?2:3);
      const sa=genSa;   // ズレ幅はスライダ値（全件共通）
      // 目的地が拠点（駅・病院など）なら着希望を多めに（到着時刻指定が現実的）。
      // 非拠点が目的地なら発希望。tは発希望なら希望乗車、着希望なら希望到着として解釈される。
      const toArr = HUBS.includes(d) && Math.random()<0.7;
      // 運行時間内に完結しない希望は棄却して引き直す：
      //  発希望＝希望発＋乗降＋最大乗車が運行終了まで／着希望＝希望着が運行開始＋所要以降かつ終了まで
      const tripMax=P.dwell+mrtFromDRT(TT[o][d],P);
      if(toArr){ if(t<T_START+tripMax || t>T_END) continue; }
      else     { if(t+tripMax>T_END) continue; }
      rows.push({line:rows.length+1,o,d,mode:toArr?"arr":"dep",t,pax,sa,err:null});   // 行き
      // 帰り便：目的施設での滞留後に逆向きで乗車（帰りは発希望＝用が済んだら乗る）。比率genRetで付与。
      if(rows.length<n && Math.random()<genRet){
        const stay=Math.round(genStay*(0.7+0.6*Math.random()));   // 滞留時間（中央値±30%程度）
        // 目的地への到着時刻：着希望ならt（希望到着）、発希望ならt＋乗降＋最大乗車の推定着
        const arriveAt = toArr ? t : (t+P.dwell+mrtFromDRT(TT[o][d],P));
        const tRet=ROUND(arriveAt+stay);
        const retMax=P.dwell+mrtFromDRT(TT[d][o],P);
        if(tRet+retMax<=T_END){                                    // 帰りの降車まで運行時間内に収まる場合のみ採用
          rows.push({line:rows.length+1,o:d,d:o,mode:"dep",t:tRet,pax,sa,err:null});
        }
      }
    }
    setErr(null);setResult(null);setParsed(rows);
    setDemandMeta({source:"generated",
      settings:{n,peakInt,hubConc,genSa,genRet,genStay}});
    setDemandViz?.(rows);
  };

  const run=()=>{
    if(!parsed)return;
    // クリアON（既定）：予約とルートを白紙にしてから流し込む。ただしセミ車の背骨アンカーは
    // 再注入する＝運行設定は維持。背骨適用→同じ需要で流し直す、が1操作で正しく行える。
    // クリアOFF：現在の状態に追加（従来動作）。同じ需要を2回流すと二重計上になる点に注意。
    let resMap, routes;
    if(clearFirst){
      resMap={};
      routes={};
      for(const k of Object.keys(state.routes)) routes[k]=[];
      for(const v of vehicles) if(v.mode==="semi"&&v.backbone&&v.backbone.length)
        routes[v.id]=anchorEvents(v.backbone);
    }else{
      resMap={...state.resMap};
      routes={};
      for(const k of Object.keys(state.routes)) routes[k]=[...state.routes[k]];
    }
    let num=state.nextNum;
    const out=[];
    for(const row of parsed){
      if(row.err){out.push({...row,ok:false,reason:row.err});continue;}
      const drt=TT[row.o][row.d];
      const mrt=mrtFromDRT(drt,P);
      const dpt=row.mode==="dep"?row.t:row.t-(mrt+P.dwell);
      const id="I"+num+"_"+Math.random().toString(36).slice(2,7);
      const r={id,num,o:row.o,d:row.d,dpt,drt,mrt,sa:row.sa,pax:row.pax,
        ipt:null,idt:null,vehicle:null,label:`予約${num}`};
      const cands=searchInsertions(routes,resMap,r,P,vehicles);
      if(cands.length>0){
        const c=cands[0];
        r.ipt=c.apt;r.idt=c.apt+P.dwell+r.mrt;r.vehicle=c.vehicle;
        routes[c.vehicle]=c.route;resMap[id]=r;num++;
        out.push({...row,ok:true,num:r.num,vehicle:c.vehicle,apt:c.apt,
          idt:r.idt,dev:c.apt-dpt,maxLoad:c.maxLoad,cap:c.cap});
      }else{
        out.push({...row,ok:false,
          reason:"全稼働車両で約束条件（San・MRT・定員・運行時間）を満たす挿入位置なし"});
      }
    }
    const rec=out.map(r=>({o:r.o,d:r.d,mode:r.mode,t:r.t,pax:r.pax,sa:r.sa,
      ok:!!r.ok,reason:r.reason,num:r.num,vehicle:r.vehicle,apt:r.apt,idt:r.idt}));
    const okN=out.filter(r=>r.ok).length;
    const lastDemand={...(demandMeta||{source:"unknown"}),
      committedBy:"自動確定（先着順）",when:Date.now(),
      okRate:out.length?okN/out.length:null,rows:rec};
    setState(s=>({...s,resMap,routes,nextNum:num,lastDemand}));
    setResult({ok:out.filter(r=>r.ok).length,total:out.length,rows:out});
  };

  // シミュレーション結果（plan＝routes/resMap）を運行に確定反映する。
  // 既存の確定予約は置き換える（A/Bは空からこの需要をさばいた“その日の計画”のため）。
  const commitPlan=(plan,demandRec)=>{
    if(!plan||!plan.resMap){return;}
    const ids=Object.keys(plan.resMap).sort((a,b)=>(plan.resMap[a].ipt??0)-(plan.resMap[b].ipt??0));
    const resMap={};let num=1;
    for(const id of ids){const r=plan.resMap[id];resMap[id]={...r,num,label:`予約${num}`};num++;}
    const routes={};for(const v of vehicles)routes[v.id]=plan.routes[v.id]?[...plan.routes[v.id]]:[];
    setState(s=>({...s,resMap,routes,nextNum:num,
      lastDemand:demandRec??s.lastDemand??null}));
    setPendingCommit(null);
    onClose();   // 運行盤・ダイヤ・確認表・分析に反映された状態を見る
  };

  const th={padding:"4px 8px",borderBottom:"2px solid #1E2A38",fontSize:11,whiteSpace:"nowrap"};
  const td={padding:"3px 8px",borderBottom:"1px solid #DDD8CB",fontSize:12,whiteSpace:"nowrap",
    fontFamily:"'SF Mono','Consolas',monospace",textAlign:"center"};
  const vn=id=>vehicles.find(v=>v.id===id)?.name??id;

  return (
  <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(10,16,26,0.55)",
    display:"flex",alignItems:"center",justifyContent:"center",zIndex:50}}>
    <div onClick={e=>e.stopPropagation()} style={{background:"#F5F3EE",borderRadius:12,
      width:"min(880px,94vw)",maxHeight:"88vh",overflow:"auto",padding:18,
      boxShadow:"0 18px 60px rgba(0,0,0,0.45)"}}>
      <div style={{display:"flex",alignItems:"center",marginBottom:10}}>
        <div style={{fontSize:16,fontWeight:700}}>予約一覧の流し込み</div>
        <button onClick={onClose} style={{marginLeft:"auto",border:"1px solid #C9C4B8",background:"#fff",
          borderRadius:6,padding:"4px 12px",cursor:"pointer",fontSize:13}}>閉じる</button>
      </div>

      {!result&&<>
        <div style={{fontSize:12,color:"#6B6453",lineHeight:1.8,marginBottom:10}}>
          CSV または Excel（先頭シート）を読み込み、上から順に「おすすめ候補」で自動確定する（先着順）。<br/>
          列：<b>発地, 着地, 希望（発/着）, 時刻, 人数, ズレ幅（分）</b>。見出し行は自動判別。停留所はBS001〜BS{String(STOPS.length).padStart(3,"0")}。
        </div>
        <pre style={{background:"#fff",border:"1px solid #D8D3C6",borderRadius:8,padding:10,
          fontSize:11.5,fontFamily:"'SF Mono','Consolas',monospace",margin:"0 0 10px"}}>{IMPORT_SAMPLE}</pre>
        <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
          <input type="file" accept=".csv,.txt,.xlsx,.xls" onChange={onFile} style={{fontSize:13}}/>
          <span style={{fontSize:11,color:"#8A8474"}}>または</span>
          <input type="number" min={10} max={500} step={10} value={nDemand}
            onChange={e=>setNDemand(Math.max(10,Math.min(500,Number(e.target.value)||100)))}
            style={{width:64,padding:"5px 6px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:13,
              fontFamily:"'SF Mono','Consolas',monospace"}}/>
          <button onClick={()=>genTest(nDemand)} style={{padding:"6px 12px",borderRadius:8,fontSize:12,cursor:"pointer",
            border:"1px solid #14202F",background:"#fff",color:"#14202F",fontWeight:700}}>
            件のテスト需要を生成
          </button>
          {(()=>{const a=vehicles.filter(v=>v.active);
            if(!a.length)return <span style={{fontSize:11,color:"#9B3B2B",fontWeight:700}}>稼働車両が0台。先に車両設定で運行時間を設定すること。</span>;
            const t0=Math.min(...a.map(v=>v.start)), t1=Math.max(...a.map(v=>v.end));
            return <span style={{fontSize:11,color:"#8A8474"}}>
              生成範囲 {fmt(t0)}〜{fmt(t1)}（車両設定の運行時間に連動。変更する場合は先に車両設定を）
            </span>;})()}
        </div>
        <div style={{marginTop:10,padding:10,background:"#fff",border:"1px solid #D8D3C6",borderRadius:8}}>
          <div style={{fontSize:11.5,color:"#6B6453",marginBottom:6}}>
            実需要モデル：朝（周辺→中心）・昼（通院）・夕（中心→周辺）の時間帯ピークと、拠点
            （{HUBS.map(h=>STOPS[h]+(HUB_NAMES[h]?`=${HUB_NAMES[h]}`:"")).join("・")}）への発着集中を反映。
            一部は目的施設での滞留後に逆向きで戻る往復需要として生成する。
          </div>
          <div style={{display:"flex",gap:18,flexWrap:"wrap"}}>
            <label style={{fontSize:11,color:"#444",display:"flex",flexDirection:"column",gap:2}}>
              時間帯ピーク強度：{Math.round(peakInt*100)}%
              <input type="range" min={0} max={1} step={0.1} value={peakInt}
                onChange={e=>setPeakInt(Number(e.target.value))} style={{width:180}}/>
              <span style={{fontSize:10,color:"#8A8474"}}>0=終日均等／高=朝夕に集中</span>
            </label>
            <label style={{fontSize:11,color:"#444",display:"flex",flexDirection:"column",gap:2}}>
              拠点集中度：{Math.round(hubConc*100)}%
              <input type="range" min={0} max={1} step={0.1} value={hubConc}
                onChange={e=>setHubConc(Number(e.target.value))} style={{width:180}}/>
              <span style={{fontSize:10,color:"#8A8474"}}>0=全停留所均等／高=拠点に発着集中</span>
            </label>
            <label style={{fontSize:11,color:"#444",display:"flex",flexDirection:"column",gap:2}}>
              ズレ幅（希望±）：±{Math.round(genSa/60)}分
              <input type="range" min={300} max={7200} step={300} value={genSa}
                onChange={e=>setGenSa(Number(e.target.value))} style={{width:180}}/>
              <span style={{fontSize:10,color:"#8A8474"}}>狭=希望時刻に厳格／広=相乗りの余地が増える</span>
            </label>
            <label style={{fontSize:11,color:"#444",display:"flex",flexDirection:"column",gap:2}}>
              帰り便の比率：{Math.round(genRet*100)}%
              <input type="range" min={0} max={1} step={0.1} value={genRet}
                onChange={e=>setGenRet(Number(e.target.value))} style={{width:180}}/>
              <span style={{fontSize:10,color:"#8A8474"}}>行きのうち何割を往復にするか（帰りは逆向き）</span>
            </label>
            <label style={{fontSize:11,color:"#444",display:"flex",flexDirection:"column",gap:2}}>
              滞留時間（中央値）：{Math.round(genStay/60)}分
              <input type="range" min={900} max={10800} step={900} value={genStay}
                onChange={e=>setGenStay(Number(e.target.value))} style={{width:180}}/>
              <span style={{fontSize:10,color:"#8A8474"}}>目的施設での滞在。帰りの出発時刻＝行きの着＋滞留（±30%）</span>
            </label>
          </div>
        </div>
        {err&&<div style={{marginTop:10,padding:"8px 10px",borderRadius:8,fontSize:12,
          background:"#F8E7E3",color:"#9B3B2B"}}>{err}</div>}
        {parsed&&<div style={{marginTop:12}}>
          <DemandPreview rows={parsed}/>
          <div style={{fontSize:13,margin:"12px 0 8px"}}>
            読込 <b>{parsed.length}件</b>（うち形式エラー {parsed.filter(r=>r.err).length}件）。
            {clearFirst
              ?<>既存予約 {Object.keys(state.resMap).length}件をクリアしてから流し込む（セミ背骨・車両設定は維持）。</>
              :<>現在の確定予約 {Object.keys(state.resMap).length}件・稼働 {vehicles.filter(v=>v.active).length}台 に<b>追加</b>で流し込む。同じ需要を2回流すと二重計上になる。</>}
          </div>
          <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,margin:"0 0 8px",cursor:"pointer"}}>
            <input type="checkbox" checked={clearFirst} onChange={e=>setClearFirst(e.target.checked)}/>
            流し込み前に既存予約をクリア（背骨・車両設定・この需要データは維持）
          </label>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <button onClick={run} style={{...btnPrimary,width:"auto",padding:"9px 22px",background:"#2E9E6B"}}>
              自動確定を実行
            </button>
            <button onClick={()=>openExport(rowsToTable(parsed,false,vehicles),`reservations_${stamp()}.csv`)}
              style={{padding:"8px 14px",borderRadius:8,fontSize:12,cursor:"pointer",
                border:"1px solid #14202F",background:"#fff",color:"#14202F",fontWeight:700}}>
              この一覧をCSV出力
            </button>
            <span style={{fontSize:11,color:"#8A8474"}}>※流し込み前の入力一覧（再現用）。Excelで開ける（UTF-8）。</span>
          </div>

          {/* 車両規模シミュレーション */}
          <div style={{marginTop:14,padding:12,background:"#fff",border:"1px solid #D8D3C6",borderRadius:10}}>
            <div style={{fontSize:13,fontWeight:700,marginBottom:6}}>車両規模シミュレーション</div>
            <div style={{fontSize:11.5,color:"#6B6453",lineHeight:1.7,marginBottom:8}}>
              この需要一覧（{parsed.filter(r=>!r.err).length}件）を、車両構成を変えながら繰り返しさばいてみて、
              何台・何人乗りなら目標の成立率に届くかを調べる。試す構成は
              <b>台数（1〜6台）×定員（4・6・8・10・12名）＝30通り</b>。
              各構成について、空の状態から需要を上から順に自動確定し、成立率を記録する。
              運行時間は全車9:00〜17:00、その他は現在の設定（寄り道ST三段式・TW{P.tw/60}分・Dwell{P.dwell}秒）。
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              <span style={{fontSize:12}}>目標成立率</span>
              <input type="number" min={50} max={100} value={target}
                onChange={e=>setTarget(Math.max(50,Math.min(100,Number(e.target.value)||90)))}
                style={{width:54,padding:"5px 6px",border:"1px solid #C9C4B8",borderRadius:6,fontSize:13,
                  fontFamily:"'SF Mono','Consolas',monospace"}}/>
              <span style={{fontSize:12}}>％</span>
              <button onClick={()=>{
                const rows=parsed.filter(r=>!r.err);
                const caps=[4,6,8,10,12], maxV=6;
                const grid=[];
                for(let nv=1;nv<=maxV;nv++){
                  const row=[];
                  for(const cap of caps){
                    const fleet=Array.from({length:nv},(_,i)=>({id:i+1,name:`${i+1}号車`,
                      color:"#888",active:true,cap,start:9*3600,end:17*3600}));
                    const routes={}; for(let i=1;i<=nv;i++)routes[i]=[];
                    const rm={}; let okc=0;
                    for(const q of rows){
                      const drt=TT[q.o][q.d], mrt=mrtFromDRT(drt,P);
                      const dpt=q.mode==="dep"?q.t:q.t-(mrt+P.dwell);
                      const r={id:"S"+nv+"_"+cap+"_"+okc+"_"+Math.random().toString(36).slice(2,6),
                        num:0,o:q.o,d:q.d,dpt,drt,mrt,sa:q.sa,pax:q.pax,ipt:null,idt:null,vehicle:null};
                      const cs=searchInsertions(routes,rm,r,P,fleet);
                      if(cs.length){const c=cs[0];r.ipt=c.apt;r.idt=c.apt+P.dwell+mrt;r.vehicle=c.vehicle;
                        routes[c.vehicle]=c.route;rm[r.id]=r;okc++;}
                    }
                    row.push(Math.round(okc/rows.length*1000)/10);
                  }
                  grid.push(row);
                }
                // 推奨：目標を満たす最小構成（台数優先→定員）
                let rec=null;
                for(let i=0;i<grid.length&&!rec;i++)
                  for(let j=0;j<caps.length;j++)
                    if(grid[i][j]>=target){rec={nv:i+1,cap:caps[j],rate:grid[i][j]};break;}
                setSizing({grid,caps,maxV,rec,target,total:rows.length});
              }}
                style={{padding:"7px 16px",borderRadius:8,fontSize:12,cursor:"pointer",
                  border:"none",background:"#14202F",color:"#fff",fontWeight:700}}>
                30通りを試して必要な車両を調べる
              </button>
              <span style={{fontSize:10.5,color:"#8A8474"}}>※確定はしない（状態は変更されない）。100件で5秒程度、件数に比例して長くなる。</span>
            </div>

            {sizing&&<div style={{marginTop:10}}>
              <table style={{borderCollapse:"collapse",background:"#fff"}}>
                <thead><tr>
                  <th style={{padding:"4px 10px",border:"1px solid #C9C4B8",fontSize:11,background:"#F0EDE4"}}>成立率%</th>
                  {sizing.caps.map(c=><th key={c} style={{padding:"4px 10px",border:"1px solid #C9C4B8",
                    fontSize:11,background:"#F0EDE4"}}>定員{c}</th>)}
                </tr></thead>
                <tbody>
                  {sizing.grid.map((row,i)=>(
                    <tr key={i}>
                      <th style={{padding:"4px 10px",border:"1px solid #C9C4B8",fontSize:11,background:"#F0EDE4"}}>{i+1}台</th>
                      {row.map((rate,j)=>{
                        const hit=rate>=sizing.target;
                        const isRec=sizing.rec&&sizing.rec.nv===i+1&&sizing.rec.cap===sizing.caps[j];
                        return (
                        <td key={j} style={{padding:"4px 10px",border:isRec?"2.5px solid #14202F":"1px solid #DDD8CB",
                          fontSize:12.5,textAlign:"right",fontWeight:hit?700:400,
                          fontFamily:"'SF Mono','Consolas',monospace",
                          background:hit?"#DCEEDF":"#FBEAE6",color:hit?"#1E6B40":"#9B3B2B"}}>
                          {rate.toFixed(1)}
                        </td>);
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{fontSize:12.5,marginTop:8,fontWeight:700}}>
                {sizing.rec
                  ?`目標${sizing.target}％を満たす最小構成：${sizing.rec.nv}台 × 定員${sizing.rec.cap}名（成立率 ${sizing.rec.rate}％）`
                  :`6台×定員12でも目標${sizing.target}％に未達。需要の時間集中の平準化、ズレ幅（San）の拡大、運行時間の延長を検討。`}
              </div>
              <div style={{fontSize:11,color:"#8A8474",marginTop:4}}>
                緑＝目標達成／赤＝未達。太枠が推奨構成（台数優先で最小）。同一需要・先着順・空の状態からの計測値。
              </div>
              <div style={{fontSize:11,color:"#6B6453",marginTop:4,lineHeight:1.7}}>
                規模シミュレーションは構成探索のため、これ自体は運行に確定できない。ある構成を運行として見るには、
                「車両設定」で台数・定員をその構成に合わせてから、下の「自動確定」または全体最適化A/Bを実行する。
              </div>
            </div>}
          </div>

          {/* 全体最適化シミュレーション＝第2段（2方式） */}
          <div style={{marginTop:14,padding:12,background:"#fff",border:"1px solid #D8D3C6",borderRadius:10}}>
            <div style={{fontSize:13,fontWeight:700,marginBottom:6}}>全体最適化シミュレーション（2方式）</div>
            <div style={{fontSize:11.5,color:"#6B6453",lineHeight:1.7,marginBottom:8}}>
              現在の稼働車両（{vehicles.filter(v=>v.active).length}台）で、この需要を先着順固定（確定したら動かさない）でさばいた場合と、
              割付を組み直した場合を並べて出す。どちらも<b>約束（約束発IPT〜IPT+TW・約束着IDT）と定員・運行時間は厳守</b>。
              <br/><b>方式A・到着ごと（当日随時受付）</b>：新規が1件入るたびに、入らなければ確定済みの割付を組み直して救済を試す。
              オンライン処理は目先の救済が後続を潰すことがあるため、先着順を下回らない床つき。
              <br/><b>方式B・一括（前日予約向け）</b>：前日に受付済み＝発着を約束済みの予約全体を、約束を守ったまま割付を一括で組み直し、
              締まって空いた余地に当日相当の予約（先着順では入らなかった分）を詰める。
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <button onClick={()=>{
              const rows=parsed.filter(r=>!r.err);
              const fleet=vehicles.filter(v=>v.active);
              if(!fleet.length){setReopt({err:"稼働車両が0台。車両設定で有効化を。"});return;}
              if(!rows.length){setReopt({err:"有効な需要が0件。"});return;}
              setReopt({running:true});
              setTimeout(()=>{
                setReopt({...reoptimizeOnArrival(rows,P,fleet,6),mode:"online",fleet:fleet.length,total:rows.length});
              },30);
            }}
              style={{padding:"7px 16px",borderRadius:8,fontSize:12,cursor:"pointer",
                border:"none",background:"#7A4FA3",color:"#fff",fontWeight:700}}>
              A・到着ごとに最適化（当日随時）
            </button>
            <button onClick={()=>{
              const rows=parsed.filter(r=>!r.err);
              const fleet=vehicles.filter(v=>v.active);
              if(!fleet.length){setReopt({err:"稼働車両が0台。車両設定で有効化を。"});return;}
              if(!rows.length){setReopt({err:"有効な需要が0件。"});return;}
              setReopt({running:true});
              setTimeout(()=>{
                setReopt({...offlineOptimize(rows,P,fleet,12),mode:"batch",fleet:fleet.length,total:rows.length});
              },30);
            }}
              style={{padding:"7px 16px",borderRadius:8,fontSize:12,cursor:"pointer",
                border:"none",background:"#2E6B9E",color:"#fff",fontWeight:700}}>
              B・一括で最適化（前日予約）
            </button>
            </div>
            <span style={{fontSize:10.5,color:"#8A8474"}}>※確定はしない（状態は変更されない）。先着順固定より計算が重い。まず50〜80件で。</span>

            {reopt&&reopt.running&&<div style={{marginTop:10,fontSize:12,color:"#6B6453"}}>計算中…（件数が多いと数十秒かかることがある）</div>}
            {reopt&&reopt.err&&<div style={{marginTop:10,padding:"8px 10px",borderRadius:8,fontSize:12,background:"#F8E7E3",color:"#9B3B2B"}}>{reopt.err}</div>}
            {reopt&&reopt.reopt&&(()=>{
              const g=reopt.greedy,z=reopt.reopt,batch=reopt.mode==="batch";
              const col=batch?"#2E6B9E":"#7A4FA3";
              const dOk=z.ok-g.ok,dSpan=g.span-z.span;
              const per=(span,ok)=>ok?Math.round(span/60/ok):0;
              const card=(title,c,ok,total,span,extra)=>(
                <div style={{flex:1,minWidth:190,border:`1px solid ${c}`,borderRadius:10,padding:"10px 12px",background:"#fff"}}>
                  <div style={{fontSize:11,fontWeight:700,color:c,letterSpacing:1}}>{title}</div>
                  <div style={{fontSize:26,fontWeight:800,fontFamily:"'SF Mono','Consolas',monospace",marginTop:2}}>
                    {ok}<span style={{fontSize:15,color:"#6B6453"}}> / {total}</span>
                    <span style={{fontSize:14,color:"#6B6453",marginLeft:6}}>{Math.round(ok/total*100)}%</span>
                  </div>
                  <div style={{fontSize:11,color:"#6B6453",marginTop:2}}>総走行 {Math.round(span/60)}分（1件あたり{per(span,ok)}分）{extra}</div>
                </div>);
              return (
              <div style={{marginTop:12}}>
                <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                  {card("先着順固定（現行）","#8A8474",g.ok,g.total,g.span,"")}
                  {card(batch?"一括最適化（前日予約）":"全体最適化（到着ごと）",col,z.ok,z.total,z.span,
                    batch?`／当日相当を ${z.rescued}件 収容`:`／割付の入替 計${z.reassign}回`)}
                </div>
                <div style={{marginTop:10,fontSize:13,fontWeight:700}}>
                  差：成立 {dOk>=0?"+":""}{dOk}件　／　総走行 {dSpan>=0?"−":"+"}{Math.abs(Math.round(dSpan/60))}分
                  <span style={{fontWeight:400,fontSize:11.5,color:"#6B6453",marginLeft:8}}>
                    {batch?"（約束を全件守ったまま、走行をほぼ増やさず当日相当を多く受け入れられたかを見る。1件あたり走行が下がれば割付効率化の効果）"
                          :"（成立のプラス・総走行のマイナスが、割付を動かせることの利得）"}
                  </span>
                </div>
                {z.fellBack&&<div style={{fontSize:11.5,color:"#8A6D1F",background:"#FBF3DC",
                  border:"1px solid #E8D9A8",borderRadius:8,padding:"7px 10px",marginTop:6,lineHeight:1.7}}>
                  この需要では、到着ごとの割付シャッフルでは先着順固定を上回れなかった（救済して詰めると後続の成立を潰し、全体では損になるケース）。
                  そのため全体最適化は先着順固定と同じ結果を採用している（下回らないことを保証）。需要の集中度・車両構成によっては上回る。
                </div>}
                <div style={{fontSize:11,color:"#8A8474",marginTop:6,lineHeight:1.7}}>
                  {batch?"前日約束分は発着の約束をそのまま維持。割付を組み直して空いた余地に当日相当（先着順では入らなかった予約）を詰めている。表の「区分」が当日＝新たに収容できた分。"
                        :"「割付の入替」は、ある新規を受けた際に既存予約の担当便が変わった延べ回数。多いほど最適化が効く一方、実運用では運転手の予定変更コストになる。"}
                </div>
                {reopt.plan&&<div style={{marginTop:8}}>
                  {pendingCommit
                    ? <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",
                        background:"#FBF3DC",border:"1px solid #E8D9A8",borderRadius:8,padding:"8px 10px"}}>
                        <span style={{fontSize:11.5,color:"#8A6D1F"}}>
                          「{pendingCommit.label}」を運行に確定する。現在の確定予約 {Object.keys(state.resMap).length}件 はこの計画に置き換わる。
                        </span>
                        <button onClick={()=>commitPlan(pendingCommit.plan,pendingCommit.demandRec)}
                          style={{padding:"6px 14px",borderRadius:8,fontSize:12,cursor:"pointer",
                            border:"none",background:col,color:"#fff",fontWeight:700}}>確定する</button>
                        <button onClick={()=>setPendingCommit(null)}
                          style={{padding:"6px 12px",borderRadius:8,fontSize:12,cursor:"pointer",
                            border:"1px solid #C9C4B8",background:"#fff"}}>やめる</button>
                      </div>
                    : <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                        <button onClick={()=>{
                          const rrows=(reopt.reopt.rows||[]).map(r=>({o:r.o,d:r.d,mode:r.mode,t:r.t,
                            pax:r.pax,sa:r.sa,ok:!!r.ok,reason:r.reason,num:r.num,vehicle:r.vehicle,apt:r.apt,idt:r.idt}));
                          const demandRec={...(demandMeta||{source:"unknown"}),
                            committedBy:batch?"一括最適化で確定":"全体最適化で確定",when:Date.now(),
                            okRate:reopt.reopt.total?reopt.reopt.ok/reopt.reopt.total:null,rows:rrows};
                          setPendingCommit({plan:reopt.plan,demandRec,
                            label:batch?"一括最適化（前日予約）":"全体最適化（到着ごと）"});
                        }}
                          style={{padding:"8px 16px",borderRadius:8,fontSize:12,cursor:"pointer",
                            border:"none",background:col,color:"#fff",fontWeight:700}}>
                          この結果を運行に確定して見る
                        </button>
                        <span style={{fontSize:10.5,color:"#9B3B2B"}}>
                          ※現在の確定予約をこの計画に置き換える。確定後は運行盤・ダイヤ・確認表・分析に反映される。
                        </span>
                      </div>}
                </div>}
                <div style={{overflow:"auto",maxHeight:"40vh",border:"1px solid #B9B2A1",borderRadius:8,marginTop:8}}>
                  <table style={{borderCollapse:"collapse",background:"#fff",width:"100%"}}>
                    <thead><tr style={{background:"#F0EDE4"}}>
                      <th style={th}>受付</th><th style={th}>区間</th><th style={th}>希望</th>
                      <th style={th}>結果</th><th style={th}>便・約束</th><th style={th}>ずれ</th>
                      <th style={th}>{batch?"区分":"既存入替"}</th>
                    </tr></thead>
                    <tbody>
                      {z.rows.map((r,i)=>(
                        <tr key={i} style={{background:r.ok?undefined:"#F8E7E3"}}>
                          <td style={td}>{i+1}</td>
                          <td style={td}>{r.o!=null&&r.d!=null?`${STOPS[r.o]}→${STOPS[r.d]}`:"—"}</td>
                          <td style={td}>{r.t!=null?`${r.mode==="dep"?"発":"着"} ${fmt(r.t)}`:"—"}</td>
                          <td style={{...td,fontWeight:700,color:r.ok?col:"#C0392B"}}>{r.ok?"成立":"不成立"}</td>
                          <td style={{...td,textAlign:"left"}}>{r.ok?`${vn(r.vehicle)} 約束発${fmt(r.apt)} 〜着${fmt(r.idt)}`:r.reason}</td>
                          <td style={td}>{r.ok?fmtDev(r.dev):"—"}</td>
                          <td style={td}>{batch?(r.ok?(r.rescued?"当日":"前日"):"—"):(r.ok?(r.reassigned>0?`${r.reassigned}件`:"—"):"—")}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>);
            })()}
          </div>
        </div>}
      </>}

      {exp&&<div style={{marginTop:12,padding:12,background:"#fff",border:"2px solid #14202F",borderRadius:10}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8,flexWrap:"wrap"}}>
          <b style={{fontSize:13}}>{exp.name}</b>
          {exp.tried===false&&<span style={{fontSize:11,color:"#9B3B2B"}}>
            この環境では自動ダウンロードが制限されている。下のコピーで取得を。</span>}
          {exp.tried===true&&<span style={{fontSize:11,color:"#8A8474"}}>
            ダウンロードを開始した（保存されない場合は下のコピーで取得を）。</span>}
          <button onClick={copyExp} style={{padding:"6px 14px",borderRadius:8,fontSize:12,cursor:"pointer",
            border:"none",background:exp.copied?"#2E9E6B":"#14202F",color:"#fff",fontWeight:700}}>
            {exp.copied?"コピー済み ✓":"クリップボードにコピー"}
          </button>
          <button onClick={()=>setExp(null)} style={{padding:"6px 12px",borderRadius:8,fontSize:12,cursor:"pointer",
            border:"1px solid #C9C4B8",background:"#fff"}}>閉じる</button>
        </div>
        <textarea ref={expRef} readOnly value={exp.csv}
          style={{width:"100%",height:130,fontSize:11,fontFamily:"'SF Mono','Consolas',monospace",
            border:"1px solid #D8D3C6",borderRadius:6,padding:8,boxSizing:"border-box",
            whiteSpace:"pre",resize:"vertical"}}/>
        <div style={{fontSize:11,color:"#8A8474",marginTop:4}}>
          コピーした内容をメモ帳等に貼り付けて「{exp.name}」として保存すればExcelで開ける。
          そのまま本画面のファイル読込にも使える（.txt/.csv）。
        </div>
      </div>}

      {result&&<>
        <div style={{display:"flex",alignItems:"baseline",gap:14,margin:"4px 0 12px"}}>
          <div style={{fontSize:34,fontWeight:800,
            fontFamily:"'SF Mono','Consolas',monospace"}}>
            予約成立 <span style={{color:"#2E9E6B"}}>{result.ok}</span>
            <span style={{fontSize:22,color:"#6B6453"}}> / {result.total}</span>
          </div>
          <div style={{fontSize:13,color:"#6B6453"}}>
            成立率 {Math.round(result.ok/result.total*100)}％ ／ 不成立 {result.total-result.ok}件
          </div>
        </div>
        <div style={{overflow:"auto",maxHeight:"56vh",border:"1px solid #B9B2A1"}}>
          <table style={{borderCollapse:"collapse",background:"#fff",width:"100%"}}>
            <thead><tr style={{background:"#F0EDE4"}}>
              <th style={th}>行</th><th style={th}>区間</th><th style={th}>希望</th>
              <th style={th}>人数</th><th style={th}>ズレ幅</th><th style={th}>結果</th>
              <th style={th}>便・約束</th><th style={th}>ずれ</th>
            </tr></thead>
            <tbody>
              {result.rows.map(r=>(
                <tr key={r.line} style={{background:r.ok?undefined:"#F8E7E3"}}>
                  <td style={td}>{r.line}</td>
                  <td style={td}>{r.o!=null&&r.d!=null?`${STOPS[r.o]}→${STOPS[r.d]}`:"—"}</td>
                  <td style={td}>{r.t!=null?`${r.mode==="dep"?"発":"着"} ${fmt(r.t)}`:"—"}</td>
                  <td style={td}>{r.pax}名</td>
                  <td style={td}>±{Math.round(r.sa/60)}分</td>
                  <td style={{...td,fontWeight:700,color:r.ok?"#2E9E6B":"#C0392B"}}>{r.ok?"成立":"不成立"}</td>
                  <td style={{...td,textAlign:"left"}}>
                    {r.ok?`#${r.num} ${vn(r.vehicle)} 約束発${fmt(r.apt)} 約束着〜${fmt(r.idt)}（車内最大${r.maxLoad}/${r.cap}名）`
                        :r.reason}
                  </td>
                  <td style={td}>{r.ok?fmtDev(r.dev):"—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
          <button onClick={()=>openExport(rowsToTable(result.rows,false,vehicles),`reservations_${stamp()}.csv`)}
            style={{padding:"8px 14px",borderRadius:8,fontSize:12,cursor:"pointer",
              border:"1px solid #14202F",background:"#fff",color:"#14202F",fontWeight:700}}>
            入力一覧をCSV出力（再現用）
          </button>
          <button onClick={()=>openExport(rowsToTable(result.rows,true,vehicles),`reservations_result_${stamp()}.csv`)}
            style={{padding:"8px 14px",borderRadius:8,fontSize:12,cursor:"pointer",
              border:"1px solid #14202F",background:"#fff",color:"#14202F",fontWeight:700}}>
            結果付きCSV出力
          </button>
          <button onClick={onShowAnalysis}
            style={{padding:"8px 14px",borderRadius:8,fontSize:12,cursor:"pointer",
              border:"none",background:"#2E9E6B",color:"#fff",fontWeight:700}}>
            運行分析を表示
          </button>
        </div>
        <div style={{fontSize:11,color:"#8A8474",marginTop:8}}>
          成立分は確定済みとして運行盤・ダイヤ・確認表に反映済み。不成立行はズレ幅の拡大・時刻変更・車両追加（車両設定）で再受付できる。
          「入力一覧」は流し込み前の形式そのままなので、再度読み込めば同条件を再現できる（先着順・同一車両設定の場合）。
        </div>
      </>}
    </div>
  </div>);
}

/* ---------- 運行分析 ----------
   時間の内訳は予約数ベースで区分：
   ・実車時間＝1予約以上を乗せて拘束されていた時間
   ・乗合時間＝2予約以上が同乗していた時間（同一予約の複数名は乗合に数えない）
   ・乗合率＝乗合時間÷実車時間
   走行距離は移動時間×25km/h（本マトリクスの生成速度）による換算値。 */
const SPEED_KMH=25;

/* ---------- 運行改善アドバイス（ルールベース。将来、未消費の寄り道時間をAIが最適化） ----------
   analyzeOpsのtotalと、任意で流し込み結果（lastDemand）から、しきい値で発火する助言を返す。
   機械学習ではなく決定論的ルール。提案・運用での読み筋を自動で言語化するだけ。 */
/* ---------- ボトルネック診断 ----------
   効率が上がらない主因を1つに絞る。判定木：
   ・空車率が過半 → A（回送が長い＝空間的に散る）／A2（待機が長い＝台数過剰・需要薄）
   ・実車はしているが乗合が薄い → C（寄り道許容に余力＝設定が保守的）／B（許容使い切り＝時間的に重ならない構造）
   ・いずれも小さい → 良好（これ以上は利用者負担とのトレードオフ）
   既存のanalyzeOps集計だけで算出。再シミュレーション不要。 */
function diagnoseBottleneck(total){
  const span=total.span||1;
  const loaded=total.loadedTime;
  const empty=total.emptyMove;
  const idle=Math.max(0, span - loaded - empty);
  const lr=loaded/span;                                    // 実車率
  const sr=loaded>0?total.sharedTime/loaded:0;             // 乗合率
  const str=total.stAllow>0?total.stUsed/total.stAllow:0;  // 寄り道消費率
  const empShare=empty/span, idleShare=idle/span;
  const p=x=>Math.round(x*100);
  let main=null, sub=null;

  if((1-lr) > 0.55){
    if(empty >= idle){
      main={key:"A", label:"需要が空間的に散っている（拾いに行く回送が長い）",
        detail:`拘束時間の${p(empShare)}%が空車回送。乗客のいる場所まで走る距離が長く、実車率は${p(lr)}%に留まる。`,
        action:"停留所配置の見直し、拠点への需要集約、または運行エリアの絞り込み。台数を増やしても回送は減りにくい。"};
      sub={label:"空車待機",detail:`空車で停まっている時間も${p(idleShare)}%ある。`};
    }else{
      main={key:"A2", label:"車両が呼ばれず待っている（需要に対し台数が多い）",
        detail:`拘束時間の${p(idleShare)}%が空車待機。実車率は${p(lr)}%。需要が時間的に薄く車両が遊んでいる。`,
        action:"稼働台数を減らす、または運行時間帯を需要ピークに寄せる。車両規模シミュレーションで最小台数を当たり直す。"};
      sub={label:"空車回送",detail:`回送も${p(empShare)}%ある。`};
    }
  }else if(sr < 0.25){
    if(str < 0.5){
      main={key:"C", label:"相乗りの余力を使い切れていない（設定が保守的）",
        detail:`実車率${p(lr)}%は確保できているが乗合率は${p(sr)}%。寄り道許容の消費は${p(str)}%で、まだ${p(1-str)}%の余力が残る。`,
        action:"San（希望時刻のズレ許容）を広げる、寄り道STの許容を増やす。余っている許容を相乗りに変換できる余地が大きい。"};
      sub={label:"寄り道余力",detail:`許容の${p(1-str)}%が未使用。`};
    }else{
      main={key:"B", label:"需要が時間的に重ならず束ねられない（構造的）",
        detail:`乗合率${p(sr)}%と低いが寄り道許容は${p(str)}%まで使い切っている。許容を広げても相乗りにならない＝同時間帯・同方向の需要が薄い。`,
        action:"設定調整では改善しにくい。受付時間帯の集約（時間帯を絞って便を立てる）、需要そのものの喚起が必要。"};
      sub={label:"実車率",detail:`実車率は${p(lr)}%。`};
    }
  }else{
    main={key:"OK", label:"主要なロスは小さい",
      detail:`実車率${p(lr)}%・乗合率${p(sr)}%・寄り道消費${p(str)}%。空車も相乗り不足も目立たない。`,
      action:"これ以上の効率化は利用者の負担（乗車時間の伸び）とのトレードオフになる。"};
    sub=null;
  }
  return {main, sub, breakdown:{loaded, empty, idle, span}};
}

function buildAdvice(total,P,lastDemand){
  if(!total||total.resCount===0)return [];
  const out=[];
  const ratio=(a,b)=>b>0?a/b:0;
  const loadRate=ratio(total.loadedTime,total.span);          // 実車率
  const shareRate=ratio(total.sharedTime,total.loadedTime);   // 乗合率
  const stRate=ratio(total.stUsed,total.stAllow);             // 寄り道消費率
  const stretch=total.resCount?total.stretchSum/total.resCount:0; // 平均倍率
  const occLoaded=ratio(total.paxSec,total.loadedTime);       // 実車時間あたり平均乗車人数
  const okRate=lastDemand&&typeof lastDemand.okRate==="number"?lastDemand.okRate:null;

  if(okRate!=null&&okRate<0.85)
    out.push({level:"warn",text:`成立率 ${Math.round(okRate*100)}% と低い。車両規模シミュレーションで必要な台数・定員を当たり直すとよい。`});
  if(loadRate<0.4)
    out.push({level:"warn",text:`実車率 ${Math.round(loadRate*100)}%（拘束時間に対し人を乗せた時間が短い）。空車回送が多い＝需要が空間的に散っている。停留所配置か台数配分の見直し余地。`});
  if(shareRate<0.2)
    out.push({level:"info",text:`乗合率 ${Math.round(shareRate*100)}% と低い。San（ズレ幅）が狭く相乗りの機会を逃しているか、需要が時間的に重なっていない。Sanを広げると改善しうる。`});
  if(stRate<0.4&&shareRate<0.3)
    out.push({level:"info",text:`寄り道時間の消費 ${Math.round(stRate*100)}% で余力が残るのに乗合率が低い。許容（寄り道枠）はあるが相乗りに変換できていない。San・定員・時間帯の重なりを確認。`});
  if(stretch>1.6)
    out.push({level:"info",text:`乗車時間の伸び（平均倍率）${stretch.toFixed(2)}倍。短距離客の負担が大きい。寄り道ST三段式の短距離側（最小ゆとり）を下げる検討。`});
  if(occLoaded>0&&occLoaded<1.4)
    out.push({level:"info",text:`実車時間あたりの平均乗車人数 ${occLoaded.toFixed(2)}人。相乗りが薄い。需要を時間的に束ねる（受付時間帯の集約・San拡大）と密度が上がる。`});
  if(stRate>0.7&&shareRate>0.4)
    out.push({level:"good",text:`寄り道消費 ${Math.round(stRate*100)}%・乗合率 ${Math.round(shareRate*100)}%。許容範囲を使い切って効率化できている。これ以上は利用者体験とのトレードオフ。`});
  if(out.length===0)
    out.push({level:"good",text:"主要指標に大きな偏りなし。現状の設定で需要を捌けている。"});
  return out;
}

/* ---------- 確定予約の個票データ（DRT実績。個別比較用） ----------
   analyzeOpsの車両別集計と同じロジックを、予約1件ごとの実績として取り出す。
   不成立分は実績が存在しないため対象外（確定予約resMapのみが母集団）。 */
function drtPassengerDetail(vehicles,sims,resMap){
  const out=[];
  for(const v of vehicles.filter(v=>v.active)){
    const sim=sims[v.id];
    if(!sim.ok||sim.events.length===0)continue;
    const ev=sim.events;
    const myRes=Object.values(resMap).filter(r=>r.vehicle===v.id);
    for(const r of myRes){
      const o=ev.find(e=>e.resId===r.id&&e.type==="O");
      const d=ev.find(e=>e.resId===r.id&&e.type==="D");
      if(!o||!d)continue;
      const ride=d.adt-o.etd;
      const dev=o.apt-r.dpt;
      out.push({id:r.id,o:r.o,d:r.d,pax:r.pax??1,dev,devAbs:Math.abs(dev),ride,direct:r.drt});
    }
  }
  return out;
}

function analyzeOps(vehicles,sims,resMap,P){
  const per=[];
  for(const v of vehicles.filter(v=>v.active)){
    const sim=sims[v.id];
    if(!sim.ok||sim.events.length===0){
      // 稼働中だが予約が乗っていない車両も、列として表示する（0埋め）
      per.push({v,used:true,empty:true,span:0,moveTime:0,emptyMove:0,loadedTime:0,sharedTime:0,
        paxSec:0,resCount:0,paxCarried:0,stUsed:0,stAllow:0,
        devSum:0,devAbsSum:0,rideSum:0,directSum:0,stretchSum:0,onTime:0,distKm:0});
      continue;
    }
    const ev=sim.events;
    const span=ev[ev.length-1].etd-ev[0].eta;
    let moveTime=0,emptyMove=0,loadedTime=0,sharedTime=0,paxSec=0;
    // イベント間の区分ごとに「同乗中の予約数」「乗車人数」を集計
    const onboard=new Set();
    const tm=e=>e.type==="O"?e.apt:(e.type==="D"?e.adt:e.apt); // アンカーはapt(=杭時刻)
    for(let i=0;i<ev.length;i++){
      const e=ev[i];
      if(e.type==="O")onboard.add(e.resId);else if(e.type==="D")onboard.delete(e.resId);
      const t0=tm(e);
      const t1=i+1<ev.length?tm(ev[i+1]):e.etd;
      const dt=Math.max(0,t1-t0);
      const nRes=onboard.size, load=e.load;
      if(nRes>=1)loadedTime+=dt;
      if(nRes>=2)sharedTime+=dt;
      paxSec+=load*dt;
      // 移動時間（このイベントのetd→次のeta）
      if(i+1<ev.length){
        const mv=Math.max(0,ev[i+1].eta-e.etd);
        moveTime+=mv;
        if(nRes===0)emptyMove+=mv;
      }
    }
    // 予約単位の集計（寄り道消費）
    const myRes=Object.values(resMap).filter(r=>r.vehicle===v.id);
    let resCount=0,paxCarried=0,stUsed=0,stAllow=0;
    let devSum=0,devAbsSum=0,rideSum=0,directSum=0,stretchSum=0,onTime=0;
    for(const r of myRes){
      const o=ev.find(e=>e.resId===r.id&&e.type==="O");
      const d=ev.find(e=>e.resId===r.id&&e.type==="D");
      if(!o||!d)continue;
      resCount++;paxCarried+=r.pax??1;
      const ride=d.adt-o.etd;                      // 実乗車時間
      stUsed+=Math.max(0,ride-r.drt);
      stAllow+=r.mrt-r.drt;
      // 希望時刻とのズレ（約束乗車−希望乗車）。発希望・着希望でdptは希望乗車に正規化済み
      const dev=o.apt-r.dpt;
      devSum+=dev; devAbsSum+=Math.abs(dev);
      if(Math.abs(dev)<=60)onTime++;              // ±1分以内を定時とみなす
      // 乗車時間の伸び率
      rideSum+=ride; directSum+=r.drt;
      stretchSum+=ride/Math.max(1,r.drt);
    }
    per.push({v,used:true,span,moveTime,emptyMove,loadedTime,sharedTime,
      paxSec,resCount,paxCarried,stUsed,stAllow,
      devSum,devAbsSum,rideSum,directSum,stretchSum,onTime,
      distKm:moveTime/3600*SPEED_KMH});
  }
  const used=per.filter(p=>p.used);
  const sum=k=>used.reduce((a,p)=>a+p[k],0);
  const total={n:used.length,span:sum("span"),moveTime:sum("moveTime"),
    emptyMove:sum("emptyMove"),loadedTime:sum("loadedTime"),sharedTime:sum("sharedTime"),
    paxSec:sum("paxSec"),resCount:sum("resCount"),paxCarried:sum("paxCarried"),
    stUsed:sum("stUsed"),stAllow:sum("stAllow"),distKm:sum("distKm"),
    devSum:sum("devSum"),devAbsSum:sum("devAbsSum"),rideSum:sum("rideSum"),
    directSum:sum("directSum"),stretchSum:sum("stretchSum"),onTime:sum("onTime")};
  return {per,total};
}

/* ---------- 運行方式の適性判定（A＋：需要構造からの3択判定＋粗い反実仮想） ----------
   DRT／セミデマンド／定時定路線 のどれが向くかを、需要構造の3軸から判定する。
   軸1 需要密度   ＝ 総需要件数 / 稼働台時（1台1時間で何件の需要があるか）
   軸2 空間集中度 ＝ HUB絡みシェア と 上位コリドーシェア の高い方（路線が引ける度合い）
   軸3 時間集中度 ＝ 最大の1時間帯が占める需要シェア（ピークの尖り）
   判定は 密度×空間集中 の2軸マトリクス＋時間集中の補正。
   反実仮想：上位コリドーを1本の定時路線とみなし、カバー率・便あたり乗客を概算。
   路線は需要上位のODから機械的に引く（人が選ばない＝恣意性を殺す）。前提は画面に明示。
   すべて既存の需要データ・analyzeOps集計から算出。再シミュレーション不要。 */
function judgeModality(demandRows, vehTotal, params){
  const rows=demandRows||[];
  const demandTotal=rows.length;
  const served=rows.filter(r=>r.ok).length;
  const failRate=demandTotal>0?(demandTotal-served)/demandTotal:0;
  const vehHours=params.vehHours||0;
  const dens=vehHours>0?demandTotal/vehHours:0;              // 件/台時

  const hubSet=new Set(HUBS);
  let hubTouch=0; const odCount={}, hourB={};
  for(const r of rows){
    if(hubSet.has(r.o)||hubSet.has(r.d))hubTouch++;
    const key=r.o<r.d?`${r.o}-${r.d}`:`${r.d}-${r.o}`;
    odCount[key]=(odCount[key]||0)+1;
    const h=Math.floor((r.t||0)/3600); hourB[h]=(hourB[h]||0)+1;
  }
  const hubShare=demandTotal>0?hubTouch/demandTotal:0;
  const odPairs=Object.entries(odCount).sort((a,b)=>b[1]-a[1]);
  const topShare=demandTotal>0?odPairs.slice(0,3).reduce((s,e)=>s+e[1],0)/demandTotal:0;
  const spatial=Math.max(hubShare,topShare);                 // 空間集中（拠点集中とコリドー集中の高い方）
  const hourVals=Object.values(hourB);
  const peakHourN=hourVals.length?Math.max(...hourVals):0;
  const temporal=demandTotal>0?peakHourN/demandTotal:0;      // 時間集中（最大1時間帯シェア）

  // 反実仮想：定時路線＝需要接触の多い上位K停留所を1本に連ねた回廊とみなす。
  // その路線上の停留所どうし（発着の両方が路線上）の需要をカバー可能とする。
  // K＝路線が停まる停留所数（現実の路線規模パラメータ。機械的に需要上位から採る）。
  const touch={};
  for(const r of rows){ touch[r.o]=(touch[r.o]||0)+1; touch[r.d]=(touch[r.d]||0)+1; }
  const rankedStops=Object.entries(touch).sort((a,b)=>b[1]-a[1]).map(e=>Number(e[0]));
  const K=Math.max(2,params.lineStopN||Math.ceil((STOPS.length||10)*0.4));
  let sel=rankedStops.slice(0,K);
  // 停留所を需要軸（最も離れた2点を結ぶ方向）へ射影し、路線として順序付ける。
  // 地図のコリドー軸と同じ考え方＝停留所の集合を「1本の路線」として並べる。
  if(sel.length>=2){
    let A=sel[0],B=sel[1],best=-1;
    for(let i=0;i<sel.length;i++)for(let j=i+1;j<sel.length;j++){
      const d=Math.hypot(POS[sel[i]][0]-POS[sel[j]][0],POS[sel[i]][1]-POS[sel[j]][1]);
      if(d>best){best=d;A=sel[i];B=sel[j];}
    }
    const ax=POS[B][0]-POS[A][0],ay=POS[B][1]-POS[A][1];
    const proj=s=>(POS[s][0]-POS[A][0])*ax+(POS[s][1]-POS[A][1])*ay;
    sel=sel.slice().sort((p,q)=>proj(p)-proj(q));
  }
  const lineOrder=sel;                       // 路線順（始発→終点）
  const lineStops=new Set(sel);
  const posInLine={}; lineOrder.forEach((s,i)=>{posInLine[s]=i;});
  const dwellSec=params.dwell??60;            // 中間停留所1か所あたりの停車時間

  // 定時路線に乗った場合の乗車時間＝始点から終点まで、路線上の各区間の直行時間（TT）を積算
  // ＋通過する中間停留所ごとの停車（dwell）。これがDRT直行（TT[o][d]）に対する伸びになる。
  let covered=0,coveredPax=0,stretchSum=0,stretchN=0;
  for(const r of rows){
    if(!(lineStops.has(r.o)&&lineStops.has(r.d)))continue;
    covered++; coveredPax+=(r.pax||1);
    const io=posInLine[r.o], id=posInLine[r.d];
    const lo=Math.min(io,id), hi=Math.max(io,id);
    let rideT=0;
    for(let i=lo;i<hi;i++) rideT+=TT[lineOrder[i]][lineOrder[i+1]];
    rideT += Math.max(0,hi-lo-1)*dwellSec;    // 通過する中間停留所の停車時間
    const direct=Math.max(1,TT[r.o][r.d]);
    stretchSum += rideT/direct; stretchN++;
  }
  const coverRate=demandTotal>0?covered/demandTotal:0;
  const fixedStretch=stretchN>0?stretchSum/stretchN:1;   // 定時路線に乗った場合の平均乗車時間の伸び倍率
  const opHours=params.opHours||8;
  const trips=params.headwayMin>0?Math.floor(opHours*60/params.headwayMin):0;  // 片方向便数
  const paxPerTrip=trips>0?coveredPax/(trips*2):0;           // 往復=trips×2便で割る
  const waitAvgMin=(params.headwayMin||0)/2;                 // 平均待ち時間（ランダム到着想定）
  const waitMaxMin=params.headwayMin||0;                     // 最大待ち時間（発車直後に来た場合）
  // 路線名は主要停留所（HUB優先→需要上位）を数個だけ表示
  const hubOnLine=[...lineStops].filter(i=>hubSet.has(i));
  const labelStops=(hubOnLine.length?hubOnLine:rankedStops.slice(0,3)).slice(0,4);
  const corridorNames=labelStops.map(i=>STOPS[i]).join("・")+(lineStops.size>labelStops.length?` ほか計${lineStops.size}停留所`:"");

  // 個別比較：確定予約（実績あり）のうち、この路線でカバーされる人だけを対象に
  // DRT実績（乗車時間・希望とのズレ）と定時試算（乗車時間・待ち時間）を1対1で比べる。
  // 母集団は「不成立込みの全需要」ではなく「実際に運んだ確定予約」に限定する
  // （不成立分は実績が存在しないため比較できない）。
  const drtDetail=params.drtDetail||[];
  const onLine=drtDetail.filter(r=>lineStops.has(r.o)&&lineStops.has(r.d));
  let compare=null;
  if(onLine.length>0){
    const n=onLine.length, paxN=onLine.reduce((s,r)=>s+r.pax,0);
    const drtRideAvg=onLine.reduce((s,r)=>s+r.ride,0)/n;
    const drtDevAvg=onLine.reduce((s,r)=>s+r.devAbs,0)/n;
    let fxSum=0;
    for(const r of onLine){
      const io=posInLine[r.o], id=posInLine[r.d];
      const lo=Math.min(io,id), hi=Math.max(io,id);
      let rideT=0; for(let i=lo;i<hi;i++) rideT+=TT[lineOrder[i]][lineOrder[i+1]];
      rideT += Math.max(0,hi-lo-1)*dwellSec;
      fxSum+=rideT;
    }
    const fxRideAvg=fxSum/n;
    compare={n,paxN,drtTotalN:drtDetail.length,
      drtRideAvgMin:drtRideAvg/60, fxRideAvgMin:fxRideAvg/60,
      drtDevAvgMin:drtDevAvg/60, fxWaitAvgMin:(params.headwayMin||0)/2, fxWaitMaxMin:params.headwayMin||0};
  }
  // 走行距離：DRTは現運用の総走行距離（全予約分・回送含む）。定時は候補路線を往復×便数、
  // 需要の有無に関わらず走る固定距離。母集団も性質も異なる「システム全体のコスト」比較として示す。
  const lineKm=(()=>{ let s=0; for(let i=0;i<lineOrder.length-1;i++) s+=TT[lineOrder[i]][lineOrder[i+1]]; return s/3600*SPEED_KMH; })();
  const fixedDailyKm=lineKm*2*trips;
  const drtDailyKm=(vehTotal||{}).distKm||0;

  // 現DRT実績（対比の左側）
  const t=vehTotal||{};
  const drtPaxPerVehHour=vehHours>0?served/vehHours:0;
  const drtLoadRate=t.span>0?t.loadedTime/t.span:0;
  const drtShareRate=t.loadedTime>0?t.sharedTime/t.loadedTime:0;

  // 判定マトリクス（閾値は仮置き＝現場キャリブレーション前提）
  const HI_D=3.0, HI_S=0.5, HI_T=0.30;
  const dHi=dens>=HI_D, sHi=spatial>=HI_S, tHi=temporal>=HI_T;
  const p=x=>Math.round(x*100);
  let v;
  if(dHi&&sHi){
    v={key:"fixed",label:"定時定路線／コミュニティバス",
      reason:`需要密度が高く（${dens.toFixed(1)}件/台時）、${p(spatial)}%が特定コリドー・拠点に集中している。決まった路線に需要が乗るため、定時便の方が1便あたり乗客を確保でき、DRTの相乗り調整コストが不要。`,
      action:"主要コリドーに定時路線を敷き、DRTは路線を補完する端末アクセス（ラストマイル）に限定する構成が有力。"};
  }else if(dHi&&!sHi){
    v={key:"full",label:"フルデマンドDRT（飽和に注意）",
      reason:`需要密度は高い（${dens.toFixed(1)}件/台時）が、起終点が多方向に散っている（コリドー集中${p(spatial)}%）。路線を引いても取りこぼしが多く、多対多を捌けるDRTが向く。`,
      action:failRate>=0.15?`ただし不成立が${p(failRate)}%と高く飽和気味。ゾーン分割（エリアを割って車両専属化）か増車・増定員を検討。`:"現状の多対多需要はDRTが適合。密度が上がり続けるならゾーン分割を視野に。"};
  }else if(!dHi&&sHi){
    v={key:"semi",label:"セミデマンド（背骨型）",
      reason:`需要は薄い（${dens.toFixed(1)}件/台時）が、${p(spatial)}%がコリドー・拠点に集中している。背骨（定時ルート）を1本通し、端をデマンドで拾う中間形態が効率的。`,
      action:"上位コリドーを背骨に設定し、外れる需要のみデマンド逸脱で拾う。実装済みのセミデマンド運行が適合する。"};
  }else{
    v={key:"full",label:"フルデマンドDRT",
      reason:`需要が薄く（${dens.toFixed(1)}件/台時）、起終点も散っている（コリドー集中${p(spatial)}%）。定時便は空気を運ぶことになり、必要な時だけ動くDRTが適合。`,
      action:"フルデマンドが妥当。さらに薄いなら個別運行（タクシー的）や需要喚起（受付時間帯の集約）も選択肢。"};
  }
  if(tHi&&(v.key==="full"||v.key==="semi"))
    v.tempNote=`時間帯の集中が高い（ピーク1時間に${p(temporal)}%）。ピーク帯だけ定時便を増発する併用も検討余地。`;

  return {demandTotal,served,failRate,vehHours,dens,spatial,hubShare,topShare,temporal,
    verdict:v,thresholds:{HI_D,HI_S,HI_T},
    cf:{coverRate,paxPerTrip,trips,corridorNames,drtPaxPerVehHour,drtLoadRate,drtShareRate,
      fixedStretch,waitAvgMin,waitMaxMin,dwellSec},
    compare, distance:{lineKm,fixedDailyKm,drtDailyKm}};
}

function SuitabilityView({demandRows,vehicles,total,demandSrc,dwell,drtDetail}){
  const [headway,setHeadway]=React.useState(30);
  const defK=Math.max(2,Math.ceil((STOPS.length||10)*0.4));
  const [lineK,setLineK]=React.useState(defK);
  const activeV=vehicles.filter(v=>v.active);
  const vehHours=activeV.reduce((s,v)=>s+(v.end-v.start),0)/3600;
  const opHours=activeV.reduce((mx,v)=>Math.max(mx,(v.end-v.start)/3600),0)||8;
  if(!demandRows||demandRows.length===0)
    return (
    <div style={{marginTop:16,fontSize:11.5,color:"#8A8474",background:"#fff",
      border:"1px dashed #D8D3C6",borderRadius:10,padding:"10px 12px"}}>
      運行方式の適性判定：判定に使う需要がない。予約を確定するか需要を流し込むと表示される。
    </div>);
  const J=judgeModality(demandRows,total,{vehHours,opHours,headwayMin:headway,lineStopN:lineK,dwell,drtDetail});
  const pc=x=>`${Math.round(x*100)}%`;
  const vc = J.verdict.key==="fixed"?{bg:"#EAF0F7",bd:"#C4D3E6",fg:"#2C4A6E",mk:"定時向き"}
           : J.verdict.key==="semi" ?{bg:"#F3EEF7",bd:"#D8C9E4",fg:"#5B3F73",mk:"中間"}
           :                         {bg:"#EAF3EC",bd:"#C4DFC9",fg:"#2C6E3C",mk:"DRT向き"};
  const meter=(label,val,scale,hi,unitTxt)=>{
    const w=Math.max(0,Math.min(100,val/scale*100));
    const hiPos=Math.max(0,Math.min(100,hi/scale*100));
    const isHi=val>=hi;
    return (
    <div style={{marginBottom:8}}>
      <div style={{display:"flex",fontSize:11,marginBottom:3,alignItems:"baseline"}}>
        <span style={{color:"#3A3526",fontWeight:700}}>{label}</span>
        <span style={{marginLeft:"auto",fontFamily:"'SF Mono','Consolas',monospace",color:"#3A3526"}}>{unitTxt}</span>
        <span style={{marginLeft:8,fontSize:10,fontWeight:700,color:isHi?"#2C4A6E":"#8A8474",width:20,textAlign:"right"}}>{isHi?"高":"低"}</span>
      </div>
      <div style={{height:10,background:"#EDEAE0",borderRadius:5,overflow:"hidden",position:"relative"}}>
        <div style={{width:`${w}%`,height:"100%",background:isHi?"#5B7DA8":"#B7AE9B"}}/>
        <div style={{position:"absolute",left:`${hiPos}%`,top:-1,bottom:-1,width:1.5,background:"#9B3B2B"}} title="判定しきい値"/>
      </div>
    </div>);
  };
  const inS={width:52,padding:"3px 5px",border:"1px solid #C9C4B8",borderRadius:5,fontSize:12,
    fontFamily:"'SF Mono','Consolas',monospace",textAlign:"right"};
  return (
  <div style={{marginTop:16,border:`1px solid ${vc.bd}`,borderRadius:10,overflow:"hidden"}}>
    <div style={{background:"#14202F",color:"#E8E4DA",fontSize:12,fontWeight:700,padding:"7px 12px",letterSpacing:1}}>
      運行方式の適性判定
      <span style={{fontWeight:400,fontSize:10.5,color:"#9AA7BA",marginLeft:8}}>そもそもDRTが適正か／定時定路線が向くか（この需要パターン下での試算）</span>
    </div>
    <div style={{background:"#fff",padding:"12px 14px"}}>

      <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:6,flexWrap:"wrap"}}>
        <span style={{fontSize:10.5,fontWeight:700,color:vc.fg,background:vc.bg,border:`1px solid ${vc.bd}`,borderRadius:5,padding:"2px 8px"}}>判定・{vc.mk}</span>
        <span style={{fontSize:14,fontWeight:700,color:"#22303F"}}>{J.verdict.label}</span>
      </div>
      <div style={{fontSize:12.5,color:"#3A424C",lineHeight:1.6,marginBottom:8}}>{J.verdict.reason}</div>
      <div style={{fontSize:12.5,color:"#22303F",lineHeight:1.6,background:vc.bg,borderLeft:`3px solid ${vc.fg}`,padding:"7px 10px",borderRadius:"0 6px 6px 0"}}>
        <span style={{fontWeight:700}}>打ち手：</span>{J.verdict.action}
      </div>
      {J.verdict.tempNote&&<div style={{fontSize:11,color:"#8A8474",marginTop:7}}>補足：{J.verdict.tempNote}</div>}

      <div style={{marginTop:14,marginBottom:6,fontSize:11.5,fontWeight:700,color:"#3A3526"}}>需要構造の3軸<span style={{fontWeight:400,fontSize:10,color:"#9B3B2B",marginLeft:8}}>赤線＝定時寄り判定のしきい値</span></div>
      {meter("需要密度（件/台時）",J.dens,6,J.thresholds.HI_D,`${J.dens.toFixed(1)} 件/台時`)}
      {meter("空間集中度（コリドー・拠点集中）",J.spatial,1,J.thresholds.HI_S,pc(J.spatial))}
      {meter("時間集中度（ピーク1時間シェア）",J.temporal,1,J.thresholds.HI_T,pc(J.temporal))}

      <div style={{marginTop:16,marginBottom:6,fontSize:11.5,fontWeight:700,color:"#3A3526"}}>
        現DRT実績 vs 定時路線の試算
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <div style={{flex:"1 1 240px",border:"1px solid #C4DFC9",borderRadius:8,overflow:"hidden"}}>
          <div style={{background:"#EAF3EC",color:"#2C6E3C",fontSize:11,fontWeight:700,padding:"5px 10px"}}>現DRT（実績）</div>
          <div style={{padding:"8px 10px",fontSize:12,lineHeight:1.9,color:"#3A3526"}}>
            台時あたり乗客：<b>{J.cf.drtPaxPerVehHour.toFixed(1)}人/台時</b><br/>
            実車率：<b>{pc(J.cf.drtLoadRate)}</b>／乗合率：<b>{pc(J.cf.drtShareRate)}</b><br/>
            捌いた需要：<b>{J.served}/{J.demandTotal}件</b>（不成立{pc(J.failRate)}）
          </div>
        </div>
        <div style={{flex:"1 1 240px",border:"1px solid #C4D3E6",borderRadius:8,overflow:"hidden"}}>
          <div style={{background:"#EAF0F7",color:"#2C4A6E",fontSize:11,fontWeight:700,padding:"5px 10px"}}>定時路線（試算）</div>
          <div style={{padding:"8px 10px",fontSize:12,lineHeight:1.9,color:"#3A3526"}}>
            <span style={{fontSize:10.5,color:"#5B7DA8",fontWeight:700}}>輸送量</span><br/>
            この路線がカバーする需要：<b>{pc(J.cf.coverRate)}</b><br/>
            便あたり平均乗客：<b>{J.cf.paxPerTrip.toFixed(1)}人/便</b>（{J.cf.trips}便/日・往復）<br/>
            <span style={{fontSize:10.5,color:"#9B3B2B",fontWeight:700}}>サービスコスト（利用者の負担）</span><br/>
            平均待ち時間：<b>{J.cf.waitAvgMin.toFixed(0)}分</b>（最大{J.cf.waitMaxMin.toFixed(0)}分）※便間隔を広げた分だけ増える<br/>
            乗車時間の伸び：<b>{J.cf.fixedStretch.toFixed(2)}倍</b>（直行比）※経由する停留所が増えるほど伸びる<br/>
            路線：<b>{J.cf.corridorNames||"—"}</b>
          </div>
        </div>
      </div>
      <div style={{fontSize:10.5,color:"#6B6453",marginTop:8,lineHeight:1.7}}>
        読み筋：輸送量とサービスコストはトレードオフの関係にある。便間隔を広げれば「便あたり乗客」は上がるが、それは便を減らした結果で待ち時間が延びただけであり、効率化ではない。路線の停留所数を増やせばカバー率は上がるが、経由地が増える分だけ乗車時間の伸びも大きくなる。判断は輸送量の数字だけでなく、待ち時間・乗車時間の伸びと合わせて見る必要がある。カバー率が低い→需要が路線に乗らず取りこぼす（DRT向きのサイン）。
      </div>

      {/* 個別比較：確定予約（実績あり）のみを対象。母集団が上の判定（不成立込み全需要）と異なる点に注意 */}
      <div style={{marginTop:16,marginBottom:6,fontSize:11.5,fontWeight:700,color:"#3A3526"}}>
        現DRT利用者 個別比較（確定予約ベース）
        <span style={{fontWeight:400,fontSize:10,color:"#9B3B2B",marginLeft:8}}>母集団が上の判定と異なる：ここは実際に運んだ確定予約のみが対象</span>
      </div>
      {!J.compare
        ?<div style={{fontSize:11.5,color:"#8A8474",background:"#F6F3EC",borderRadius:8,padding:"8px 10px"}}>
          確定予約の中に、この路線でカバーされる人がいない。予約を確定するか、路線の停留所数を増やすと表示される。
        </div>
        :<>
        <div style={{fontSize:12,color:"#4A4636",marginBottom:8}}>
          現在の確定予約 {J.compare.drtTotalN}件のうち、この路線でカバーされるのは <b>{J.compare.n}件・{J.compare.paxN}人</b>。
        </div>
        <div style={{overflow:"auto"}}>
        <table style={{borderCollapse:"collapse",width:"100%",fontSize:12}}>
          <thead><tr style={{background:"#F0EDE4"}}>
            <th style={{padding:"5px 10px",textAlign:"left",borderBottom:"2px solid #1E2A38",fontSize:11}}>項目</th>
            <th style={{padding:"5px 10px",textAlign:"right",borderBottom:"2px solid #1E2A38",fontSize:11,color:"#2C6E3C"}}>現DRT実績</th>
            <th style={{padding:"5px 10px",textAlign:"right",borderBottom:"2px solid #1E2A38",fontSize:11,color:"#2C4A6E"}}>定時路線試算</th>
          </tr></thead>
          <tbody>
            <tr>
              <td style={{padding:"4px 10px",borderBottom:"1px solid #EBE7DC"}}>乗車時間（平均）</td>
              <td style={{padding:"4px 10px",textAlign:"right",borderBottom:"1px solid #EBE7DC",fontFamily:"monospace"}}>{J.compare.drtRideAvgMin.toFixed(1)}分</td>
              <td style={{padding:"4px 10px",textAlign:"right",borderBottom:"1px solid #EBE7DC",fontFamily:"monospace"}}>{J.compare.fxRideAvgMin.toFixed(1)}分</td>
            </tr>
            <tr>
              <td style={{padding:"4px 10px",borderBottom:"1px solid #EBE7DC"}}>希望とのズレ／待ち時間（平均）</td>
              <td style={{padding:"4px 10px",textAlign:"right",borderBottom:"1px solid #EBE7DC",fontFamily:"monospace"}}>{J.compare.drtDevAvgMin.toFixed(1)}分</td>
              <td style={{padding:"4px 10px",textAlign:"right",borderBottom:"1px solid #EBE7DC",fontFamily:"monospace"}}>{J.compare.fxWaitAvgMin.toFixed(1)}分（最大{J.compare.fxWaitMaxMin.toFixed(0)}分）</td>
            </tr>
          </tbody>
        </table>
        </div>
        <div style={{fontSize:10.5,color:"#8A8474",marginTop:6,lineHeight:1.6}}>
          「希望とのズレ／待ち時間」はDRT・定時で性質が異なる指標を並べている。DRT側は乗車位置まで迎えに来る約束時刻と希望の差、定時側は停留所で便を待つ物理的な待ち時間（便間隔÷2の期待値）。数字の大小だけでなく体験の質の違いも踏まえて読む必要がある。
        </div>
      </>}

      {/* 車両側の走行距離比較（システム全体のコスト。母集団が異なる点に注意） */}
      <div style={{marginTop:16,marginBottom:6,fontSize:11.5,fontWeight:700,color:"#3A3526"}}>
        車両側の走行距離（1日あたり）
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",fontSize:12,color:"#3A3526"}}>
        <div style={{flex:"1 1 200px",background:"#EAF3EC",borderRadius:8,padding:"8px 10px"}}>
          現DRT（全稼働・回送含む）：<b>{J.distance.drtDailyKm.toFixed(1)}km</b>
        </div>
        <div style={{flex:"1 1 200px",background:"#EAF0F7",borderRadius:8,padding:"8px 10px"}}>
          定時路線（往復×{J.cf.trips}便・需要に関わらず固定）：<b>{J.distance.fixedDailyKm.toFixed(1)}km</b>
        </div>
      </div>
      <div style={{fontSize:10.5,color:"#8A8474",marginTop:6,lineHeight:1.6}}>
        DRT側は現在の全需要（この路線対象外の予約も含む）を運ぶための総走行距離。定時側はこの候補路線を毎日往復運行するのに必要な距離で、乗客がいてもいなくても発生する固定コスト。母集団が異なるため単純な優劣比較ではなく、「DRTは需要に応じて伸縮する変動コスト」「定時は需要に関わらない固定コスト」という性質の違いとして読む。
      </div>

      <div style={{marginTop:14,display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",
        background:"#F6F3EC",borderRadius:8,padding:"8px 10px"}}>
        <span style={{fontSize:11,fontWeight:700,color:"#3A3526"}}>試算の前提</span>
        <label style={{fontSize:11,color:"#4A4636"}}>便間隔
          <input type="number" min={5} max={120} step={5} value={headway}
            onChange={e=>setHeadway(Math.max(5,Math.min(120,Number(e.target.value)||30)))} style={{...inS,marginLeft:5}}/>分</label>
        <label style={{fontSize:11,color:"#4A4636"}}>路線が停まる停留所数
          <input type="number" min={2} max={STOPS.length} step={1} value={lineK}
            onChange={e=>setLineK(Math.max(2,Math.min(STOPS.length,Number(e.target.value)||defK)))} style={{...inS,marginLeft:5}}/>停</label>
      </div>
      <div style={{fontSize:10.5,color:"#8A8474",marginTop:8,lineHeight:1.7}}>
        前提：判定は「{demandSrc}」の需要{J.demandTotal}件に基づく（不成立を含む）。定時路線は需要上位のODから機械的に引いた仮想路線で、人手で選んでいない。稼働{J.vehHours.toFixed(1)}台時・運行{opHours.toFixed(1)}時間帯・便間隔{headway}分での試算。1シナリオの試算であり地域全体の結論ではない。しきい値（密度{J.thresholds.HI_D}・集中{pc(J.thresholds.HI_S)}・時間{pc(J.thresholds.HI_T)}）は仮置きで現場の実態に合わせ調整が要る。
      </div>
    </div>
  </div>);
}

function AnalysisTab({vehicles,sims,resMap,P,lastDemand}){
  const {per,total}=analyzeOps(vehicles,sims,resMap,P);
  // セミ整合性チェック：mode=semiなのに実際のルートにアンカーが無い車両を検出。
  // 予約全消去やネットワーク再読込の後など、routesだけ空になった「見かけフル」状態を炙り出す。
  const semiBroken=vehicles.filter(v=>v.active&&v.mode==="semi"
    &&(!sims[v.id]||!sims[v.id].events||sims[v.id].events.every(e=>e.type!=="ANCHOR")));
  const advice=buildAdvice(total,P,lastDemand);
  const diag=diagnoseBottleneck(total);
  const avgLr=total.span>0?total.loadedTime/total.span:0;
  const dk=diag.main.key;
  const dc = dk==="OK" ? {bd:"#BFE0CD",bg:"#E4F2E9",fg:"#23694A",mk:"良好"}
           : dk==="C"  ? {bd:"#CFD9E5",bg:"#EEF2F7",fg:"#3A5572",mk:"改善余地"}
           :             {bd:"#E5C3B9",bg:"#F8E7E3",fg:"#9B3B2B",mk:"要対応"};
  // 適性判定に使う需要：不成立を含む直近の需要を優先。無ければ確定予約で代替。
  const hasDemand=lastDemand&&lastDemand.rows&&lastDemand.rows.length>0;
  const demandRows=hasDemand
    ? lastDemand.rows.map(r=>({o:r.o,d:r.d,t:r.t,pax:r.pax??1,ok:!!r.ok}))
    : Object.values(resMap).map(r=>({o:r.o,d:r.d,t:r.dpt,pax:r.pax??1,ok:true}));
  const demandSrc=hasDemand
    ?(lastDemand.source==="generated"?"生成需要":lastDemand.source==="file"?"取込需要":"直近の需要")+"（不成立含む）"
    :"確定予約（成立分のみ）";
  // 個別比較用：確定予約の実績（不成立分は実績が無いため対象外）
  const drtDetail=drtPassengerDetail(vehicles,sims,resMap);
  const hm=s=>{const h=Math.floor(s/3600),m=Math.round(s%3600/60);return h>0?`${h}時間${String(m).padStart(2,"0")}分`:`${m}分`;};
  const hmS=s=>{const h=Math.floor(s/3600),m=Math.round(s%3600/60);return h>0?`${h}:${String(m).padStart(2,"0")}`:`${m}分`;};
  const pct=(a,b)=>b>0?`${Math.round(a/b*100)}%`:"—";
  const th={padding:"4px 7px",borderBottom:"2px solid #1E2A38",fontSize:10.5,whiteSpace:"nowrap",textAlign:"right"};
  const td={padding:"4px 7px",borderBottom:"1px solid #DDD8CB",fontSize:11.5,whiteSpace:"nowrap",textAlign:"right",
    fontFamily:"'SF Mono','Consolas',monospace",fontVariantNumeric:"tabular-nums"};
  if(total.n===0)return <div style={{fontSize:13,color:"#6B6453",padding:"16px 0"}}>運行している車両がない。予約を確定すると分析が表示される。</div>;

  const ROWS=[
    ["__eff","運行効率（事業者目線）"],
    ["拘束時間",p=>hm(p.span),"最初の乗車から最後の降車まで車両が運用に拘束された時間。"],
    ["走行距離",p=>`${p.distKm.toFixed(1)}km`,`移動時間×${SPEED_KMH}km/h の換算値。回送を含む。`],
    ["実車時間",p=>`${hm(p.loadedTime)}（${pct(p.loadedTime,p.span)}）`,"1予約以上を乗せていた時間。括弧は拘束時間に対する割合＝実車率。"],
    ["空車回送",p=>`${hm(p.emptyMove)}（${pct(p.emptyMove,p.moveTime)}）`,"誰も乗せずに移動した時間。括弧は移動時間に占める割合。少ないほど効率的。"],
    ["運んだ予約・人数",p=>`${p.resCount}件・${p.paxCarried}人`,"この車両が完了した予約件数と延べ乗客数。"],
    ["延べ乗車人時",p=>`${(p.paxSec/3600).toFixed(1)}人時`,"乗客数×乗車時間の合計。輸送量の実体。"],
    ["平均乗車人数",p=>`${(p.paxSec/Math.max(1,p.span)).toFixed(2)}人`,"延べ乗車人時÷拘束時間。常時何人乗せて走っていたかの平均（空車時間を含む＝車両稼働効率）。"],
    ["平均乗車人数（実車あたり）",p=>`${(p.paxSec/Math.max(1,p.loadedTime)).toFixed(2)}人`,"延べ乗車人時÷実車時間。人を乗せている間に平均何人乗っていたか＝相乗りの濃さ（回送を除く）。"],
    ["乗合時間",p=>hm(p.sharedTime),"別々の予約が2件以上同乗していた時間。同一予約の複数名は乗合に数えない。"],
    ["乗合率（時間ベース）",p=>pct(p.sharedTime,p.loadedTime),"乗合時間÷実車時間。実車中どれだけ相乗りが成立していたか。オンデマンド交通の中核評価指標。"],
    ["寄り道時間の消費",p=>p.stAllow?`${pct(p.stUsed,p.stAllow)}（${hmS(p.stUsed)}／${hmS(p.stAllow)}）`:"—","括弧は 消費／許容（時:分）。消費＝乗客が負担した遠回り（実乗車−直行）の合計、許容＝STn合計。消費率が高いほど利用者の我慢を使い切って乗合を成立させている。"],
    ["__svc","サービス品質（利用者目線）"],
    ["希望時刻とのズレ（平均）",p=>p.resCount?`${(p.devSum/p.resCount/60>=0?"+":"−")}${Math.abs(p.devSum/p.resCount/60).toFixed(1)}分（絶対${(p.devAbsSum/p.resCount/60).toFixed(1)}分）`:"—","約束乗車−希望乗車の平均。符号付きは早発/遅発の偏り、絶対値は希望からの平均的なずれ幅。0に近いほど希望に忠実。"],
    ["希望ズレの比率",p=>p.resCount?pct(p.devAbsSum,p.directSum):"—","ズレ絶対値の合計÷直行時間の合計。移動の長さに対してどれだけ待たされたかの相対指標。短距離トリップで数分ずれると比率は大きく出る。"],
    ["定時率（±1分）",p=>p.resCount?pct(p.onTime,p.resCount):"—","希望乗車時刻の±1分以内に乗車できた予約の割合。"],
    ["乗車時間の伸び（平均倍率）",p=>p.resCount?`${(p.stretchSum/p.resCount).toFixed(2)}倍`:"—","実乗車時間÷直行時間の予約平均。1.0なら直行と同じ、1.5なら直行の1.5倍の時間車内にいた。乗合の遠回りが利用者にかけた負担。"],
    ["乗車時間の伸び（全体）",p=>p.directSum?`${(p.rideSum/p.directSum).toFixed(2)}倍`:"—","実乗車時間の合計÷直行時間の合計。長いトリップの影響を反映した全体倍率。"],
  ];

  return (
  <div>
    <div style={{fontSize:12,color:"#6B6453",marginBottom:10}}>
      現在の確定予約に基づく運行の分析。稼働 {total.n}台・予約 {total.resCount}件・延べ {total.paxCarried}人。
    </div>
    {semiBroken.length>0&&(
      <div style={{background:"#F8E7E3",border:"1px solid #E5C3B9",borderRadius:8,
        padding:"9px 12px",marginBottom:10,fontSize:12,color:"#9B3B2B",lineHeight:1.6}}>
        <b>注意：{semiBroken.map(v=>v.name).join("・")} はセミ設定だが背骨（アンカー）が運行に反映されていない。</b>
        フルデマンドとして集計されている。予約全消去やネットワーク再読込の後は背骨だけが残り
        ルートが空になっている可能性がある。コリドーパネルまたは車両設定から背骨を再適用し、
        その後で需要を自動確定すること。
      </div>
    )}

    {/* ボトルネック診断：効率が上がらない主因を1つに絞る */}
    <div style={{marginBottom:14,border:`1px solid ${dc.bd}`,borderRadius:10,overflow:"hidden"}}>
      <div style={{background:"#14202F",color:"#E8E4DA",fontSize:12,fontWeight:700,padding:"7px 12px",letterSpacing:1}}>
        ボトルネック診断
        <span style={{fontWeight:400,fontSize:10.5,color:"#9AA7BA",marginLeft:8}}>効率が上がらない主因を1つに絞る</span>
      </div>
      <div style={{background:"#fff",padding:"12px 14px"}}>
        <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:6,flexWrap:"wrap"}}>
          <span style={{fontSize:10.5,fontWeight:700,color:dc.fg,background:dc.bg,border:`1px solid ${dc.bd}`,borderRadius:5,padding:"2px 8px"}}>主因・{dc.mk}</span>
          <span style={{fontSize:14,fontWeight:700,color:"#22303F"}}>{diag.main.label}</span>
        </div>
        <div style={{fontSize:12.5,color:"#3A424C",lineHeight:1.6,marginBottom:8}}>{diag.main.detail}</div>
        <div style={{fontSize:12.5,color:"#22303F",lineHeight:1.6,background:"#F6F3EC",borderLeft:`3px solid ${dc.fg}`,padding:"7px 10px",borderRadius:"0 6px 6px 0"}}>
          <span style={{fontWeight:700}}>打ち手：</span>{diag.main.action}
        </div>
        {diag.sub&&<div style={{fontSize:11,color:"#8A8474",marginTop:7}}>次点：{diag.sub.label} — {diag.sub.detail}</div>}
      </div>
    </div>

    {/* 拘束時間の使われ方（実車／空車回送／空車待機の排他分解） */}
    <div style={{marginBottom:14}}>
      <div style={{fontSize:11.5,fontWeight:700,color:"#3A3526",marginBottom:5}}>拘束時間の使われ方（全体 {hm(total.span)}）</div>
      <div style={{display:"flex",height:26,borderRadius:6,overflow:"hidden",border:"1px solid #C9C4B8"}}>
        {[["実車",diag.breakdown.loaded,"#4E8C6A"],["空車回送",diag.breakdown.empty,"#C36B52"],["空車待機",diag.breakdown.idle,"#B7AE9B"]].map(([lb,v,c],i)=>{
          const w=total.span>0?v/total.span*100:0;
          return w<=0?null:(
          <div key={i} title={`${lb} ${hm(v)}`} style={{width:`${w}%`,background:c,display:"flex",
            alignItems:"center",justifyContent:"center",color:"#fff",fontSize:10,fontWeight:700}}>
            {w>=12?`${Math.round(w)}%`:""}
          </div>);
        })}
      </div>
      <div style={{display:"flex",gap:14,marginTop:5,fontSize:10.5,color:"#6B6453",flexWrap:"wrap"}}>
        <span><span style={{display:"inline-block",width:9,height:9,background:"#4E8C6A",borderRadius:2,marginRight:4}}/>実車 {hm(diag.breakdown.loaded)}（{pct(diag.breakdown.loaded,total.span)}）</span>
        <span><span style={{display:"inline-block",width:9,height:9,background:"#C36B52",borderRadius:2,marginRight:4}}/>空車回送 {hm(diag.breakdown.empty)}（{pct(diag.breakdown.empty,total.span)}）</span>
        <span><span style={{display:"inline-block",width:9,height:9,background:"#B7AE9B",borderRadius:2,marginRight:4}}/>空車待機 {hm(diag.breakdown.idle)}（{pct(diag.breakdown.idle,total.span)}）</span>
      </div>
    </div>

    {/* 車両別ばらつき：需要不足か配分偏りかを切り分ける */}
    <div style={{marginBottom:14}}>
      <div style={{fontSize:11.5,fontWeight:700,color:"#3A3526",marginBottom:6}}>車両別の実車率・乗合率</div>
      <div style={{display:"flex",flexDirection:"column",gap:6}}>
        {per.filter(p=>p.used).map(p=>{
          const lr=p.span>0?p.loadedTime/p.span:0;
          const sr=p.loadedTime>0?p.sharedTime/p.loadedTime:0;
          const off=p.span>0&&Math.abs(lr-avgLr)>=0.15;
          const bar=(val,c)=>(
            <div style={{flex:1,height:13,background:"#EDEAE0",borderRadius:4,overflow:"hidden",minWidth:60}}>
              <div style={{width:`${Math.round(val*100)}%`,height:"100%",background:c}}/>
            </div>);
          return (
          <div key={p.v.id} style={{display:"flex",alignItems:"center",gap:7,padding:"3px 6px",borderRadius:6,
            background:off?"#FBF4EF":"transparent",border:off?"1px solid #E8D3C6":"1px solid transparent"}}>
            <div style={{width:112,fontSize:11,color:"#3A3526",display:"flex",alignItems:"center",gap:4,flexShrink:0}}>
              <Dot c={p.v.color}/>{p.v.name}
              {off&&<span style={{fontSize:9,color:"#9B3B2B",fontWeight:700,marginLeft:2}}>外れ</span>}
              {p.empty&&<span style={{fontSize:9,color:"#A89F8B"}}>(予約なし)</span>}
            </div>
            <span style={{width:30,fontSize:9,color:"#8A8474",textAlign:"right",flexShrink:0}}>実車</span>
            {bar(lr,"#4E8C6A")}
            <span style={{width:32,fontSize:10.5,fontFamily:"monospace",textAlign:"right",flexShrink:0}}>{Math.round(lr*100)}%</span>
            <span style={{width:30,fontSize:9,color:"#8A8474",textAlign:"right",flexShrink:0}}>乗合</span>
            {bar(sr,"#5B7DA8")}
            <span style={{width:32,fontSize:10.5,fontFamily:"monospace",textAlign:"right",flexShrink:0}}>{Math.round(sr*100)}%</span>
          </div>);
        })}
      </div>
      <div style={{fontSize:10.5,color:"#8A8474",marginTop:6,lineHeight:1.6}}>
        全体平均 実車率{pct(total.loadedTime,total.span)}・乗合率{pct(total.sharedTime,total.loadedTime)}。
        平均から実車率が15ポイント以上外れた車両を「外れ」と表示。特定車両だけ低いなら需要不足ではなく配分の偏り＝全体最適化A/Bの再配分で改善余地。全車一様に低いなら需要側の要因。
      </div>
    </div>

    {/* 運行改善アドバイス（ルールベース） */}
    <div style={{marginBottom:14,border:"1px solid #D8D3C6",borderRadius:10,overflow:"hidden"}}>
      <div style={{background:"#14202F",color:"#E8E4DA",fontSize:12,fontWeight:700,
        padding:"7px 12px",letterSpacing:1}}>
        運行改善アドバイス（補足）
        <span style={{fontWeight:400,fontSize:10.5,color:"#9AA7BA",marginLeft:8}}>
          主因の外にある個別ヒント（ルールベース）
        </span>
      </div>
      <div style={{background:"#fff",padding:"4px 0"}}>
        {advice.map((a,i)=>{
          const col=a.level==="warn"?{bg:"#F8E7E3",bd:"#E5C3B9",fg:"#9B3B2B",mk:"要改善"}
                   :a.level==="good"?{bg:"#E4F2E9",bd:"#BFE0CD",fg:"#23694A",mk:"良好"}
                   :{bg:"#EEF2F7",bd:"#CFD9E5",fg:"#3A5572",mk:"ヒント"};
          return (
          <div key={i} style={{display:"flex",gap:10,alignItems:"flex-start",padding:"7px 12px"}}>
            <span style={{flexShrink:0,fontSize:10,fontWeight:700,color:col.fg,background:col.bg,
              border:`1px solid ${col.bd}`,borderRadius:5,padding:"2px 7px",marginTop:1}}>{col.mk}</span>
            <span style={{fontSize:12.5,color:"#2A323C",lineHeight:1.6}}>{a.text}</span>
          </div>);
        })}
      </div>
    </div>

    <div style={{overflow:"auto"}}>
    <table style={{borderCollapse:"collapse",background:"#fff",border:"1px solid #B9B2A1"}}>
      <thead><tr style={{background:"#F0EDE4"}}>
        <th style={{...th,textAlign:"left",position:"sticky",left:0,zIndex:2,background:"#F0EDE4"}}>指標</th>
        {per.filter(p=>p.used).map(p=>(
          <th key={p.v.id} style={th}><Dot c={p.v.color}/>{p.v.name}
            {p.empty&&<span style={{fontWeight:400,fontSize:9.5,color:"#A89F8B",marginLeft:3}}>(予約なし)</span>}
          </th>))}
        <th style={{...th,background:"#E8E4DA"}}>合計/全体</th>
        <th style={{...th,textAlign:"left",fontWeight:400,color:"#8A8474"}}>説明</th>
      </tr></thead>
      <tbody>
        {ROWS.map((row,i)=>{
          if(row[0]==="__svc"||row[0]==="__eff")return (
            <tr key={i}><td colSpan={per.filter(p=>p.used).length+3}
              style={{padding:"7px 10px",background:"#14202F",color:"#E8E4DA",fontSize:11,
                fontWeight:700,letterSpacing:1}}>{row[1]}</td></tr>);
          const [name,fn,desc]=row;
          return (
          <tr key={i}>
            <td style={{...td,textAlign:"left",fontFamily:"inherit",fontWeight:700,
              position:"sticky",left:0,zIndex:1,background:"#fff"}}>{name}</td>
            {per.filter(p=>p.used).map(p=>(
              <td key={p.v.id} style={td}>{fn(p)}</td>))}
            <td style={{...td,background:"#F6F3EC",fontWeight:700}}>{fn(total)}</td>
            <td style={{...td,textAlign:"left",fontFamily:"inherit",fontSize:10.5,color:"#6B6453",
              whiteSpace:"normal",minWidth:150,maxWidth:230,lineHeight:1.45}}>{desc}</td>
          </tr>);
        })}
      </tbody>
    </table>
    </div>
    <div style={{fontSize:11,color:"#8A8474",marginTop:10,lineHeight:1.8}}>
      読み方の目安：実車率が低く空車回送が多い→需要が空間的に散っており停留所配置か台数配分の見直し余地。
      乗合率が低い→時間的に需要が重なっていないか、San（ズレ幅）が狭く相乗りの機会を逃している。
      寄り道消費率が高いのに乗合率も高い→エンジンは許容範囲を使い切って効率化しており、これ以上は利用者体験とのトレードオフ。
    </div>

    {/* 運行方式の適性判定（DRT／セミ／定時） */}
    <SuitabilityView demandRows={demandRows} vehicles={vehicles} total={total} demandSrc={demandSrc} dwell={P.dwell} drtDetail={drtDetail}/>

    {/* 流し込み需要の記録（生成設定＋不成立を含む一覧） */}
    <DemandRecord lastDemand={lastDemand} vehicles={vehicles}/>
  </div>);
}

/* ---------- 流し込み需要の記録（運行分析タブ末尾） ----------
   直近に自動確定した需要の「生成設定/取込元」と、不成立を含む全件一覧を表示。
   どの設定で流したか・何が成立しなかったかを後から追えるようにする。 */
function DemandRecord({lastDemand,vehicles}){
  const [open,setOpen]=React.useState(false);
  if(!lastDemand||!lastDemand.rows||lastDemand.rows.length===0)
    return (
    <div style={{marginTop:14,fontSize:11.5,color:"#8A8474",background:"#fff",
      border:"1px dashed #D8D3C6",borderRadius:10,padding:"10px 12px"}}>
      流し込み需要の記録なし。「予約一覧の流し込み」から自動確定すると、生成設定と不成立を含む一覧がここに残る。
    </div>);
  const rows=lastDemand.rows;
  const ok=rows.filter(r=>r.ok).length;
  const s=lastDemand.settings;
  const vn=id=>vehicles.find(v=>v.id===id)?.name??(id?`${id}号車`:"—");
  const td={padding:"3px 8px",borderBottom:"1px solid #EBE7DC",fontSize:11.5,whiteSpace:"nowrap",
    fontFamily:"'SF Mono','Consolas',monospace",textAlign:"center"};
  const th={padding:"4px 8px",borderBottom:"2px solid #1E2A38",fontSize:10.5,whiteSpace:"nowrap",background:"#F0EDE4"};
  const when=lastDemand.when?new Date(lastDemand.when):null;
  return (
  <div style={{marginTop:14,border:"1px solid #D8D3C6",borderRadius:10,overflow:"hidden",background:"#fff"}}>
    <div style={{background:"#F0EDE4",padding:"8px 12px",fontSize:12,fontWeight:700,color:"#3A3526"}}>
      流し込み需要の記録
      <span style={{fontWeight:400,color:"#6B6453",marginLeft:8,fontSize:11}}>
        {lastDemand.committedBy??"自動確定"}／成立 {ok}/{rows.length}件
        {when?`／${when.getHours()}:${String(when.getMinutes()).padStart(2,"0")}流し込み`:""}
      </span>
    </div>
    <div style={{padding:"8px 12px",fontSize:11.5,color:"#4A4636",lineHeight:1.8}}>
      {lastDemand.source==="generated"&&s?
        `生成設定：件数${s.n}件／時間帯ピーク強度${Math.round((s.peakInt??0)*100)}%／拠点集中度${Math.round((s.hubConc??0)*100)}%／ズレ幅±${Math.round((s.genSa??0)/60)}分／帰り便${Math.round((s.genRet??0)*100)}%／滞留${Math.round((s.genStay??0)/60)}分（中央値）`
       :lastDemand.source==="file"?
        `取込元：ファイル ${lastDemand.name??""}`
       :"取込元：不明"}
    </div>
    <div style={{padding:"0 12px 10px"}}>
      <button onClick={()=>setOpen(o=>!o)} style={{fontSize:11.5,border:"1px solid #C9C4B8",
        background:"#F5F3EE",borderRadius:6,padding:"4px 12px",cursor:"pointer",color:"#3A3526"}}>
        {open?"需要一覧を隠す":`需要一覧を見る（全${rows.length}件・不成立${rows.length-ok}件）`}
      </button>
      {open&&<div style={{maxHeight:280,overflow:"auto",marginTop:8,border:"1px solid #E3DFD3",borderRadius:6}}>
        <table style={{borderCollapse:"collapse",width:"100%"}}>
          <thead><tr>
            <th style={{...th,textAlign:"right"}}>#</th>
            <th style={th}>状態</th><th style={th}>発→着</th><th style={th}>希望</th>
            <th style={th}>希望時刻</th><th style={th}>人数</th><th style={th}>San</th>
            <th style={{...th,textAlign:"left"}}>号車／約束（実）・不成立理由</th>
          </tr></thead>
          <tbody>
            {rows.map((r,i)=>(
            <tr key={i} style={{background:r.ok?"#fff":"#FBF1EE"}}>
              <td style={{...td,textAlign:"right",color:"#8A8474"}}>{r.num??"—"}</td>
              <td style={{...td,fontWeight:700,color:r.ok?"#23694A":"#9B3B2B"}}>{r.ok?"成立":"不成立"}</td>
              <td style={td}>{STOPS[r.o]}→{STOPS[r.d]}</td>
              <td style={td}>{r.mode==="arr"?"着":"発"}</td>
              <td style={td}>{fmt(r.t)}</td>
              <td style={td}>{r.pax??1}</td>
              <td style={td}>±{Math.round((r.sa??0)/60)}分</td>
              <td style={{...td,textAlign:"left",whiteSpace:"normal",fontFamily:"inherit",
                color:r.ok?"#23694A":"#9B3B2B",minWidth:200}}>
                {r.ok
                  ?`${vn(r.vehicle)}・${r.mode==="arr"?(r.idt!=null?`約束着${fmt(r.idt)}`:"成立"):(r.apt!=null?`約束発${fmt(r.apt)}`:"成立")}`
                  :r.reason}
              </td>
            </tr>))}
          </tbody>
        </table>
      </div>}
    </div>
  </div>);
}

/* ---------- 需要プレビュー（流し込み前の確認） ----------
   生成・読込した需要の「時間帯分布」「拠点への発着集中」「明細一覧」を表示。
   スライダ調整の結果がどんな需要になったかを流し込み前に目視確認する。 */
function DemandPreview({rows}){
  const valid=rows.filter(r=>!r.err);
  // 時間帯ヒストグラム（30分ビン、9:00〜17:00）
  const bins=Array.from({length:16},(_,i)=>({h0:9*3600+i*1800,n:0}));
  for(const r of valid){
    const idx=Math.floor((r.t-9*3600)/1800);
    if(idx>=0&&idx<bins.length)bins[idx].n++;
  }
  const maxBin=Math.max(1,...bins.map(b=>b.n));
  const avgBin=valid.length/bins.length;   // 平らな場合の1ビンあたり期待件数
  // 停留所別の発着集計
  const dep=Array(STOPS.length).fill(0), arr=Array(STOPS.length).fill(0);
  for(const r of valid){dep[r.o]++;arr[r.d]++;}
  const totPax=valid.reduce((a,r)=>a+r.pax,0);
  const hubTraffic=HUBS.reduce((a,h)=>a+dep[h]+arr[h],0);
  const hubShare=valid.length?Math.round(hubTraffic/(valid.length*2)*100):0;
  // 上位発着停留所
  const traffic=STOPS.map((s,i)=>({i,s,total:dep[i]+arr[i],dep:dep[i],arr:arr[i]}))
    .sort((a,b)=>b.total-a.total).slice(0,6);

  const [showList,setShowList]=React.useState(false);
  const fmtBinLabel=t=>{const h=Math.floor(t/3600),m=t%3600/60;return m===0?`${h}`:"";};
  const th={padding:"3px 7px",borderBottom:"1px solid #C9C4B8",fontSize:10.5,whiteSpace:"nowrap",background:"#F0EDE4"};
  const td={padding:"3px 7px",borderBottom:"1px solid #EBE7DC",fontSize:11.5,whiteSpace:"nowrap",
    fontFamily:"'SF Mono','Consolas',monospace",textAlign:"center"};

  return (
  <div style={{background:"#fff",border:"1px solid #D8D3C6",borderRadius:10,padding:12}}>
    <div style={{fontSize:13,fontWeight:700,marginBottom:8}}>
      生成された需要の確認　<span style={{fontWeight:400,color:"#6B6453",fontSize:12}}>
      {valid.length}件・延べ{totPax}人／拠点発着シェア {hubShare}%（拠点{HUBS.length}停留所）</span>
    </div>

    {/* 時間帯分布 */}
    <div style={{fontSize:11,color:"#6B6453",marginBottom:4}}>時間帯分布（30分ごとの件数）</div>
    <svg viewBox="0 0 520 90" style={{width:"100%",height:90,marginBottom:10}}>
      {bins.map((b,i)=>{
        const x=10+i*31.5, h=b.n/maxBin*60;
        const hour=b.h0/3600;
        // ピーク帯（共通定義）に該当し、かつ実データが平均を有意に上回るビンを橙に。
        // ピーク強度0で全体が平らなら橙にならない（色＝実際の山）。
        const peak=isPeakHour(hour+0.25)&&b.n>avgBin*1.25;
        return (<g key={i}>
          <rect x={x} y={72-h} width={26} height={h} rx={2}
            fill={peak?"#E0853E":"#9BB0C9"} opacity="0.9"/>
          {b.n>0&&<text x={x+13} y={70-h} textAnchor="middle" fontSize="9" fill="#444"
            fontFamily="'SF Mono','Consolas',monospace">{b.n}</text>}
          {b.h0%3600===0&&<text x={x+13} y={85} textAnchor="middle" fontSize="9" fill="#8A8474"
            fontFamily="'SF Mono','Consolas',monospace">{fmtBinLabel(b.h0)}時</text>}
        </g>);
      })}
    </svg>

    {/* 発着の多い停留所 */}
    <div style={{fontSize:11,color:"#6B6453",marginBottom:4}}>発着の多い停留所（上位6）</div>
    <table style={{borderCollapse:"collapse",background:"#fff",marginBottom:8}}>
      <thead><tr>
        <th style={th}>停留所</th><th style={th}>発</th><th style={th}>着</th><th style={th}>計</th><th style={th}>種別</th>
      </tr></thead>
      <tbody>
        {traffic.map(t=>(
          <tr key={t.i}>
            <td style={{...td,fontWeight:700}}>{t.s}</td>
            <td style={td}>{t.dep}</td><td style={td}>{t.arr}</td>
            <td style={{...td,fontWeight:700}}>{t.total}</td>
            <td style={{...td,color:HUBS.includes(t.i)?"#1E6B40":"#8A8474"}}>
              {HUBS.includes(t.i)?`拠点${HUB_NAMES[t.i]?"・"+HUB_NAMES[t.i]:""}`:"一般"}</td>
          </tr>
        ))}
      </tbody>
    </table>

    {/* 明細一覧（折りたたみ） */}
    <button onClick={()=>setShowList(s=>!s)}
      style={{padding:"5px 12px",borderRadius:7,fontSize:11.5,cursor:"pointer",
        border:"1px solid #C9C4B8",background:"#F5F3EE",color:"#444",fontWeight:700}}>
      {showList?"明細を隠す":`明細一覧を見る（${rows.length}行）`}
    </button>
    {showList&&<div style={{maxHeight:240,overflow:"auto",border:"1px solid #D8D3C6",borderRadius:8,marginTop:8}}>
      <table style={{borderCollapse:"collapse",background:"#fff",width:"100%"}}>
        <thead><tr>
          <th style={{...th,position:"sticky",top:0}}>#</th>
          <th style={{...th,position:"sticky",top:0}}>発地</th>
          <th style={{...th,position:"sticky",top:0}}>着地</th>
          <th style={{...th,position:"sticky",top:0}}>希望</th>
          <th style={{...th,position:"sticky",top:0}}>時刻</th>
          <th style={{...th,position:"sticky",top:0}}>人数</th>
          <th style={{...th,position:"sticky",top:0}}>ズレ幅</th>
        </tr></thead>
        <tbody>
          {rows.map((r,i)=>(
            <tr key={i} style={{background:r.err?"#F8E7E3":undefined}}>
              <td style={td}>{i+1}</td>
              <td style={td}>{r.o!=null?STOPS[r.o]:"—"}{r.o!=null&&HUBS.includes(r.o)?"●":""}</td>
              <td style={td}>{r.d!=null?STOPS[r.d]:"—"}{r.d!=null&&HUBS.includes(r.d)?"●":""}</td>
              <td style={td}>{r.mode==="arr"?"着":"発"}</td>
              <td style={td}>{r.t!=null?fmt(r.t):"—"}</td>
              <td style={td}>{r.pax}名</td>
              <td style={td}>±{Math.round(r.sa/60)}分</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>}
    <div style={{fontSize:10.5,color:"#8A8474",marginTop:6}}>
      橙＝ピーク時間帯で件数が平均を上回るビン／●＝拠点停留所。スライダを変えて再生成すると、この分布が変わる。
    </div>
  </div>);
}
