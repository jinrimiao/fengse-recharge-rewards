'use client';

import {useCallback,useEffect,useMemo,useRef,useState} from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import {CheckCircle2,Download,FileSpreadsheet,Gamepad2,Pencil,Plus,Settings2,Trash2,Upload} from 'lucide-react';
import {Button} from '@/components/ui/button';
import {Input} from '@/components/ui/input';
import {Label} from '@/components/ui/label';
import {Card,CardContent,CardHeader,CardTitle} from '@/components/ui/card';
import {SITE_VERSION} from './version';

type Reward={item:string;quantity:number};
type Tier={amount:number;rewards:Reward[]};
type RewardRule='highest'|'all';
type Game={id:string;name:string;tiers:Tier[];allTiers:Tier[];rewardRule:RewardRule};
type RegionName='亚洲'|'美洲'|'欧洲';
type Player={region:RegionName;server:string;roleId:string;amount:number};
type RegionIssue={sheet:string;row:number;server:string;roleId:string;reason:string};
type OutputRow=Player&Tier;
type PendingImport={gameId:string;gameName:string;rule:RewardRule;tiers:Tier[];tierCount:number;rewardCount:number;fileName:string};
type GameBackup={format:string;version:number;createdAt:string;games:Game[];activeGameId:string};
type PendingBackupRestore={backup:GameBackup;fileName:string;gameCount:number;tierCount:number;rewardCount:number};
type GameDialog={mode:'add'|'rename';value:string};

const STORAGE_KEY='fengse-reward-games-v1';
const BACKUP_FORMAT='fengse-recharge-rewards-all-games-backup';
const BACKUP_VERSION=2;
const LEGACY_KEYS=['fengse-reward-tiers','rewardTiers','tiers'];
const DEFAULT_GAME_NAME='风色传说';
const initial:Tier[]=[[70,'24156'],[140,'24157'],[300,'24158'],[400,'24159'],[600,'24160'],[700,'24161'],[1500,'24162'],[3000,'24163'],[4500,'24164']].map(([amount,item])=>({amount:Number(amount),rewards:[{item:String(item),quantity:1}]}));
const cloneTiers=(tiers:Tier[])=>tiers.map(t=>({amount:t.amount,rewards:t.rewards.map(r=>({...r,item:String(r.item)}))}));
const RULE_LABEL:Record<RewardRule,string>={highest:'最高档位奖励',all:'全部达标档位奖励'};
const upgradeGame=(game:{id:string;name:string;tiers:Tier[];allTiers?:Tier[];rewardRule?:RewardRule}):Game=>({id:game.id,name:game.name,tiers:cloneTiers(game.tiers),allTiers:cloneTiers(validTiers(game.allTiers)?game.allTiers:initial),rewardRule:game.rewardRule==='all'?'all':'highest'});
const createId=()=>globalThis.crypto?.randomUUID?.()??'game-'+Date.now()+'-'+Math.random().toString(36).slice(2);
const stringValue=(v:unknown)=>typeof v==='string'?v:typeof v==='number'||typeof v==='boolean'?String(v):'';
const norm=(v:unknown)=>stringValue(v).trim().toLowerCase().replace(/[\s_()（）$：:]/g,'');
const col=(hs:unknown[],rs:RegExp[])=>hs.findIndex(h=>rs.some(r=>r.test(norm(h))));
const money=(v:unknown)=>{if(typeof v==='number')return v;const m=stringValue(v).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);return m?Number(m[0]):NaN};
const REGION_PATTERNS:Record<RegionName,RegExp[]>={
 亚洲:[/(?:^|[-_\s./:])(?:asia|as)(?=$|[-_\s./:]|\d)/,/(?:^|[-_\s./:])亚洲(?=$|[-_\s./:]|\d)/],
 美洲:[/(?:^|[-_\s./:])(?:america|usa|us|na)(?=$|[-_\s./:]|\d)/,/(?:^|[-_\s./:])(?:美洲|北美)(?=$|[-_\s./:]|\d)/],
 欧洲:[/(?:^|[-_\s./:])(?:europe|eu)(?=$|[-_\s./:]|\d)/,/(?:^|[-_\s./:])欧洲(?=$|[-_\s./:]|\d)/],
};
const matchRegion=(value:string)=>{const text=value.trim().toLowerCase(),matches=(Object.entries(REGION_PATTERNS) as [RegionName,RegExp[]][]).filter(([,patterns])=>patterns.some(pattern=>pattern.test(text))).map(([name])=>name);return matches.length===1?{status:'matched' as const,region:matches[0],matches}:matches.length>1?{status:'conflict' as const,region:null,matches}:{status:'unrecognized' as const,region:null,matches}};
const identifyRegion=(sheet:string,server:string,hasServerColumn:boolean)=>{const source=hasServerColumn&&server.trim()?{label:'服务器名称',value:server}:{label:'工作表名称',value:sheet},matched=matchRegion(source.value);return matched.status==='matched'?{region:matched.region,reason:''}:{region:null,reason:matched.status==='conflict'?source.label+'包含多个大区标识：'+matched.matches.join('、'):source.label+'无法识别大区'}};
const validTiers=(value:unknown):value is Tier[]=>Array.isArray(value)&&value.length>0&&value.every(t=>t&&typeof t==='object'&&Number.isFinite((t as Tier).amount)&&Array.isArray((t as Tier).rewards)&&(t as Tier).rewards.length>0&&(t as Tier).rewards.every(r=>r&&stringValue(r.item).trim()&&Number.isInteger(Number(r.quantity))&&Number(r.quantity)>0));
const safeName=(name:string)=>name.replace(/[\\/:*?"<>|]/g,'_').trim()||'游戏';
const isObject=(value:unknown):value is Record<string,unknown>=>Boolean(value)&&typeof value==='object'&&!Array.isArray(value);
const hasControlCharacter=(value:string)=>{for(let index=0;index<value.length;index++)if(value.charCodeAt(index)<32)return true;return false};
const requireFields=(value:Record<string,unknown>,required:string[],label:string)=>{const keys=Object.keys(value),missing=required.filter(key=>!keys.includes(key)),extra=keys.filter(key=>!required.includes(key));if(missing.length)throw new Error(label+'缺少必要字段：'+missing.join('、'));if(extra.length)throw new Error(label+'包含不支持的字段：'+extra.join('、'))};

function validateBackup(value:unknown):GameBackup{
 if(!isObject(value))throw new Error('备份根对象格式无效');
 requireFields(value,['format','version','createdAt','games','activeGameId'],'备份文件');
 if(value.format!==BACKUP_FORMAT)throw new Error('不是本工具支持的全部游戏备份文件');
 if(value.version!==1&&value.version!==BACKUP_VERSION)throw new Error('备份版本不兼容，当前支持版本 1 和 '+BACKUP_VERSION);
 if(typeof value.createdAt!=='string'||!value.createdAt||!Number.isFinite(Date.parse(value.createdAt)))throw new Error('备份创建时间无效');
 if(!Array.isArray(value.games)||!value.games.length)throw new Error('备份中必须至少包含一个游戏');
 const ids=new Set<string>(),names=new Set<string>();
 const games=value.games.map((rawGame,index)=>{
  const label='第 '+(index+1)+' 个游戏';if(!isObject(rawGame))throw new Error(label+'格式无效');requireFields(rawGame,value.version===1?['id','name','tiers']:['id','name','tiers','allTiers','rewardRule'],label);
  if(typeof rawGame.id!=='string'||!rawGame.id.trim()||rawGame.id.length>200||hasControlCharacter(rawGame.id))throw new Error(label+'的唯一 ID 无效');
  if(ids.has(rawGame.id))throw new Error('存在重复游戏 ID：'+rawGame.id);ids.add(rawGame.id);
  if(typeof rawGame.name!=='string'||rawGame.name!==rawGame.name.trim()||!rawGame.name||rawGame.name.length>100||hasControlCharacter(rawGame.name))throw new Error(label+'的游戏名称无效');
  const normalizedName=rawGame.name.toLowerCase();if(names.has(normalizedName))throw new Error('存在重复游戏名称：'+rawGame.name);names.add(normalizedName);
  const parseTiers=(rawTiers:unknown,ruleLabel:string)=>{
   if(!Array.isArray(rawTiers)||!rawTiers.length)throw new Error(label+'的'+ruleLabel+'必须至少包含一个奖励档位');
   return rawTiers.map((rawTier,tierIndex)=>{
   const tierLabel=label+'第 '+(tierIndex+1)+' 个档位';if(!isObject(rawTier))throw new Error(tierLabel+'格式无效');requireFields(rawTier,['amount','rewards'],tierLabel);
   if(typeof rawTier.amount!=='number'||!Number.isFinite(rawTier.amount)||rawTier.amount<=0)throw new Error(tierLabel+'的储值金额必须是有效的正数');
   if(!Array.isArray(rawTier.rewards)||!rawTier.rewards.length)throw new Error(tierLabel+'必须至少包含一个物品');
   const rewards=rawTier.rewards.map((rawReward,rewardIndex)=>{
    const rewardLabel=tierLabel+'第 '+(rewardIndex+1)+' 个物品';if(!isObject(rawReward))throw new Error(rewardLabel+'格式无效');requireFields(rawReward,['item','quantity'],rewardLabel);
    if(typeof rawReward.item!=='string'||!rawReward.item.trim()||rawReward.item.length>200||hasControlCharacter(rawReward.item))throw new Error(rewardLabel+'的物品 ID 无效，物品 ID 必须以文本保存');
    if(typeof rawReward.quantity!=='number'||!Number.isSafeInteger(rawReward.quantity)||rawReward.quantity<=0)throw new Error(rewardLabel+'的数量必须是有效的正整数');
    return{item:rawReward.item,quantity:rawReward.quantity};
   });
   return{amount:rawTier.amount,rewards};
  });
  };
  const tiers=parseTiers(rawGame.tiers,'最高档位奖励'),allTiers=value.version===1?cloneTiers(initial):parseTiers(rawGame.allTiers,'全部达标档位奖励');
  if(value.version!==1&&rawGame.rewardRule!=='highest'&&rawGame.rewardRule!=='all')throw new Error(label+'的当前奖励规则无效');
  return{id:rawGame.id,name:rawGame.name,tiers,allTiers,rewardRule:value.version===1?'highest':rawGame.rewardRule as RewardRule};
 });
 if(typeof value.activeGameId!=='string'||!ids.has(value.activeGameId))throw new Error('当前选中的游戏 ID 无效或不在游戏列表中');
 return{format:BACKUP_FORMAT,version:BACKUP_VERSION,createdAt:value.createdAt,games,activeGameId:value.activeGameId};
}

function downloadBuffer(data:ArrayBuffer|Uint8Array,name:string,type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'){
 const bytes=data instanceof ArrayBuffer?new Uint8Array(data):Uint8Array.from(data);
 const blob=new Blob([bytes],{type}),url=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=url;a.download=name;a.click();URL.revokeObjectURL(url);
}

function defaultSendTemplate(){
 const title='Recharge Reward',content='Dear Adventurer, \nThis is your offline cumulative recharge reward. Please claim it promptly. Thank you for your support!';
 const ws=XLSX.utils.aoa_to_sheet([['区服','角色ID','','标题','内容','物品ID','数量'],['','','',title,content,'','']]);
 ws['!cols']=[{wch:10},{wch:14},{wch:14},{wch:20},{wch:52},{wch:12},{wch:9}];
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Sheet1');
 return XLSX.write(wb,{bookType:'xlsx',type:'array'}) as ArrayBuffer;
}

export default function Home(){
 const [games,setGames]=useState<Game[]>([]);
 const [activeGameId,setActiveGameId]=useState('');
 const [hydrated,setHydrated]=useState(false);
 const [players,setPlayers]=useState<Player[]>([]);
 const [sourceName,setSourceName]=useState('');
 const [template,setTemplate]=useState<ArrayBuffer|null>(null);
 const [templateName,setTemplateName]=useState('默认发奖模板');
 const [message,setMessage]=useState('上传累计储值表后即可预览');
 const [regionIssues,setRegionIssues]=useState<RegionIssue[]>([]);
 const [pendingImport,setPendingImport]=useState<PendingImport|null>(null);
 const [pendingBackupRestore,setPendingBackupRestore]=useState<PendingBackupRestore|null>(null);
 const [gameDialog,setGameDialog]=useState<GameDialog|null>(null);
 const [headers,setHeaders]=useState({server:'区服',role:'角色ID',item:'物品ID',quantity:'数量',amount:'累计储值',tier:'达成档位'});
 const sourceRef=useRef<HTMLInputElement>(null),templateRef=useRef<HTMLInputElement>(null),tierTemplateRef=useRef<HTMLInputElement>(null),backupRef=useRef<HTMLInputElement>(null);
 const activeGame=games.find(g=>g.id===activeGameId)??games[0];
 const activeRule:RewardRule=activeGame?.rewardRule??'highest';
 const tiers=activeGame?(activeRule==='all'?activeGame.allTiers:activeGame.tiers):initial;

 useEffect(()=>{
  const timer=setTimeout(()=>{
  let restored:Game[]|null=null,selected='';
  try{
   const raw=localStorage.getItem(STORAGE_KEY);
   if(raw){
    const saved=JSON.parse(raw) as {games?:Array<Partial<Game>&{id:string;name:string;tiers:Tier[]}>;activeGameId?:string};
    if(Array.isArray(saved.games)&&saved.games.length&&saved.games.every(g=>g&&typeof g.id==='string'&&typeof g.name==='string'&&validTiers(g.tiers)&&(!g.allTiers||validTiers(g.allTiers)))){
     restored=saved.games.map(upgradeGame);selected=String(saved.activeGameId??'');
    }
   }
   if(!restored){
    for(const key of LEGACY_KEYS){
     const legacy=localStorage.getItem(key);if(!legacy)continue;
     const parsed=JSON.parse(legacy);
     if(validTiers(parsed)){restored=[upgradeGame({id:createId(),name:DEFAULT_GAME_NAME,tiers:parsed})];break}
    }
   }
  }catch{}
  const next=restored??[upgradeGame({id:createId(),name:DEFAULT_GAME_NAME,tiers:initial})];
  setGames(next);setActiveGameId(next.some(g=>g.id===selected)?selected:next[0].id);setHydrated(true);
  },0);
  return()=>clearTimeout(timer);
 },[]);

 useEffect(()=>{if(hydrated&&games.length)localStorage.setItem(STORAGE_KEY,JSON.stringify({games,activeGameId}))},[games,activeGameId,hydrated]);

 const updateCurrentTiers=useCallback((updater:Tier[]|((current:Tier[])=>Tier[]))=>{
  setGames(current=>current.map(g=>{if(g.id!==activeGameId)return g;const key=g.rewardRule==='all'?'allTiers':'tiers',next=typeof updater==='function'?updater(g[key]):updater;return{...g,[key]:cloneTiers(next)}}));
 },[activeGameId]);

 const outputs=useMemo<OutputRow[]>(()=>players.flatMap(p=>{const reached=[...tiers].filter(t=>p.amount>=t.amount).sort((a,b)=>a.amount-b.amount);if(activeRule==='all')return reached.map(t=>({...t,...p}));const hit=reached.at(-1);return hit?[{...hit,...p}]:[]}),[players,tiers,activeRule]);
 const stats=useMemo(()=>['亚洲','美洲','欧洲'].map(r=>({region:r,players:players.filter(p=>p.region===r&&tiers.some(t=>p.amount>=t.amount)).length,rewards:outputs.filter(x=>x.region===r).length})),[players,outputs,tiers]);
 const regionCounts=useMemo(()=>({亚洲:players.filter(p=>p.region==='亚洲').length,美洲:players.filter(p=>p.region==='美洲').length,欧洲:players.filter(p=>p.region==='欧洲').length}),[players]);

 useEffect(()=>{
  const context=(document as Document&{modelContext?:{registerTool:(tool:unknown,options?:{signal?:AbortSignal})=>void|Promise<void>}}).modelContext;
  if(!context?.registerTool)return;
  const lifecycle=new AbortController();
  void Promise.resolve(context.registerTool({name:'set_reward_tiers',title:'设置当前游戏累计储值奖档',description:'批量替换当前游戏的累计储值金额及每档多个物品奖励。',inputSchema:{type:'object',properties:{tiers:{type:'array',items:{type:'object',properties:{amount:{type:'number',minimum:0},rewards:{type:'array',minItems:1,items:{type:'object',properties:{item:{type:'string',minLength:1},quantity:{type:'number',minimum:1}},required:['item','quantity'],additionalProperties:false}}},required:['amount','rewards'],additionalProperties:false}}},required:['tiers'],additionalProperties:false},annotations:{readOnlyHint:false,untrustedContentHint:false},execute(input:unknown){const value=input as {tiers?:Tier[]};if(!validTiers(value.tiers))throw new Error('奖档格式无效');updateCurrentTiers(value.tiers);setMessage('已更新当前游戏的 '+value.tiers.length+' 个奖档');return{updated:value.tiers.length}}},{signal:lifecycle.signal})).catch(()=>{});
  return()=>lifecycle.abort();
 },[updateCurrentTiers]);

 function submitGameDialog(){
  if(!gameDialog)return;const entered=gameDialog.value.trim();
  if(!entered){setMessage('游戏名称不能为空');return}
  if(games.some(g=>(gameDialog.mode==='add'||g.id!==activeGameId)&&g.name.toLowerCase()===entered.toLowerCase())){setMessage('游戏名称不能重复');return}
  if(gameDialog.mode==='add'){
   const game=upgradeGame({id:createId(),name:entered,tiers:initial});
   setGames(current=>[...current,game]);setActiveGameId(game.id);setPendingImport(null);setMessage('已新增并切换到游戏：'+entered);
  }else{
   setGames(current=>current.map(g=>g.id===activeGameId?{...g,name:entered}:g));setMessage('游戏已重命名为：'+entered);
  }
  setGameDialog(null);
 }

 function backupAllGames(){
  try{
   const createdAt=new Date().toISOString(),backup=validateBackup({format:BACKUP_FORMAT,version:BACKUP_VERSION,createdAt,games,activeGameId});
   const date=createdAt.slice(0,10).replace(/-/g,''),bytes=new TextEncoder().encode(JSON.stringify(backup,null,2));
   downloadBuffer(bytes,'累计储值奖励_全部游戏备份_'+date+'.json','application/json;charset=utf-8');
   setMessage('已备份全部 '+games.length+' 个游戏配置');
  }catch(e){setMessage(e instanceof Error?'备份失败：'+e.message:'备份失败，请检查游戏配置')}
 }

 async function parseGameBackup(file?:File){
  if(!file)return;
  try{
   let parsed:unknown;try{parsed=JSON.parse(await file.text())}catch{throw new Error('文件不是合法的 JSON')}
   const backup=validateBackup(parsed),tierCount=backup.games.reduce((sum,game)=>sum+game.tiers.length+game.allTiers.length,0),rewardCount=backup.games.reduce((sum,game)=>sum+[...game.tiers,...game.allTiers].reduce((n,tier)=>n+tier.rewards.length,0),0);
   setPendingBackupRestore({backup,fileName:file.name,gameCount:backup.games.length,tierCount,rewardCount});
   setMessage('游戏备份已完成解析和校验，等待确认恢复');
  }catch(e){setPendingBackupRestore(null);setMessage(e instanceof Error?'恢复失败：'+e.message:'恢复失败：备份文件无法读取')}
  finally{if(backupRef.current)backupRef.current.value=''}
 }

 function confirmGameBackupRestore(){
  if(!pendingBackupRestore)return;
  const restoredGames=pendingBackupRestore.backup.games.map(upgradeGame),restoredActiveId=restoredGames.some(game=>game.id===pendingBackupRestore.backup.activeGameId)?pendingBackupRestore.backup.activeGameId:restoredGames[0].id;
  const serialized=JSON.stringify({games:restoredGames,activeGameId:restoredActiveId}),previous=localStorage.getItem(STORAGE_KEY);
  try{
   localStorage.setItem(STORAGE_KEY,serialized);
   const saved=localStorage.getItem(STORAGE_KEY);if(saved!==serialized)throw new Error('写入后校验失败');
   const verified=JSON.parse(saved) as {games?:Game[];activeGameId?:string};
   if(JSON.stringify(verified)!==serialized)throw new Error('重新读取的配置与备份不一致');
   setGames(restoredGames);setActiveGameId(restoredActiveId);setPendingBackupRestore(null);setPendingImport(null);setGameDialog(null);
   setMessage('已完整恢复 '+restoredGames.length+' 个游戏配置');
  }catch(e){
   let rollback='，原配置已保留';
   try{if(previous===null)localStorage.removeItem(STORAGE_KEY);else localStorage.setItem(STORAGE_KEY,previous);if(localStorage.getItem(STORAGE_KEY)!==previous)throw new Error('rollback')}catch{rollback='，且浏览器未能确认原配置回滚，请保留当前页面并重新导入备份'}
   setMessage('恢复失败：'+(e instanceof Error?e.message:'浏览器存储写入失败')+rollback);
  }
 }

 async function loadSource(file?:File){
  if(!file)return;
  try{
   const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}),parsed:Player[]=[],issues:RegionIssue[]=[];
   for(const name of wb.SheetNames){
    const rows=XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name],{header:1,raw:true,defval:null});
    const hr=rows.findIndex(r=>r.some(v=>/角色id|role.?id/i.test(stringValue(v)))&&r.some(v=>/金额|付费|储值|amount/i.test(stringValue(v))));
    if(hr<0)continue;
    const hs=rows[hr],sc=col(hs,[/区服|服务器|server/]),rc=col(hs,[/角色id|roleid|玩家id|playerid/]),ac=col(hs,[/金额|总付费|累计储值|累储|amount/]);
    if(rc<0||ac<0)continue;
    for(let r=hr+1;r<rows.length;r++){const row=rows[r],roleId=stringValue(row[rc]).trim(),amount=money(row[ac]);if(!roleId||!Number.isFinite(amount))continue;const server=sc>=0?stringValue(row[sc]).trim():'',identified=identifyRegion(name,server,sc>=0);if(identified.region)parsed.push({region:identified.region,server,roleId,amount});else issues.push({sheet:name,row:r+1,server,roleId,reason:identified.reason})}
   }
   if(!parsed.length&&!issues.length)throw new Error('未识别到“角色ID”和“金额/总付费”列');
   setPlayers(parsed);setRegionIssues(issues);setSourceName(file.name);setMessage(issues.length?'已读取 '+(parsed.length+issues.length)+' 条有效记录，其中 '+issues.length+' 条未识别或冲突，请修正后重新上传':'已读取 '+parsed.length+' 名玩家，全部大区识别成功');
  }catch(e){setMessage(e instanceof Error?e.message:'文件解析失败')}
 }

 async function loadTemplate(file?:File){
  if(!file)return;
  try{
   const buf=await file.arrayBuffer(),wb=XLSX.read(buf,{type:'array'}),rows=XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''}),first=(rows.find(r=>r.some(Boolean))??[]).map(String);
   setTemplate(buf);setTemplateName(file.name);
   if(first.length>=4)setHeaders(h=>({...h,server:first[0]||h.server,role:first[1]||h.role,item:(first.length>=7?first[5]:first[2])||h.item,quantity:(first.length>=7?first[6]:first[3])||h.quantity}));
   setMessage('发送模板已载入，将保留模板工作表并写入奖励数据');
  }catch{setMessage('发送模板读取失败')}
 }

 function downloadStatsTemplate(){
  const wb=XLSX.utils.book_new();
  for(const name of ['亚洲','美洲','欧洲']){const ws=XLSX.utils.aoa_to_sheet([['区服','角色ID','累计储值']]);ws['!cols']=[{wch:12},{wch:18},{wch:16}];XLSX.utils.book_append_sheet(wb,ws,name)}
  downloadBuffer(XLSX.write(wb,{bookType:'xlsx',type:'array'}) as ArrayBuffer,'累计储值统计表模板.xlsx');setMessage('统计表模板已下载');
 }

 function downloadSendTemplate(){
  downloadBuffer(template?template.slice(0):defaultSendTemplate(),template?templateName:'默认发送模板.xlsx');setMessage('已下载'+(template?'当前自定义':'默认')+'发送模板');
 }

 function downloadTierTemplate(){
  if(!activeGame)return;
  const rows=[['储值金额（USD）','物品ID','数量'],...tiers.flatMap(t=>t.rewards.map(r=>[t.amount,r.item,r.quantity]))];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  for(let r=1;r<rows.length;r++){const cell=ws[XLSX.utils.encode_cell({r,c:1})];if(cell){cell.t='s';cell.v=String(cell.v);cell.z='@'}}
  ws['!cols']=[{wch:20},{wch:22},{wch:12}];
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'累计储值奖档');
  downloadBuffer(XLSX.write(wb,{bookType:'xlsx',type:'array'}) as ArrayBuffer,safeName(activeGame.name)+'_'+RULE_LABEL[activeRule]+'.xlsx');setMessage('已下载 '+activeGame.name+' 的'+RULE_LABEL[activeRule]+'模板');
 }

 async function parseTierTemplate(file?:File){
  if(!file||!activeGame)return;
  try{
   const wb=XLSX.read(await file.arrayBuffer(),{type:'array',raw:true}),ws=wb.Sheets[wb.SheetNames[0]];
   if(!ws)throw new Error('文件中没有可读取的工作表');
   const rows=XLSX.utils.sheet_to_json<unknown[]>(ws,{header:1,raw:true,defval:null});
   const hr=rows.findIndex(r=>col(r,[/储值金额|储值|金额|amount/])>=0&&col(r,[/物品id|itemid|道具id/])>=0&&col(r,[/^数量$|quantity|qty/])>=0);
   if(hr<0)throw new Error('未找到“储值金额（USD） / 物品ID / 数量”表头');
   const hs=rows[hr],ac=col(hs,[/储值金额|储值|金额|amount/]),ic=col(hs,[/物品id|itemid|道具id/]),qc=col(hs,[/^数量$|quantity|qty/]);
   const grouped=new Map<number,Tier>(),ordered:Tier[]=[];let rewardCount=0;
   for(let r=hr+1;r<rows.length;r++){
    const row=rows[r],a=row?.[ac],i=row?.[ic],q=row?.[qc];
    if([a,i,q].every(v=>v===null||v===undefined||stringValue(v).trim()===''))continue;
    const line=r+1,amount=money(a);
    if(!Number.isFinite(amount)||amount<=0)throw new Error('第 '+line+' 行：储值金额必须是有效的正数');
    const itemCell=ws[XLSX.utils.encode_cell({r,c:ic})],itemRaw=itemCell?.v??i;let item='';
    if(itemCell?.t==='n'||typeof itemRaw==='number'){const n=Number(itemRaw),digits=String(Math.trunc(Math.abs(n))).length,display=String(itemCell?.w??'').trim();if(!Number.isSafeInteger(n)||digits>15)throw new Error('第 '+line+' 行：物品 ID 为数值且可能已被 Excel 损失精度，请将该单元格设为文本后重试');item=/^\d+$/.test(display)&&Number(display)===n?display:String(n)}else item=String(itemRaw??'').trim();
    if(!item)throw new Error('第 '+line+' 行：物品 ID 不能为空');
    const quantity=typeof q==='number'?q:Number(stringValue(q).trim());
    if(!Number.isInteger(quantity)||quantity<=0)throw new Error('第 '+line+' 行：数量必须是有效的正整数');
    let tier=grouped.get(amount);if(!tier){tier={amount,rewards:[]};grouped.set(amount,tier);ordered.push(tier)}
    tier.rewards.push({item,quantity});rewardCount++;
   }
   if(!ordered.length)throw new Error('没有识别到有效的奖档记录');
   setPendingImport({gameId:activeGame.id,gameName:activeGame.name,rule:activeRule,tiers:ordered,tierCount:ordered.length,rewardCount,fileName:file.name});
   setMessage('已校验 '+ordered.length+' 个档位、'+rewardCount+' 条物品记录，等待确认导入');
  }catch(e){setPendingImport(null);setMessage(e instanceof Error?e.message:'奖档模板解析失败')}
  finally{if(tierTemplateRef.current)tierTemplateRef.current.value=''}
 }

 function confirmTierImport(){
  if(!pendingImport)return;
  setGames(current=>current.map(g=>g.id===pendingImport.gameId?{...g,[pendingImport.rule==='all'?'allTiers':'tiers']:cloneTiers(pendingImport.tiers)}:g));
  setMessage('已将 '+pendingImport.tierCount+' 个档位导入 '+pendingImport.gameName+' 的'+RULE_LABEL[pendingImport.rule]);setPendingImport(null);
 }

 function makeBook(rg:string){
  const rows=outputs.filter(r=>r.region===rg),rewardCount=Math.max(1,...rows.map(r=>r.rewards.length)),lastCol=XLSX.utils.encode_col(4+rewardCount*2);
  let wb:XLSX.WorkBook,ws:XLSX.WorkSheet;
  const title='Recharge Reward',content='Dear Adventurer, \nThis is your offline cumulative recharge reward. Please claim it promptly. Thank you for your support!',rewardHeaders=Array.from({length:rewardCount},()=>[headers.item||'物品',headers.quantity||'数量']).flat(),rowValues=(r:OutputRow,mailTitle:string,mailContent:string)=>[Number(r.server.match(/(\d+)$/)?.[1]??r.server),r.roleId,'累储：'+r.amount.toFixed(2),mailTitle,mailContent,...r.rewards.flatMap(x=>[x.item,x.quantity])];
  if(template){
   wb=XLSX.read(template.slice(0),{type:'array'});ws=wb.Sheets[wb.SheetNames[0]];
   const data=XLSX.utils.sheet_to_json<unknown[]>(ws,{header:1,defval:''}),example=data[1]??[];
   if(ws['!ref']){const ref=XLSX.utils.decode_range(ws['!ref']);for(let rr=1;rr<=ref.e.r;rr++)for(let cc=0;cc<=ref.e.c;cc++)delete ws[XLSX.utils.encode_cell({r:rr,c:cc})];for(let cc=5;cc<=ref.e.c;cc++)delete ws[XLSX.utils.encode_cell({r:0,c:cc})]}
   XLSX.utils.sheet_add_aoa(ws,[rewardHeaders],{origin:{r:0,c:5}});
   rows.forEach((r,i)=>XLSX.utils.sheet_add_aoa(ws,[rowValues(r,stringValue(example[3])||title,stringValue(example[4])||content)],{origin:{r:i+1,c:0}}));
   if(ws['!cols'])ws['!cols']=[...ws['!cols'].slice(0,5),...Array.from({length:rewardCount*2},(_,i)=>ws['!cols']?.[5+i]??{wch:9})];
   ws['!ref']='A1:'+lastCol+Math.max(2,rows.length+1);
  }else{
   const hs=['区服','角色ID','','标题','内容',...rewardHeaders];ws=XLSX.utils.aoa_to_sheet([hs,...rows.map(r=>rowValues(r,title,content))]);ws['!cols']=[{wch:10},{wch:14},{wch:14},{wch:20},{wch:52},...Array.from({length:rewardCount*2},()=>({wch:9}))];wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Sheet1');
  }
  return XLSX.write(wb,{bookType:'xlsx',type:'array'}) as ArrayBuffer;
 }

 async function download(){
  if(regionIssues.length){setMessage('存在未识别或冲突记录，请修正统计表后重新上传');return}
  if(!players.length){setMessage('请先上传累计储值表');return}
  if(!validTiers(tiers)){setMessage('请检查奖档及物品配置');return}
  const zip=new JSZip();['亚洲','美洲','欧洲'].forEach(r=>zip.file('累计储值奖励_'+r+'.xlsx',makeBook(r)));
  const blob=await zip.generateAsync({type:'blob'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='累计储值奖励_三大区.zip';a.click();URL.revokeObjectURL(url);setMessage('三个大区的发奖文件已生成');
 }

 if(!hydrated)return <main className="min-h-screen bg-[#f5f7fb]"/>;

 return <main className="min-h-screen bg-[#f5f7fb] text-slate-900"><div className="mx-auto max-w-7xl px-4 py-8 sm:px-6 lg:px-8">
  <header className="mb-8 flex flex-col justify-between gap-4 border-b border-slate-200 pb-7 md:flex-row md:items-end"><div><p className="mb-2 text-sm font-semibold tracking-wide text-indigo-600">风色传说 · 运营工具</p><h1 className="text-3xl font-bold tracking-tight sm:text-4xl">累计储值奖励生成器</h1><p className="mt-2 text-slate-500">上传统计表，核对奖档，一次导出亚洲、美洲、欧洲三份发奖文件。</p></div><div className="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-800"><CheckCircle2 className="mr-2 inline h-4 w-4"/>数据仅在当前浏览器处理</div></header>

  <Card className="mb-6"><CardHeader><CardTitle className="flex items-center gap-2 text-lg"><Gamepad2 className="h-5 w-5 text-indigo-600"/>游戏管理</CardTitle></CardHeader><CardContent><div className="flex flex-col gap-3 sm:flex-row sm:flex-wrap sm:items-end"><div className="min-w-0 flex-1 space-y-2"><Label htmlFor="game-select">选择游戏</Label><select id="game-select" className="flex h-9 w-full rounded-md border border-input bg-transparent px-3 py-1 text-sm shadow-xs outline-none focus-visible:border-ring focus-visible:ring-3 focus-visible:ring-ring/50" value={activeGameId} onChange={e=>{setActiveGameId(e.target.value);setPendingImport(null);setGameDialog(null);const game=games.find(g=>g.id===e.target.value);setMessage('已切换到游戏：'+(game?.name??'')+'（'+RULE_LABEL[game?.rewardRule??'highest']+'）')}}>{games.map(g=><option key={g.id} value={g.id}>{g.name}</option>)}</select></div><Button variant="outline" onClick={()=>setGameDialog({mode:'add',value:''})}><Plus className="mr-1 h-4 w-4"/>新增游戏</Button><Button variant="outline" onClick={()=>activeGame&&setGameDialog({mode:'rename',value:activeGame.name})}><Pencil className="mr-1 h-4 w-4"/>重命名</Button><Button variant="outline" onClick={backupAllGames}><Download className="mr-1 h-4 w-4"/>备份全部游戏</Button><input ref={backupRef} className="hidden" type="file" accept=".json,application/json" onChange={e=>parseGameBackup(e.target.files?.[0])}/><Button variant="outline" onClick={()=>backupRef.current?.click()}><Upload className="mr-1 h-4 w-4"/>恢复游戏备份</Button></div>{gameDialog&&<div className="mt-4 rounded-xl border border-indigo-200 bg-indigo-50 p-4"><Label htmlFor="game-name">{gameDialog.mode==='add'?'新游戏名称':'新的游戏名称'}</Label><div className="mt-2 flex flex-col gap-2 sm:flex-row"><Input id="game-name" value={gameDialog.value} onChange={e=>setGameDialog({...gameDialog,value:e.target.value})} onKeyDown={e=>{if(e.key==='Enter')submitGameDialog()}}/><Button onClick={submitGameDialog}>{gameDialog.mode==='add'?'确认新增':'确认重命名'}</Button><Button variant="outline" onClick={()=>setGameDialog(null)}>取消</Button></div></div>}{pendingBackupRestore&&<div className="mt-4 rounded-xl border border-amber-200 bg-amber-50 p-4"><div className="font-semibold text-amber-950">全部游戏备份已完成解析和校验</div><div className="mt-2 grid gap-1 text-sm text-amber-900 sm:grid-cols-2 lg:grid-cols-4"><span>创建时间：{new Date(pendingBackupRestore.backup.createdAt).toLocaleString()}</span><span>游戏：{pendingBackupRestore.gameCount} 个</span><span>奖档：{pendingBackupRestore.tierCount} 个</span><span>物品记录：{pendingBackupRestore.rewardCount} 条</span></div><p className="mt-3 text-sm text-amber-950">恢复备份将替换当前浏览器保存的全部游戏及累计储值奖档配置，是否确认？</p><div className="mt-3 flex gap-2"><Button size="sm" onClick={confirmGameBackupRestore}>确认恢复</Button><Button variant="outline" size="sm" onClick={()=>{setPendingBackupRestore(null);setMessage('已取消恢复，当前游戏配置未修改')}}>取消</Button></div></div>}</CardContent></Card>

  <section className="grid gap-5 lg:grid-cols-[1fr_1fr_1.1fr]">
   <Card><CardHeader><CardTitle className="flex items-center gap-2 text-lg"><span className="step">1</span>上传统计表</CardTitle></CardHeader><CardContent><input ref={sourceRef} className="hidden" type="file" accept=".xlsx,.xls,.csv" onChange={e=>loadSource(e.target.files?.[0])}/><button onClick={()=>sourceRef.current?.click()} className="upload-zone"><Upload className="h-7 w-7 text-indigo-600"/><span className="font-semibold">选择累计储值表</span><span className="text-xs text-slate-500">支持 XLSX、XLS、CSV</span></button><p className="mt-3 truncate text-sm text-slate-600">{sourceName||'尚未选择文件'}</p><Button variant="outline" size="sm" className="mt-3 w-full" onClick={downloadStatsTemplate}><Download className="mr-1 h-4 w-4"/>下载统计表模板</Button></CardContent></Card>
   <Card><CardHeader><CardTitle className="flex items-center gap-2 text-lg"><span className="step">2</span>发送模板</CardTitle></CardHeader><CardContent><input ref={templateRef} className="hidden" type="file" accept=".xlsx,.xls" onChange={e=>loadTemplate(e.target.files?.[0])}/><button onClick={()=>templateRef.current?.click()} className="upload-zone"><FileSpreadsheet className="h-7 w-7 text-indigo-600"/><span className="font-semibold">上传或替换模板</span><span className="text-xs text-slate-500">不上传则使用默认格式</span></button><p className="mt-3 truncate text-sm text-slate-600">{templateName}</p><Button variant="outline" size="sm" className="mt-3 w-full" onClick={downloadSendTemplate}><Download className="mr-1 h-4 w-4"/>下载当前发送模板</Button></CardContent></Card>
   <Card className="bg-slate-950 text-white"><CardHeader><CardTitle className="flex items-center gap-2 text-lg"><span className="step dark-step">3</span>生成文件</CardTitle></CardHeader><CardContent><div className="grid grid-cols-3 gap-2">{stats.map(s=><div key={s.region} className="rounded-lg bg-white/7 p-3"><div className="text-xs text-slate-400">{s.region}</div><div className="mt-1 text-xl font-bold">{s.players}</div><div className="text-[11px] text-slate-400">玩家 / {s.rewards} 条奖励</div></div>)}</div><div className="mt-3 rounded-lg bg-white/7 px-3 py-2 text-xs text-slate-300">识别记录：亚洲 {regionCounts.亚洲} · 美洲 {regionCounts.美洲} · 欧洲 {regionCounts.欧洲} · 未识别/冲突 {regionIssues.length}</div><Button onClick={download} disabled={regionIssues.length>0} className="mt-4 w-full bg-indigo-500 text-white hover:bg-indigo-400"><Download className="mr-2 h-4 w-4"/>下载三大区 ZIP</Button><p className="mt-3 text-xs text-slate-300" aria-live="polite">{message}</p></CardContent></Card>
  </section>

  {regionIssues.length>0&&<Card className="mt-6 border-rose-200 bg-rose-50"><CardHeader><CardTitle className="text-lg text-rose-950">未识别或冲突记录（{regionIssues.length}）</CardTitle><p className="text-sm text-rose-800">以下记录不会分配到任何大区，已暂停生成 ZIP。请修正统计表后重新上传。</p></CardHeader><CardContent><div className="overflow-x-auto"><table className="w-full min-w-[760px] text-left text-sm"><thead><tr className="border-b border-rose-200 text-rose-950"><th className="px-2 py-2">工作表</th><th className="px-2 py-2">原始行号</th><th className="px-2 py-2">服务器名称</th><th className="px-2 py-2">角色 ID</th><th className="px-2 py-2">原因</th></tr></thead><tbody>{regionIssues.map((issue,index)=><tr key={issue.sheet+'-'+issue.row+'-'+index} className="border-b border-rose-100 text-rose-900"><td className="px-2 py-2">{issue.sheet}</td><td className="px-2 py-2">{issue.row}</td><td className="px-2 py-2">{issue.server||'（空）'}</td><td className="px-2 py-2">{issue.roleId}</td><td className="px-2 py-2">{issue.reason}</td></tr>)}</tbody></table></div></CardContent></Card>}

  <Card className="mt-6"><CardHeader className="gap-4 sm:flex-row sm:items-center sm:justify-between"><div><CardTitle className="flex flex-wrap items-center gap-2"><span className="flex items-center gap-2"><Settings2 className="h-5 w-5 text-indigo-600"/>累计储值奖档</span><span className="inline-flex rounded-lg border border-slate-200 bg-slate-100 p-1 text-sm">{(['highest','all'] as RewardRule[]).map(rule=><button key={rule} type="button" className={'rounded-md px-3 py-1.5 transition '+(activeRule===rule?'bg-white font-semibold text-indigo-700 shadow-sm':'text-slate-500 hover:text-slate-800')} aria-pressed={activeRule===rule} onClick={()=>{setGames(current=>current.map(g=>g.id===activeGameId?{...g,rewardRule:rule}:g));setPendingImport(null);setMessage('已切换为'+RULE_LABEL[rule])}}>{RULE_LABEL[rule]}</button>)}</span></CardTitle><p className="mt-1 text-sm text-slate-500">当前游戏：{activeGame?.name}。{activeRule==='highest'?'每条玩家记录只发达到的最高档位。':'每个达到的档位分别生成一封邮件。'}每个档位可配置多个物品。</p></div><div className="flex flex-wrap gap-2"><input ref={tierTemplateRef} className="hidden" type="file" accept=".xlsx" onChange={e=>parseTierTemplate(e.target.files?.[0])}/><Button variant="outline" size="sm" onClick={downloadTierTemplate}><Download className="mr-1 h-4 w-4"/>下载奖档模板</Button><Button variant="outline" size="sm" onClick={()=>tierTemplateRef.current?.click()}><Upload className="mr-1 h-4 w-4"/>上传奖档模板</Button><Button variant="outline" size="sm" onClick={()=>updateCurrentTiers(current=>[...current,{amount:0,rewards:[{item:'',quantity:1}]}])}><Plus className="mr-1 h-4 w-4"/>添加档位</Button></div></CardHeader><CardContent>
   {pendingImport&&<div className="mb-5 rounded-xl border border-indigo-200 bg-indigo-50 p-4"><div className="font-semibold text-indigo-950">奖档模板已完成解析和校验</div><div className="mt-2 grid gap-1 text-sm text-indigo-900 sm:grid-cols-4"><span>目标游戏：{pendingImport.gameName}</span><span>奖励规则：{RULE_LABEL[pendingImport.rule]}</span><span>识别档位：{pendingImport.tierCount} 个</span><span>物品记录：{pendingImport.rewardCount} 条</span></div><p className="mt-3 text-sm text-indigo-900">导入奖档模板将只替换该游戏的{RULE_LABEL[pendingImport.rule]}配置，是否确认？</p><div className="mt-3 flex gap-2"><Button size="sm" onClick={confirmTierImport}>确认导入</Button><Button variant="outline" size="sm" onClick={()=>{setPendingImport(null);setMessage('已取消导入，原配置未修改')}}>取消</Button></div></div>}
   <div className="space-y-4">{tiers.map((t,i)=><div key={i} className="rounded-xl border border-slate-200 bg-slate-50/60 p-4"><div className="mb-3 flex flex-wrap items-end justify-between gap-3"><div className="w-full max-w-xs space-y-2"><Label>储值金额（USD）</Label><Input type="number" min="0" value={t.amount} onChange={e=>updateCurrentTiers(current=>current.map((x,j)=>j===i?{...x,amount:Number(e.target.value)}:x))}/></div><Button variant="ghost" size="sm" aria-label="删除档位" onClick={()=>updateCurrentTiers(current=>current.filter((_,j)=>j!==i))}><Trash2 className="mr-1 h-4 w-4 text-slate-400"/>删除档位</Button></div><div className="space-y-2"><div className="grid grid-cols-[1fr_1fr_40px] gap-3 px-1 text-sm font-medium text-slate-600"><span>物品 ID</span><span>数量</span><span/></div>{t.rewards.map((reward,k)=><div key={k} className="grid grid-cols-[1fr_1fr_40px] gap-3"><Input value={reward.item} placeholder="输入物品 ID" onChange={e=>updateCurrentTiers(current=>current.map((x,j)=>j===i?{...x,rewards:x.rewards.map((r,n)=>n===k?{...r,item:e.target.value}:r)}:x))}/><Input type="number" min="1" value={reward.quantity} onChange={e=>updateCurrentTiers(current=>current.map((x,j)=>j===i?{...x,rewards:x.rewards.map((r,n)=>n===k?{...r,quantity:Number(e.target.value)}:r)}:x))}/><Button variant="ghost" size="icon" aria-label="删除物品" disabled={t.rewards.length===1} onClick={()=>updateCurrentTiers(current=>current.map((x,j)=>j===i?{...x,rewards:x.rewards.filter((_,n)=>n!==k)}:x))}><Trash2 className="h-4 w-4 text-slate-400"/></Button></div>)}<Button variant="outline" size="sm" className="mt-2" onClick={()=>updateCurrentTiers(current=>current.map((x,j)=>j===i?{...x,rewards:[...x.rewards,{item:'',quantity:1}]}:x))}><Plus className="mr-1 h-4 w-4"/>添加物品</Button></div></div>)}</div>
  </CardContent></Card>

  <Card className="mt-6"><CardHeader><CardTitle className="text-lg">输出列名</CardTitle><p className="text-sm text-slate-500">可按发送系统要求修改。上传模板时，名称相同的列会自动填充。</p></CardHeader><CardContent className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">{Object.entries(headers).map(([k,v])=><div key={k} className="space-y-2"><Label>{({server:'区服列',role:'角色列',item:'物品列',quantity:'数量列',amount:'累计金额列',tier:'达成档位列'} as Record<string,string>)[k]}</Label><Input value={v} onChange={e=>setHeaders({...headers,[k]:e.target.value})}/></div>)}</CardContent></Card>
  <section className="mt-6 rounded-xl border border-slate-200 bg-white p-5 text-sm text-slate-600"><strong className="text-slate-900">识别规则</strong><span className="ml-2">优先根据区服/服务器列识别 Asia、AS、亚洲、America、US、USA、NA、美洲、北美、Europe、EU、欧洲；服务器为空时再使用工作表名称。未识别或冲突记录会暂停生成。低于最低档位的玩家不会写入。</span></section>
  <footer className="pt-5 text-center text-xs text-slate-400">当前版本：{SITE_VERSION.version} ｜ 最后更新时间：{SITE_VERSION.releasedAt}</footer>
 </div></main>;
}
