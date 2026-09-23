const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');

const app = express();
const PORT = process.env.PORT || 3050;
const BUILD_VERSION='V1.4.2';

const ANNUAL_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'annual');
const MONTHLY_SOURCE_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'monthly_source');
const MONTHLY_WORK_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'monthly_work');
const MONTHLY_TEMPLATE_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'monthly_templates');
const SYSTEM_TEMPLATE_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'system_templates');
const BLANK_ANNUAL_TEMPLATE = path.join(SYSTEM_TEMPLATE_DIR, 'Smart_PM_BLANK_ANNUAL_TEMPLATE.xlsx');
const BLANK_MONTHLY_TEMPLATE = path.join(SYSTEM_TEMPLATE_DIR, 'Smart_PM_BLANK_MONTHLY_TEMPLATE.xlsx');
const EXCEL_UPLOAD_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'excel_uploads');
for(const d of [ANNUAL_DIR,MONTHLY_WORK_DIR,MONTHLY_TEMPLATE_DIR,EXCEL_UPLOAD_DIR,SYSTEM_TEMPLATE_DIR]){if(!fs.existsSync(d))fs.mkdirSync(d,{recursive:true});}

const TEST_LINE = 'Layer Cake';
const DEFAULT_TEST_DATE = new Date(2026, 6, 3, 12, 0, 0); // initial test
function getPmDate(req){
  const token=(req && req.headers && req.headers.authorization || '').replace(/^Bearer\s+/i,'');
  const selected=token ? pmViewByToken.get(token) : null;
  if(selected && /^\d{4}-\d{2}-\d{2}$/.test(selected)){
    return new Date(selected+'T12:00:00');
  }

  // Default view after server start/login:
  // always use the real current PM week from the computer/server date.
  // PM week starts Friday and ends Thursday.
  const now=new Date();
  const currentPmFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
  const daysSinceFriday=(currentPmFriday.getDay()-5+7)%7;
  currentPmFriday.setDate(currentPmFriday.getDate()-daysSinceFriday);
  return currentPmFriday;
}

const CONFIG_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'config');
const BACKUP_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'backups');
const BRANDING_DIR = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'branding');
for(const d of [CONFIG_DIR,BACKUP_DIR,BRANDING_DIR])if(!fs.existsSync(d))fs.mkdirSync(d,{recursive:true});
const sessions = new Map();
const pmViewByToken = new Map();

// Optional PM issue/comments log. Reporting an issue does NOT complete the PM task.
const PM_ISSUES_FILE = path.join(process.env.PM_DATA_ROOT || path.join(__dirname,'data'), 'pm_issues.json');
function loadPmIssues(){
  try{
    if(!fs.existsSync(PM_ISSUES_FILE)) return {version:1,items:[]};
    const x=JSON.parse(fs.readFileSync(PM_ISSUES_FILE,'utf8'));
    return {version:1,items:Array.isArray(x.items)?x.items:[]};
  }catch(_){ return {version:1,items:[]}; }
}
function savePmIssues(x){
  fs.writeFileSync(PM_ISSUES_FILE,JSON.stringify(x,null,2),'utf8');
}
function pmIssueKey(line,taskId,weekHeader){ return `${line}||${taskId}||${weekHeader}`; }
function latestPmIssue(line,taskId,weekHeader){
  const key=pmIssueKey(line,taskId,weekHeader);
  const items=loadPmIssues().items.filter(x=>x.key===key);
  return items.length?items[items.length-1]:null;
}


function pmDataStartRow(ws){
  for(let r=1;r<=Math.min(20,ws.rowCount||20);r++){
    const h=String(ws.getCell(r,8).value||'').trim().toLowerCase();
    if(h==='row type')return r+1;
  }
  return 5;
}

function setThinBorder(cell){
  cell.border={
    top:{style:'thin',color:{argb:'FFB8C2CC'}},
    bottom:{style:'thin',color:{argb:'FFB8C2CC'}},
    left:{style:'thin',color:{argb:'FFB8C2CC'}},
    right:{style:'thin',color:{argb:'FFB8C2CC'}}
  };
}
function styleRange(ws,r1,c1,r2,c2,{fill=null,fontColor='FF000000',bold=false,size=10,center=true}={}){
  for(let r=r1;r<=r2;r++){
    for(let c=c1;c<=c2;c++){
      const cell=ws.getCell(r,c);
      if(fill)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:fill}};
      cell.font={bold,color:{argb:fontColor},size};
      cell.alignment={horizontal:center?'center':'left',vertical:'middle',wrapText:true};
      setThinBorder(cell);
    }
  }
}
function buildProfessionalAnnualWorkbook(line='',year=new Date().getFullYear()){
  const wb=new ExcelJS.Workbook();
  wb.creator='Smart PM System';
  wb.created=new Date();
  const ws=wb.addWorksheet('Full',{views:[{state:'frozen',ySplit:6,xSplit:3}]});

  ws.mergeCells('A1:C1'); ws.mergeCells('D1:O1'); ws.mergeCells('P1:BO1');
  ws.getCell('A1').value='COMPANY LOGO';
  ws.getCell('D1').value='COMPANY NAME';
  ws.getCell('P1').value='ANNUAL PREVENTIVE MAINTENANCE MASTER PLAN';
  styleRange(ws,1,1,1,3,{fill:'FFEAF2F8',fontColor:'FF5B6573',bold:true,size:10});
  styleRange(ws,1,4,1,15,{fill:'FF1F4E78',fontColor:'FFFFFFFF',bold:true,size:15});
  styleRange(ws,1,16,1,67,{fill:'FF1F4E78',fontColor:'FFFFFFFF',bold:true,size:14});
  ws.getRow(1).height=34;

  const kpi=['Total Machines','Total PM Tasks','Weekly','Monthly','Quarterly','Semi Annual','Annual','Running Hours'];
  kpi.forEach((v,i)=>ws.getCell(2,i+1).value=v);
  styleRange(ws,2,1,2,8,{fill:'FFD9EAF7',bold:true,size:10});
  for(let c=1;c<=8;c++)ws.getCell(3,c).value=0;
  styleRange(ws,3,1,3,8,{fill:'FFF7FBFE',bold:true,size:12});

  const meta=['Line Name','Annual Year','Prepared By','Approved By','Revision','Last Update'];
  meta.forEach((v,i)=>ws.getCell(2,10+i).value=v);
  styleRange(ws,2,10,2,15,{fill:'FFE2F0D9',bold:true,size:10});
  ws.getCell('J3').value=line; ws.getCell('K3').value=year;
  styleRange(ws,3,10,3,15,{fill:'FFFFFFFF',size:10});

  const base=['Section','Machine No.','Machine Name','Machine Identified','Manufacturer','Part','Maintenance',
              'Row Type','PM Code','Frequency Value','Frequency Text','Calculation Mode','PLC Tag / Counter',
              'First PM Date','Notes'];
  const weeks=Array.from({length:52},(_,i)=>`W${i+1}`);
  [...base,...weeks].forEach((v,i)=>ws.getCell(6,i+1).value=v);
  styleRange(ws,6,1,6,67,{fill:'FF5B9BD5',fontColor:'FFFFFFFF',bold:true,size:9});
  ws.getRow(6).height=30;

  const widths=[16,13,22,18,14,18,38,12,9,13,16,16,20,14,24];
  widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
  // User-facing PM table: A:G + K (Freq.). Keep technical columns hidden for system compatibility.
  [8,9,10,12,13,14,15].forEach(c=>ws.getColumn(c).hidden=true);
  ws.getCell(6,11).value='Freq.';
  for(let c=16;c<=67;c++)ws.getColumn(c).width=8;

  // Pre-format blank task area; no machine/task data.
  // This is the REAL template used when a new production line is created.
  for(let r=7;r<=506;r++){
    ws.getRow(r).height=18;
    for(let c=1;c<=67;c++){
      const cell=ws.getCell(r,c);
      cell.alignment={vertical:'middle',wrapText:true};
      setThinBorder(cell);
      if(c>=16){
        cell.fill={type:'pattern',pattern:'none'};
        cell.numFmt=';;;'; // Annual PM schedule shows color only, never the numeric code.
      }
    }
  }

  ws.getCell(3,1).value={formula:'IFERROR(SUMPRODUCT((B7:B506<>"")/COUNTIFS(B7:B506,B7:B506&"")),0)'};
  ws.getCell(3,2).value={formula:'COUNTIF(H7:H506,"planed")'};
  ws.getCell(3,3).value={formula:'COUNTIFS(H7:H506,"planed",K7:K506,"Weekly")'};
  ws.getCell(3,4).value={formula:'COUNTIFS(H7:H506,"planed",K7:K506,"Monthly")'};
  ws.getCell(3,5).value={formula:'COUNTIFS(H7:H506,"planed",K7:K506,"Quarterly")'};
  ws.getCell(3,6).value={formula:'COUNTIFS(H7:H506,"planed",K7:K506,"Semi Annual")'};
  ws.getCell(3,7).value={formula:'COUNTIFS(H7:H506,"planed",K7:K506,"Annual")'};
  ws.getCell(3,8).value={formula:'COUNTIFS(H7:H506,"planed",L7:L506,"Hours")'};

  prepareAnnualCalendarHeaders(ws,line,year);
  formatProfessionalAnnualLayout(ws,7,7);
  // New Annual file is now born with the same color-only schedule format used later by Add/Edit.
  for(let r=7;r<=506;r++){
    for(let c=16;c<=67;c++){
      ws.getCell(r,c).fill={type:'pattern',pattern:'none'};
      ws.getCell(r,c).numFmt=';;;';
    }
  }
  return {wb,ws};
}
function buildProfessionalMonthlyWorkbook(line='',year=new Date().getFullYear(),month=1){
  const wb=new ExcelJS.Workbook();
  wb.creator='Smart PM System';
  wb.created=new Date();
  const ws=wb.addWorksheet('Sheet1',{views:[{state:'frozen',ySplit:6,xSplit:3}]});

  ws.mergeCells('A1:C1'); ws.mergeCells('D1:J1'); ws.mergeCells('K1:AK1');
  ws.getCell('A1').value='COMPANY LOGO';
  ws.getCell('D1').value='COMPANY NAME';
  ws.getCell('K1').value='MONTHLY PREVENTIVE MAINTENANCE EXECUTION REPORT';
  styleRange(ws,1,1,1,3,{fill:'FFEAF2F8',fontColor:'FF5B6573',bold:true,size:10});
  styleRange(ws,1,4,1,10,{fill:'FF1F4E78',fontColor:'FFFFFFFF',bold:true,size:15});
  styleRange(ws,1,11,1,37,{fill:'FF1F4E78',fontColor:'FFFFFFFF',bold:true,size:13});
  ws.getRow(1).height=34;

  const kpi=['Total Machines','Planned Tasks','Done Tasks','Remaining','Achievement %','Calendar Tasks','Hours Tasks','PLC Tasks'];
  kpi.forEach((v,i)=>ws.getCell(2,i+1).value=v);
  styleRange(ws,2,1,2,8,{fill:'FFD9EAF7',bold:true,size:10});
  for(let c=1;c<=8;c++)ws.getCell(3,c).value=0;
  styleRange(ws,3,1,3,8,{fill:'FFF7FBFE',bold:true,size:12});
  ws.getCell('E3').numFmt='0.0%';

  const wk=['W1 Plan','W1 Done','W1 %','W2 Plan','W2 Done','W2 %','W3 Plan','W3 Done','W3 %',
            'W4 Plan','W4 Done','W4 %','W5 Plan','W5 Done','W5 %','Month'];
  wk.forEach((v,i)=>ws.getCell(2,10+i).value=v);
  styleRange(ws,2,10,2,25,{fill:'FFE2F0D9',bold:true,size:9});
  styleRange(ws,3,10,3,25,{fill:'FFFFFFFF',bold:true,size:10});

  ws.mergeCells('P4:T4');
  ws.getCell('P4').value='PM WEEK SCHEDULE';
  styleRange(ws,4,16,4,20,{fill:'FFD9EAF7',bold:true,size:10});

  ws.mergeCells('U4:Y4');
  ws.getCell('U4').value='EXECUTED BY / EXECUTION DATE';
  styleRange(ws,4,21,4,25,{fill:'FFFCE4D6',bold:true,size:10});

  const headers=['Section','Machine No.','Machine Name','Machine Identified','Manufacturer','Part','Maintenance',
    'Row Type','PM Code','Frequency Value','Frequency Text','Calculation Mode','PLC Tag / Counter','First PM Date','Notes',
    'W1','W2','W3','W4','W5','Audit W1','Audit W2','Audit W3','Audit W4','Audit W5'];
  headers.forEach((v,i)=>ws.getCell(6,i+1).value=v);
  styleRange(ws,6,1,6,25,{fill:'FF5B9BD5',fontColor:'FFFFFFFF',bold:true,size:9});
  ws.getRow(6).height=30;

  const widths=[16,13,22,18,14,18,38,12,9,13,16,16,20,14,24];
  widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
  // User-facing PM table: A:G + K (Freq.). Keep technical columns hidden for system compatibility.
  [8,9,10,12,13,14,15].forEach(c=>ws.getColumn(c).hidden=true);
  ws.getCell(6,11).value='Freq.';
  for(let c=16;c<=20;c++)ws.getColumn(c).width=9;
  for(let c=21;c<=25;c++)ws.getColumn(c).width=18;

  // Hidden calculation area: AL:AW
  for(let c=38;c<=49;c++)ws.getColumn(c).width=10;

  // Pre-format the REAL Monthly template used when a new month is created.
  for(let r=7;r<=507;r++){
    ws.getRow(r).height=18;
    for(let c=1;c<=25;c++){
      const cell=ws.getCell(r,c);
      cell.alignment={vertical:'middle',wrapText:true};
      setThinBorder(cell);

      if(c>=16 && c<=20){
        // W1-W5 starts with NO stale fill and NO visible numeric value.
        // After tasks are populated, normalizeMonthlyScheduleDisplay() decides:
        // planned row = PM color only; done row = green check for 6; zero = blank.
        cell.fill={type:'pattern',pattern:'none'};
        cell.numFmt=';;;';
      }
      if(c>=21 && c<=25){
        cell.fill={type:'pattern',pattern:'none'};
      }
    }
  }

  ws.getCell(3,25).value=`${line} — ${monthShort(month)} ${year}`;
  restoreMonthlyKpiHeaders(ws);
  rebuildMonthlyTaskFormulas(ws);
  hideMonthlyHelperColumns(ws);
  formatProfessionalMonthlyLayout(ws,7,7);
  forceExcelRecalculation(wb);
  return {wb,ws};
}
function brandingLogoFullPath(branding){
  const p=String(branding?.companyLogo||branding?.leftLogo||branding?.rightLogo||'').trim();
  if(!p)return '';
  if(path.isAbsolute(p))return p;
  return path.join(__dirname,p.replace(/^[/\\]+/,''));
}
function excelImageExtension(file){
  const e=path.extname(file||'').toLowerCase();
  if(e==='.png')return 'png';
  if(e==='.jpg'||e==='.jpeg')return 'jpeg';
  return '';
}
function brandingDefaults(){
  return {
    companyName:'Smart PM Maintenance',
    factoryName:'',
    headerTitle:'Preventive Maintenance Management System',
    theme:{primary:'#c90000',header:'#c90000',background:'#f4f7fb',card:'#ffffff'},
    logoMode:'left',companyLogo:'',leftLogo:'',rightLogo:'',
    logos:[
      {id:'logo1',file:'',position:'left',size:'medium',width:92,height:48,enabled:true},
      {id:'logo2',file:'',position:'center',size:'medium',width:92,height:48,enabled:false},
      {id:'logo3',file:'',position:'right',size:'medium',width:92,height:48,enabled:false}
    ],
    // Excel has its own logo sizing.  The logo is kept completely inside the
    // merged A1:C1 cell and row 1 grows automatically with the selected height.
    excelLogo:{slot:1,width:160,height:80,autoFitCell:true,padding:8},
    headerZones:{
      left:{logoSlot:1,showLogo:true,showCompany:false,showFactory:false,showHeaderTitle:false,showPageTitle:false,align:'left',logoWidth:92,logoHeight:48,companySize:'medium',factorySize:'small',headerTitleSize:'small',pageTitleSize:'large'},
      center:{logoSlot:2,showLogo:true,showCompany:true,showFactory:true,showHeaderTitle:true,showPageTitle:true,align:'center',logoWidth:92,logoHeight:48,companySize:'medium',factorySize:'small',headerTitleSize:'small',pageTitleSize:'large'},
      right:{logoSlot:3,showLogo:true,showCompany:false,showFactory:false,showHeaderTitle:false,showPageTitle:false,align:'right',logoWidth:92,logoHeight:48,companySize:'medium',factorySize:'small',headerTitleSize:'small',pageTitleSize:'large'}
    },
    textLayout:{
      company:{position:'left',size:'medium',enabled:true},
      factory:{position:'left',size:'small',enabled:true,mode:'underCompany'},
      pageTitle:{position:'center',size:'large',enabled:true},
      headerTitle:{position:'center',size:'small',enabled:true}
    }
  };
}
function normalizeLogoSlot(slot,index){
  const sizes={small:{width:64,height:34},medium:{width:92,height:48},large:{width:130,height:68}};
  const size=['small','medium','large','custom'].includes(String(slot?.size||'').toLowerCase())?String(slot.size).toLowerCase():'medium';
  const preset=sizes[size]||sizes.medium;
  const position=['left','center','right'].includes(String(slot?.position||'').toLowerCase())?String(slot.position).toLowerCase():(index===0?'left':index===1?'center':'right');
  const width=Math.max(30,Math.min(220,Number(slot?.width||preset.width)||preset.width));
  const height=Math.max(20,Math.min(120,Number(slot?.height||preset.height)||preset.height));
  const file=String(slot?.file||'').trim();
  return {id:'logo'+(index+1),file,position,size,width,height,enabled:slot?.enabled!==false&&!!file};
}

function normalizeExcelLogo(x){
  x=x||{};
  return {
    slot:[1,2,3].includes(Number(x.slot))?Number(x.slot):1,
    // Excel logo sizing is independent from the web header.  Wider limits are
    // allowed here because the merged A1:C1 logo cell can now grow with it.
    width:Math.max(40,Math.min(700,Number(x.width)||160)),
    height:Math.max(25,Math.min(350,Number(x.height)||80)),
    autoFitCell:x.autoFitCell!==false,
    padding:Math.max(0,Math.min(50,Number(x.padding)||8))
  };
}
function normalizeTextControl(ctrl,defaults={}){
  const positions=['left','center','right'],sizes=['small','medium','large'];
  return {
    position:positions.includes(String(ctrl?.position||'').toLowerCase())?String(ctrl.position).toLowerCase():(defaults.position||'left'),
    size:sizes.includes(String(ctrl?.size||'').toLowerCase())?String(ctrl.size).toLowerCase():(defaults.size||'medium'),
    enabled:ctrl?.enabled!==false
  };
}
function normalizeTextLayout(layout){
  const raw=layout||{};
  return {
    company:normalizeTextControl(raw.company,{position:'left',size:'medium'}),
    factory:{...normalizeTextControl(raw.factory,{position:'left',size:'small'}),mode:['underCompany','independent'].includes(String(raw.factory?.mode||''))?String(raw.factory.mode):'underCompany'},
    pageTitle:normalizeTextControl(raw.pageTitle,{position:'center',size:'large'}),
    headerTitle:normalizeTextControl(raw.headerTitle,{position:'center',size:'small'})
  };
}
function normalizeHeaderZone(zone,name){
  const sizes=['small','medium','large'],aligns=['left','center','right'];
  const def=name==='left'?'left':name==='right'?'right':'center';
  return {
    logoSlot:[0,1,2,3].includes(Number(zone?.logoSlot))?Number(zone.logoSlot):(name==='left'?1:name==='center'?2:3),
    showLogo:zone?.showLogo!==false,
    showCompany:!!zone?.showCompany,
    showFactory:!!zone?.showFactory,
    showHeaderTitle:!!zone?.showHeaderTitle,
    showPageTitle:!!zone?.showPageTitle,
    align:aligns.includes(String(zone?.align||''))?String(zone.align):def,
    logoWidth:Math.max(30,Math.min(220,Number(zone?.logoWidth||92)||92)),
    logoHeight:Math.max(20,Math.min(120,Number(zone?.logoHeight||48)||48)),
    companySize:sizes.includes(String(zone?.companySize||''))?String(zone.companySize):'medium',
    factorySize:sizes.includes(String(zone?.factorySize||''))?String(zone.factorySize):'small',
    headerTitleSize:sizes.includes(String(zone?.headerTitleSize||''))?String(zone.headerTitleSize):'small',
    pageTitleSize:sizes.includes(String(zone?.pageTitleSize||''))?String(zone.pageTitleSize):'large'
  };
}
function normalizeHeaderZones(z){
  z=z||{};
  return {left:normalizeHeaderZone(z.left,'left'),center:normalizeHeaderZone(z.center,'center'),right:normalizeHeaderZone(z.right,'right')};
}
function currentBranding(){
  const raw={...brandingDefaults(),...readJson('branding.json',{})};
  let logos=Array.isArray(raw.logos)?raw.logos:[];
  if(!logos.length&&raw.companyLogo)logos=[{file:raw.companyLogo,position:String(raw.logoMode||'left').toLowerCase(),size:'medium',enabled:true}];
  while(logos.length<3)logos.push({});
  raw.logos=logos.slice(0,3).map(normalizeLogoSlot);raw.excelLogo=normalizeExcelLogo(raw.excelLogo);raw.textLayout=normalizeTextLayout(raw.textLayout);raw.headerZones=normalizeHeaderZones(raw.headerZones);
  if(raw.logos[0]?.file)raw.companyLogo=raw.logos[0].file;
  if(raw.logos[0]?.position)raw.logoMode=raw.logos[0].position;
  return raw;
}
function brandingLogoFullPathValue(value){
  const p=String(value||'').trim();if(!p)return '';if(path.isAbsolute(p))return p;return path.join(__dirname,p.replace(/^[/\\]+/,''));
}
function brandingLogoFullPath(branding){
  const b=branding||currentBranding();const primary=(Array.isArray(b.logos)&&b.logos[0]?.file)||b.companyLogo||b.leftLogo||b.rightLogo||'';return brandingLogoFullPathValue(primary);
}
function brandingPublicPayload(){
  const b=currentBranding();
  const logos=(b.logos||[]).map((slot,index)=>{const full=brandingLogoFullPathValue(slot.file);let version='';try{if(full&&fs.existsSync(full))version=String(Math.floor(fs.statSync(full).mtimeMs))}catch(_){}return {id:'logo'+(index+1),position:slot.position,size:slot.size,width:slot.width,height:slot.height,enabled:slot.enabled&&!!(full&&fs.existsSync(full)),hasLogo:!!(full&&fs.existsSync(full)),logoUrl:(full&&fs.existsSync(full))?`/api/branding/logo/${index+1}?v=${version}`:'',file:String(slot.file||'')};});
  return {companyName:String(b.companyName||'Smart PM Maintenance'),factoryName:String(b.factoryName||''),headerTitle:String(b.headerTitle||'Preventive Maintenance Management System'),theme:{primary:String(b.theme?.primary||'#c90000'),header:String(b.theme?.header||'#c90000'),background:String(b.theme?.background||'#f4f7fb'),card:String(b.theme?.card||'#ffffff')},textLayout:b.textLayout,headerZones:b.headerZones,excelLogo:normalizeExcelLogo(b.excelLogo),logos,logoMode:logos[0]?.position||'left',hasLogo:!!logos[0]?.hasLogo,logoUrl:logos[0]?.logoUrl||'',companyLogo:String(b.companyLogo||'')};
}

function detectImageExtension(buffer,contentType='',original=''){
  if(Buffer.isBuffer(buffer)&&buffer.length>=8){
    if(buffer[0]===0x89&&buffer[1]===0x50&&buffer[2]===0x4E&&buffer[3]===0x47)return '.png';
    if(buffer[0]===0xFF&&buffer[1]===0xD8&&buffer[2]===0xFF)return '.jpg';
  }
  const ext=path.extname(String(original||'')).toLowerCase();
  if(ext==='.png')return '.png';
  if(ext==='.jpg'||ext==='.jpeg')return '.jpg';
  const ct=String(contentType||'').toLowerCase();
  if(ct.includes('png'))return '.png';
  if(ct.includes('jpeg')||ct.includes('jpg'))return '.jpg';
  return '';
}
function clearWorkbookBrandingImages(ws){try{if(Array.isArray(ws._media))ws._media=ws._media.filter(x=>x?.type!=='image')}catch(_){}}
function excelLogoPlacement(kind,width,height,padding=0){
  const w=Math.max(40,Math.min(700,Number(width)||160));
  const h=Math.max(25,Math.min(350,Number(height)||80));
  const p=Math.max(0,Math.min(50,Number(padding)||0));
  // The image is anchored inside the merged A1:C1 logo cell.
  // Small offsets are proportional to the selected padding, while the actual
  // row/column growth is handled by fitExcelLogoCell().
  const colOffset=Math.min(0.45, p/70);
  const rowOffset=Math.min(0.35, p/70);
  return {tl:{col:colOffset,row:rowOffset},ext:{width:w,height:h},editAs:'oneCell'};
}
function excelColumnPixels(width){
  // Practical Excel/ExcelJS approximation for Calibri 11 columns.
  // It is intentionally conservative so an image never overflows its cell.
  const w=Math.max(0,Number(width)||0);
  return Math.floor(w*7+5);
}
function fitExcelLogoCell(ws,excelLogo){
  if(!excelLogo?.autoFitCell)return;
  const w=Math.max(40,Math.min(700,Number(excelLogo.width)||160));
  const h=Math.max(25,Math.min(350,Number(excelLogo.height)||80));
  const padding=Math.max(0,Math.min(50,Number(excelLogo.padding)||8));

  // HEIGHT: Excel row height is points; image dimensions are pixels
  // (96 dpi => 0.75 pt/px).  Row 1 therefore grows/shrinks with the logo.
  const requiredPoints=Math.ceil((h + padding*2) * 0.75);
  ws.getRow(1).height=Math.max(34,requiredPoints);

  // WIDTH: A1:C1 is one merged logo cell, so its visible width is the sum of
  // columns A, B and C.  Grow those three columns when the selected logo needs
  // more room.  Keep the professional template widths as the minimum so a
  // smaller logo never damages the PM table layout.
  const base=[16,13,22];
  const basePixels=base.reduce((n,x)=>n+excelColumnPixels(x),0);
  const requiredPixels=w + padding*2 + 12;
  if(requiredPixels>basePixels){
    const scale=requiredPixels/basePixels;
    for(let i=0;i<3;i++) ws.getColumn(i+1).width=Math.round(base[i]*scale*100)/100;
  }else{
    for(let i=0;i<3;i++) ws.getColumn(i+1).width=base[i];
  }
}
function applyBrandingToWorkbook(wb,ws,kind){
  const b=currentBranding(),company=String(b.companyName||b.factoryName||'Smart PM Maintenance').trim(),factory=String(b.factoryName||'').trim(),customTitle=String(b.headerTitle||'').trim(),isAnnual=kind==='annual';
  if(isAnnual){ws.getCell(1,4).value=company;ws.getCell(1,16).value=customTitle||'ANNUAL PREVENTIVE MAINTENANCE MASTER PLAN';if(factory)ws.getCell(3,12).value=factory;}else{ws.getCell(1,4).value=company;ws.getCell(1,11).value=customTitle||'MONTHLY PREVENTIVE MAINTENANCE EXECUTION REPORT';if(factory)ws.getCell(3,9).value=factory;}
  clearWorkbookBrandingImages(ws);ws.getCell(1,1).value=null;

  const excelLogo=normalizeExcelLogo(b.excelLogo);
  fitExcelLogoCell(ws,excelLogo);
  const slot=b.logos?.[excelLogo.slot-1];
  let added=false;
  if(slot?.file){
    const logo=brandingLogoFullPathValue(slot.file),ext=excelImageExtension(logo);
    if(logo&&ext&&fs.existsSync(logo)){
      try{
        const id=wb.addImage({filename:logo,extension:ext});
        ws.addImage(id,excelLogoPlacement(kind,excelLogo.width,excelLogo.height,excelLogo.padding));
        added=true;
      }catch(e){console.error('[EXCEL BRANDING LOGO]',slot.id,e.message)}
    }
  }
  if(!added)ws.getCell(1,1).value='COMPANY LOGO';
}

async function buildCanonicalAnnualTemplate(line='',year=new Date().getFullYear()){
  const built=buildProfessionalAnnualWorkbook(line,year);
  applyBrandingToWorkbook(built.wb,built.ws,'annual');
  forceExcelRecalculation(built.wb);
  return built;
}

async function buildCanonicalMonthlyTemplate(line='',year=new Date().getFullYear(),month=1){
  const built=buildProfessionalMonthlyWorkbook(line,year,month);
  applyBrandingToWorkbook(built.wb,built.ws,'monthly');
  normalizeMonthlyScheduleDisplay(built.ws,pmDataStartRow(built.ws),built.ws.rowCount);
  forceExcelRecalculation(built.wb);
  return built;
}

async function refreshCanonicalSystemTemplates(){
  // The packaged reference files are NOT an independent source anymore.
  // They are regenerated from the same builders used by a real customer installation.
  fs.mkdirSync(SYSTEM_TEMPLATE_DIR,{recursive:true});

  const now=new Date(),year=now.getFullYear(),month=now.getMonth()+1;

  const annual=await buildCanonicalAnnualTemplate('',year);
  await safeWriteWorkbook(annual.wb,BLANK_ANNUAL_TEMPLATE);

  const monthly=await buildCanonicalMonthlyTemplate('',year,month);
  await safeWriteWorkbook(monthly.wb,BLANK_MONTHLY_TEMPLATE);

  return {
    annual:path.basename(BLANK_ANNUAL_TEMPLATE),
    monthly:path.basename(BLANK_MONTHLY_TEMPLATE)
  };
}

function safeDownloadName(v,fallback){
  const s=String(v||'').trim().replace(/[<>:"/\\|?*\x00-\x1F]/g,'_');
  return s||fallback;
}


function readJson(name, fallback) {
  try { return JSON.parse(fs.readFileSync(path.join(CONFIG_DIR,name),'utf8')); }
  catch(e) { return fallback; }
}
function writeJson(name, value) {
  fs.writeFileSync(path.join(CONFIG_DIR,name), JSON.stringify(value,null,2), 'utf8');
}
function audit(action, user, details={}) {
  const list = readJson('audit.json', []);
  list.unshift({time:new Date().toISOString(), user:user?.name||user?.username||'System', role:user?.role||'System', action, details});
  writeJson('audit.json', list.slice(0,2000));
}


app.use(express.json());
let requestQueue122=Promise.resolve();
app.use('/api',(req,res,next)=>{const previous=requestQueue122;let release;requestQueue122=new Promise(r=>release=r);res.once('finish',()=>{try{if(req.user&&req.method!=='GET'&&res.statusCode<400)audit('SYSTEM_MUTATION',req.user,{method:req.method,path:req.originalUrl});}finally{release();}});res.once('close',release);previous.then(()=>{if(!res.destroyed)next();});});
app.use((req,res,next)=>{
  if(/\.(html|js|css)$/i.test(req.path) || req.path==='/' || req.path==='/branding.html'){
    res.setHeader('Cache-Control','no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma','no-cache');
    res.setHeader('Expires','0');
    res.setHeader('Surrogate-Control','no-store');
  }
  next();
});
app.use(express.static(path.join(__dirname, 'public'),{etag:false,lastModified:false,maxAge:0}));

// Public branding is intentionally available before login so the Login screen
// can show the customer's identity.
app.get('/api/version',(req,res)=>res.json({build:BUILD_VERSION,branding:true}));
app.get('/api/branding',(req,res)=>{
  try{res.json(brandingPublicPayload())}
  catch(e){res.status(500).json({error:e.message})}
});
app.get('/api/branding/logo/:slot',(req,res)=>{try{const n=Number(req.params.slot);if(!Number.isInteger(n)||n<1||n>3)return res.status(404).end();const slot=currentBranding().logos?.[n-1],logo=brandingLogoFullPathValue(slot?.file);if(!logo||!fs.existsSync(logo))return res.status(404).end();res.setHeader('Cache-Control','no-store, max-age=0');res.sendFile(path.resolve(logo));}catch(e){res.status(404).end()}});
app.get('/api/branding/logo',(req,res)=>{try{const slot=currentBranding().logos?.[0],logo=brandingLogoFullPathValue(slot?.file);if(!logo||!fs.existsSync(logo))return res.status(404).end();res.setHeader('Cache-Control','no-store, max-age=0');res.sendFile(path.resolve(logo));}catch(e){res.status(404).end()}});


const CODE_INFO = {
  1: { name: 'Frequent / Weekly' },
  2: { name: 'Monthly / 500-1000 hr' },
  3: { name: 'Quarterly / 2000-3000 hr' },
  4: { name: 'Semi Annual / 5000-6000 hr' },
  5: { name: 'Annual / Long Interval' }
};

function safeText(ws, row, col) {
  try {
    const cell = ws.getCell(row, col);
    try {
      const t = cell.text;
      if (t !== null && t !== undefined) return String(t);
    } catch (_) {}
    const v = cell.value;
    if (v == null) return '';
    if (v instanceof Date) return v;
    if (typeof v === 'object') {
      if (v.text != null) return String(v.text);
      if (v.result != null) return String(v.result);
      if (Array.isArray(v.richText)) return v.richText.map(x => x.text || '').join('');
      return '';
    }
    return String(v);
  } catch (_) { return ''; }
}
function colNumber(letters){
  let n=0;
  for(const ch of String(letters).toUpperCase()){
    if(ch<'A'||ch>'Z')continue;
    n=n*26+(ch.charCodeAt(0)-64);
  }
  return n;
}

function excelCol(n) {
  let s='';
  while(n>0){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}
  return s;
}
function fmtDate(d) {
  if (!(d instanceof Date)) return '';
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}
function excelSerialToDate(n) {
  if (typeof n !== 'number' || !isFinite(n)) return null;
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(epoch.getTime() + n*86400000);
}
function getWeek(date = getPmDate()) {
  const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12);
  while(start.getDay() !== 5) start.setDate(start.getDate()-1);
  const end = new Date(start); end.setDate(end.getDate()+6);

  const first = new Date(date.getFullYear(),0,1,12);
  while(first.getDay() !== 5) first.setDate(first.getDate()+1);
  const idx = Math.max(0, Math.round((start-first)/604800000));
  return {
    start,end,
    header:`${start.getDate()}--${end.getDate()}`,
    annualCol:16+idx
  };
}

function tokenUser(req) {
  const auth = req.headers.authorization || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7) : '';
  return sessions.get(token) || null;
}
function requireAuth(req,res,next) {
  const user = tokenUser(req);
  if(!user) return res.status(401).json({error:'Login required'});
  req.user=user;
  // Viewer / Review Only is enforced on the server, not only in the UI.
  // Viewer may navigate the selected PM period (/api/pm/view) and logout,
  // but cannot create, edit, complete, defer, import, delete or change settings.
  if(String(user.role||'').toLowerCase()==='viewer' && req.method!=='GET'){
    const allowed = req.path==='/pm/view' || req.path==='/logout';
    if(!allowed) return res.status(403).json({error:'Viewer access is read-only'});
  }
  next();
}
app.post('/api/login',(req,res)=>{
  const {username,password}=req.body||{};
  const USERS=readJson('users.json',[]); const user=USERS.find(u=>u.active!==false && u.username===username && u.password===password);
  if(!user) return res.status(401).json({error:'Invalid username or password'});
  const token=crypto.randomBytes(24).toString('hex');
  sessions.set(token,{id:user.id,username:user.username,name:user.name,role:user.role}); audit('LOGIN',{id:user.id,username:user.username,name:user.name,role:user.role});
  res.json({token,user:{username:user.username,name:user.name,role:user.role}});
});
app.get('/api/me',requireAuth,(req,res)=>res.json(req.user));
app.post('/api/logout',requireAuth,(req,res)=>{
  const token=(req.headers.authorization||'').slice(7);
  audit('LOGOUT',req.user); sessions.delete(token); pmViewByToken.delete(token); res.json({ok:true});
});

function configuredPmLines(){
  // FINAL RULE:
  // Dashboard / PM operation lines come ONLY from Lines & Machines system config.
  // Annual/Monthly Excel files are data/history and must never recreate a deleted line.
  try{
    const x=loadPmMaster();
    return [...new Set(
      (x.lines||[])
        .filter(v=>v && v.enabled!==false && v.active!==false)
        .map(v=>String(v.name||'').trim())
        .filter(Boolean)
    )];
  }catch(_){
    return [];
  }
}

async function openAnnual(req,line) {
  const d=getPmDate(req);
  const file=`${line} PM ${d.getFullYear()}.xlsx`;
  const fullPath=path.join(ANNUAL_DIR,file);
  if(!fs.existsSync(fullPath)) throw new Error(`Annual PM file not found: ${file}`);
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(fullPath);
  return {file,fullPath,wb,ws:wb.getWorksheet('Full')||wb.worksheets[0]};
}
async function openMonthly(req,line) {
  const d=getPmDate(req);
  const year=d.getFullYear(), month=d.getMonth()+1;
  const file=monthlyFileName(line,year,month);
  const fullPath=findExistingMonthlyFile(line,year,month);
  if(!fs.existsSync(fullPath)) throw new Error(`Monthly PM file not found: ${file}. Create the month from Schedule & System.`);
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(fullPath);
  return {file,fullPath,wb,ws:wb.getWorksheet('Sheet1')||wb.worksheets[0]};
}

function normalizeWeekHeader(v){
  return String(v??'')
    .trim()
    .replace(/\s/g,'')
    .replace(/[–—−]/g,'-')
    .replace(/-+/g,'-');
}

function weekHeaderMatches(value,target){
  return normalizeWeekHeader(value)===normalizeWeekHeader(target);
}

function detectMonthlyWeekRow(ws){
  // Support all Monthly layouts used by the project.
  // Search the header area instead of assuming row 4 or row 5.
  for(let r=2;r<=6;r++){
    let count=0;
    for(let c=16;c<=20;c++){
      const h=normalizeWeekHeader(safeText(ws,r,c));
      if(/^\d{1,2}-\d{1,2}$/.test(h))count++;
    }
    if(count>=1)return r;
  }
  return null;
}

function monthlyWeeks(ws){
  const weeks=[];
  const row=detectMonthlyWeekRow(ws);
  if(!row)return weeks;

  for(let c=16;c<=20;c++){
    const h=normalizeWeekHeader(safeText(ws,row,c));
    if(/^\d{1,2}-\d{1,2}$/.test(h)){
      weeks.push({header:h,dataCol:c,headerRow:row});
    }
  }
  return weeks;
}

function findWeekColumns(ws,header){
  const target=normalizeWeekHeader(header);
  const row=detectMonthlyWeekRow(ws);
  if(!row)throw new Error(`Monthly PM week header row not found`);

  let dataCol=null;
  for(let c=16;c<=20;c++){
    if(weekHeaderMatches(safeText(ws,row,c),target)){
      dataCol=c;
      break;
    }
  }
  if(!dataCol)throw new Error(`PM week ${target} not found in monthly Plan/Done columns`);

  // Audit layout changed during development. Detect it dynamically.
  let auditCol=null;

  // Current Professional layout U:Y.
  for(let c=21;c<=25;c++){
    if(weekHeaderMatches(safeText(ws,row,c),target)){
      auditCol=c; break;
    }
  }

  // Legacy layout AG:AK.
  if(!auditCol){
    for(let c=33;c<=37;c++){
      if(weekHeaderMatches(safeText(ws,row,c),target)){
        auditCol=c; break;
      }
    }
  }

  // Some old files have audit date labels one row above/below PM labels.
  if(!auditCol){
    for(let r=Math.max(2,row-2);r<=Math.min(6,row+2)&&!auditCol;r++){
      for(const [c1,c2] of [[21,25],[33,37]]){
        for(let c=c1;c<=c2;c++){
          if(weekHeaderMatches(safeText(ws,r,c),target)){
            auditCol=c; break;
          }
        }
        if(auditCol)break;
      }
    }
  }

  // Reading the Dashboard only needs dataCol. Keep auditCol nullable;
  // Done/Undo can report a precise audit-layout error if an old file lacks it.
  return {dataCol,auditCol,headerRow:row};
}

function readDateValue(cell) {
  const v=cell.value;
  if(v instanceof Date) return fmtDate(v);
  if(typeof v==='number') return fmtDate(excelSerialToDate(v));
  return String(cell.text||'').trim();
}

async function readPlan(req,line) {
  const monthly=await openMonthly(req,line);
  const week=getWeek(getPmDate(req));
  const {dataCol,auditCol}=findWeekColumns(monthly.ws,week.header);
  const machines=new Map(), tasks=[];
  const systemTaskById=new Map((loadSystemTasks().tasks||[]).map(t=>[String(t.taskId||''),t]));
  let cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};

  for(let r=pmDataStartRow(monthly.ws);r<=monthly.ws.rowCount;r++){
    if(String(safeText(monthly.ws,r,8)).trim().toLowerCase()!=='planed') continue;
    const vals={section:String(safeText(monthly.ws,r,1)||'').trim(),machineNo:String(safeText(monthly.ws,r,2)||'').trim(),machineName:String(safeText(monthly.ws,r,3)||'').trim(),machineIdentified:String(safeText(monthly.ws,r,4)||'').trim(),type:String(safeText(monthly.ws,r,5)||'').trim()};
    if(vals.machineNo && vals.machineNo!==cur.machineNo)cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
    if(vals.section)cur.section=vals.section;if(vals.machineNo)cur.machineNo=vals.machineNo;if(vals.machineName)cur.machineName=vals.machineName;if(vals.machineIdentified&&vals.machineIdentified!=='0 0')cur.machineIdentified=vals.machineIdentified;if(vals.type)cur.type=vals.type;
    const code=Number(monthly.ws.getCell(r,dataCol).value);
    if(![1,2,3,4,5].includes(code)) continue;
    const done=Number(monthly.ws.getCell(r+1,dataCol).value)===6;
    const doneBy=done?String(safeText(monthly.ws,r,auditCol)||'').trim():'';
    const doneDate=done?readDateValue(monthly.ws.getCell(r+1,auditCol)):'';
    const key=cur.machineNo||cur.machineIdentified||`ROW-${r}`;
    if(!machines.has(key))machines.set(key,{id:key,machineNo:cur.machineNo,machineName:cur.machineName,taskCount:0,doneCount:0});
    const m=machines.get(key);m.taskCount++;if(done)m.doneCount++;
    const taskId=stableTaskId(monthly.ws,r);
    const issue=latestPmIssue(line,taskId,week.header);
    const masterTask=systemTaskById.get(String(taskId||''))||{};
    tasks.push({row:r,taskId,line,section:cur.section,machineNo:cur.machineNo,machineName:cur.machineName,machineIdentified:cur.machineIdentified,type:cur.type,part:String(safeText(monthly.ws,r,6)||'').trim(),maintenance:String(safeText(monthly.ws,r,7)||'').trim(),arabicMaintenance:String(masterTask.arabicMaintenance||'').trim(),frequency:String(safeText(monthly.ws,r,11)||'').trim(),code,codeName:CODE_INFO[code]?.name||'',done,doneBy,doneDate,issue:issue?{comment:issue.comment,reportedBy:issue.reportedBy,reportedAt:issue.reportedAt}:null});
  }
  const weekList=monthlyWeeks(monthly.ws).map((w,i)=>{
    let planned=0,achieved=0;
    for(let r=pmDataStartRow(monthly.ws);r<=monthly.ws.rowCount;r++){
      if(String(safeText(monthly.ws,r,8)).trim().toLowerCase()!=='planed')continue;
      const p=Number(monthly.ws.getCell(r,w.dataCol).value);
      if([1,2,3,4,5].includes(p)){planned++;if(Number(monthly.ws.getCell(r+1,w.dataCol).value)===6)achieved++;}
    }
    return {week:i+1,header:w.header,planned,achieved,rate:planned?Math.round((achieved/planned)*1000)/10:0};
  });
  return {line,available:true,week:{start:fmtDate(week.start),end:fmtDate(week.end),header:week.header,dataColumn:excelCol(dataCol),auditColumn:excelCol(auditCol)},monthWeeks:weekList,machines:Array.from(machines.values()),tasks};
}

app.get('/api/lines',requireAuth,async(req,res)=>{
  try{
    const filter=String(req.query.line||'ALL');
    let names=configuredPmLines();
    if(filter && filter.toUpperCase()!=='ALL')names=names.filter(x=>x===filter);
    const out=[];
    for(const line of names){
      try{
        const d=await readPlan(req,line);
        out.push({line:d.line,available:true,machines:d.machines.length,tasks:d.tasks.length,done:d.tasks.filter(x=>x.done).length,week:d.week,monthWeeks:d.monthWeeks});
      }catch(e){
        const msg=String(e.message||'');
        const missing=/Monthly PM file not found:/i.test(msg);
        out.push({
          line,
          available:false,
          missingMonthly:missing,
          machines:0,tasks:0,done:0,
          error:msg,
          week:null,monthWeeks:[]
        });
      }
    }
    res.json(out);
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});
app.get('/api/line/:line',requireAuth,async(req,res)=>{
  try{
    const line=decodeURIComponent(req.params.line);
    if(!configuredPmLines().includes(line))return res.status(404).json({error:'PM line not found'});
    res.json(await readPlan(req,line));
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.post('/api/task/:line/:row/done',requireAuth,requireCurrentPmMonth,async(req,res)=>{
  try{
    const row=Number(req.params.row);
    const week=getWeek(getPmDate(req));
    const line=decodeURIComponent(req.params.line);
    if(!configuredPmLines().includes(line))return res.status(404).json({error:'PM line not found'});
    const monthly=await openMonthly(req,line);
    await monthlyState122(line,week.start.getFullYear(),week.start.getMonth()+1);
    if(String(safeText(monthly.ws,row,8)).toLowerCase()!=='planed'||req.body?.taskId&&stableTaskId(monthly.ws,row)!==req.body.taskId)return res.status(409).json({error:'Task changed; refresh and retry'});
    const {dataCol,auditCol}=findWeekColumns(monthly.ws,week.header);

    if(String(safeText(monthly.ws,row,8)).trim().toLowerCase()!=='planed')throw new Error('Not a planned task row');
    const code=Number(monthly.ws.getCell(row,dataCol).value);
    if(![1,2,3,4,5].includes(code))throw new Error('Task is not planned for this week');

    const now=new Date();
    const currentFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
    currentFriday.setDate(currentFriday.getDate()-((currentFriday.getDay()-5+7)%7));

    const weekStart=new Date(week.start.getFullYear(),week.start.getMonth(),week.start.getDate(),12,0,0);
    const weekEnd=new Date(week.end.getFullYear(),week.end.getMonth(),week.end.getDate(),12,0,0);
    const isHistorical=weekEnd.getTime()<currentFriday.getTime();

    let executionDate;

    if(isHistorical){
      // Historical PM editing is allowed for every authenticated user.

      const raw=String((req.body&&req.body.executionDate)||'');
      if(!/^\d{4}-\d{2}-\d{2}$/.test(raw)){
        return res.status(400).json({error:'Admin must select the actual execution date for this historical PM.'});
      }

      const [yy,mm,dd]=raw.split('-').map(Number);
      executionDate=dateOnly(yy,mm,dd);

      if(
        executionDate.getFullYear()!==yy ||
        executionDate.getMonth()!==mm-1 ||
        executionDate.getDate()!==dd ||
        executionDate.getTime()<weekStart.getTime() ||
        executionDate.getTime()>weekEnd.getTime()
      ){
        return res.status(400).json({error:`Execution date must be inside PM week ${fmtDate(weekStart)} to ${fmtDate(weekEnd)}.`});
      }
    }else{
      // Current PM: use actual computer/server execution date automatically.
      executionDate=dateOnly(now.getFullYear(),now.getMonth()+1,now.getDate());
    }

    monthly.ws.getCell(row+1,dataCol).value=6;
    monthly.ws.getCell(row,auditCol).value=req.user.name;
    monthly.ws.getCell(row,auditCol).numFmt='General';
    monthly.ws.getCell(row+1,auditCol).value=executionDate;
    monthly.ws.getCell(row+1,auditCol).numFmt='dd/mm/yyyy';

    // Monthly display only: keep 6 numeric for KPI logic, show it as a green check.
    normalizeMonthlyScheduleDisplay(monthly.ws,pmDataStartRow(monthly.ws),monthly.ws.rowCount);

    await saveExecution122(line,week.start.getFullYear(),week.start.getMonth()+1,monthly.ws);
    await safeWriteWorkbook(monthly.wb,monthly.fullPath);

    // A completed Running Hours / PLC Hours PM resets its runtime counter from
    // the actual execution date.  We store the reset point in System Master so
    // future Year Planner builds continue the counter instead of restarting on
    // 1 January.  Current Annual/Monthly history is not rewritten here.
    try{
      const linkedId=stableTaskId(monthly.ws,row);
      if(linkedId){
        const taskStore=loadSystemTasks();
        const masterTask=taskStore.tasks.find(t=>t.taskId===linkedId);
        if(masterTask && ['hours','plc_hours'].includes(normalizePlannerCalculationMethod(masterTask))){
          masterTask.lastExecutionDate=fmtIsoDateLocal(executionDate);
          masterTask.runningHoursCounterStartDate=fmtIsoDateLocal(executionDate);
          masterTask.updatedAt=new Date().toISOString();
          saveSystemTasks(taskStore);
        }
      }
    }catch(counterErr){console.warn('Running-hours counter reset warning:',counterErr.message);}

    audit(isHistorical?'HISTORICAL_PM_DONE':'PM_DONE',req.user,{
      line:req.params.line,
      row,
      pmWeek:week.header,
      executionDate:fmtDate(executionDate),
      correctionTime:new Date().toISOString()
    });

    res.json({
      ok:true,
      value:6,
      user:req.user.name,
      date:fmtDate(executionDate),
      historical:isHistorical,
      dataColumn:excelCol(dataCol),
      auditColumn:excelCol(auditCol)
    });
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.post('/api/task/:line/:row/issue',requireAuth,requireCurrentPmMonth,async(req,res)=>{
  try{
    const row=Number(req.params.row);
    const line=decodeURIComponent(req.params.line);
    if(!configuredPmLines().includes(line))return res.status(404).json({error:'PM line not found'});
    const comment=String(req.body?.comment||'').trim();
    if(!comment)return res.status(400).json({error:'Enter the issue / comment before saving.'});
    if(comment.length>2000)return res.status(400).json({error:'Issue / comment is too long (maximum 2000 characters).'});
    const week=getWeek(getPmDate(req));
    const monthly=await openMonthly(req,line);
    if(String(safeText(monthly.ws,row,8)).toLowerCase()!=='planed'||req.body?.taskId&&stableTaskId(monthly.ws,row)!==req.body.taskId)return res.status(409).json({error:'Task changed; refresh and retry'});
    const taskId=stableTaskId(monthly.ws,row);
    const now=new Date();
    const currentFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
    currentFriday.setDate(currentFriday.getDate()-((currentFriday.getDay()-5+7)%7));
    const weekStart=new Date(week.start.getFullYear(),week.start.getMonth(),week.start.getDate(),12,0,0);
    const weekEnd=new Date(week.end.getFullYear(),week.end.getMonth(),week.end.getDate(),12,0,0);
    if(weekStart.getTime()>currentFriday.getTime())return res.status(403).json({error:'Future PM weeks cannot have issues reported.'});
    // Historical issue reporting is allowed for every authenticated user.
    const x=loadPmIssues();
    const item={id:crypto.randomUUID(),key:pmIssueKey(line,taskId,week.header),line,row,taskId,pmWeek:week.header,comment,reportedBy:req.user.name,reportedAt:now.toISOString()};
    x.items.push(item); savePmIssues(x);
    audit('PM_ISSUE_REPORTED',req.user,{line,row,taskId,pmWeek:week.header,comment});
    res.json({ok:true,issue:{comment:item.comment,reportedBy:item.reportedBy,reportedAt:item.reportedAt}});
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.post('/api/task/:line/:row/undo',requireAuth,requireCurrentPmMonth,async(req,res)=>{
  try{
    // Undo/correction is allowed for every authenticated user, including historical periods.
    const row=Number(req.params.row);
    const week=getWeek(getPmDate(req));
    const line=decodeURIComponent(req.params.line);
    if(!configuredPmLines().includes(line))return res.status(404).json({error:'PM line not found'});
    const monthly=await openMonthly(req,line);
    if(String(safeText(monthly.ws,row,8)).toLowerCase()!=='planed'||req.body?.taskId&&stableTaskId(monthly.ws,row)!==req.body.taskId)return res.status(409).json({error:'Task changed; refresh and retry'});
    const {dataCol,auditCol}=findWeekColumns(monthly.ws,week.header);

    monthly.ws.getCell(row+1,dataCol).value=0;
    monthly.ws.getCell(row,auditCol).value='';
    monthly.ws.getCell(row+1,auditCol).value='';
    monthly.ws.getCell(row+1,auditCol).numFmt='dd/mm/yyyy';

    // 0 remains numeric for logic but is visually empty.
    normalizeMonthlyScheduleDisplay(monthly.ws,pmDataStartRow(monthly.ws),monthly.ws.rowCount);

    await saveExecution122(line,week.start.getFullYear(),week.start.getMonth()+1,monthly.ws);
    await safeWriteWorkbook(monthly.wb,monthly.fullPath);
    audit('PM_UNDO',req.user,{
      line:req.params.line,
      row,
      pmWeek:week.header,
      correctionTime:new Date().toISOString()
    });
    res.json({ok:true});
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});


function requireAdmin(req,res,next){
  if(!req.user || req.user.role!=='Admin') return res.status(403).json({error:'Admin access required'});
  next();
}

function safeUploadToken(v){return /^[a-f0-9]{32}\.xlsx$/i.test(String(v||''))?String(v):'';}
function stagedUploadPath(token){const t=safeUploadToken(token);return t?path.join(EXCEL_UPLOAD_DIR,t):'';}
async function validateXlsxBuffer(buf){
  if(!Buffer.isBuffer(buf)||buf.length<1000)throw new Error('Excel file is empty or invalid');
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  if(!wb.worksheets.length)throw new Error('Excel workbook has no worksheets');
  return wb;
}
app.post('/api/pm-master/upload-xlsx',requireAuth,requireAdmin,express.raw({type:'application/octet-stream',limit:'30mb'}),async(req,res)=>{
  try{
    const original=String(req.headers['x-file-name']||'upload.xlsx');
    if(!/\.xlsx$/i.test(original))return res.status(400).json({error:'Only .xlsx files are accepted'});
    await validateXlsxBuffer(req.body);
    const token=crypto.randomBytes(16).toString('hex')+'.xlsx';
    const full=path.join(EXCEL_UPLOAD_DIR,token);
    fs.writeFileSync(full,req.body);
    audit('PM_EXCEL_UPLOAD_STAGED',req.user,{original,size:req.body.length,token});
    res.json({ok:true,token,original,size:req.body.length});
  }catch(e){console.error(e);res.status(400).json({error:e.message});}
});


// ---------- CANONICAL BLANK EXCEL DOWNLOADS ----------
// These are generated LIVE from the same code used to create customer files.
// Therefore a downloaded template can never drift away from the system format.
app.get('/api/templates/annual',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const year=Number(req.query.year||new Date().getFullYear());
    const line=cleanName(req.query.line||'');
    if(year<2020||year>2100)return res.status(400).json({error:'Invalid year'});
    const built=await buildCanonicalAnnualTemplate(line,year);
    const buf=await built.wb.xlsx.writeBuffer();
    const name=safeDownloadName(line||'Smart_PM_BLANK_ANNUAL', 'Smart_PM_BLANK_ANNUAL');
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="${name}_${year}.xlsx"`);
    res.send(Buffer.from(buf));
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.get('/api/templates/monthly',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const year=Number(req.query.year||new Date().getFullYear());
    const month=Number(req.query.month||new Date().getMonth()+1);
    const line=cleanName(req.query.line||'');
    if(year<2020||year>2100||month<1||month>12)return res.status(400).json({error:'Invalid year or month'});
    const built=await buildCanonicalMonthlyTemplate(line,year,month);
    const buf=await built.wb.xlsx.writeBuffer();
    const name=safeDownloadName(line||'Smart_PM_BLANK_MONTHLY', 'Smart_PM_BLANK_MONTHLY');
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition',`attachment; filename="${name}_${year}_${String(month).padStart(2,'0')}.xlsx"`);
    res.send(Buffer.from(buf));
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.get('/api/pm-master/excel-sources',requireAuth,requireAdmin,(req,res)=>{
  try{
    const year=Number(req.query.year||new Date().getFullYear());
    const annual=fs.existsSync(ANNUAL_DIR)?fs.readdirSync(ANNUAL_DIR).filter(f=>new RegExp(` PM ${year}\\.xlsx$`,'i').test(f)).map(file=>({file,line:file.replace(new RegExp(` PM ${year}\\.xlsx$`,'i'),''),year})):[];
    const monthlyTemplates=fs.existsSync(MONTHLY_TEMPLATE_DIR)?fs.readdirSync(MONTHLY_TEMPLATE_DIR).filter(f=>/ PM_TEMPLATE\.xlsx$/i.test(f)).map(file=>({file,line:file.replace(/ PM_TEMPLATE\.xlsx$/i,'')})):[];
    res.json({year,annual,monthlyTemplates});
  }catch(e){res.status(500).json({error:e.message});}
});



// ---------- MULTI EMAIL ACCOUNTS ----------
app.get('/api/settings/email', requireAuth, requireAdmin, (req,res)=>{
  const cfg=readJson('email.json',{accounts:[],fallbackEnabled:true,weekly:{},monthly:{}});
  // Never send SMTP passwords back to the browser.
  res.json({
    ...cfg,
    accounts:(cfg.accounts||[]).map(a=>({...a,smtpPassword:a.smtpPassword?'********':''}))
  });
});

app.put('/api/settings/email', requireAuth, requireAdmin, (req,res)=>{
  const old=readJson('email.json',{accounts:[]});
  const next=req.body||{};
  const oldById=new Map((old.accounts||[]).map(a=>[String(a.id),a]));
  next.accounts=(next.accounts||[]).map(a=>{
    const prev=oldById.get(String(a.id));
    // Mask means keep existing secret.
    if(a.smtpPassword==='********') a.smtpPassword=prev?.smtpPassword||'';
    return a;
  });

  // Ensure only one default account.
  let defaultSeen=false;
  next.accounts.forEach(a=>{
    if(a.isDefault && !defaultSeen){defaultSeen=true;}
    else if(a.isDefault){a.isDefault=false;}
  });
  if(!defaultSeen && next.accounts.length) next.accounts[0].isDefault=true;

  writeJson('email.json',next);
  audit('SETTINGS_EMAIL',req.user,{accounts:next.accounts.map(a=>({id:a.id,name:a.name,enabled:a.enabled,isDefault:a.isDefault})),fallbackEnabled:!!next.fallbackEnabled});
  res.json({ok:true});
});

function buildTransport(account){
  const port = Number(account.smtpPort) || 587;
  return nodemailer.createTransport({
    host: String(account.smtpHost || '').trim(),
    port,
    secure: port === 465 ? true : !!account.secure,
    auth: {
      user: String(account.smtpUser || '').trim(),
      pass: String(account.smtpPassword || '').replace(/\s+/g,'')
    },
    connectionTimeout: 20000,
    greetingTimeout: 20000,
    socketTimeout: 30000
  });
}

app.post('/api/settings/email/test/:id', requireAuth, requireAdmin, async (req,res)=>{
  try{
    const cfg=readJson('email.json',{accounts:[]});
    const account=(cfg.accounts||[]).find(a=>String(a.id)===String(req.params.id));
    if(!account) return res.status(404).json({error:'Email account not found'});
    if(!account.smtpHost||!account.smtpUser||!account.smtpPassword) return res.status(400).json({error:'SMTP Host, User and Password are required'});
    const to=String(req.body?.to||'').trim();
    if(!to) return res.status(400).json({error:'Test recipient email is required'});

    const transporter=buildTransport(account);
    await transporter.verify();
    const info=await transporter.sendMail({
      from:`"${account.fromName||'Smart PM Maintenance'}" <${account.smtpUser}>`,
      to,
      subject:'Smart PM Maintenance - Test Email',
      text:`SMTP test succeeded using account: ${account.name||account.smtpUser}`
    });
    audit('EMAIL_TEST_SUCCESS',req.user,{account:account.name||account.id,to,messageId:info.messageId});
    res.json({ok:true,message:'Test email sent successfully'});
  }catch(e){
    audit('EMAIL_TEST_FAILED',req.user,{accountId:req.params.id,error:e.message});
    res.status(500).json({error:e.message});
  }
});





// ---------- PM VIEW / HISTORY ----------
function listAvailablePmMonths(){
  const map=new Map();

  for(const fullPath of listMonthlyExcelFiles()){
    const f=path.basename(fullPath);
    const m=f.match(/^(.+?)\s+PM_(\d{4})_(\d{2})\.xlsx$/i);
    if(!m)continue;

    const line=m[1].trim();
    const year=Number(m[2]);
    const month=Number(m[3]);
    const key=`${year}-${String(month).padStart(2,'0')}`;

    if(!map.has(key)){
      map.set(key,{key,year,month,lines:[]});
    }
    if(!map.get(key).lines.includes(line))map.get(key).lines.push(line);
  }

  return [...map.values()].sort((a,b)=>b.key.localeCompare(a.key));
}


app.get('/api/system/monthly-storage-status',requireAuth,requireAdmin,(req,res)=>{
  try{
    const files=listMonthlyExcelFiles();
    res.json({
      root:MONTHLY_WORK_DIR,
      count:files.length,
      files:files.map(f=>path.relative(MONTHLY_WORK_DIR,f)),
      months:listAvailablePmMonths()
    });
  }catch(e){
    res.status(500).json({error:e.message});
  }
});

app.get('/api/pm/history',requireAuth,(req,res)=>{
  res.json({months:listAvailablePmMonths()});
});

app.get('/api/pm/access-mode',requireAuth,(req,res)=>{
  const role=String((req.user&&req.user.role)||'').toLowerCase();
  const selectedWeek=getWeek(getPmDate(req));

  const now=new Date();
  const currentFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
  currentFriday.setDate(currentFriday.getDate()-((currentFriday.getDay()-5+7)%7));

  const selectedStart=new Date(selectedWeek.start.getFullYear(),selectedWeek.start.getMonth(),selectedWeek.start.getDate(),12,0,0);
  const selectedEnd=new Date(selectedWeek.end.getFullYear(),selectedWeek.end.getMonth(),selectedWeek.end.getDate(),12,0,0);

  const isCurrentWeek=selectedStart.getTime()===currentFriday.getTime();
  const isPastWeek=selectedEnd.getTime()<currentFriday.getTime();
  const isFutureWeek=selectedStart.getTime()>currentFriday.getTime();

  const iso=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;

  res.json({
    role,
    isAdmin:role==='admin',
    isCurrentWeek,
    isPastWeek,
    isFutureWeek,
    canHistoricalEdit:isPastWeek,
    weekStart:iso(selectedStart),
    weekEnd:iso(selectedEnd)
  });
});

app.get('/api/pm/view',requireAuth,(req,res)=>{
  const d=getPmDate(req);

  // Current PM week from actual computer/server date.
  const now=new Date();
  const currentFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
  const daysSinceFriday=(currentFriday.getDay()-5+7)%7;
  currentFriday.setDate(currentFriday.getDate()-daysSinceFriday);

  const currentKey=`${currentFriday.getFullYear()}-${String(currentFriday.getMonth()+1).padStart(2,'0')}`;
  const viewKey=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;

  // Only the PM month containing the real current PM week is editable.
  const readOnly=false; // Past periods remain editable; future-week completion is still blocked by requireCurrentPmMonth.

  const fridays=fridaysInMonth(d.getFullYear(),d.getMonth()+1);

  res.json({
    date:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`,
    year:d.getFullYear(),
    month:d.getMonth()+1,
    readOnly,
    mode:readOnly?'history':'active',
    currentPmDate:`${currentFriday.getFullYear()}-${String(currentFriday.getMonth()+1).padStart(2,'0')}-${String(currentFriday.getDate()).padStart(2,'0')}`,
    weeks:fridays.map((f,i)=>{
      const e=new Date(f); e.setDate(e.getDate()+6);
      return {
        index:i+1,
        start:`${f.getFullYear()}-${String(f.getMonth()+1).padStart(2,'0')}-${String(f.getDate()).padStart(2,'0')}`,
        label:`W${i+1} ${f.getDate()}--${e.getDate()}`
      };
    })
  });
});

app.post('/api/pm/view',requireAuth,(req,res)=>{
  const year=Number(req.body?.year), month=Number(req.body?.month), weekIndex=Number(req.body?.weekIndex||1);
  const fridays=fridaysInMonth(year,month);
  if(!year||month<1||month>12||weekIndex<1||weekIndex>fridays.length) return res.status(400).json({error:'Invalid PM month/week'});
  const f=fridays[weekIndex-1];
  const selected=`${f.getFullYear()}-${String(f.getMonth()+1).padStart(2,'0')}-${String(f.getDate()).padStart(2,'0')}`;
  const token=(req.headers.authorization||'').replace(/^Bearer\s+/i,'');
  pmViewByToken.set(token,selected);
  res.json({ok:true,date:selected});
});


function requireCurrentPmMonth(req,res,next){
  const selectedWeek=getWeek(getPmDate(req));

  const now=new Date();
  const currentFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
  currentFriday.setDate(currentFriday.getDate()-((currentFriday.getDay()-5+7)%7));

  const selectedStart=new Date(selectedWeek.start.getFullYear(),selectedWeek.start.getMonth(),selectedWeek.start.getDate(),12,0,0);
  const selectedEnd=new Date(selectedWeek.end.getFullYear(),selectedWeek.end.getMonth(),selectedWeek.end.getDate(),12,0,0);

  if(selectedStart.getTime()>currentFriday.getTime()){
    return res.status(403).json({error:'Future PM weeks cannot be completed.'});
  }

  // Historical weeks are editable for every authenticated user.

  next();
}

app.post('/api/pm/view-current',requireAuth,(req,res)=>{
  const token=(req.headers.authorization||'').replace(/^Bearer\s+/i,'');

  // Current Active must follow the computer/server date, not activePmDate test setting.
  // PM week runs Friday -> Thursday, so find the most recent Friday.
  const now=new Date();
  const currentPmFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
  const daysSinceFriday=(currentPmFriday.getDay()-5+7)%7;
  currentPmFriday.setDate(currentPmFriday.getDate()-daysSinceFriday);

  const selected=`${currentPmFriday.getFullYear()}-${String(currentPmFriday.getMonth()+1).padStart(2,'0')}-${String(currentPmFriday.getDate()).padStart(2,'0')}`;
  pmViewByToken.set(token,selected);

  res.json({
    ok:true,
    date:selected,
    year:currentPmFriday.getFullYear(),
    month:currentPmFriday.getMonth()+1
  });
});

app.post('/api/email/resend-month',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const year=Number(req.body?.year), month=Number(req.body?.month);
    const fridays=fridaysInMonth(year,month);
    if(!fridays.length) return res.status(400).json({error:'Invalid month'});
    const last=fridays[fridays.length-1];
    const th=new Date(last); th.setDate(th.getDate()+6);
    const r=await sendReportEmail('monthly',th,{force:true});
    audit('MONTHLY_EXCEL_RESENT',req.user,{year,month,files:r.files,account:r.account});
    res.json(r);
  }catch(e){res.status(500).json({error:e.message});}
});




// ---------- STEP 43: OLD EXCEL -> PROFESSIONAL TEMPLATE MIGRATION ----------

function simplePmHeaderMap(ws){
  const map={};
  for(let c=1;c<=Math.min(ws.columnCount||20,30);c++){
    // safeText protects the importer from blank/null cells and Excel merged-cell values.
    const h=String(safeText(ws,1,c)||'').trim().toLowerCase().replace(/\s+/g,' ');
    if(h)map[h]=c;
  }
  const pick=(...names)=>{for(const n of names)if(map[n])return map[n];return 0};
  const out={
    section:pick('section'),machineNo:pick('machine no','machine no.','machine number'),machineName:pick('machine name'),
    machineIdentified:pick('machine identified','machine id','machine identification'),manufacturer:pick('manufacturer','type'),
    part:pick('part'),maintenance:pick('maintenance','maintenance task','pm task'),arabicMaintenance:pick('arabic maintenance','maintenance arabic','arabic pm'),
    frequency:pick('freq.','freq','frequency','oem'), maintenanceType:pick('type of maintenance','maintenance type','discipline','e/m','type')
  };
  return out.machineNo&&out.maintenance&&out.frequency?out:null;
}
function simpleFrequencyInfo(text){
  const raw=String(text||'').trim(),lower=raw.toLowerCase().replace(/,/g,'');
  const hm=lower.match(/(\d+)\s*(?:hr|hrs|hour|hours)\b/);
  if(hm)return {pmCode:1,frequencyValue:Number(hm[1]),frequencyText:`${Number(hm[1])} hr`,calculationMode:'Running Hours'};
  if(lower.includes('week'))return {pmCode:1,frequencyValue:1,frequencyText:'Weekly',calculationMode:'Calendar'};
  if(lower.includes('quarter')||/\b3\s*month/.test(lower))return {pmCode:3,frequencyValue:3,frequencyText:'Quarterly',calculationMode:'Calendar'};
  if(lower.includes('semi')||/\b6\s*month/.test(lower))return {pmCode:4,frequencyValue:6,frequencyText:'Semi Annual',calculationMode:'Calendar'};
  if(lower.includes('annual')||lower.includes('year')||/\b12\s*month/.test(lower))return {pmCode:5,frequencyValue:12,frequencyText:'Annual',calculationMode:'Calendar'};
  const mm=lower.match(/(\d+)\s*month/);
  if(lower.includes('month')||lower.includes('montly'))return {pmCode:2,frequencyValue:mm?Number(mm[1]):1,frequencyText:mm&&Number(mm[1])>1?`${Number(mm[1])} Months`:'Monthly',calculationMode:'Calendar'};
  return {pmCode:2,frequencyValue:1,frequencyText:raw||'Monthly',calculationMode:'Calendar'};
}
function simplePmRowsFromWorkbook(wb){
  const rows=[];
  for(const ws of wb.worksheets){
    const h=simplePmHeaderMap(ws); if(!h)continue;
    let cur={section:'',machineNo:'',machineName:'',machineIdentified:'',manufacturer:'',part:''};
    let autoMachine=0, lastMachineName='';
    let arabicMap={}; try{arabicMap=JSON.parse(fs.readFileSync(path.join(PM_MASTER_DIR,'arabic_task_map.json'),'utf8'));}catch(_e){}
    for(let r=2;r<=ws.rowCount;r++){
      const get=c=>c?String(safeText(ws,r,c)||'').trim():'';
      const incomingName=get(h.machineName), incomingNo=get(h.machineNo);
      if(incomingName && incomingName!==lastMachineName){ autoMachine++; lastMachineName=incomingName; cur.part=''; }
      if(incomingNo && incomingNo!==cur.machineNo)cur.part='';
      for(const k of ['section','machineName','manufacturer']){const v=get(h[k]);if(v)cur[k]=v;}
      if(incomingNo)cur.machineNo=incomingNo; else if(incomingName)cur.machineNo='P'+String(autoMachine).padStart(2,'0');
      const ident=get(h.machineIdentified); if(ident)cur.machineIdentified=ident; else if(incomingName)cur.machineIdentified=(cur.machineNo+' '+incomingName).trim();
      const part=get(h.part); if(part)cur.part=part;
      const maintenance=get(h.maintenance); if(!maintenance)continue;
      const mt=(get(h.maintenanceType)||'').toUpperCase();
      const arabicMaintenance=get(h.arabicMaintenance)||String(arabicMap[maintenance]||'');
      rows.push({...cur,part:cur.part,maintenance,arabicMaintenance,maintenanceType:(mt==='E'||mt.startsWith('ELEC'))?'E':(mt==='M'||mt.startsWith('MECH'))?'M':mt,frequency:get(h.frequency),sourceSheet:ws.name,sourceRow:r});
    }
  }
  return rows;
}
function simplePmWorkbookToLegacy(wb,year){
  const src=simplePmRowsFromWorkbook(wb); if(!src.length)return null;
  const out=new ExcelJS.Workbook(),ws=out.addWorksheet('Full');
  const fridays=annualFridays(Number(year));
  // Balance imported calendar PM across the real PM Fridays. This avoids putting
  // every Monthly/Quarterly task into the same week while preserving frequency.
  const balance={weekly:0,monthly:0,quarterly:0,semiannual:0,annual:0};
  const weeksByMonth=Array.from({length:12},()=>[]);
  fridays.forEach((d,i)=>weeksByMonth[d.getMonth()].push(i+1));
  function balancedCalendarWeeks(fi){
    const raw=String(fi.frequencyText||'').toLowerCase();
    let kind=fi.pmCode===1?'weekly':fi.pmCode===2?'monthly':fi.pmCode===3?'quarterly':fi.pmCode===4?'semiannual':'annual';
    if(raw.includes('week'))kind='weekly'; else if(raw.includes('quarter'))kind='quarterly'; else if(raw.includes('semi')||raw.includes('6 month'))kind='semiannual'; else if(raw.includes('annual')||raw.includes('year'))kind='annual'; else if(raw.includes('month'))kind='monthly';
    if(kind==='weekly')return fridays.map((_,i)=>i+1);
    const idx=balance[kind]++;
    const interval=kind==='monthly'?1:kind==='quarterly'?3:kind==='semiannual'?6:12;
    const phase=idx%interval;
    const slot=Math.floor(idx/interval);
    const outWeeks=[];
    for(let m=phase;m<12;m+=interval){
      const arr=weeksByMonth[m]; if(!arr.length)continue;
      outWeeks.push(arr[slot%arr.length]);
    }
    return outWeeks;
  }
  let r=1;
  for(const x of src){
    const fi=simpleFrequencyInfo(x.frequency);
    ws.getCell(r,1).value=x.section;ws.getCell(r,2).value=x.machineNo;ws.getCell(r,3).value=x.machineName;ws.getCell(r,4).value=x.machineIdentified;ws.getCell(r,5).value=x.manufacturer;
    ws.getCell(r,6).value=x.part;ws.getCell(r,7).value=x.maintenance;ws.getCell(r,8).value='planed';
    ws.getCell(r,15).value=fi.frequencyText;
    if(fi.calculationMode==='Running Hours')ws.getCell(r,15).value=fi.frequencyText;
    else if(fi.pmCode===1)ws.getCell(r,10).value='X'; else if(fi.pmCode===2)ws.getCell(r,11).value='X'; else if(fi.pmCode===3)ws.getCell(r,12).value='X'; else if(fi.pmCode===4)ws.getCell(r,13).value='X'; else if(fi.pmCode===5)ws.getCell(r,14).value='X';
    if(fi.calculationMode==='Calendar'){
      for(const w of balancedCalendarWeeks(fi))ws.getCell(r,15+w).value=fi.pmCode;
    }
    ws.getCell(r+1,1).value=x.section;ws.getCell(r+1,2).value=x.machineNo;ws.getCell(r+1,3).value=x.machineName;ws.getCell(r+1,4).value=x.machineIdentified;ws.getCell(r+1,5).value=x.manufacturer;ws.getCell(r+1,8).value='done';
    r+=2;
  }
  return {wb:out,ws,rows:src};
}
function enrichArabicMaintenance(line,year,rows){
  if(!rows||!rows.length)return;
  const norm=v=>String(v||'').trim().replace(/\s+/g,' ').toLowerCase();
  const key=(m,p,t)=>[norm(m),norm(p),norm(t)].join('|');
  const meta=new Map(rows.map(x=>[key(x.machineNo,x.part,x.maintenance),x]));
  const store=loadSystemTasks();let changed=false;
  for(const t of store.tasks){if(same122(t.line,line)&&Number(t.year)===Number(year)){
    const v=meta.get(key(t.machineNo,t.part,t.maintenance)); if(!v)continue;
    if(v.arabicMaintenance&&t.arabicMaintenance!==v.arabicMaintenance){t.arabicMaintenance=v.arabicMaintenance;changed=true;}
    if(v.maintenanceType&&t.maintenanceType!==v.maintenanceType){t.maintenanceType=v.maintenanceType;changed=true;}
  }}
  if(changed)saveSystemTasks(store);
}
function migrationTaskRows(ws){
  const out=[];
  for(let r=1;r<=ws.rowCount;r++){
    const rt=String(ws.getCell(r,8).value||'').trim().toLowerCase();
    if(rt==='planed'||rt==='planned'){
      out.push({
        row:r,
        section:String(ws.getCell(r,1).value||'').trim(),
        machineNo:String(ws.getCell(r,2).value||'').trim(),
        machineName:String(ws.getCell(r,3).value||'').trim(),
        machineIdentified:String(ws.getCell(r,4).value||'').trim(),
        type:String(ws.getCell(r,5).value||'').trim(),
        part:String(ws.getCell(r,6).value||'').trim(),
        task:String(ws.getCell(r,7).value||'').trim()
      });
    }
  }
  return out;
}
function migrationBackupPath(sourcePath){
  const dir=path.join(BACKUP_DIR,'excel_migration');
  fs.mkdirSync(dir,{recursive:true});
  const stamp=new Date().toISOString().replace(/[:.]/g,'-');
  return path.join(dir,`${stamp}_${path.basename(sourcePath)}`);
}
async function loadStagedWorkbook(token){
  const p=stagedUploadPath(token);
  if(!p||!fs.existsSync(p))throw new Error('Upload the old Excel file first');
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(p);
  return {p,wb};
}
function clearDataArea(ws,start,maxCol){
  const last=Math.max(ws.rowCount,start+600);
  for(let r=start;r<=last;r++)for(let c=1;c<=maxCol;c++)ws.getCell(r,c).value=null;
}

function migrationCellValue(cell){
  const v=cell.value;
  if(v===null||v===undefined)return null;
  if(v instanceof Date)return new Date(v.getTime());
  if(typeof v==='string'||typeof v==='number'||typeof v==='boolean')return v;
  if(typeof v==='object'){
    // Old formulas/shared-formulas are never copied into the new professional workbook.
    // Preserve the last calculated result when available.
    if(Object.prototype.hasOwnProperty.call(v,'result'))return v.result??null;
    if(Object.prototype.hasOwnProperty.call(v,'text'))return String(v.text||'');
    if(Object.prototype.hasOwnProperty.call(v,'richText'))return (v.richText||[]).map(x=>x.text||'').join('');
    if(Object.prototype.hasOwnProperty.call(v,'hyperlink'))return String(v.text||v.hyperlink||'');
    try{return String(cell.text||'');}catch(_){return null;}
  }
  return null;
}

function legacyText(v){return String(v??'').trim();}
function legacyFrequencyInfo(oldWs,planRow,kind){
  // Legacy layout:
  // I Daily | J Weekly | K Monthly | L Quarterly | M Semi annual | N Annual | O Freq.
  const names=[
    {col:9,text:'Daily',months:0},
    {col:10,text:'Weekly',months:0},
    {col:11,text:'Monthly',months:1},
    {col:12,text:'Quarterly',months:3},
    {col:13,text:'Semi Annual',months:6},
    {col:14,text:'Annual',months:12}
  ];
  let selected=null;
  for(const x of names){
    const cell=oldWs.getCell(planRow,x.col);
    const hasValue=legacyText(migrationCellValue(cell))!=='';
    const fill=cell.fill&&cell.fill.fgColor&&cell.fill.fgColor.argb;
    if(hasValue||fill){selected=x;break;}
  }

  const freqRaw=legacyText(migrationCellValue(oldWs.getCell(planRow,15)));
  const scheduleCodes=[];
  if(kind==='annual'){
    for(let c=16;c<=67;c++){
      const n=Number(migrationCellValue(oldWs.getCell(planRow,c)));
      if([1,2,3,4,5].includes(n))scheduleCodes.push(n);
    }
  }else{
    for(let c=16;c<=20;c++){
      const n=Number(migrationCellValue(oldWs.getCell(planRow,c)));
      if([1,2,3,4,5,6].includes(n))scheduleCodes.push(n);
    }
  }
  const pmCode=scheduleCodes.find(n=>n>=1&&n<=5)||null;

  const lower=freqRaw.toLowerCase();
  // V1.3.7: normalize thousands separators ONLY while importing legacy Excel
  // Running Hours text. This deliberately does not change the Year Planner or
  // the runtime-hours calculation used by V1.3.5. Examples:
  //   10,000 hr -> 10000 hr   |   6,000 hours -> 6000 hr
  const legacyHoursMatch=lower.match(/(\d{1,3}(?:,\d{3})+|\d+)\s*(?:hr|hrs|hour|hours)\b/);
  if(legacyHoursMatch){
    const normalizedHours=legacyHoursMatch[1].replace(/,/g,'');
    const hoursValue=Number(normalizedHours);
    if(hoursValue>0){
      return {pmCode:pmCode||1,frequencyValue:hoursValue,frequencyText:`${normalizedHours} hr`,calculationMode:'Running Hours'};
    }
  }

  if(selected){
    let frequencyValue=1;
    if(selected.text==='Monthly'){
      const m=lower.match(/(\d+)\s*month/);
      frequencyValue=m?Number(m[1]):1;
    }else if(selected.text==='Quarterly')frequencyValue=3;
    else if(selected.text==='Semi Annual')frequencyValue=6;
    else if(selected.text==='Annual')frequencyValue=12;
    return {pmCode:pmCode||1,frequencyValue,frequencyText:selected.text,calculationMode:'Calendar'};
  }

  // Fallback to legacy Freq text when the category cells are empty.
  if(lower.includes('week'))return {pmCode:pmCode||1,frequencyValue:1,frequencyText:'Weekly',calculationMode:'Calendar'};
  if(lower.includes('month')||lower.includes('montly')){
    const m=lower.match(/(\d+)\s*month/);
    return {pmCode:pmCode||2,frequencyValue:m?Number(m[1]):1,frequencyText:'Monthly',calculationMode:'Calendar'};
  }
  if(lower.includes('quarter'))return {pmCode:pmCode||3,frequencyValue:3,frequencyText:'Quarterly',calculationMode:'Calendar'};
  if(lower.includes('semi'))return {pmCode:pmCode||4,frequencyValue:6,frequencyText:'Semi Annual',calculationMode:'Calendar'};
  if(lower.includes('annual')||lower.includes('year'))return {pmCode:pmCode||5,frequencyValue:12,frequencyText:'Annual',calculationMode:'Calendar'};
  return {pmCode:pmCode||null,frequencyValue:null,frequencyText:freqRaw||'',calculationMode:''};
}
function mergeMigratedTaskPairs(ws,startRow,lastRow){
  // Merge Part + Maintenance Task only for a real planned/done pair.
  // This is safe for Annual, Monthly and the blank reference templates.
  for(let r=startRow;r<=lastRow;r++){
    const rt=String(safeText(ws,r,8)||'').trim().toLowerCase();
    if(rt!=='planed')continue;
    const d=r+1;
    if(d>lastRow || String(safeText(ws,d,8)||'').trim().toLowerCase()!=='done')continue;

    for(const c of [6,7]){
      try{ws.unMergeCells(r,c,d,c)}catch(_){}
      const v=ws.getCell(r,c).value||ws.getCell(d,c).value;
      ws.getCell(r,c).value=v||null;
      ws.getCell(d,c).value=null;
      try{ws.mergeCells(r,c,d,c)}catch(_){}
      ws.getCell(r,c).alignment={
        vertical:'middle',
        horizontal:(c===7?'left':'center'),
        wrapText:true
      };
    }
  }
}
function copyTaskPairs(oldWs,newWs,kind){
  const oldTasks=migrationTaskRows(oldWs);
  const newStart=pmDataStartRow(newWs);
  const maxCol=kind==='annual'?67:25;
  clearDataArea(newWs,newStart,maxCol);
  let nr=newStart,count=0;

  // Carry merged legacy identity values forward.
  let cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};

  for(const t of oldTasks){
    const pr=t.row,dr=pr+1;
    const effective={
      section:t.section||cur.section,
      machineNo:t.machineNo||cur.machineNo,
      machineName:t.machineName||cur.machineName,
      machineIdentified:t.machineIdentified||cur.machineIdentified,
      type:t.type||cur.type
    };
    Object.assign(cur,effective);
    const fi=legacyFrequencyInfo(oldWs,pr,kind);

    // Planned row
    for(let c=1;c<=7;c++)newWs.getCell(nr,c).value=migrationCellValue(oldWs.getCell(pr,c));
    newWs.getCell(nr,1).value=effective.section||null;
    newWs.getCell(nr,2).value=effective.machineNo||null;
    newWs.getCell(nr,3).value=effective.machineName||null;
    newWs.getCell(nr,4).value=effective.machineIdentified||null;
    newWs.getCell(nr,5).value=effective.type||null;
    newWs.getCell(nr,8).value='planed';
    newWs.getCell(nr,9).value=fi.pmCode;
    newWs.getCell(nr,10).value=fi.frequencyValue;
    newWs.getCell(nr,11).value=fi.frequencyText||null;
    newWs.getCell(nr,12).value=fi.calculationMode||null;
    newWs.getCell(nr,13).value=null;
    newWs.getCell(nr,14).value=null;
    newWs.getCell(nr,15).value=null;

    // Done row: identity is filled temporarily; merge rebuild handles visual grouping.
    for(let c=1;c<=7;c++)newWs.getCell(nr+1,c).value=migrationCellValue(oldWs.getCell(dr,c));
    newWs.getCell(nr+1,1).value=effective.section||null;
    newWs.getCell(nr+1,2).value=effective.machineNo||null;
    newWs.getCell(nr+1,3).value=effective.machineName||null;
    newWs.getCell(nr+1,4).value=effective.machineIdentified||null;
    newWs.getCell(nr+1,5).value=effective.type||null;
    newWs.getCell(nr+1,8).value='done';

    if(kind==='annual'){
      // Preserve the exact legacy annual PM week plan P:BO.
      for(let c=16;c<=67;c++){
        newWs.getCell(nr,c).value=migrationCellValue(oldWs.getCell(pr,c));
        newWs.getCell(nr+1,c).value=migrationCellValue(oldWs.getCell(dr,c));
      }
    }else{
      // Preserve W1-W5 execution values.
      for(let c=16;c<=20;c++){
        newWs.getCell(nr,c).value=migrationCellValue(oldWs.getCell(pr,c));
        newWs.getCell(nr+1,c).value=migrationCellValue(oldWs.getCell(dr,c));
      }
      // Preserve legacy username/execution-date audit into Professional U:Y.
      for(let i=0;i<5;i++){
        newWs.getCell(nr,21+i).value=migrationCellValue(oldWs.getCell(pr,33+i));
        newWs.getCell(nr+1,21+i).value=migrationCellValue(oldWs.getCell(dr,33+i));
      }
    }
    nr+=2;count++;
  }
  return {tasks:count,newStart,lastRow:nr-1};
}
function applyAnnualProfessionalKpis(ws){
  const start=pmDataStartRow(ws);
  const last=Math.max(start,ws.rowCount);
  const H=`H${start}:H${last}`, B=`B${start}:B${last}`;
  const I=`I${start}:I${last}`, L=`L${start}:L${last}`;

  ws.getCell(3,1).value={formula:`IFERROR(SUMPRODUCT((${H}="planed")*(${B}<>"")/COUNTIFS(${B},${B}&"",${H},"planed")),0)`};
  ws.getCell(3,2).value={formula:`COUNTIF(${H},"planed")`};

  // Calendar KPI is driven by PM Code, not free-text Frequency Text.
  // This makes legacy spellings such as "Montly" and values such as "2month"
  // classify correctly without affecting Running Hours tasks.
  ws.getCell(3,3).value={formula:`COUNTIFS(${H},"planed",${L},"Calendar",${I},1)`};
  ws.getCell(3,4).value={formula:`COUNTIFS(${H},"planed",${L},"Calendar",${I},2)`};
  ws.getCell(3,5).value={formula:`COUNTIFS(${H},"planed",${L},"Calendar",${I},3)`};
  ws.getCell(3,6).value={formula:`COUNTIFS(${H},"planed",${L},"Calendar",${I},4)`};
  ws.getCell(3,7).value={formula:`COUNTIFS(${H},"planed",${L},"Calendar",${I},5)`};
  ws.getCell(3,8).value={formula:`COUNTIFS(${H},"planed",${L},"Running Hours")`};
}

function pmColorForCode(v){
  const code=Number(v);
  return {
    1:'FF92D050',
    2:'FF00B0F0',
    3:'FF0070C0',
    4:'FF275417',
    5:'FF002060'
  }[code]||null;
}




// STEP 96: keep Monthly numeric planning cells numeric.
// ExcelJS can deserialize a numeric value as a Date when a copied style carries a date numFmt.
// Convert those accidental 1900 dates back to their Excel serial and enforce the correct formats.
function excelDateBackToSerial(v){
  if(!(v instanceof Date))return v;
  return Math.round((Date.UTC(v.getFullYear(),v.getMonth(),v.getDate())-Date.UTC(1899,11,30))/86400000);
}
function normalizeMonthlyPlanningNumberFormats(ws,startRow,lastRow){
  const first=Math.max(7,Number(startRow)||7), last=Math.max(first,Number(lastRow)||first);
  for(let r=first;r<=last;r++){
    const kind=String(safeText(ws,r,8)||'').trim().toLowerCase();
    // I = PM Code, J = Frequency Value: always numeric/general, never Date.
    for(const c of [9,10]){
      const cell=ws.getCell(r,c);
      if(cell.value instanceof Date)cell.value=excelDateBackToSerial(cell.value);
      cell.numFmt=(c===9?'0':'0.##');
    }
    // P:T = W1..W5.
    // Planned rows show PM COLOR only; Done rows hide zero but show Done=6.
    for(let c=16;c<=20;c++){
      const cell=ws.getCell(r,c);
      if(cell.value instanceof Date)cell.value=excelDateBackToSerial(cell.value);
      cell.numFmt=(kind==='planed'?';;;':'[Green][=6]"✓ Done";;;');
    }
    // U:Y audit: planned row is username text; done row is execution date.
    for(let c=21;c<=25;c++){
      const cell=ws.getCell(r,c);
      cell.numFmt=(kind==='done'?'dd/mm/yyyy':'General');
    }
  }
}


function normalizeMonthlyScheduleDisplay(ws,startRow,lastRow){
  const first=Math.max(pmDataStartRow(ws),Number(startRow)||pmDataStartRow(ws));
  const last=Math.max(first,Number(lastRow)||first);

  for(let r=first;r<=last;r++){
    const kind=String(safeText(ws,r,8)||'').trim().toLowerCase();

    for(let c=16;c<=20;c++){
      const cell=ws.getCell(r,c);
      const n=Number(cell.value);

      // Always remove inherited/stale fill first.
      // This is critical when a new Monthly task row is copied from another task.
      cell.fill={type:'pattern',pattern:'none'};
      cell.alignment={...(cell.alignment||{}),horizontal:'center',vertical:'middle'};

      if(kind==='planed'){
        // Planned row: display ONLY the PM color of the week actually planned.
        // The numeric code remains stored but is visually hidden.
        cell.numFmt=';;;';

        if([1,2,3,4,5].includes(n)){
          const argb=pmColorForCode(n);
          if(argb){
            cell.fill={type:'pattern',pattern:'solid',fgColor:{argb}};
            cell.font={
              ...(cell.font||{}),
              color:{argb:(n>=3?'FFFFFFFF':'FF000000')},
              bold:true
            };
          }
        }
      }else if(kind==='done'){
        // Done row: no background color.
        // 0 is empty; 6 is displayed as a green check while remaining numeric 6 internally.
        cell.numFmt='[Green][=6]"✓ Done";;;';
        cell.font={...(cell.font||{}),color:{argb:'FF008000'},bold:true};
      }
    }
  }
}

function formatProfessionalMonthlyLayout(ws,startRow,lastRow){
  const last=Math.max(startRow,lastRow||startRow);

  // Header / KPI heights.
  ws.getRow(1).height=36;
  ws.getRow(2).height=28;
  ws.getRow(3).height=24;
  ws.getRow(4).height=20;
  ws.getRow(5).height=22;
  ws.getRow(6).height=32;

  // Week date rows are directly BELOW their group titles.
  for(let c=16;c<=20;c++){
    const cell=ws.getCell(5,c);
    cell.alignment={horizontal:'center',vertical:'middle'};
    cell.font={name:'Arial',size:9,bold:true};
    cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEAF4E2'}};
    setThinBorder(cell);
  }
  for(let c=21;c<=25;c++){
    const cell=ws.getCell(5,c);
    cell.alignment={horizontal:'center',vertical:'middle'};
    cell.font={name:'Arial',size:9,bold:true};
    cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}};
    setThinBorder(cell);
  }

  // Group titles use colors different from the KPI band.
  const pmTitle=ws.getCell('P4');
  pmTitle.alignment={horizontal:'center',vertical:'middle'};
  pmTitle.font={name:'Arial',size:10,bold:true};
  pmTitle.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD9EAF7'}};

  const title=ws.getCell('U4');
  title.alignment={horizontal:'center',vertical:'middle'};
  title.font={name:'Arial',size:10,bold:true};
  title.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFCE4D6'}};

  // Make the KPI block visually self-contained.
  for(let c=1;c<=8;c++){
    ws.getCell(2,c).alignment={horizontal:'center',vertical:'middle',wrapText:true};
    ws.getCell(3,c).alignment={horizontal:'center',vertical:'middle'};
    setThinBorder(ws.getCell(2,c)); setThinBorder(ws.getCell(3,c));
  }
  for(let c=10;c<=25;c++){
    ws.getCell(2,c).alignment={horizontal:'center',vertical:'middle',wrapText:true};
    ws.getCell(3,c).alignment={horizontal:'center',vertical:'middle'};
    setThinBorder(ws.getCell(2,c)); setThinBorder(ws.getCell(3,c));
  }

  // Column header alignment.
  for(let c=1;c<=25;c++){
    const cell=ws.getCell(6,c);
    cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};
  }

  // Data rows: readable task text, compact done row, centered PM/Audit.
  for(let r=startRow;r<=last;r++){
    const kind=String(safeText(ws,r,8)||'').trim().toLowerCase();
    ws.getRow(r).height=kind==='planed'?28:18;

    for(let c=1;c<=15;c++){
      ws.getCell(r,c).alignment={
        vertical:'middle',
        horizontal:(c===6||c===7||c===15)?'left':'center',
        wrapText:true
      };
    }
    for(let c=16;c<=25;c++){
      ws.getCell(r,c).alignment={horizontal:'center',vertical:'middle',wrapText:true};
    }
  }

  // Professional column widths.
  const widths={
    1:16,2:12,3:21,4:18,5:14,6:20,7:40,8:11,
    9:9,10:13,11:16,12:16,13:20,14:14,15:22
  };
  for(const [c,w] of Object.entries(widths))ws.getColumn(Number(c)).width=w;
  for(let c=16;c<=20;c++)ws.getColumn(c).width=9;
  for(let c=21;c<=25;c++)ws.getColumn(c).width=18;

  // Final Monthly-only visual pass:
  // - planned weeks = correct PM color only
  // - non-planned weeks = blank/no fill
  // - Done 6 = green check
  // - Done 0 = blank
  normalizeMonthlyPlanningNumberFormats(ws,startRow,last);
  normalizeMonthlyScheduleDisplay(ws,startRow,last);

  // Freeze task heading and machine identity columns.
  ws.views=[{state:'frozen',ySplit:6,xSplit:3}];

  // STEP138: formatting is called after branding in multiple flows. Keep the
  // Excel logo container size as the final operation so row/column defaults
  // cannot overwrite the user's Excel logo settings.
  fitExcelLogoCell(ws,normalizeExcelLogo(currentBranding().excelLogo));
}


function formatProfessionalAnnualLayout(ws,startRow,lastRow){
  const first=Math.max(7,Number(startRow)||7);
  const last=Math.max(first,Number(lastRow)||first);

  ws.getRow(1).height=36;
  ws.getRow(2).height=28;
  ws.getRow(3).height=24;
  ws.getRow(4).height=22;
  ws.getRow(6).height=32;

  const widths={
    1:16,2:13,3:22,4:18,5:14,6:20,7:40,8:11,
    9:9,10:13,11:16,12:16,13:20,14:14,15:22
  };
  for(const [c,w] of Object.entries(widths))ws.getColumn(Number(c)).width=w;
  for(let c=16;c<=67;c++)ws.getColumn(c).width=8;

  for(let c=1;c<=67;c++){
    const h=ws.getCell(6,c);
    h.alignment={horizontal:'center',vertical:'middle',wrapText:true};
    setThinBorder(h);
  }

  for(let r=first;r<=last;r++){
    const kind=String(safeText(ws,r,8)||'').trim().toLowerCase();
    ws.getRow(r).height=kind==='planed'?28:18;

    for(let c=1;c<=15;c++){
      const cell=ws.getCell(r,c);
      cell.alignment={
        vertical:'middle',
        horizontal:(c===6||c===7||c===15)?'left':'center',
        wrapText:true
      };
      setThinBorder(cell);
    }

    for(let c=16;c<=67;c++){
      const cell=ws.getCell(r,c);
      cell.alignment={horizontal:'center',vertical:'middle'};
      setThinBorder(cell);

      // Planned PM codes are represented by color only.
      if(kind==='planed')cell.numFmt=';;;';
      // Done/blank row: hide zero but keep any real value visible.
      else if(kind==='done')cell.numFmt='0;-0;;@';
    }
  }

  applyPmScheduleColors(ws,'annual',first,last);
  ws.getColumn(68).hidden=true; // BP Machine Link
  ws.getColumn(69).hidden=true; // BQ Task ID
  ws.views=[{state:'frozen',ySplit:6,xSplit:3}];

  // STEP138: re-apply Excel logo cell sizing LAST. The old code reset row 1
  // to 36 and columns A:C to template widths after branding, so the cell did
  // not visibly change even when the logo image did.
  fitExcelLogoCell(ws,normalizeExcelLogo(currentBranding().excelLogo));
}

function restoreMonthlyKpiHeaders(ws){
  const headers={
    J2:'W1 Plan', K2:'W1 Done', L2:'W1 %',
    M2:'W2 Plan', N2:'W2 Done', O2:'W2 %',
    P2:'W3 Plan', Q2:'W3 Done', R2:'W3 %',
    S2:'W4 Plan', T2:'W4 Done', U2:'W4 %',
    V2:'W5 Plan', W2:'W5 Done', X2:'W5 %',
    Y2:'Month'
  };
  for(const [addr,text] of Object.entries(headers))ws.getCell(addr).value=text;
}

function forceExcelRecalculation(workbook){
  // Ask Excel to recalculate all formulas when the workbook is opened.
  workbook.calcProperties=workbook.calcProperties||{};
  workbook.calcProperties.calcMode='auto';
  workbook.calcProperties.fullCalcOnLoad=true;
  workbook.calcProperties.forceFullCalc=true;
}

function hideMonthlyHelperColumns(ws){
  // Visible report ends at Y:
  // P:T = W1-W5
  // U:Y = Audit W1-W5
  //
  // Z:AK are unused and hidden.
  for(let c=26;c<=37;c++)ws.getColumn(c).hidden=true;

  // AL:AW are calculation helpers and remain fully hidden.
  for(let c=38;c<=49;c++)ws.getColumn(c).hidden=true;
  // AX:AY are stable Machine Link / Task ID. Never shown to operators.
  for(let c=50;c<=51;c++)ws.getColumn(c).hidden=true;
}

function applyPmScheduleColors(ws,kind,startRow,lastRow){
  const firstCol=16; // P
  const lastCol=kind==='annual'?67:20; // BO for Annual, T for Monthly
  for(let r=startRow;r<=lastRow;r++){
    const rowType=String(safeText(ws,r,8)||'').trim().toLowerCase();
    for(let c=firstCol;c<=lastCol;c++){
      const cell=ws.getCell(r,c);
      const n=Number(cell.value);

      // Always remove stale schedule fill first.
      // This prevents old green/blue colors remaining after a week is cleared.
      cell.fill={type:'pattern',pattern:'none'};
      cell.alignment={...(cell.alignment||{}),horizontal:'center',vertical:'middle'};

      if(rowType==='done'){
        // Keep numeric 6 internally for all calculations, but show only a green check.
        // Zero/blank stays visually empty.
        cell.numFmt='[Green][=6]"✓ Done";;;';
        cell.font={...(cell.font||{}),color:{argb:'FF008000'},bold:true};
        continue;
      }

      // Planned rows: retain numeric PM code internally but display color only.
      cell.numFmt=';;;';
      const argb=pmColorForCode(cell.value);
      if(!argb)continue;
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb}};
      cell.font={
        ...(cell.font||{}),
        color:{argb:(n>=3?'FFFFFFFF':'FF000000')},
        bold:true
      };
    }
  }
}
function verifyMigratedTaskCount(ws,expected){
  const actual=migrationTaskRows(ws).length;
  if(actual!==expected){
    throw new Error(`Migration verification failed: expected ${expected} tasks but generated ${actual}`);
  }
  return actual;
}

function syncMigratedAnnualToSystemMaster(line,year,ws){
  const master=loadPmMaster();
  const lineEntry=pmLineEntry(master,line);
  if(!lineEntry)throw new Error(`Line not found in Lines & Machines: ${line}`);

  const migrated=migrationTaskRows(ws);
  if(!migrated.length)throw new Error('No migrated Annual tasks were found for System Master import');

  // SAFE UPDATE MODE (V1.6): importing an updated Annual Excel must continue the
  // existing PM program. Existing task IDs and execution/history are preserved.
  // Rows in the upload add/update tasks; missing rows are NOT treated as deletes.
  // Use the normal Delete Task / Delete Line controls for intentional deletion.
  const norm=v=>String(v||'').trim().replace(/\s+/g,' ').toLowerCase();
  const taskKey=(machineNo,part,maintenance)=>[norm(machineNo),norm(part),norm(maintenance)].join('|');

  const store=loadSystemTasks();
  const existingForYear=store.tasks.filter(t=>same122(t.line,line)&&Number(t.year)===Number(year));
  const existingByKey=new Map();
  for(const t of existingForYear){
    const k=taskKey(t.machineNo,t.part,t.maintenance);
    if(k&&!existingByKey.has(k))existingByKey.set(k,t);
  }

  // Merge machines rather than replacing the list. This prevents an updated Excel
  // containing only new/changed machines from deleting the rest of the line.
  const machines=Array.isArray(lineEntry.machines)?lineEntry.machines.map(m=>({...m})):[];
  const machineMap=new Map(machines.map(m=>[norm(m.machineNo),m]));
  let cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:'',manufacturer:''};
  for(const t of migrated){
    const effective={
      section:String(t.section||cur.section||'').trim(),
      machineNo:String(t.machineNo||cur.machineNo||'').trim(),
      machineName:String(t.machineName||cur.machineName||'').trim(),
      machineIdentified:String(t.machineIdentified||cur.machineIdentified||'').trim(),
      type:String(t.type||cur.type||'').trim(),
      manufacturer:String(t.manufacturer||cur.manufacturer||'').trim()
    };
    Object.assign(cur,effective);
    if(!effective.machineNo)continue;
    const key=norm(effective.machineNo);
    if(!machineMap.has(key)){
      const m={...effective,expectedWeeklyHours:120,enabled:true};
      machines.push(m);machineMap.set(key,m);
    }else{
      const m=machineMap.get(key);
      // Uploaded non-blank identity fields update the machine; blank cells never erase data.
      for(const k of ['section','machineName','machineIdentified','type','manufacturer'])if(effective[k])m[k]=effective[k];
    }
  }
  if(!machines.length)throw new Error('No machines were found in the migrated Annual file');

  const now=new Date().toISOString();
  const imported=[]; const touchedIds=new Set();
  cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:'',manufacturer:''};
  for(const t of migrated){
    const pr=t.row;
    const effective={
      section:String(t.section||cur.section||'').trim(),
      machineNo:String(t.machineNo||cur.machineNo||'').trim(),
      machineName:String(t.machineName||cur.machineName||'').trim(),
      machineIdentified:String(t.machineIdentified||cur.machineIdentified||'').trim(),
      type:String(t.type||cur.type||'').trim(),
      manufacturer:String(t.manufacturer||cur.manufacturer||'').trim()
    };
    Object.assign(cur,effective);
    if(!effective.machineNo)throw new Error(`Task row ${pr} has no Machine No.`);

    let weeks=[]; const scheduleCodes={};
    for(let c=16;c<=67;c++){
      const code=Number(ws.getCell(pr,c).value);
      if([1,2,3,4,5].includes(code)){const w=c-15;weeks.push(w);scheduleCodes[w]=code;}
    }
    let dominant=weeks.length?Number(scheduleCodes[weeks[0]])||1:(Number(ws.getCell(pr,9).value)||1);
    const calc=String(safeText(ws,pr,12)||'Calendar').trim()||'Calendar';
    const freqValue=Number(ws.getCell(pr,10).value)||0;
    const freqText=String(safeText(ws,pr,11)||'').trim();
    const calcLower=calc.toLowerCase();
    const importedHoursMethod=calcLower.includes('running')||calcLower.includes('plc')||/\bhr\b|hour/.test(freqText.toLowerCase());
    let importedFirstPmDate=String(safeText(ws,pr,14)||'').trim();
    if(importedHoursMethod && freqValue>0){
      const mh=Number(machineMap.get(norm(effective.machineNo))?.expectedWeeklyHours||120);
      const hs=generateRunningHoursScheduleForYear(Number(year),`${Number(year)}-01-01`,freqValue,mh);
      weeks=[...hs.weeks];for(const k of Object.keys(scheduleCodes))delete scheduleCodes[k];
      for(const w of weeks)scheduleCodes[w]=hs.derivedCode;dominant=hs.derivedCode||dominant;importedFirstPmDate=hs.firstDueDate||'';
    }

    const part=String(safeText(ws,pr,6)||'').trim();
    const maintenance=String(safeText(ws,pr,7)||'').trim();
    const oldTask=existingByKey.get(taskKey(effective.machineNo,part,maintenance));
    const taskId=oldTask?.taskId||taskLinkId();
    touchedIds.add(taskId);
    writeStableTaskLink(ws,pr,pr+1,effective.machineNo,taskId);

    imported.push({
      ...(oldTask||{}),
      taskId,line,year:Number(year),machineNo:effective.machineNo,part,maintenance,
      frequency:freqText,code:dominant,effectiveCode:dominant,
      calculationMethod:calcLower.includes('plc')?'plc_hours':importedHoursMethod?'hours':'calendar',
      calendarFrequency:importedHoursMethod?'':freqText,intervalHours:importedHoursMethod?freqValue:0,
      expectedWeeklyHours:Number(oldTask?.expectedWeeklyHours||0),useMachineExpectedHours:oldTask?.useMachineExpectedHours!==false,
      runningHoursCounterStartDate:importedHoursMethod?(oldTask?.runningHoursCounterStartDate||`${Number(year)}-01-01`):'',
      firstPmDate:importedHoursMethod?(oldTask?.firstPmDate||importedFirstPmDate):String(safeText(ws,pr,14)||oldTask?.firstPmDate||'').trim(),
      scheduleWeeks:weeks,scheduleCodes,notes:String(safeText(ws,pr,15)||oldTask?.notes||'').trim(),
      active:oldTask?.active!==false,isActive:oldTask?.isActive!==false,
      createdAt:oldTask?.createdAt||now,updatedAt:now
    });
  }

  // Preserve every existing task that was not present in the upload. This is the
  // key protection against an Excel update resetting/deleting the PM master.
  const untouched=existingForYear.filter(t=>!touchedIds.has(t.taskId));
  const otherTasks=store.tasks.filter(t=>!(same122(t.line,line)&&Number(t.year)===Number(year)));

  lineEntry.machines=machines;lineEntry.enabled=true;lineEntry.fresh122=false;lineEntry.createdYear=Number(year);savePmMaster(master);
  store.tasks=[...otherTasks,...untouched,...imported];
  store.initialized=store.initialized||{};store.initialized[String(line).toLowerCase()+'|'+Number(year)]=true;saveSystemTasks(store);
  // Do NOT purge row/task execution state here: reused taskIds keep Done/Deferred/history.

  return {machines:machines.length,tasks:imported.length,updatedTasks:imported.filter(t=>existingByKey.has(taskKey(t.machineNo,t.part,t.maintenance))).length,newTasks:imported.filter(t=>!existingByKey.has(taskKey(t.machineNo,t.part,t.maintenance))).length,preservedTasks:untouched.length,replacedTasks:0};
}

async function writeMigratedWorkbookSafe(wb,finalPath){
  const dir=path.dirname(finalPath);
  fs.mkdirSync(dir,{recursive:true});
  const base=path.basename(finalPath);
  const tmp=path.join(dir,`.${base}.migration.${process.pid}.${Date.now()}.tmp.xlsx`);
  let backup=null;
  try{
    // Migration workbook is already a fresh Professional ExcelJS workbook.
    // Do not pass it through legacy shared-formula normalization.
    await wb.xlsx.writeFile(tmp);

    const st=fs.statSync(tmp);
    if(!st||st.size<1000)throw new Error('Generated Professional Excel is empty');

    // Validate the generated XLSX before replacing the active workbook.
    const verify=new ExcelJS.Workbook();
    await verify.xlsx.readFile(tmp);
    if(!verify.worksheets||!verify.worksheets.length){
      throw new Error('Generated Professional Excel has no worksheet');
    }

    if(fs.existsSync(finalPath)){
      backup=migrationBackupPath(finalPath);
      fs.copyFileSync(finalPath,backup);
    }

    fs.copyFileSync(tmp,finalPath);
    fs.unlinkSync(tmp);
    return {backup};
  }catch(e){
    try{if(fs.existsSync(tmp))fs.unlinkSync(tmp)}catch(_){}
    throw e;
  }
}
app.post('/api/excel-migration/preview',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const kind=String(req.body?.kind||'').toLowerCase();
  if(!['annual','monthly'].includes(kind))return res.status(400).json({error:'Select Annual or Monthly'});
  const {wb}=await loadStagedWorkbook(req.body?.uploadToken);
  const simple=kind==='annual'?simplePmWorkbookToLegacy(wb,Number(req.body?.year)||new Date().getFullYear()):null;
  const ws=simple?.ws||(kind==='annual'?(wb.getWorksheet('Full')||wb.worksheets[0]):(wb.getWorksheet('Sheet1')||wb.worksheets[0]));
  if(!ws)throw new Error('Worksheet not found');
  let tasks=migrationTaskRows(ws);
  // Fallback for older PM files: discover planned rows directly if helper parsing returns nothing.
  if(!tasks.length){
    const found=[];
    for(let r=1;r<=ws.rowCount;r++){
      const rt=String(ws.getCell(r,8).value||'').trim().toLowerCase();
      if(rt==='planed'||rt==='planned'){
        found.push({row:r,machineNo:String(ws.getCell(r,2).value||'').trim()});
      }
    }
    tasks=found;
  }
  const machines=[...new Set(tasks.map(t=>String(t.machineNo||'').trim()).filter(Boolean))];
  const done=kind==='monthly'?tasks.reduce((n,t)=>n+[16,17,18,19,20].filter(c=>Number(ws.getCell(t.row+1,c).value)===6).length,0):0;
  res.json({ok:true,kind,tasks:tasks.length,machines:machines.length,done,rows:ws.rowCount,sheet:ws.name});
 }catch(e){console.error(e);res.status(400).json({error:e.message});}
});
app.post('/api/excel-migration/convert',requireAuth,requireAdmin,async(req,res)=>{
 let stage='start';
 try{
  const b=req.body||{},kind=String(b.kind||'').toLowerCase(),line=cleanName(b.line),year=Number(b.year);
  if(!['annual','monthly'].includes(kind))return res.status(400).json({error:'Select Annual or Monthly'});
  if(!line||!year)return res.status(400).json({error:'Line and year are required'});

  stage='load old uploaded workbook';
  const old=await loadStagedWorkbook(b.uploadToken);
  const simple=kind==='annual'?simplePmWorkbookToLegacy(old.wb,year):null;
  const oldWs=simple?.ws||(kind==='annual'?(old.wb.getWorksheet('Full')||old.wb.worksheets[0]):(old.wb.getWorksheet('Sheet1')||old.wb.worksheets[0]));
  if(!oldWs)throw new Error('Old worksheet not found');

  let templatePath,targetPath,newWb,newWs,month=0;
  if(kind==='annual'){
    stage='build professional annual template';
    targetPath=annualPathFor(line,year);
    ({wb:newWb,ws:newWs}=buildProfessionalAnnualWorkbook(line,year));
  }else{
    month=Number(b.month);
    if(!Number.isInteger(month)||month<1||month>12)return res.status(400).json({error:'Select month'});
    stage='build professional monthly template';
    targetPath=monthlyPathFor(line,year,month);
    ({wb:newWb,ws:newWs}=buildProfessionalMonthlyWorkbook(line,year,month));
  }

  stage='copy old PM data';
  const copied=copyTaskPairs(oldWs,newWs,kind);
  if(!copied.tasks)throw new Error('No PM tasks were found in the uploaded Excel');

  stage='apply company branding';
  applyBrandingToWorkbook(newWb,newWs,kind);

  if(kind==='annual'){
    stage='rebuild annual KPI';
    applyAnnualProfessionalKpis(newWs);
  }else{
    stage='rebuild monthly KPI';
    newWs.getCell(3,25).value=`${line} — ${monthShort(month)} ${year}`;
    newWs.getCell(2,16).value=year;
    newWs.getCell(3,16).value=monthShort(month);
    rebuildMonthlyTaskFormulas(newWs);
  }

  stage='apply PM schedule colors';
  applyPmScheduleColors(newWs,kind,copied.newStart,copied.lastRow);

  stage='rebuild machine groups';
  try{rebuildTaskGroupMerges(newWs)}catch(e){console.warn('[MIGRATION MERGE]',e.message)}

  stage='merge planned/done task pairs';
  try{mergeMigratedTaskPairs(newWs,copied.newStart,copied.lastRow)}catch(e){console.warn('[MIGRATION TASK PAIR MERGE]',e.message)}

  stage='verify migrated task count';
  verifyMigratedTaskCount(newWs,copied.tasks);

  let masterSync=null;
  if(kind==='annual'){
    stage='import machines and tasks into System Master';
    masterSync=syncMigratedAnnualToSystemMaster(line,year,newWs);
    if(simple?.rows)enrichArabicMaintenance(line,year,simple.rows);
  }

  stage='write professional Excel';
  const writeResult=await writeMigratedWorkbookSafe(newWb,targetPath);
  const backup=writeResult.backup;

  if(kind==='annual'){
    stage='rebuild Annual from System Master';
    await rebuildAnnualFromSystemMaster(line,year);
  }

  stage='finish';
  try{fs.unlinkSync(old.p)}catch(_){}
  audit('PM_EXCEL_MIGRATED',req.user,{kind,line,year,month:month||null,tasks:copied.tasks,target:path.basename(targetPath),backup:backup?path.basename(backup):null});
  res.json({
    ok:true,kind,line,year,month:month||null,tasks:copied.tasks,
    machines:masterSync?masterSync.machines:null,
    systemTasks:masterSync?masterSync.tasks:null,
    replacedTasks:masterSync?masterSync.replacedTasks:null,
    target:path.basename(targetPath),
    backup:backup?path.basename(backup):null,
    message:masterSync
      ?`Converted successfully: ${masterSync.machines} machines and ${masterSync.tasks} PM tasks imported into System Master`
      :`Converted successfully: ${copied.tasks} PM tasks`
  });
 }catch(e){
  console.error('[EXCEL MIGRATION CONVERT]',stage,e&&e.stack?e.stack:e);
  res.status(500).json({error:`${stage}: ${e.message||e}`});
 }
});

// ---------- PM TASK MANAGEMENT (ADMIN) ----------
function validPmLine(line){ return configuredPmLines().includes(line); }

async function openAnnualForLine(line,year){
  if(!validPmLine(line)) throw new Error('Invalid PM line');
  const file=`${line} PM ${year}.xlsx`;
  const fullPath=path.join(ANNUAL_DIR,file);
  if(!fs.existsSync(fullPath)) throw new Error(`Annual file not found: ${file}`);
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(fullPath);
  return {file,fullPath,wb,ws:wb.getWorksheet('Full')||wb.worksheets[0]};
}

function taskRowsFromAnnual(ws){
  const out=[];
  let cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
  for(let r=pmDataStartRow(ws);r<=ws.rowCount;r++){
    const activity=String(safeText(ws,r,8)||'').trim().toLowerCase();
    const vals={
      section:String(safeText(ws,r,1)||'').trim(),
      machineNo:String(safeText(ws,r,2)||'').trim(),
      machineName:String(safeText(ws,r,3)||'').trim(),
      machineIdentified:String(safeText(ws,r,4)||'').trim(),
      type:String(safeText(ws,r,5)||'').trim()
    };
    const linkedMachine=stableMachineLink(ws,r);
    if(linkedMachine){
      const master=authoritativeMachineIdentity(String(safeText(ws,3,10)||'').trim(),linkedMachine);
      vals.machineNo=linkedMachine;
      if(master){vals.section=master.section;vals.machineName=master.machineName;vals.machineIdentified=master.machineIdentified;vals.type=master.type;}
    }
    if(vals.machineNo && vals.machineNo!==cur.machineNo)cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
    if(vals.section)cur.section=vals.section;
    if(vals.machineNo)cur.machineNo=vals.machineNo;
    if(vals.machineName)cur.machineName=vals.machineName;
    if(vals.machineIdentified && vals.machineIdentified!=='0 0')cur.machineIdentified=vals.machineIdentified;
    if(vals.type)cur.type=vals.type;
    if(activity!=='planed') continue;

    const codes=[];
    for(let c=16;c<=67;c++){
      const n=Number(ws.getCell(r,c).value);
      if([1,2,3,4,5].includes(n)) codes.push({col:c,code:n});
    }
    const counts={1:0,2:0,3:0,4:0,5:0};
    codes.forEach(x=>counts[x.code]++);
    let dominantCode=1, max=-1;
    for(let k=1;k<=5;k++) if(counts[k]>max){max=counts[k];dominantCode=k;}

    out.push({
      row:r,taskId:stableTaskId(ws,r),
      section:cur.section,
      machineNo:cur.machineNo,
      machineName:cur.machineName,
      machineIdentified:cur.machineIdentified,
      type:cur.type,
      part:String(safeText(ws,r,6)||'').trim(),
      maintenance:String(safeText(ws,r,7)||'').trim(),
      frequency:String(safeText(ws,r,11)||'').trim(),
      isActive:true,
      dominantCode,
      plannedWeeks:codes.length
    });
  }
  return out;
}


function machinePlaceholderRows(ws){
  const out=[];
  for(let r=pmDataStartRow(ws);r<=ws.rowCount;r++){
    if(String(safeText(ws,r,8)||'').trim().toLowerCase()!=='machine')continue;
    out.push({
      row:r,
      section:String(safeText(ws,r,1)||'').trim(),
      machineNo:String(safeText(ws,r,2)||'').trim(),
      machineName:String(safeText(ws,r,3)||'').trim(),
      machineIdentified:String(safeText(ws,r,4)||'').trim(),
      type:String(safeText(ws,r,5)||'').trim()
    });
  }
  return out;
}

function machineFromMaster(line,machineNo){
  const x=loadPmMaster(),l=pmLineEntry(x,line);
  return pmMachineEntry(l,machineNo)||null;
}

function authoritativeMachineIdentity(line,machineNo){
  const m=machineFromMaster(line,machineNo);
  if(!m)return null;
  return {
    section:String(m.section||'').trim(),
    machineNo:String(m.machineNo||'').trim(),
    machineName:String(m.machineName||'').trim(),
    machineIdentified:String(m.machineIdentified||'').trim(),
    type:String(m.type||'').trim(),
    expectedWeeklyHours:Number(m.expectedWeeklyHours||120)
  };
}

function machineExpectedWeeklyHours(line,machineNo){
  const m=machineFromMaster(line,machineNo);
  return Number(m?.expectedWeeklyHours||120);
}
function effectiveTaskExpectedWeeklyHours(t){
  if(t?.useMachineExpectedHours===true || !(Number(t?.expectedWeeklyHours)>0)){
    return machineExpectedWeeklyHours(t?.line,t?.machineNo);
  }
  return Number(t.expectedWeeklyHours||0);
}
function recalcInheritedHoursTasksForMachine(line,machineNo){
  const store=loadSystemTasks();
  let changed=0;
  for(const t of store.tasks||[]){
    if(!same122(t.line,line)||!same122(t.machineNo,machineNo))continue;
    const method=normalizePlannerCalculationMethod(t);
    if(method!=='hours'&&method!=='plc_hours')continue;
    if(!(t.useMachineExpectedHours===true || !(Number(t.expectedWeeklyHours)>0)))continue;
    const year=Number(t.year), date=String(t.firstPmDate||'');
    if(!year||!/^\d{4}-\d{2}-\d{2}$/.test(date)||!(Number(t.intervalHours)>0))continue;
    try{
      const schedule=generateAnnualSchedule(year,date,{calculationMethod:method,calendarFrequency:t.calendarFrequency,intervalHours:Number(t.intervalHours),expectedWeeklyHours:machineExpectedWeeklyHours(line,machineNo)});
      t.scheduleWeeks=schedule.weeks;t.effectiveCode=schedule.derivedCode||t.effectiveCode||t.code;t.updatedAt=new Date().toISOString();changed++;
    }catch(_){}
  }
  if(changed)saveSystemTasks(store);
  return changed;
}

function excelLinkColumns(ws){
  const monthly=String(ws?.name||'').toLowerCase()==='sheet1';
  return monthly ? {machine:50,task:51} : {machine:68,task:69}; // Monthly AX:AY / Annual BP:BQ
}
function taskLinkId(){
  try{return crypto.randomUUID()}catch(_){return crypto.randomBytes(12).toString('hex')}
}
function writeStableTaskLink(ws,planRow,doneRow,machineNo,taskId=''){
  const c=excelLinkColumns(ws),id=String(taskId||'').trim()||taskLinkId(),no=String(machineNo||'').trim();
  for(const r of [planRow,doneRow]){
    ws.getCell(r,c.machine).value=no||null;
    ws.getCell(r,c.task).value=id;
  }
  ws.getColumn(c.machine).hidden=true;
  ws.getColumn(c.task).hidden=true;
  return id;
}
function stableMachineLink(ws,row){
  const c=excelLinkColumns(ws);
  return String(safeText(ws,row,c.machine)||'').trim();
}
function stableTaskId(ws,row){
  const c=excelLinkColumns(ws);
  return String(safeText(ws,row,c.task)||'').trim();
}
function writeStableMachineOnlyLink(ws,row,machineNo){
  const c=excelLinkColumns(ws);
  ws.getCell(row,c.machine).value=String(machineNo||'').trim()||null;
  ws.getCell(row,c.task).value=null;
  ws.getColumn(c.machine).hidden=true;
  ws.getColumn(c.task).hidden=true;
}
function backfillStableTaskLinks(ws){
  const c=excelLinkColumns(ws);
  let curMachine='',curTask='';
  for(let r=pmDataStartRow(ws);r<=ws.rowCount;r++){
    const rt=String(safeText(ws,r,8)||'').trim().toLowerCase();
    const rawNo=String(safeText(ws,r,2)||'').trim();
    if(rawNo)curMachine=rawNo;
    if(rt==='planed'){
      const existingNo=String(safeText(ws,r,c.machine)||'').trim();
      if(!existingNo&&curMachine)ws.getCell(r,c.machine).value=curMachine;
      let id=String(safeText(ws,r,c.task)||'').trim();
      if(!id){id=taskLinkId();ws.getCell(r,c.task).value=id}
      curTask=id;
      if(r+1<=ws.rowCount&&String(safeText(ws,r+1,8)||'').trim().toLowerCase()==='done'){
        ws.getCell(r+1,c.machine).value=String(safeText(ws,r,c.machine)||'').trim()||curMachine||null;
        ws.getCell(r+1,c.task).value=id;
      }
    }else if(rt==='machine'){
      const existingNo=String(safeText(ws,r,c.machine)||'').trim();
      if(!existingNo&&rawNo)ws.getCell(r,c.machine).value=rawNo;
      ws.getCell(r,c.task).value=null;
    }
  }
  ws.getColumn(c.machine).hidden=true;
  ws.getColumn(c.task).hidden=true;
}
function rebuildVisibleMachineIdentityFromStableLinks(ws,line){
  backfillStableTaskLinks(ws);
  const c=excelLinkColumns(ws),start=pmDataStartRow(ws);
  // Remove only A:E data merges. They are display-only.
  const ranges=[];
  if(ws._merges)for(const key of Object.keys(ws._merges)){
    const m=ws._merges[key]?.model;
    if(m&&m.left<=5&&m.right>=1&&m.bottom>=start)ranges.push(key);
  }
  for(const rg of ranges)try{ws.unMergeCells(rg)}catch(_){}

  for(let r=start;r<=ws.rowCount;r++){
    const rt=String(safeText(ws,r,8)||'').trim().toLowerCase();
    if(rt!=='planed'&&rt!=='done'&&rt!=='machine')continue;
    const no=String(safeText(ws,r,c.machine)||'').trim();
    if(!no)continue;
    const identity=authoritativeMachineIdentity(line,no);
    if(!identity)continue;
    ws.getCell(r,1).value=identity.section||null;
    ws.getCell(r,2).value=identity.machineNo||null;
    ws.getCell(r,3).value=identity.machineName||null;
    ws.getCell(r,4).value=identity.machineIdentified||null;
    ws.getCell(r,5).value=identity.type||null;
  }
  rebuildTaskGroupMerges(ws);
}


function materializeMachineIdentityForTask(ws,planRow,identity){
  if(!identity)return;
  const doneRow=planRow+1;
  // Unmerge only A:E ranges touching this task pair.
  const ranges=[];
  if(ws._merges)for(const key of Object.keys(ws._merges)){
    const m=ws._merges[key]?.model;
    if(!m)continue;
    if(m.left<=5&&m.right>=1&&m.bottom>=planRow&&m.top<=doneRow)ranges.push(key);
  }
  for(const rg of ranges)try{ws.unMergeCells(rg)}catch(_){}
  for(const r of [planRow,doneRow]){
    ws.getCell(r,1).value=identity.section||null;
    ws.getCell(r,2).value=identity.machineNo||null;
    ws.getCell(r,3).value=identity.machineName||null;
    ws.getCell(r,4).value=identity.machineIdentified||null;
    ws.getCell(r,5).value=identity.type||null;
  }
}

function validateTaskMachineIdentity(ws,planRow,expected){
  const got={
    section:String(safeText(ws,planRow,1)||'').trim(),
    machineNo:String(safeText(ws,planRow,2)||'').trim(),
    machineName:String(safeText(ws,planRow,3)||'').trim(),
    machineIdentified:String(safeText(ws,planRow,4)||'').trim(),
    type:String(safeText(ws,planRow,5)||'').trim()
  };
  return ['section','machineNo','machineName','machineIdentified','type'].every(k=>
    String(got[k]||'')===String(expected?.[k]||'')
  );
}


function lastLogicalDataRow(ws){
  let last=pmDataStartRow(ws)-1;
  for(let r=pmDataStartRow(ws);r<=ws.rowCount;r++){
    const rt=String(safeText(ws,r,8)||'').trim().toLowerCase();
    if(rt==='planed'||rt==='done'||rt==='machine')last=r;
  }
  return last;
}

function styleMachinePlaceholderRow(ws,row){
  ws.getRow(row).height=24;
  for(let c=1;c<=Math.min(ws.columnCount||67,67);c++){
    const cell=ws.getCell(row,c);
    cell.alignment={vertical:'middle',horizontal:(c===6||c===7||c===15)?'left':'center',wrapText:true};
    try{setThinBorder(cell)}catch(_){}
  }
  // No PM plan / no Done value for a machine-only row.
  const lastCol=String(ws.name||'').toLowerCase()==='full'?67:20;
  for(let c=16;c<=lastCol;c++){
    const cell=ws.getCell(row,c);
    cell.value=null;
    cell.fill={type:'pattern',pattern:'none'};
    cell.numFmt=';;;';
  }
}

function writeMachinePlaceholder(ws,row,m){
  ws.getCell(row,1).value=String(m.section||'').trim()||null;
  ws.getCell(row,2).value=String(m.machineNo||'').trim()||null;
  ws.getCell(row,3).value=String(m.machineName||'').trim()||null;
  ws.getCell(row,4).value=String(m.machineIdentified||'').trim()||null;
  ws.getCell(row,5).value=String(m.type||'').trim()||null;
  ws.getCell(row,6).value=null;
  ws.getCell(row,7).value=null;
  ws.getCell(row,8).value='machine';
  writeStableMachineOnlyLink(ws,row,m.machineNo);
  for(let c=9;c<=15;c++)ws.getCell(row,c).value=null;
  if(String(ws.name||'').toLowerCase()==='sheet1'){
    for(let c=21;c<=25;c++)ws.getCell(row,c).value=null;
  }
  styleMachinePlaceholderRow(ws,row);
}

function removeMachinePlaceholder(ws,machineNo){
  const row=machinePlaceholderRows(ws).find(x=>String(x.machineNo||'').toLowerCase()===String(machineNo||'').toLowerCase())?.row;
  if(!row)return false;
  try{snapshotAndUnmerge(ws)}catch(_){}
  ws.spliceRows(row,1);
  return true;
}

function ensureMachinePlaceholderInWorksheet(ws,m){
  if(!m||!String(m.machineNo||'').trim())return {added:false,reason:'invalid machine'};
  const no=String(m.machineNo).trim();

  const tasks=(String(ws.name||'').toLowerCase()==='full'?taskRowsFromAnnual(ws):taskRowsFromMonthly(ws))
    .filter(t=>String(t.machineNo||'').toLowerCase()===no.toLowerCase());
  if(tasks.length){
    removeMachinePlaceholder(ws,no);
    return {added:false,reason:'machine has tasks'};
  }

  const existing=machinePlaceholderRows(ws).find(x=>String(x.machineNo||'').toLowerCase()===no.toLowerCase());
  if(existing){
    writeMachinePlaceholder(ws,existing.row,m);
    return {added:false,updated:true,row:existing.row};
  }

  const row=Math.max(pmDataStartRow(ws),lastLogicalDataRow(ws)+1);
  writeMachinePlaceholder(ws,row,m);
  return {added:true,row};
}

function ensureAllMasterMachinesInWorksheet(line,ws){
  const x=loadPmMaster(),l=pmLineEntry(x,line);
  const results=[];
  for(const m of (l?.machines||[])){
    if(m.enabled===false)continue;
    results.push({machineNo:m.machineNo,...ensureMachinePlaceholderInWorksheet(ws,m)});
  }
  return results;
}

async function ensureMachineInAnnualExcel(line,year,m){
  const fullPath=annualPathFor(line,year);
  if(!fs.existsSync(fullPath))return {updated:false,reason:'Annual Excel not found'};
  const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(fullPath);normalizeWorkbookFormulas(wb);
  const ws=wb.getWorksheet('Full')||wb.worksheets[0];
  const result=ensureMachinePlaceholderInWorksheet(ws,m);
  formatProfessionalAnnualLayout(ws,pmDataStartRow(ws),Math.max(pmDataStartRow(ws),lastLogicalDataRow(ws)));
  await safeWriteWorkbook(wb,fullPath);
  return {updated:true,...result};
}

async function ensureMachineInMonthlyExcel(line,year,month,m){
  const fullPath=monthlyPathFor(line,year,month);
  if(!fs.existsSync(fullPath))return {updated:false,reason:'Monthly Excel not found'};
  const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(fullPath);normalizeWorkbookFormulas(wb);
  const ws=wb.getWorksheet('Sheet1')||wb.worksheets[0];
  const result=ensureMachinePlaceholderInWorksheet(ws,m);
  rebuildMonthlyTaskFormulas(ws);
  formatProfessionalMonthlyLayout(ws,pmDataStartRow(ws),Math.max(pmDataStartRow(ws),lastLogicalDataRow(ws)));
  await safeWriteWorkbook(wb,fullPath);
  return {updated:true,...result};
}



// ===== STEP 16: Line -> Machine -> Task master flow =====
const PM_MASTER_DIR=process.env.PM_DATA_ROOT || path.join(__dirname,'data');
if(!fs.existsSync(PM_MASTER_DIR))fs.mkdirSync(PM_MASTER_DIR,{recursive:true});
const PM_MASTER_FILE=path.join(PM_MASTER_DIR,'pm_master.json');
function loadPmMaster(){
  let x={lines:[],removedLines:[]}; try{x=JSON.parse(fs.readFileSync(PM_MASTER_FILE,'utf8'));}catch(e){if(e.code!=='ENOENT')throw e;}
  if(!Array.isArray(x.lines))x.lines=[];
  if(!Array.isArray(x.removedLines))x.removedLines=[];
  // Excel-only mode: production lines are created only by an explicit Excel import.
  // Never seed/recreate legacy lines at server startup. This makes Delete Line permanent.
  for(const l of x.lines){
    if(!Array.isArray(l.machines))l.machines=[];
    if(l.enabled===undefined)l.enabled=true;
    for(const m of l.machines){
      if(m.enabled===undefined)m.enabled=true;
      if(m.type===undefined)m.type='';
      if(!(Number(m.expectedWeeklyHours)>0))m.expectedWeeklyHours=120;
    }
  }
  return x;
}

function savePmMaster(x){atomicJson122(PM_MASTER_FILE,x);}

// ===== STEP 118: SYSTEM MASTER -> EXCEL REPORT =====
// Excel is no longer the database for Annual machine/task structure.
// Machines live in pm_master.json. Tasks live in pm_tasks.json.
// Annual Excel is rebuilt from those two masters after structural changes.
const PM_TASKS_FILE=path.join(PM_MASTER_DIR,'pm_tasks.json');
function loadSystemTasks(){
  let x={version:1,tasks:[]};
  try{x=JSON.parse(fs.readFileSync(PM_TASKS_FILE,'utf8'));}catch(e){if(e.code!=='ENOENT')throw e;}
  if(!x||!Array.isArray(x.tasks))throw Error('Invalid system task master');
  return x;
}
function saveSystemTasks(x){
  atomicJson122(PM_TASKS_FILE,x);
}

function archiveRemovedSystemTasks(tasks,reason='line removed'){
  if(!Array.isArray(tasks)||!tasks.length)return 0;
  const file=path.join(PM_MASTER_DIR,'pm_tasks_archive.json');
  let x={version:1,items:[]};
  try{x=JSON.parse(fs.readFileSync(file,'utf8'))}catch(_){}
  if(!x||!Array.isArray(x.items))x={version:1,items:[]};
  const removedAt=new Date().toISOString();
  for(const t of tasks)x.items.push({...t,archivedAt:removedAt,archiveReason:reason});
  fs.writeFileSync(file,JSON.stringify(x,null,2),'utf8');
  return tasks.length;
}

function purgeSystemTasksForLine(line,{year=null,archive=true,reason='line lifecycle reset'}={}){
  const store=loadSystemTasks();
  const key=String(line||'').toLowerCase();
  const removed=[],kept=[];
  for(const t of store.tasks){
    const sameLine=String(t.line||'').toLowerCase()===key;
    const sameYear=(year===null||year===undefined)?true:Number(t.year)===Number(year);
    if(sameLine&&sameYear)removed.push(t);
    else kept.push(t);
  }
  if(archive)archiveRemovedSystemTasks(removed,reason);
  store.tasks=kept;
  store.initialized=store.initialized||{};
  for(const t of removed)store.initialized[key+'|'+Number(t.year)]=true;
  saveSystemTasks(store);
  return removed.length;
}

function purgeRowStateForLine(line){
  const key=String(line||'');
  for(const fileName of ['task_state.json','task_planning_state.json']){
    const file=path.join(PM_MASTER_DIR,fileName);
    let x={};try{x=JSON.parse(fs.readFileSync(file,'utf8'))}catch(_){continue}
    let changed=false;
    for(const k of Object.keys(x)){
      const p=String(k).split('|');
      if(String(p[0]||'').toLowerCase()===key.toLowerCase()){
        delete x[k];changed=true;
      }
    }
    if(changed)fs.writeFileSync(file,JSON.stringify(x,null,2),'utf8');
  }
}

function freshLineHasNoStructuralSource(name,year){
  const annualFile=annualPathFor(name,year);
  return !fs.existsSync(annualFile);
}

function systemTasksFor(line,year){
  return loadSystemTasks().tasks.filter(t=>String(t.line||'').toLowerCase()===String(line||'').toLowerCase()&&Number(t.year)===Number(year)&&t.active!==false);
}
function scheduleWeeksFromAnnualRow(ws,row){
  const out=[];
  for(let c=16;c<=67;c++)if([1,2,3,4,5].includes(Number(ws.getCell(row,c).value)))out.push(c-15);
  return out;
}
async function ensureSystemTasksBootstrapped(line,year){
  const store=loadSystemTasks();
  const existing=store.tasks.filter(t=>String(t.line||'').toLowerCase()===String(line).toLowerCase()&&Number(t.year)===Number(year));
  const scope=String(line).toLowerCase()+'|'+Number(year);
  store.initialized=store.initialized||{};
  if(existing.length||store.initialized[scope]||pmLineEntry(loadPmMaster(),line)?.fresh122){store.initialized[scope]=true;saveSystemTasks(store);return store;}
  store.initialized[scope]=true;

  const fullPath=annualPathFor(line,year);
  if(!fs.existsSync(fullPath)){saveSystemTasks(store);return store;}
  try{
    const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(fullPath);
    const ws=wb.getWorksheet('Full')||wb.worksheets[0];
    backfillStableTaskLinks(ws);
    const mx=loadPmMaster(), ml=pmLineEntry(mx,line);
    for(const t of [...taskRowsFromAnnual(ws),...machinePlaceholderRows(ws)]){if(ml&&t.machineNo&&!pmMachineEntry(ml,t.machineNo))ml.machines.push({section:t.section||'',machineNo:t.machineNo,machineName:t.machineName||'',machineIdentified:t.machineIdentified||'',type:t.type||'',enabled:true});}
    if(ml)savePmMaster(mx);
    for(const t of taskRowsFromAnnual(ws)){
      let taskId=stableTaskId(ws,t.row)||taskLinkId();
      if(store.tasks.some(v=>v.taskId===taskId))taskId=taskLinkId();
      writeStableTaskLink(ws,t.row,t.row+1,t.machineNo,taskId);
      const meta=loadTaskPlanningMeta(line,year,t.row)||{};
      store.tasks.push({
        taskId,line,year,machineNo:t.machineNo,
        part:t.part||'',maintenance:t.maintenance||'',frequency:t.frequency||'',
        code:Number(t.dominantCode)||1,effectiveCode:Number(t.dominantCode)||1,
        calculationMethod:meta.calculationMethod||String(safeText(ws,t.row,12)||'Calendar'),
        calendarFrequency:meta.calendarFrequency||String(t.frequency||'weekly'),
        intervalHours:Number(meta.intervalHours||0),
        expectedWeeklyHours:Number(meta.expectedWeeklyHours||0),
        firstPmDate:meta.firstPmDate||String(safeText(ws,t.row,14)||''),
        scheduleWeeks:scheduleWeeksFromAnnualRow(ws,t.row),
        scheduleCodes:Object.fromEntries(scheduleWeeksFromAnnualRow(ws,t.row).map(w=>[w,Number(ws.getCell(t.row,15+w).value)])),
        notes:String(safeText(ws,t.row,15)||''),
        active:true,isActive:!readLegacyActive122(line,year,t.row),createdAt:new Date().toISOString()
      });
    }
    saveSystemTasks(store);
    await safeWriteWorkbook(wb,fullPath);
  }catch(e){throw new Error('Task master migration failed: '+e.message)}
  return store;
}

// ===== V1.3.4: MULTI-YEAR PM PLANNER + CONTINUOUS RUNNING HOURS =====
const YEAR_PLANNER_FILE=path.join(PM_MASTER_DIR,'year_planner.json');
function loadYearPlannerSettings(){
  let x={autoRollover:true,sourceStrategy:'latest_prior'};
  try{x={...x,...JSON.parse(fs.readFileSync(YEAR_PLANNER_FILE,'utf8'))}}catch(e){if(e.code!=='ENOENT')throw e;}
  x.autoRollover=x.autoRollover!==false;
  return x;
}
function saveYearPlannerSettings(x){atomicJson122(YEAR_PLANNER_FILE,{...loadYearPlannerSettings(),...(x||{}),updatedAt:new Date().toISOString()});}
function plannerYearsForLine(line){
  const years=new Set();
  for(const t of loadSystemTasks().tasks){
    if(same122(t.line,line)&&Number.isInteger(Number(t.year)))years.add(Number(t.year));
  }
  for(const f of annualFilesForLine(line))years.add(Number(f.year));
  return [...years].filter(y=>y>=2000&&y<=2200).sort((a,b)=>a-b);
}
function normalizePlannerFrequency(t){
  // Running-hours tasks are NOT calendar-weekly tasks.  Their PM code/color
  // must never be used to invent a calendar frequency such as "weekly".
  const method=normalizePlannerCalculationMethod(t);
  if(method==='hours'||method==='plc_hours')return '';
  const raw=String(t.calendarFrequency||t.frequency||'').trim().toLowerCase();
  if(raw.includes('week'))return 'weekly';
  if(raw.includes('month')&&!raw.includes('quarter')&&!raw.includes('semi'))return 'monthly';
  if(raw.includes('quarter')||raw.includes('3 month'))return 'quarterly';
  if(raw.includes('semi')||raw.includes('half')||raw.includes('6 month'))return 'semiannual';
  if(raw.includes('annual')||raw.includes('year'))return 'annual';
  const code=Number(t.effectiveCode||t.code||0);
  return code===1?'weekly':code===2?'monthly':code===3?'quarterly':code===4?'semiannual':'annual';
}
function clampDateToYear(dateText,targetYear){
  const m=String(dateText||'').match(/^\d{4}-(\d{2})-(\d{2})$/);
  if(!m)return '';
  const month=Math.min(12,Math.max(1,Number(m[1]))),day=Math.max(1,Number(m[2]));
  const maxDay=new Date(targetYear,month,0).getDate();
  return `${targetYear}-${String(month).padStart(2,'0')}-${String(Math.min(day,maxDay)).padStart(2,'0')}`;
}
function plannerAnchorFromSource(t,sourceYear,targetYear){
  const shifted=clampDateToYear(t.firstPmDate,targetYear);
  if(shifted)return shifted;
  const srcFridays=annualFridays(sourceYear);
  const firstWeek=[...(t.scheduleWeeks||[])].map(Number).filter(w=>w>=1&&w<=srcFridays.length).sort((a,b)=>a-b)[0];
  if(firstWeek){
    const d=srcFridays[firstWeek-1];
    const maxDay=new Date(targetYear,d.getMonth()+1,0).getDate();
    return `${targetYear}-${String(d.getMonth()+1).padStart(2,'0')}-${String(Math.min(d.getDate(),maxDay)).padStart(2,'0')}`;
  }
  return `${targetYear}-01-01`;
}
function mapSourceWeeksToTarget(t,sourceYear,targetYear){
  const src=annualFridays(sourceYear),dst=annualFridays(targetYear),out=[];
  for(const w of (t.scheduleWeeks||[])){
    const d=src[Number(w)-1]; if(!d)continue;
    const td=dateOnly(targetYear,d.getMonth()+1,Math.min(d.getDate(),new Date(targetYear,d.getMonth()+1,0).getDate()));
    const tw=nearestAnnualWeekIndex(dst,td); if(tw&&!out.includes(tw))out.push(tw);
  }
  return out.sort((a,b)=>a-b);
}
function normalizePlannerCalculationMethod(t){
  const raw=String(t.calculationMethod||'').trim().toLowerCase().replace(/[_-]+/g,' ').replace(/\s+/g,' ');
  const freq=String(t.frequency||'').trim().toLowerCase();
  const interval=Number(t.intervalHours||0);
  if(raw==='plc hours'||raw==='plc hour'||raw.includes('plc'))return 'plc_hours';
  if(raw==='hours'||raw==='hour'||raw==='running hours'||raw==='running hour'||raw.includes('running'))return 'hours';
  // Legacy/imported Annual files can contain "10000 hr" while the method field
  // is blank or was accidentally stored as Calendar.  The hour expression is
  // authoritative and must not be converted to weekly from the PM color code.
  if(/(^|\s)\d+(?:\.\d+)?\s*(?:hr|hrs|hour|hours)\b/.test(freq) || (interval>0 && /hr|hour/.test(freq)))return 'hours';
  if(!raw||raw==='calendar'||raw==='calendar based'||raw==='date'||raw==='date based')return 'calendar';
  // Older imported Excel files sometimes stored the frequency text in this field.
  if(/week|month|quarter|semi|annual|year/.test(raw))return 'calendar';
  return 'calendar';
}
function plannerFrequencyLabel(t){
  const method=normalizePlannerCalculationMethod(t);
  if(method==='hours'||method==='plc_hours'){
    const n=Number(t.intervalHours||String(t.frequency||'').match(/\d+(?:\.\d+)?/)?.[0]||0);
    return n>0?`${n} hr`:(String(t.frequency||'Running Hours').trim()||'Running Hours');
  }
  return normalizePlannerFrequency(t)||String(t.frequency||'');
}
function parseIsoDateLocal(text){
  const m=String(text||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(!m)return null;
  const y=Number(m[1]),mo=Number(m[2]),d=Number(m[3]);
  const x=dateOnly(y,mo,d);
  if(x.getFullYear()!==y||x.getMonth()!==mo-1||x.getDate()!==d)return null;
  return x;
}
function fmtIsoDateLocal(d){
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function pmFridayOnOrAfter(date){
  const d=new Date(date.getFullYear(),date.getMonth(),date.getDate(),12);
  d.setDate(d.getDate()+((5-d.getDay()+7)%7));
  return d;
}
function runningHoursStartForTask(t,sourceYear){
  // Imported Annual files begin their running-hours counter on 1 January of the
  // imported year unless a later real execution/reset date has been recorded.
  const explicit=String(t.runningHoursCounterStartDate||t.lastExecutionDate||'').trim();
  if(parseIsoDateLocal(explicit))return explicit;
  return `${Number(sourceYear)}-01-01`;
}
function generateRunningHoursScheduleForYear(targetYear,counterStartDate,intervalHours,expectedWeeklyHours){
  const start=parseIsoDateLocal(counterStartDate);
  intervalHours=Number(intervalHours||0);expectedWeeklyHours=Number(expectedWeeklyHours||0);
  if(!start)throw new Error('Invalid Running Hours counter start date');
  if(!(intervalHours>0))throw new Error('Enter Interval Hours');
  if(!(expectedWeeklyHours>0))throw new Error('Enter Expected Running Hours per Week');

  // We never schedule before the interval is actually reached.  If an interval
  // is 83.3 weeks, the PM becomes due in week 84, not at the beginning of week 83.
  const intervalWeeks=Math.max(1,Math.ceil(intervalHours/expectedWeeklyHours));
  const stepDays=intervalWeeks*7;
  let due=new Date(start.getFullYear(),start.getMonth(),start.getDate(),12);
  due.setDate(due.getDate()+stepDays);

  // Fast-forward to the first due point that can affect targetYear.
  const targetStart=dateOnly(Number(targetYear),1,1);
  while(due.getTime()<targetStart.getTime())due.setDate(due.getDate()+stepDays);

  const targetEnd=dateOnly(Number(targetYear)+1,1,1);
  const fridays=annualFridays(Number(targetYear));
  const weeks=[],dueDates=[];
  while(due.getTime()<targetEnd.getTime()){
    const pmFriday=pmFridayOnOrAfter(due);
    if(pmFriday.getFullYear()===Number(targetYear)){
      const wi=fridays.findIndex(x=>x.getTime()===pmFriday.getTime());
      if(wi>=0 && !weeks.includes(wi+1)){
        weeks.push(wi+1);
        dueDates.push(fmtIsoDateLocal(due));
      }
    }
    due.setDate(due.getDate()+stepDays);
  }
  return {
    weeks, dueDates, intervalWeeks,
    firstDueDate:dueDates[0]||'',
    derivedCode:intervalWeeks<=1?1:intervalWeeks<=6?2:intervalWeeks<=17?3:intervalWeeks<=35?4:5,
    description:`${intervalHours} hr / ~${intervalWeeks} week(s) at ${expectedWeeklyHours} hr/week`
  };
}

function projectTaskToYear(t,sourceYear,targetYear,{newIds=false}={}){
  const calculationMethod=normalizePlannerCalculationMethod(t);
  const calendarFrequency=normalizePlannerFrequency(t);
  const intervalHours=Number(t.intervalHours||0),expectedWeeklyHours=effectiveTaskExpectedWeeklyHours(t);
  let schedule=null,warning='',firstPmDate='',counterStartDate='';

  if(calculationMethod==='hours'||calculationMethod==='plc_hours'){
    counterStartDate=runningHoursStartForTask(t,sourceYear);
    try{
      schedule=generateRunningHoursScheduleForYear(targetYear,counterStartDate,intervalHours,expectedWeeklyHours);
      // For Running Hours, First PM Date means the first CALCULATED due date in
      // the target year. It is never forced to 1 January / Week 1.
      firstPmDate=schedule.firstDueDate||'';
    }catch(e){warning=e.message;}
  }else{
    firstPmDate=plannerAnchorFromSource(t,sourceYear,targetYear);
    try{
      schedule=generateAnnualSchedule(targetYear,firstPmDate,{
        calculationMethod,calendarFrequency,intervalHours,expectedWeeklyHours
      });
    }catch(e){warning=e.message;}
  }

  // Calendar tasks may preserve a legacy approved source pattern as a fallback.
  // Running-hours tasks MUST NOT map old Week Codes because the due week is
  // calculated from counter start + runtime interval.
  let weeks=(schedule&&Array.isArray(schedule.weeks))?schedule.weeks:[];
  if(!weeks.length && calculationMethod!=='hours' && calculationMethod!=='plc_hours'){
    weeks=mapSourceWeeksToTarget(t,sourceYear,targetYear);
  }
  // Zero occurrences in a target year is valid for a long running-hours interval.
  if(!weeks.length && calculationMethod!=='hours' && calculationMethod!=='plc_hours')warning=warning||'No target weeks could be generated';

  const oldWeeks=[...(t.scheduleWeeks||[])].map(Number).filter(Boolean).sort((a,b)=>a-b);
  const scheduleCodes={};
  weeks.forEach((w,i)=>{
    const sourceWeek=oldWeeks[Math.min(i,Math.max(0,oldWeeks.length-1))];
    scheduleCodes[w]=Number(t.scheduleCodes?.[sourceWeek]||schedule?.derivedCode||t.effectiveCode||t.code||1);
  });
  return {
    ...t,
    taskId:newIds?taskLinkId():t.taskId,
    line:t.line,year:targetYear,firstPmDate,
    runningHoursCounterStartDate:counterStartDate||t.runningHoursCounterStartDate||'',
    calculationMethod,calendarFrequency,
    effectiveCode:Number(schedule?.derivedCode||t.effectiveCode||t.code||1),
    scheduleWeeks:weeks,scheduleCodes,
    plannerIntervalWeeks:Number(schedule?.intervalWeeks||0),
    plannerDueDates:Array.isArray(schedule?.dueDates)?schedule.dueDates:[],
    active:t.active!==false,isActive:t.isActive!==false,
    sourceTaskId:t.taskId,rolledFromYear:sourceYear,
    createdAt:newIds?new Date().toISOString():t.createdAt,
    updatedAt:new Date().toISOString(),
    plannerWarning:warning||undefined
  };
}
async function previewYearPlan(line,sourceYear,targetYear){
  if(!pmLineEntry(loadPmMaster(),line))throw new Error('Line not found');
  sourceYear=Number(sourceYear);targetYear=Number(targetYear);
  if(!sourceYear||!targetYear||sourceYear===targetYear)throw new Error('Select different source and target years');
  await ensureSystemTasksBootstrapped(line,sourceYear);
  const src=systemTasksFor(line,sourceYear);
  if(!src.length)throw new Error(`No active PM tasks found for ${line} ${sourceYear}`);
  const projected=src.map(t=>projectTaskToYear(t,sourceYear,targetYear));
  const warnings=projected.filter(t=>t.plannerWarning).map(t=>({machineNo:t.machineNo,maintenance:t.maintenance,warning:t.plannerWarning}));

  // Machine count and preview must come from Machine Management, not from tasks.
  // This keeps machines with 0 PM tasks visible and prevents duplicate machine counts.
  const masterMachines=orderedMasterMachines(line);
  const previewTask=t=>({
    machineNo:t.machineNo,
    maintenance:t.maintenance,
    frequency:plannerFrequencyLabel(t),
    firstPmDate:t.firstPmDate||'—',
    occurrences:(t.scheduleWeeks||[]).length,
    intervalWeeks:Number(t.plannerIntervalWeeks||0),
    pmCode:Number(t.effectiveCode||t.code||0),
    warning:t.plannerWarning||''
  });
  const machinePreview=masterMachines.map(m=>({
    machineNo:String(m.machineNo||'').trim(),
    machineName:String(m.machineName||'').trim(),
    machineIdentified:String(m.machineIdentified||'').trim(),
    tasks:projected.filter(t=>same122(t.machineNo,m.machineNo)).map(previewTask)
  }));

  return {
    line,sourceYear,targetYear,machines:masterMachines.length,
    tasks:projected.length,plannedOccurrences:projected.reduce((n,t)=>n+(t.scheduleWeeks||[]).length,0),warnings,
    existingTargetTasks:systemTasksFor(line,targetYear).length,targetAnnualExists:fs.existsSync(annualPathFor(line,targetYear)),targetMonthlyFiles:monthlyFilesForLine(line).filter(x=>Number(x.year)===targetYear).length,
    machinePreview,
    // Keep the flat sample for backwards compatibility with older clients.
    sample:[...projected].sort((a,b)=>Number(Boolean(b.plannerWarning))-Number(Boolean(a.plannerWarning))).map(previewTask)
  };
}
async function buildYearPlan(line,sourceYear,targetYear,{overwrite=false,user=null}={}){
  const preview=await previewYearPlan(line,sourceYear,targetYear);
  if(preview.warnings?.length)throw new Error(`Cannot build target year while ${preview.warnings.length} planning warning(s) remain. Fix warnings and Preview again.`);
  if((preview.existingTargetTasks||preview.targetAnnualExists)&&!overwrite)throw new Error(`Target year ${targetYear} already exists. Enable overwrite to rebuild it.`);
  await ensureSystemTasksBootstrapped(line,sourceYear);
  const source=systemTasksFor(line,sourceYear);
  const clones=source.map(t=>projectTaskToYear(t,sourceYear,targetYear,{newIds:true}));
  const store=loadSystemTasks();
  const removed=store.tasks.filter(t=>same122(t.line,line)&&Number(t.year)===targetYear);
  if(removed.length)archiveRemovedSystemTasks(removed,`Year Planner overwrite ${line} ${targetYear}`);
  store.tasks=store.tasks.filter(t=>!(same122(t.line,line)&&Number(t.year)===targetYear));
  store.tasks.push(...clones.map(({plannerWarning,...t})=>t));
  store.initialized=store.initialized||{};store.initialized[String(line).toLowerCase()+'|'+targetYear]=true;
  saveSystemTasks(store);
  const rendered=await rebuildAnnualFromSystemMaster(line,targetYear);
  const rebuiltMonthly=[];
  if(overwrite){
    for(const mf of monthlyFilesForLine(line).filter(x=>Number(x.year)===targetYear)){
      try{const r=await createMonthlyForLine(line,targetYear,Number(mf.month),{overwrite:true,preserveExecution:true});rebuiltMonthly.push({month:Number(mf.month),ok:true,file:r.file});}
      catch(e){rebuiltMonthly.push({month:Number(mf.month),ok:false,error:e.message});}
    }
  }
  // Post-build verification: compare what was persisted with the Preview plan
  // and confirm the generated Annual Excel file exists before reporting success.
  const persisted=systemTasksFor(line,targetYear);
  const persistedOccurrences=persisted.reduce((n,t)=>n+(Array.isArray(t.scheduleWeeks)?t.scheduleWeeks.length:0),0);
  const annualFile=annualPathFor(line,targetYear);
  const verification={
    tasksMatch:persisted.length===preview.tasks,
    occurrencesMatch:persistedOccurrences===preview.plannedOccurrences,
    annualFileExists:fs.existsSync(annualFile),
    persistedTasks:persisted.length,
    persistedOccurrences,
    expectedTasks:preview.tasks,
    expectedOccurrences:preview.plannedOccurrences
  };
  verification.ok=verification.tasksMatch&&verification.occurrencesMatch&&verification.annualFileExists;
  if(!verification.ok)throw new Error(`Year built but verification failed: tasks ${verification.persistedTasks}/${verification.expectedTasks}, occurrences ${verification.persistedOccurrences}/${verification.expectedOccurrences}, Annual file ${verification.annualFileExists?'OK':'missing'}`);
  const result={...preview,overwrittenTasks:removed.length,file:rendered.file,built:true,rebuiltMonthly,verification};
  if(user)audit('YEAR_PLANNER_BUILD',user,result);
  return result;
}
async function autoBuildCurrentYearPlans(){
  const cfg=loadYearPlannerSettings(); if(!cfg.autoRollover)return [];
  const currentYear=new Date().getFullYear(),results=[];
  for(const l of loadPmMaster().lines.filter(x=>x.enabled!==false)){
    const line=l.name,existing=systemTasksFor(line,currentYear).length||fs.existsSync(annualPathFor(line,currentYear));
    if(existing)continue;
    const sourceYears=plannerYearsForLine(line).filter(y=>y<currentYear);
    const sourceYear=sourceYears[sourceYears.length-1]; if(!sourceYear)continue;
    try{results.push({ok:true,line,sourceYear,targetYear:currentYear,...await buildYearPlan(line,sourceYear,currentYear,{overwrite:false})});}
    catch(e){results.push({ok:false,line,sourceYear,targetYear:currentYear,error:e.message});}
  }
  return results;
}

app.get('/api/year-planner',requireAuth,requireAdmin,(req,res)=>{
  const master=loadPmMaster();
  res.json({settings:loadYearPlannerSettings(),currentYear:new Date().getFullYear(),lines:master.lines.filter(l=>l.enabled!==false).map(l=>({name:l.name,years:plannerYearsForLine(l.name)}))});
});
app.put('/api/year-planner/settings',requireAuth,requireAdmin,(req,res)=>{
  saveYearPlannerSettings({autoRollover:req.body?.autoRollover!==false});
  audit('YEAR_PLANNER_SETTINGS',req.user,{autoRollover:req.body?.autoRollover!==false});
  res.json({ok:true,settings:loadYearPlannerSettings()});
});
app.post('/api/year-planner/preview',requireAuth,requireAdmin,async(req,res)=>{
  try{res.json(await previewYearPlan(cleanName(req.body?.line),Number(req.body?.sourceYear),Number(req.body?.targetYear)));}
  catch(e){res.status(400).json({error:e.message});}
});
app.post('/api/year-planner/build',requireAuth,requireAdmin,async(req,res)=>{
  try{res.json(await buildYearPlan(cleanName(req.body?.line),Number(req.body?.sourceYear),Number(req.body?.targetYear),{overwrite:!!req.body?.overwrite,user:req.user}));}
  catch(e){res.status(400).json({error:e.message});}
});
app.post('/api/year-planner/auto-run',requireAuth,requireAdmin,async(req,res)=>{
  try{const results=await autoBuildCurrentYearPlans();audit('YEAR_PLANNER_AUTO_RUN',req.user,{results});res.json({ok:true,results});}
  catch(e){res.status(500).json({error:e.message});}
});

function orderedMasterMachines(line){
  const x=loadPmMaster(),l=pmLineEntry(x,line);
  return (l?.machines||[]).filter(m=>m.enabled!==false);
}
async function rebuildAnnualFromSystemMaster(line,year){
  await ensureSystemTasksBootstrapped(line,year);

  const tasks=systemTasksFor(line,year);
  const machines=orderedMasterMachines(line);
  const built=await buildCanonicalAnnualTemplate(line,year);
  const wb=built.wb,ws=built.ws;
  const startRow=pmDataStartRow(ws);

  // Deterministic renderer: Excel is output only, never the data source.
  const oldDataMerges=[];
  if(ws._merges){
    for(const key of Object.keys(ws._merges)){
      const m=ws._merges[key]?.model;
      if(!m)continue;
      if(m.bottom>=startRow && m.left<=7 && m.right>=1)oldDataMerges.push(key);
    }
  }
  for(const rg of oldDataMerges)try{ws.unMergeCells(rg)}catch(_){}

  const clearLast=Math.max(ws.rowCount,startRow+500);
  for(let r=startRow;r<=clearLast;r++){
    for(let c=1;c<=69;c++)ws.getCell(r,c).value=null;
    ws.getRow(r).height=18;
  }

  let r=startRow;
  const machineBlocks=[];

  for(const m of machines){
    const machineTasks=tasks.filter(t=>
      String(t.machineNo||'').toLowerCase()===String(m.machineNo||'').toLowerCase()
    );

    const machineStart=r;

    if(!machineTasks.length){
      writeMachinePlaceholder(ws,r,m);
      machineBlocks.push({
        start:r,end:r,
        section:String(m.section||'').trim(),
        machineNo:String(m.machineNo||'').trim(),
        machineName:String(m.machineName||'').trim(),
        machineIdentified:String(m.machineIdentified||'').trim(),
        type:String(m.type||'').trim(),
        taskCount:0
      });
      r+=1;
      continue;
    }

    for(const t of machineTasks){
      const planRow=r,doneRow=r+1;

      for(const rr of [planRow,doneRow]){
        ws.getCell(rr,1).value=String(m.section||'').trim()||null;
        ws.getCell(rr,2).value=String(m.machineNo||'').trim()||null;
        ws.getCell(rr,3).value=String(m.machineName||'').trim()||null;
        ws.getCell(rr,4).value=String(m.machineIdentified||'').trim()||null;
        ws.getCell(rr,5).value=String(m.type||'').trim()||null;
      }

      ws.getCell(planRow,6).value=String(t.part||'').trim()||null;
      ws.getCell(planRow,7).value=String(t.maintenance||'').trim()||null;
      ws.getCell(planRow,8).value='planed';

      ws.getCell(doneRow,6).value=null;
      ws.getCell(doneRow,7).value=null;
      ws.getCell(doneRow,8).value='done';

      applyProfessionalTaskMeta(ws,planRow,doneRow,{
        frequency:t.frequency||'',
        calculationMethod:t.calculationMethod||'Calendar',
        calendarFrequency:t.calendarFrequency||t.frequency||'',
        intervalHours:Number(t.intervalHours||0),
        expectedWeeklyHours:effectiveTaskExpectedWeeklyHours(t),
        firstPmDate:t.firstPmDate||'',
        notes:t.notes||''
      },Number(t.effectiveCode||t.code)||1);

      for(let c=16;c<=67;c++){
        ws.getCell(planRow,c).value=null;
        ws.getCell(doneRow,c).value=null;
        clearCellFill(ws.getCell(planRow,c));
        clearCellFill(ws.getCell(doneRow,c));
      }

      for(const w of (Array.isArray(t.scheduleWeeks)?t.scheduleWeeks:[])){
        const c=15+Number(w);
        if(c>=16&&c<=67){
          const code=Number(t.scheduleCodes?.[w]||t.effectiveCode||t.code)||1;
          ws.getCell(planRow,c).value=code;
          setPmCodeFill(ws.getCell(planRow,c),code);
        }
      }

      writeStableTaskLink(ws,planRow,doneRow,m.machineNo,t.taskId);

      try{ws.mergeCells(planRow,6,doneRow,6)}catch(_){}
      try{ws.mergeCells(planRow,7,doneRow,7)}catch(_){}

      r+=2;
    }

    machineBlocks.push({
      start:machineStart,end:r-1,
      section:String(m.section||'').trim(),
      machineNo:String(m.machineNo||'').trim(),
      machineName:String(m.machineName||'').trim(),
      machineIdentified:String(m.machineIdentified||'').trim(),
      type:String(m.type||'').trim(),
      taskCount:machineTasks.length
    });
  }

  const last=Math.max(startRow,r-1);

  // Build machine merges directly from master blocks.
  for(const b of machineBlocks){
    if(b.taskCount<=0)continue;

    for(let c=2;c<=5;c++){
      const value=c===2?b.machineNo:c===3?b.machineName:c===4?b.machineIdentified:b.type;
      if(b.end>b.start){
        try{ws.mergeCells(b.start,c,b.end,c)}catch(_){}
      }
      ws.getCell(b.start,c).value=value||null;
    }
  }

  // Merge Section only across contiguous machines with same Section.
  let i=0;
  while(i<machineBlocks.length){
    let j=i;
    while(j+1<machineBlocks.length &&
          String(machineBlocks[j+1].section||'')===String(machineBlocks[i].section||'')){
      j++;
    }
    const top=machineBlocks[i].start,bottom=machineBlocks[j].end;
    const section=machineBlocks[i].section;
    if(section){
      if(bottom>top){
        try{ws.mergeCells(top,1,bottom,1)}catch(_){}
      }
      ws.getCell(top,1).value=section;
    }
    i=j+1;
  }

  for(const b of machineBlocks){
    if(b.taskCount!==0)continue;
    ws.getCell(b.start,1).value=b.section||null;
    ws.getCell(b.start,2).value=b.machineNo||null;
    ws.getCell(b.start,3).value=b.machineName||null;
    ws.getCell(b.start,4).value=b.machineIdentified||null;
    ws.getCell(b.start,5).value=b.type||null;
    ws.getCell(b.start,8).value='machine';
    writeStableMachineOnlyLink(ws,b.start,b.machineNo);
  }

  formatProfessionalAnnualLayout(ws,startRow,last);
  applyAnnualProfessionalKpis(ws);
  ws.getCell(3,1).value=machines.length;ws.getCell(3,2).value=machineBlocks.reduce((n,b)=>n+b.taskCount,0);
  ws.getColumn(68).hidden=true;
  ws.getColumn(69).hidden=true;
  forceExcelRecalculation(wb);

  // Hard validation before save.
  for(const b of machineBlocks){
    const vals=[
      String(safeText(ws,b.start,1)||'').trim(),
      String(safeText(ws,b.start,2)||'').trim(),
      String(safeText(ws,b.start,3)||'').trim(),
      String(safeText(ws,b.start,4)||'').trim(),
      String(safeText(ws,b.start,5)||'').trim()
    ];
    const expected=[b.section,b.machineNo,b.machineName,b.machineIdentified,b.type].map(x=>String(x||'').trim());
    if(vals.join('|')!==expected.join('|')){
      throw new Error(`Annual render validation failed for machine ${b.machineNo}: expected ${expected.join(' | ')} but found ${vals.join(' | ')}`);
    }
  }

  const fullPath=annualPathFor(line,year);
  fs.mkdirSync(path.dirname(fullPath),{recursive:true});
  await safeWriteWorkbook(wb,fullPath);

  return {
    file:path.basename(fullPath),
    machines:machines.length,
    tasks:tasks.length,
    machineBlocks:machineBlocks.length,
    lastRow:last,
    mode:'DETERMINISTIC_SYSTEM_MASTER_RENDER'
  };
}

function cleanName(v){return String(v||'').trim().replace(/[\\/:*?"<>|]/g,'-');}
function annualPathFor(line,year){return path.join(ANNUAL_DIR,`${line} PM ${year}.xlsx`);}
function monthlyPathFor(line,year,month){
  const exact=findExistingMonthlyFile(line,year,month);
  if(fs.existsSync(exact))return exact;
  for(const alias of pmLineAliases(line)){
    const p=findExistingMonthlyFile(alias,year,month);
    if(fs.existsSync(p))return p;
  }
  return exact;
}


function readLegacyActive122(line,year,row){try{return JSON.parse(fs.readFileSync(path.join(PM_MASTER_DIR,'task_state.json'),'utf8'))[line+'|'+year+'|'+row]?.deactivated;}catch(e){if(e.code!=='ENOENT')throw e;return false;}}
// Step122: one system source for structure, month overrides and execution.
function same122(a,b){return String(a||'').toLowerCase()===String(b||'').toLowerCase();}
function atomicJson122(file,value){const tmp=file+'.tmp';fs.writeFileSync(tmp,JSON.stringify(value,null,2));fs.renameSync(tmp,file);}
function monthlyStore122(){const f=path.join(PM_MASTER_DIR,'monthly_state.json');if(!fs.existsSync(f))return {periods:{}};return JSON.parse(fs.readFileSync(f,'utf8'));}
function periodKey122(line,year,month){if(!pmLineEntry(loadPmMaster(),line))throw Error('Line not found');if(!Number.isInteger(year)||year<2000||year>2200||!Number.isInteger(month)||month<1||month>12)throw Error('Invalid PM period');return line.toLowerCase()+'|'+year+'|'+month;}
function saveMonthlyState122(line,year,month,p){const x=monthlyStore122();x.periods[periodKey122(line,year,month)]=p;atomicJson122(path.join(PM_MASTER_DIR,'monthly_state.json'),x);}
function resetMonthlyState122(line){const x=monthlyStore122();for(const k of Object.keys(x.periods))if(k.startsWith(line.toLowerCase()+'|'))delete x.periods[k];atomicJson122(path.join(PM_MASTER_DIR,'monthly_state.json'),x);}
function renameMonthlyState122(oldLine,newLine){const x=monthlyStore122();for(const k of Object.keys(x.periods))if(k.startsWith(oldLine.toLowerCase()+'|')){x.periods[newLine.toLowerCase()+k.slice(oldLine.length)]=x.periods[k];delete x.periods[k];}atomicJson122(path.join(PM_MASTER_DIR,'monthly_state.json'),x);}
function renameMonthlyMachine122(line,oldNo,newNo){const x=monthlyStore122();for(const [k,p] of Object.entries(x.periods))if(k.startsWith(line.toLowerCase()+'|'))for(const t of [...p.extras,...Object.values(p.overrides),...(p.renderedTasks||[])])if(same122(t.machineNo,oldNo))t.machineNo=newNo;atomicJson122(path.join(PM_MASTER_DIR,'monthly_state.json'),x);}
function execution122(ws){const x={};for(const t of taskRowsFromMonthly(ws)){const id=stableTaskId(ws,t.row);if(!id)throw Error('Missing stable task ID: migrate Monthly first');x[id]={done:[16,17,18,19,20].map(c=>Number(ws.getCell(t.row+1,c).value)||0),user:[21,22,23,24,25].map(c=>safeText(ws,t.row,c)||''),date:[21,22,23,24,25].map(c=>{const v=ws.getCell(t.row+1,c).value;return v instanceof Date?v.toISOString():v||'';})};}return x;}
async function monthlyState122(line,year,month){const key=periodKey122(line,year,month),all=monthlyStore122();if(all.periods[key])return all.periods[key];const p={extras:[],overrides:{},excluded:[],execution:{}};const file=findExistingMonthlyFile(line,year,month);const master=pmLineEntry(loadPmMaster(),line);if(file&&fs.existsSync(file)&&!master.fresh122){const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(file);const ws=wb.getWorksheet('Sheet1')||wb.worksheets[0],annual=systemTasksFor(line,year);for(const t of taskRowsFromMonthly(ws)){let id=stableTaskId(ws,t.row);if(!id){const matches=annual.filter(v=>same122(v.machineNo,t.machineNo)&&v.part===t.part&&v.maintenance===t.maintenance);if(matches.length>1)throw Error('Ambiguous legacy Task IDs; explicit migration required');id=matches[0]?.taskId||taskLinkId();}writeStableTaskLink(ws,t.row,t.row+1,t.machineNo,id);const item={...t,taskId:id,code:t.dominantCode,effectiveCode:t.dominantCode,scheduleWeeks:t.weekIndexes.map(w=>annualColumnForFriday(fridaysInMonth(year,month)[w-1])-15)};if(annual.some(v=>v.taskId===id))p.overrides[id]=item;else p.extras.push(item);}p.execution=execution122(ws);p.renderedTasks=monthlyTasks122(line,year,month,p);await safeWriteWorkbook(wb,file);}saveMonthlyState122(line,year,month,p);return p;}
async function saveExecution122(line,year,month,ws){const p=await monthlyState122(line,year,month);p.execution={...p.execution,...execution122(ws)};saveMonthlyState122(line,year,month,p);}
function monthlyTasks122(line,year,month,p){return [...systemTasksFor(line,year).map(t=>({...t,...p.overrides[t.taskId]})),...p.extras].filter(t=>!p.excluded.includes(t.taskId));}
function rows122(line,tasks,machines=orderedMasterMachines(line)){let row=7;const out=[];for(const m of machines){const mt=tasks.filter(t=>same122(t.machineNo,m.machineNo));if(!mt.length){row++;continue;}for(const t of mt){out.push({...t,...m,row,dominantCode:t.effectiveCode||t.code||1,plannedWeeks:(t.scheduleWeeks||[]).length,isActive:t.isActive!==false,planningMeta:t});row+=2;}}return out;}
function annualRows122(line,year){return rows122(line,systemTasksFor(line,year));}
function monthlyRows122(line,year,month,p){const fridays=fridaysInMonth(year,month);return rows122(line,p.renderedTasks||monthlyTasks122(line,year,month,p),p.machines||orderedMasterMachines(line)).map(t=>{const weekIndexes=fridays.map((f,i)=>(t.scheduleWeeks||[]).includes(annualColumnForFriday(f)-15)?i+1:null).filter(Boolean);return {...t,weekIndexes,plannedWeeks:weekIndexes.length};});}
function resolve122(rows,row,id){const t=rows.find(t=>t.row===Number(row));if(!t||id&&t.taskId!==id){const e=Error('Task changed; refresh the page and retry');e.status=409;throw e;}return t;}
function resolveAnnual122(line,year,row,id){return resolve122(annualRows122(line,year),row,id);}
function resolveMonthly122(line,year,month,p,row,id){return resolve122(monthlyRows122(line,year,month,p),row,id);}
function validateTask122(line,b,before){if(![1,2,3,4,5].includes(Number(b.code)))throw Error('PM code must be 1 to 5');if(!authoritativeMachineIdentity(line,b.machineNo||before?.machineNo))throw Error('Machine must be registered in Machine Management');if(!String(b.maintenance??before?.maintenance??'').trim())throw Error('Maintenance Task is required');}


app.post('/api/pm-master/lines/:line/rebuild-annual/:year',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),year=Number(req.params.year);
  if(!pmLineEntry(loadPmMaster(),line))return res.status(404).json({error:'Line not found'});
  const result=await rebuildAnnualFromSystemMaster(line,year);
  audit('ANNUAL_REBUILD_FROM_SYSTEM_MASTER',req.user,{line,year,result});
  res.json({ok:true,result});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});


app.post('/api/pm-master/lines/:line/reset-active-master/:year',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),year=Number(req.params.year);
  const l=pmLineEntry(loadPmMaster(),line);
  if(!l)return res.status(404).json({error:'Line not found'});

  const purgedTasks=purgeSystemTasksForLine(line,{
    year,
    archive:true,
    reason:`Manual active-master reset for ${line} ${year}`
  });
  purgeRowStateForLine(line);

  const rebuilt=await rebuildAnnualFromSystemMaster(line,year);
  audit('PM_LINE_ACTIVE_MASTER_RESET',req.user,{line,year,purgedTasks,rebuilt});
  res.json({ok:true,purgedTasks,rebuilt});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});


app.post('/api/pm-master/lines/:line/rebuild-monthly/:year/:month',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),year=Number(req.params.year),month=Number(req.params.month);
  if(!pmLineEntry(loadPmMaster(),line))return res.status(404).json({error:'Line not found'});
  if(month<1||month>12)return res.status(400).json({error:'Invalid month'});
  const result=await createMonthlyForLine(line,year,month,{overwrite:true,preserveExecution:true});
  audit('MONTHLY_REBUILD_FROM_SYSTEM_MASTER',req.user,{line,year,month,result});
  res.json({ok:true,result});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.get('/api/pm-master',requireAuth,(req,res)=>res.json(loadPmMaster()));

app.get('/api/pm-master/excel-machines/:line/:year',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year);await ensureSystemTasksBootstrapped(line,year);const tasks=annualRows122(line,year);res.json({line,year,machines:orderedMasterMachines(line).map(m=>({...m,types:m.type?[m.type]:[],tasks:tasks.filter(t=>same122(t.machineNo,m.machineNo))}))});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});



function annualFilesForLine(line){
  if(!fs.existsSync(ANNUAL_DIR))return [];
  const esc=String(line).replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  const rx=new RegExp('^'+esc+' PM (\\d{4})\\.xlsx$','i');
  return fs.readdirSync(ANNUAL_DIR).map(file=>{const m=file.match(rx);return m?{file,year:Number(m[1]),fullPath:path.join(ANNUAL_DIR,file)}:null}).filter(Boolean).sort((a,b)=>a.year-b.year);
}
function pmLineAliases(line){
  try{
    const x=loadPmMaster(),l=pmLineEntry(x,line);
    return [...new Set([line,...(Array.isArray(l?.historyNames)?l.historyNames:[])].filter(Boolean))];
  }catch(_){return [line];}
}
function monthlyFilesForLine(line){
  if(!fs.existsSync(MONTHLY_WORK_DIR))return [];
  const aliases=pmLineAliases(line);
  const out=[];
  const allFiles=listMonthlyExcelFiles();
  for(const alias of aliases){
    const escAlias=String(alias).replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    const rx=new RegExp('^'+escAlias+' PM_(\\d{4})_(\\d{2})\\.xlsx$','i');
    for(const fullPath of allFiles){
      const file=path.basename(fullPath),m=file.match(rx);
      if(m)out.push({file,year:Number(m[1]),month:Number(m[2]),fullPath,sourceLine:alias});
    }
  }
  const seen=new Set();
  return out.filter(x=>{const k=x.fullPath.toLowerCase();if(seen.has(k))return false;seen.add(k);return true;}).sort((a,b)=>a.year-b.year||a.month-b.month);
}
function pmLineEntry(x,line){return (x.lines||[]).find(v=>String(v.name||'').toLowerCase()===String(line||'').toLowerCase());}
function pmMachineEntry(l,machineNo){return (l?.machines||[]).find(v=>String(v.machineNo||'').toLowerCase()===String(machineNo||'').toLowerCase());}
function moveStateLineKey(filePath,oldLine,newLine){
  let st={};try{st=JSON.parse(fs.readFileSync(filePath,'utf8'));}catch(_){return;}
  const out={};
  for(const [k,v] of Object.entries(st)){
    const p=k.split('|');
    if(p.length>=3&&String(p[0]).toLowerCase()===String(oldLine).toLowerCase())p[0]=newLine;
    out[p.join('|')]=v;
  }
  fs.writeFileSync(filePath,JSON.stringify(out,null,2),'utf8');
}
function countAnnualTasksInFile(fullPath){
  return (async()=>{const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(fullPath);const ws=wb.getWorksheet('Full')||wb.worksheets[0];return taskRowsFromAnnual(ws).length;})();
}
async function machineExistsInMonthly(fullPath,machineNo){
  const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(fullPath);const ws=wb.getWorksheet('Sheet1')||wb.worksheets[0];
  let cur='';
  for(let r=pmDataStartRow(ws);r<=ws.rowCount;r++){
    const v=String(safeText(ws,r,2)||'').trim();if(v)cur=v;
    const act=String(safeText(ws,r,8)||'').trim().toLowerCase();
    if(act==='planed'&&cur.toLowerCase()===String(machineNo).toLowerCase())return true;
  }
  return false;
}
async function lineUsage(line){
  const annual=annualFilesForLine(line),monthly=monthlyFilesForLine(line);let tasks=0;
  for(const f of annual){try{tasks+=await countAnnualTasksInFile(f.fullPath)}catch(_){}}
  return {annualYears:annual.map(x=>x.year),annualFiles:annual.length,monthlyFiles:monthly.length,tasks};
}
function removeLineFromMaster(x,line){
  x.lines=(x.lines||[]).filter(v=>String(v.name||'').toLowerCase()!==String(line||'').toLowerCase());
  if(!Array.isArray(x.removedLines))x.removedLines=[];
  if(!x.removedLines.some(v=>String(v).toLowerCase()===String(line).toLowerCase()))x.removedLines.push(line);
}
function renameLineFiles(oldLine,newLine,opt={}){
  const planned=[];
  // Annual Excel is the Master Plan: rename annual workbooks with the line.
  for(const f of annualFilesForLine(oldLine))planned.push([f.fullPath,path.join(ANNUAL_DIR,`${newLine} PM ${f.year}.xlsx`)]);
  const oldTpl=path.join(MONTHLY_TEMPLATE_DIR,`${oldLine} PM_TEMPLATE.xlsx`),newTpl=path.join(MONTHLY_TEMPLATE_DIR,`${newLine} PM_TEMPLATE.xlsx`);
  if(fs.existsSync(oldTpl))planned.push([oldTpl,newTpl]);

  // Historical Monthly files are NOT renamed automatically.
  // Admin may explicitly rename only the selected current Monthly workbook.
  if(opt.applyCurrentMonthly && opt.year && opt.month){
    const oldMonthly=path.join(MONTHLY_WORK_DIR,`${oldLine} PM_${opt.year}_${String(opt.month).padStart(2,'0')}.xlsx`);
    const newMonthly=path.join(MONTHLY_WORK_DIR,`${newLine} PM_${opt.year}_${String(opt.month).padStart(2,'0')}.xlsx`);
    if(fs.existsSync(oldMonthly))planned.push([oldMonthly,newMonthly]);
  }

  for(const [,to] of planned)if(fs.existsSync(to))throw new Error(`Cannot rename: target file already exists: ${path.basename(to)}`);
  for(const [from,to] of planned)fs.renameSync(from,to);
  return planned.length;
}


app.get('/api/pm-master/lines/:line/monthly-files',requireAuth,requireAdmin,(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line);
  const files=monthlyFilesForLine(line).map(x=>({year:x.year,month:x.month,file:x.file,sourceLine:x.sourceLine||line}));
  res.json({line,files});
 }catch(e){res.status(500).json({error:e.message});}
});

app.get('/api/pm-master/lines/:line/usage',requireAuth,requireAdmin,async(req,res)=>{
 try{const line=decodeURIComponent(req.params.line);res.json({line,...await lineUsage(line)});}catch(e){res.status(500).json({error:e.message});}
});

app.put('/api/pm-master/lines/:line',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const oldLine=decodeURIComponent(req.params.line),b=req.body||{},newName=cleanName(b.name||oldLine);
  if(!newName)return res.status(400).json({error:'Line name is required'});
  const x=loadPmMaster(),l=pmLineEntry(x,oldLine);if(!l)return res.status(404).json({error:'Line not found'});
  const renamed=newName.toLowerCase()!==String(oldLine).toLowerCase();
  if(renamed&&x.lines.some(v=>v!==l&&String(v.name||'').toLowerCase()===newName.toLowerCase()))return res.status(400).json({error:'Another line already has this name'});
  let filesRenamed=0;
  if(renamed){
    // Annual Excel follows the renamed line because Annual is the Master Plan.
    // Old Monthly history stays untouched unless the Admin explicitly selects current Monthly.
    filesRenamed=renameLineFiles(oldLine,newName,{
      applyCurrentMonthly:!!b.applyCurrentMonthly,
      year:Number(b.monthlyYear||new Date().getFullYear()),
      month:Number(b.monthlyMonth||new Date().getMonth()+1)
    });
    moveStateLineKey(path.join(PM_MASTER_DIR,'task_state.json'),oldLine,newName);
    moveStateLineKey(path.join(PM_MASTER_DIR,'task_planning_state.json'),oldLine,newName);
    if(!Array.isArray(l.historyNames))l.historyNames=[];
    if(!l.historyNames.some(v=>String(v).toLowerCase()===String(oldLine).toLowerCase()))l.historyNames.push(oldLine);
    const ts=loadSystemTasks();for(const t of ts.tasks)if(same122(t.line,oldLine))t.line=newName;
    for(const k of Object.keys(ts.initialized||{}))if(k.startsWith(oldLine.toLowerCase()+'|')){ts.initialized[newName.toLowerCase()+k.slice(oldLine.length)]=true;delete ts.initialized[k];}saveSystemTasks(ts);
    renameMonthlyState122(oldLine,newName);
    if(!x.removedLines.some(v=>same122(v,oldLine)))x.removedLines.push(oldLine);
    l.name=newName;
    if(Array.isArray(x.removedLines))x.removedLines=x.removedLines.filter(v=>String(v).toLowerCase()!==newName.toLowerCase());
  }
  if(b.enabled!==undefined)l.enabled=!!b.enabled;
  savePmMaster(x);
  if(renamed){for(const y of new Set(loadSystemTasks().tasks.filter(t=>same122(t.line,newName)).map(t=>Number(t.year)).concat([Number(l.createdYear||new Date().getFullYear())]))){await rebuildAnnualFromSystemMaster(newName,y);}if(b.applyCurrentMonthly)await createMonthlyForLine(newName,Number(b.monthlyYear||new Date().getFullYear()),Number(b.monthlyMonth||new Date().getMonth()+1),{overwrite:true});}
  audit('PM_LINE_EDIT',req.user,{oldLine,newLine:l.name,enabled:l.enabled,filesRenamed,annualMasterUpdated:renamed,currentMonthlyUpdated:renamed&&!!b.applyCurrentMonthly});

    // Every production line owns its Monthly storage folder from the moment
    // the line is created. This is independent from Machines/Tasks/Monthly creation.
    fs.mkdirSync(monthlyLineFolder(l.name),{recursive:true});

  res.json({ok:true,line:l,filesRenamed,annualMasterUpdated:renamed,currentMonthlyUpdated:renamed&&!!b.applyCurrentMonthly});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.post('/api/pm-master/lines/:line/deactivate',requireAuth,requireAdmin,(req,res)=>{
 try{const line=decodeURIComponent(req.params.line),x=loadPmMaster(),l=pmLineEntry(x,line);if(!l)return res.status(404).json({error:'Line not found'});l.enabled=false;savePmMaster(x);audit('PM_LINE_DEACTIVATE',req.user,{line});res.json({ok:true});}catch(e){res.status(500).json({error:e.message});}
});
app.post('/api/pm-master/lines/:line/activate',requireAuth,requireAdmin,(req,res)=>{
 try{const line=decodeURIComponent(req.params.line),x=loadPmMaster(),l=pmLineEntry(x,line);if(!l)return res.status(404).json({error:'Line not found'});l.enabled=true;savePmMaster(x);audit('PM_LINE_ACTIVATE',req.user,{line});res.json({ok:true});}catch(e){res.status(500).json({error:e.message});}
});

app.delete('/api/pm-master/lines/:line',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),mode=String(req.body?.mode||'system'),x=loadPmMaster(),l=pmLineEntry(x,line);
  if(!l)return res.status(404).json({error:'Line not found'});
  if(!['system','excel','folder'].includes(mode))return res.status(400).json({error:'Invalid delete mode'});
  const usage=await lineUsage(line);
  let annualDeleted=0,monthlyDeleted=0,foldersDeleted=0;

  // Excel deletion modes remove Annual + Monthly workbooks. System-only keeps every file.
  if(mode==='excel'||mode==='folder'){
    for(const f of annualFilesForLine(line)){
      if(fs.existsSync(f.fullPath)){fs.unlinkSync(f.fullPath);annualDeleted++;}
    }
    const tpl=path.join(MONTHLY_TEMPLATE_DIR,`${line} PM_TEMPLATE.xlsx`);
    if(fs.existsSync(tpl)){fs.unlinkSync(tpl);annualDeleted++;}

    for(const f of monthlyFilesForLine(line)){
      if(fs.existsSync(f.fullPath)){fs.unlinkSync(f.fullPath);monthlyDeleted++;}
    }
  }

  // Complete deletion removes the production-line Monthly folder(s) after files are deleted.
  // Historical aliases belong to the same logical line and are included only in this explicit mode.
  if(mode==='folder'){
    for(const alias of pmLineAliases(line)){
      const folder=monthlyLineFolder(alias);
      if(fs.existsSync(folder)){
        fs.rmSync(folder,{recursive:true,force:true});
        foldersDeleted++;
      }
    }
  }

  // A deleted line must leave NO active machine/task configuration behind.
  const purgedTasks=purgeSystemTasksForLine(line,{
    archive:true,
    reason:`Line deleted from active system (${mode})`
  });
  purgeRowStateForLine(line);
  removeLineFromMaster(x,line);savePmMaster(x);
  audit('PM_LINE_DELETE',req.user,{line,mode,usage,annualDeleted,monthlyDeleted,foldersDeleted,purgedTasks});
  res.json({ok:true,mode,usage,annualDeleted,monthlyDeleted,foldersDeleted,purgedTasks,
    excelPreserved:mode==='system',folderPreserved:mode!=='folder'});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});


async function updateMachineInMonthly(line,year,month,oldNo,vals){
  const fullPath=monthlyPathFor(line,year,month);
  if(!fs.existsSync(fullPath))throw new Error(`Monthly Excel not found: ${year}-${String(month).padStart(2,'0')}`);
  const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(fullPath);normalizeWorkbookFormulas(wb);
  const ws=wb.getWorksheet('Sheet1')||wb.worksheets[0],tasks=taskRowsFromMonthly(ws);
  const matches=tasks.filter(t=>String(t.machineNo||'').toLowerCase()===String(oldNo).toLowerCase());
  snapshotAndUnmerge(ws);
  if(matches.length){
    for(const t of matches){
      for(const r of [t.row,t.row+1]){
        ws.getCell(r,1).value=vals.section||null;
        ws.getCell(r,2).value=vals.machineNo||null;
        ws.getCell(r,3).value=vals.machineName||null;
        ws.getCell(r,4).value=vals.machineIdentified||null;
        ws.getCell(r,5).value=vals.type||null;
      }
    }
    removeMachinePlaceholder(ws,oldNo);
  }else{
    const ph=machinePlaceholderRows(ws).find(x=>String(x.machineNo||'').toLowerCase()===String(oldNo).toLowerCase());
    if(ph)writeMachinePlaceholder(ws,ph.row,vals);
    else ensureMachinePlaceholderInWorksheet(ws,vals);
  }
  rebuildTaskGroupMerges(ws);rebuildMonthlyTaskFormulas(ws);
  formatProfessionalMonthlyLayout(ws,pmDataStartRow(ws),Math.max(pmDataStartRow(ws),lastLogicalDataRow(ws)));
  await safeWriteWorkbook(wb,fullPath);
  return matches.length;
}
async function deleteMachineFromMonthly(line,year,month,no){
  const fullPath=monthlyPathFor(line,year,month);
  if(!fs.existsSync(fullPath))throw new Error(`Monthly Excel not found: ${year}-${String(month).padStart(2,'0')}`);
  const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(fullPath);normalizeWorkbookFormulas(wb);
  const ws=wb.getWorksheet('Sheet1')||wb.worksheets[0],tasks=taskRowsFromMonthly(ws);
  const matches=tasks.filter(t=>String(t.machineNo||'').toLowerCase()===String(no).toLowerCase());
  if(!matches.length)throw new Error(`Machine ${no} not found in Monthly ${year}-${String(month).padStart(2,'0')}`);
  const first=Math.min(...matches.map(t=>t.row)),last=Math.max(...matches.map(t=>t.row))+1;
  snapshotAndUnmerge(ws);ws.spliceRows(first,last-first+1);rebuildTaskGroupMerges(ws);rebuildMonthlyTaskFormulas(ws);
  await safeWriteWorkbook(wb,fullPath);
  return matches.length;
}

app.put('/api/pm-master/lines/:line/machines/:machineNo',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line,machineNo}=req.params,b=req.body||{},year=Number(b.year||new Date().getFullYear()),x=loadPmMaster(),l=pmLineEntry(x,line),m=pmMachineEntry(l,machineNo);if(!m)throw Error('Machine not found');const no=String(b.machineNo||machineNo).trim();if(!no||!String(b.machineName||'').trim())throw Error('Machine No. and Name required');if(l.machines.some(v=>v!==m&&same122(v.machineNo,no)))throw Error('Machine No. already exists');await ensureSystemTasksBootstrapped(line,year);const ts=loadSystemTasks();for(const t of ts.tasks)if(same122(t.line,line)&&same122(t.machineNo,machineNo))t.machineNo=no;saveSystemTasks(ts);Object.assign(m,{machineNo:no,machineName:String(b.machineName).trim(),section:String(b.section||'').trim(),machineIdentified:String(b.machineIdentified||'').trim(),type:String(b.type||'').trim(),expectedWeeklyHours:Number(b.expectedWeeklyHours||m.expectedWeeklyHours||120)});savePmMaster(x);renameMonthlyMachine122(line,machineNo,no);recalcInheritedHoursTasksForMachine(line,no);const years=new Set([year,...ts.tasks.filter(t=>same122(t.line,line)).map(t=>Number(t.year))]);for(const y of years)await rebuildAnnualFromSystemMaster(line,y);if(b.applyCurrentMonthly)await createMonthlyForLine(line,Number(b.monthlyYear||year),Number(b.monthlyMonth),{overwrite:true});res.json({ok:true});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

app.post('/api/pm-master/lines/:line/machines/:machineNo/deactivate',requireAuth,requireAdmin,(req,res)=>{
 try{const line=decodeURIComponent(req.params.line),no=decodeURIComponent(req.params.machineNo),x=loadPmMaster(),l=pmLineEntry(x,line),m=pmMachineEntry(l,no);if(!m)return res.status(404).json({error:'Machine not found in system master'});m.enabled=false;savePmMaster(x);audit('PM_MACHINE_DEACTIVATE',req.user,{line,machineNo:no});res.json({ok:true});}catch(e){res.status(500).json({error:e.message});}
});
app.post('/api/pm-master/lines/:line/machines/:machineNo/activate',requireAuth,requireAdmin,(req,res)=>{
 try{const line=decodeURIComponent(req.params.line),no=decodeURIComponent(req.params.machineNo),x=loadPmMaster(),l=pmLineEntry(x,line),m=pmMachineEntry(l,no);if(!m)return res.status(404).json({error:'Machine not found in system master'});m.enabled=true;savePmMaster(x);audit('PM_MACHINE_ACTIVATE',req.user,{line,machineNo:no});res.json({ok:true});}catch(e){res.status(500).json({error:e.message});}
});


app.get('/api/pm-master/lines/:line/machines/:machineNo/usage',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),no=decodeURIComponent(req.params.machineNo);
  const year=Number(req.query.year||new Date().getFullYear());

  // Active task count comes from system master, never from Excel.
  const annualTasks=systemTasksFor(line,year).filter(t=>
    String(t.machineNo||'').toLowerCase()===String(no).toLowerCase()
  ).length;

  let monthlyTasks=0,monthlyFiles=0;
  for(const f of availableMonthlyFiles(line,year)){
    try{
      const m=await openMonthlyForLine(line,f.year,f.month);
      const count=taskRowsFromMonthly(m.ws).filter(t=>String(t.machineNo||'').toLowerCase()===String(no).toLowerCase()).length;
      if(count){monthlyTasks+=count;monthlyFiles++;}
    }catch(_){}
  }

  res.json({line,machineNo:no,year,annualTasks,monthlyTasks,monthlyFiles,totalTasks:annualTasks+monthlyTasks});
 }catch(e){res.status(500).json({error:e.message});}
});


async function rebuildExistingMonthlyFilesForLineYear(line,year,{preserveExecution=true}={}){
  const results=[];
  for(const f of availableMonthlyFiles(line,year)){
    try{
      const rebuilt=await createMonthlyForLine(line,year,f.month,{
        overwrite:true,
        preserveExecution
      });
      results.push({year,month:f.month,file:f.file,ok:true,rebuilt});
    }catch(e){
      results.push({year,month:f.month,file:f.file,ok:false,error:e.message});
    }
  }
  return results;
}

app.delete('/api/pm-master/lines/:line/machines/:machineNo',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line);
  const machineNo=decodeURIComponent(req.params.machineNo);
  const b=req.body||{};
  const year=Number(b.year||new Date().getFullYear());

  const x=loadPmMaster(),l=pmLineEntry(x,line);
  const machine=pmMachineEntry(l,machineNo);
  if(!machine)return res.status(404).json({error:'Machine not found'});

  await ensureSystemTasksBootstrapped(line,year);

  // Active Task Master is authoritative. Machine cannot be deleted while tasks remain.
  const assigned=loadSystemTasks().tasks.filter(t=>
    same122(t.line,line) &&
    Number(t.year)===year &&
    same122(t.machineNo,machineNo) &&
    t.active!==false
  );
  if(assigned.length){
    return res.status(409).json({
      error:'Delete or move assigned tasks first',
      taskCount:assigned.length
    });
  }

  // Remove machine from System Master first.
  l.machines=(l.machines||[]).filter(m=>!same122(m.machineNo,machineNo));
  savePmMaster(x);

  // Annual is regenerated from current System Master.
  const annual=await rebuildAnnualFromSystemMaster(line,year);

  // IMPORTANT:
  // Existing Monthly files are structural reports too.
  // A machine deleted from the active system must disappear from every existing
  // Monthly workbook for the selected Annual year. Execution data for all
  // remaining Task IDs is preserved by createMonthlyForLine().
  const monthly=rebuildExistingMonthlyFilesForLineYear
    ? await rebuildExistingMonthlyFilesForLineYear(line,year,{preserveExecution:true})
    : [];

  const failed=monthly.filter(x=>!x.ok);
  if(failed.length){
    audit('PM_MACHINE_DELETE_PARTIAL',req.user,{
      line,year,machineNo,annual,
      monthly,
      error:'One or more Monthly files could not be rebuilt'
    });
    return res.status(500).json({
      error:'Machine deleted from system and Annual, but one or more Monthly files could not be rebuilt',
      machineDeleted:true,
      annual,
      monthly
    });
  }

  audit('PM_MACHINE_DELETE',req.user,{
    line,year,machineNo,
    annual,
    monthlyFilesRebuilt:monthly.length
  });

  res.json({
    ok:true,
    machineDeleted:true,
    annual,
    monthlyFilesRebuilt:monthly.length,
    monthly
  });
 }catch(e){
  console.error(e);
  res.status(e.status||500).json({error:e.message});
 }
});


function prepareAnnualCalendarHeaders(ws,line,year){
  // Branding is written in row 1; line/year are metadata.
  ws.getCell(3,10).value=line;
  ws.getCell(3,11).value=year;
  // Clear annual calendar header area P:BO first.
  for(let c=16;c<=67;c++){
    ws.getCell(2,c).value=null;
    ws.getCell(3,c).value=null;
    ws.getCell(4,c).value=null;
  }
  let d=new Date(year,0,1,12);
  while(d.getDay()!==5)d.setDate(d.getDate()+1);
  let col=16,lastMonth=-1;
  while(d.getFullYear()===year && col<=67){
    const month=d.getMonth();
    const th=new Date(d);th.setDate(th.getDate()+6);
    if(month!==lastMonth){
      ws.getCell(2,col).value=year;
      ws.getCell(3,col).value=monthShort(month+1);
      lastMonth=month;
    }
    ws.getCell(4,col).value=`${d.getDate()}--${th.getDate()}`;
    d.setDate(d.getDate()+7);col++;
  }
}
app.post('/api/pm-master/lines',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const b=req.body||{},name=cleanName(b.name),year=Number(b.year||new Date().getFullYear());
  if(!name)return res.status(400).json({error:'Line name is required'});

  const x=loadPmMaster();
  if(x.lines.some(v=>String(v.name||'').toLowerCase()===name.toLowerCase())){
    return res.status(400).json({error:'Line already exists'});
  }

  const annualSource=String(b.annualSource||'blank_auto');
  const monthlySource=String(b.monthlySource||'blank_auto');
  const annualFile=annualPathFor(name,year);
  const newMonthlyTemplate=path.join(MONTHLY_TEMPLATE_DIR,`${name} PM_TEMPLATE.xlsx`);

  // IMPORTANT LIFECYCLE RULE:
  // If this is a genuinely fresh line (no Annual exists), old orphan tasks with the
  // same line name must NOT come back from pm_tasks.json.
  // This is exactly what caused deleted test Tasks to reappear in Step119.
  const freshLine=annualSource==='blank_auto'||annualSource==='new_from_template';
  if(freshLine){const oldFolder=monthlyLineFolder(name);if(fs.existsSync(oldFolder)&&fs.readdirSync(oldFolder).length){const archive=path.join(PM_MASTER_DIR,'monthly_archive',name+'-'+Date.now());fs.mkdirSync(path.dirname(archive),{recursive:true});fs.renameSync(oldFolder,archive);}const ts=loadSystemTasks();ts.initialized=ts.initialized||{};ts.initialized[name.toLowerCase()+'|'+year]=true;saveSystemTasks(ts);resetMonthlyState122(name);}
  let orphanTasksPurged=0;
  if(freshLine){
    orphanTasksPurged=purgeSystemTasksForLine(name,{
      archive:true,
      reason:`Fresh line ${name} created for ${year}; removing orphan active tasks from earlier deleted line`
    });
    purgeRowStateForLine(name);
  }

  // A line always owns its Monthly folder immediately.
  fs.mkdirSync(monthlyLineFolder(name),{recursive:true});

  let restoredAnnual=false;
  let restoredMonthlyTemplate=false;
  let annualSourceName='Smart PM Blank Annual';
  let monthlyTemplateSourceName='Smart PM Blank Monthly';

  // ------------------------------------------------------
  // ANNUAL
  // If Remove from System Only was used, the Annual should
  // already exist. Reuse it instead of treating it as an error.
  // ------------------------------------------------------
  if(fs.existsSync(annualFile)&&!freshLine&&!['copy_existing','upload'].includes(annualSource)){
    restoredAnnual=true;
    annualSourceName='Existing Annual Excel';
  }else{
    if(annualSource==='blank_auto'){
      const built=await buildCanonicalAnnualTemplate(name,year),wb=built.wb,ws=built.ws;
      await safeWriteWorkbook(wb,annualFile);

    }else if(annualSource==='new_from_template'){
      annualSourceName=String(b.annualTemplateLine||'').trim();
      const source=annualPathFor(annualSourceName,year);
      if(!annualSourceName||!fs.existsSync(source))return res.status(400).json({error:'Select an existing Annual layout'});

      const wb=new ExcelJS.Workbook();
      await wb.xlsx.readFile(source);
      normalizeWorkbookFormulas(wb);

      const ws=wb.getWorksheet('Full')||wb.worksheets[0],tasks=taskRowsFromAnnual(ws);
      if(tasks.length){
        const first=Math.min(...tasks.map(t=>t.row)),last=Math.max(...tasks.map(t=>t.row))+1;
        ws.spliceRows(first,last-first+1);
      }

      prepareAnnualCalendarHeaders(ws,name,year);
      applyBrandingToWorkbook(wb,ws,'annual');
      await safeWriteWorkbook(wb,annualFile);

    }else if(annualSource==='copy_existing'){
      annualSourceName=String(b.annualTemplateLine||'').trim();
      const source=annualPathFor(annualSourceName,year);
      if(!annualSourceName||!fs.existsSync(source))return res.status(400).json({error:'Select an existing Annual Excel'});

      const wb=new ExcelJS.Workbook();
      await wb.xlsx.readFile(source);
      normalizeWorkbookFormulas(wb);
      const ws=wb.getWorksheet('Full')||wb.worksheets[0];
      prepareAnnualCalendarHeaders(ws,name,year);
      applyBrandingToWorkbook(wb,ws,'annual');
      await safeWriteWorkbook(wb,annualFile);

    }else if(annualSource==='upload'){
      const staged=stagedUploadPath(b.annualUploadToken);
      if(!staged||!fs.existsSync(staged))return res.status(400).json({error:'Upload the Annual Excel first'});

      const wb=new ExcelJS.Workbook();
      await wb.xlsx.readFile(staged);
      normalizeWorkbookFormulas(wb);
      prepareAnnualCalendarHeaders(wb.getWorksheet('Full')||wb.worksheets[0],name,year);
      await safeWriteWorkbook(wb,annualFile);
      try{fs.unlinkSync(staged)}catch(_){}
      annualSourceName='Uploaded Excel';

    }else{
      return res.status(400).json({error:'Invalid Annual Excel source'});
    }
  }

  // ------------------------------------------------------
  // MONTHLY TEMPLATE
  // Same rule: if it already exists from an earlier System-Only
  // removal, reuse it and do not block restoring the line.
  // ------------------------------------------------------
  if(fs.existsSync(newMonthlyTemplate)){
    restoredMonthlyTemplate=true;
    monthlyTemplateSourceName='Existing Monthly Template';
  }else{
    if(monthlySource==='blank_auto'){
      const built=await buildCanonicalMonthlyTemplate(name,year,1);
      await safeWriteWorkbook(built.wb,newMonthlyTemplate);

    }else if(monthlySource==='system_template'){
      monthlyTemplateSourceName=String(b.monthlyTemplateLine||'').trim();
      const source=path.join(MONTHLY_TEMPLATE_DIR,`${monthlyTemplateSourceName} PM_TEMPLATE.xlsx`);
      if(!monthlyTemplateSourceName||!fs.existsSync(source)){
        if(!restoredAnnual){try{fs.unlinkSync(annualFile)}catch(_){}}
        return res.status(400).json({error:'Select an existing Monthly layout template'});
      }
      fs.copyFileSync(source,newMonthlyTemplate);

    }else if(monthlySource==='upload'){
      const staged=stagedUploadPath(b.monthlyUploadToken);
      if(!staged||!fs.existsSync(staged)){
        if(!restoredAnnual){try{fs.unlinkSync(annualFile)}catch(_){}}
        return res.status(400).json({error:'Upload the Monthly template first'});
      }

      const wb=new ExcelJS.Workbook();
      await wb.xlsx.readFile(staged);
      normalizeWorkbookFormulas(wb);
      await safeWriteWorkbook(wb,newMonthlyTemplate);
      try{fs.unlinkSync(staged)}catch(_){}
      monthlyTemplateSourceName='Uploaded Excel';

    }else{
      if(!restoredAnnual){try{fs.unlinkSync(annualFile)}catch(_){}}
      return res.status(400).json({error:'Invalid Monthly template source'});
    }
  }

  if(Array.isArray(x.removedLines)){
    x.removedLines=x.removedLines.filter(v=>String(v).toLowerCase()!==name.toLowerCase());
  }

  x.lines.push({
    name,
    enabled:true,
    machines:[],
    createdYear:year,
    fresh122:freshLine,
    annualSource:restoredAnnual?'existing':annualSource,
    annualTemplateLine:annualSourceName,
    monthlySource:restoredMonthlyTemplate?'existing':monthlySource,
    templateLine:monthlyTemplateSourceName
  });

  savePmMaster(x);

  audit(restoredAnnual?'PM_LINE_RESTORE':'PM_LINE_ADD',req.user,{
    line:name,
    year,
    restoredAnnual,
    restoredMonthlyTemplate,
    annualSourceName,
    monthlyTemplateSourceName,
    monthlyFolder:monthlyLineFolder(name),
    orphanTasksPurged
  });

  res.json({
    ok:true,
    line:name,
    year,
    restored:restoredAnnual,
    annualFile:path.basename(annualFile),
    monthlyTemplate:path.basename(newMonthlyTemplate),
    monthlyFolder:monthlyLineFolder(name),
    orphanTasksPurged
  });

 }catch(e){
  console.error(e);
  res.status(500).json({error:e.message});
 }
});

app.post('/api/pm-master/lines/:line/machines',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),b=req.body||{},year=Number(b.year||new Date().getFullYear());
  const no=String(b.machineNo||'').trim(),name=String(b.machineName||'').trim();
  if(!no||!name)return res.status(400).json({error:'Machine No. and Machine Name are required'});
  const x=loadPmMaster(),l=pmLineEntry(x,line);
  if(!l)return res.status(404).json({error:'Line not found'});
  if((l.machines||[]).some(m=>String(m.machineNo||'').toLowerCase()===no.toLowerCase()))return res.status(400).json({error:'Machine No. already exists'});
  const machineObj={machineNo:no,machineName:name,machineIdentified:String(b.machineIdentified||'').trim(),section:String(b.section||'').trim(),type:String(b.type||'').trim(),expectedWeeklyHours:Number(b.expectedWeeklyHours||120),enabled:true};
  l.machines.push(machineObj);
  savePmMaster(x);

  await ensureSystemTasksBootstrapped(line,year);
  const annualSync=await rebuildAnnualFromSystemMaster(line,year);
  let monthlySync={updated:false};
  if(b.applyCurrentMonthly){
    const my=Number(b.monthlyYear||year),mm=Number(b.monthlyMonth||0);
    if(mm>=1&&mm<=12)monthlySync=await createMonthlyForLine(line,my,mm,{overwrite:true});
  }

  audit('PM_MACHINE_ADD',req.user,{line,year,machineNo:no,machineName:name,annualSync,monthlySync});
  res.json({ok:true,machineNo:no,annualSync,monthlySync});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.post('/api/pm-master/lines/:line/create-current-month',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),b=req.body||{},now=new Date(),year=Number(b.year||now.getFullYear()),month=Number(b.month||now.getMonth()+1);
  const target=monthlyPathFor(line,year,month);
  const existed=fs.existsSync(target);
  if(typeof createMonthlyForLine!=='function')return res.status(500).json({error:'Monthly creator function is not available'});
  // Always refresh Current Monthly after an Annual import/update so newly imported
  // tasks appear immediately. createMonthlyForLine(overwrite:true) keeps execution
  // state/history by stable task ID instead of resetting completed work.
  await createMonthlyForLine(line,year,month,{overwrite:true});
  audit('PM_MONTHLY_CREATE',req.user,{line,year,month,file:path.basename(target),refreshed:existed});
  res.json({ok:true,file:path.basename(target),alreadyExists:false,updatedExisting:existed});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.get('/api/task-management-current-period',requireAuth,requireAdmin,(req,res)=>{
  const now=new Date();
  const currentFriday=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12,0,0);
  currentFriday.setDate(currentFriday.getDate()-((currentFriday.getDay()-5+7)%7));
  const year=currentFriday.getFullYear(), month=currentFriday.getMonth()+1;
  const fridays=fridaysInMonth(year,month);
  res.json({
    year,month,
    weeks:fridays.map((f,i)=>{
      const e=new Date(f);e.setDate(e.getDate()+6);
      return {index:i+1,label:`W${i+1} ${f.getDate()}--${e.getDate()}`};
    })
  });
});



// ===== STEP 27: Independent Monthly Task Management (ADMIN ONLY) =====
async function openMonthlyForLine(line,year,month){
  if(!validPmLine(line)) throw new Error('Invalid PM line');
  if(!Number.isInteger(Number(month))||Number(month)<1||Number(month)>12) throw new Error('Invalid month');
  const file=monthlyFileName(line,Number(year),Number(month));
  const fullPath=findExistingMonthlyFile(line,Number(year),Number(month));
  if(!fs.existsSync(fullPath)) throw new Error(`Monthly file not found: ${file}`);
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(fullPath);
  normalizeWorkbookFormulas(wb);
  const ws=wb.getWorksheet('Sheet1')||wb.worksheets[0];
  if(!ws)throw new Error(`No worksheet found in ${file}`);
  return {file,fullPath,wb,ws};
}

function taskRowsFromMonthly(ws){
  const out=[];
  let cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
  for(let r=pmDataStartRow(ws);r<=ws.rowCount;r++){
    const activity=String(safeText(ws,r,8)||'').trim().toLowerCase();
    const vals={
      section:String(safeText(ws,r,1)||'').trim(),
      machineNo:String(safeText(ws,r,2)||'').trim(),
      machineName:String(safeText(ws,r,3)||'').trim(),
      machineIdentified:String(safeText(ws,r,4)||'').trim(),
      type:String(safeText(ws,r,5)||'').trim()
    };
    const linkedMachine=stableMachineLink(ws,r);
    if(linkedMachine)vals.machineNo=linkedMachine;
    if(vals.machineNo && vals.machineNo!==cur.machineNo)cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
    if(vals.section)cur.section=vals.section;
    if(vals.machineNo)cur.machineNo=vals.machineNo;
    if(vals.machineName)cur.machineName=vals.machineName;
    if(vals.machineIdentified && vals.machineIdentified!=='0 0')cur.machineIdentified=vals.machineIdentified;
    if(vals.type)cur.type=vals.type;
    if(activity!=='planed')continue;
    const codes=[];
    for(let c=16;c<=20;c++){
      const n=Number(ws.getCell(r,c).value);
      if([1,2,3,4,5].includes(n))codes.push({week:c-15,code:n});
    }
    const counts={1:0,2:0,3:0,4:0,5:0};codes.forEach(x=>counts[x.code]++);
    let dominantCode=1,max=-1;for(let k=1;k<=5;k++)if(counts[k]>max){max=counts[k];dominantCode=k;}
    out.push({
      row:r,taskId:stableTaskId(ws,r),section:cur.section,machineNo:cur.machineNo,machineName:cur.machineName,
      machineIdentified:cur.machineIdentified,type:cur.type,
      part:String(safeText(ws,r,6)||'').trim(),maintenance:String(safeText(ws,r,7)||'').trim(),
      frequency:String(safeText(ws,r,11)||'').trim(),dominantCode,plannedWeeks:codes.length,
      weekIndexes:codes.map(x=>x.week)
    });
  }
  return out;
}

function monthlyMasterFromTasks(tasks){
  const sections=[...new Set(tasks.map(t=>t.section).filter(Boolean))];
  const map=new Map();
  for(const t of tasks){
    const key=[t.section,t.machineName,t.machineNo,t.machineIdentified].join('|');
    if(!map.has(key))map.set(key,{section:t.section,machineName:t.machineName,machineNo:t.machineNo,machineIdentified:t.machineIdentified,types:[],tasks:[]});
    const m=map.get(key);if(t.type&&!m.types.includes(t.type))m.types.push(t.type);
    m.tasks.push({row:t.row,type:t.type,part:t.part,maintenance:t.maintenance});
  }
  return {sections,machines:[...map.values()]};
}


function annualWeeksForMonth(ws,row,year,month){
  const out=[];let d=new Date(year,0,1,12);while(d.getDay()!==5)d.setDate(d.getDate()+1);let i=1;
  while(d.getFullYear()===year){
    if(d.getMonth()+1===month && [1,2,3,4,5].includes(Number(ws.getCell(row,15+i).value))){
      const mf=fridaysInMonth(year,month),mi=mf.findIndex(x=>x.getDate()===d.getDate())+1;if(mi>0)out.push(mi);
    }
    i++;d.setDate(d.getDate()+7);
  } return out;
}
function exactTaskMatch(tasks,t){
 return tasks.find(x=>sameTaskIdentity(x,t))||tasks.find(x=>String(x.machineNo||'').trim()===String(t.machineNo||'').trim()&&String(x.type||'').trim()===String(t.type||'').trim()&&String(x.maintenance||'').trim()===String(t.maintenance||'').trim());
}
async function syncAnnualTaskToMonthly(line,year,month,annualWs,annualTask,b,mode,before){
 const m=await openMonthlyForLine(line,year,month),tasks=taskRowsFromMonthly(m.ws),found=exactTaskMatch(tasks,before||annualTask);
 const weeks=annualWeeksForMonth(annualWs,annualTask.row,year,month);
 if(mode==='edit'){
   if(!found)return {updated:false,reason:'Task not found in selected Monthly Excel',file:m.file};

   const edited={
     ...b,
     section:String(b.section||'').trim() || found.section || '',
     machineNo:String(b.machineNo||'').trim() || found.machineNo || '',
     machineName:String(b.machineName||'').trim() || found.machineName || '',
     machineIdentified:String(b.machineIdentified||'').trim() || found.machineIdentified || '',
     type:String(b.type||'').trim() || found.type || '',
     part:b.part!==undefined ? String(b.part).trim() : (found.part||''),
     maintenance:b.maintenance!==undefined ? String(b.maintenance).trim() : (found.maintenance||''),
     frequency:b.frequency!==undefined ? String(b.frequency).trim() : (found.frequency||'')
   };

   const identityChanged=
     edited.section!==String(found.section||'') ||
     edited.machineNo!==String(found.machineNo||'') ||
     edited.machineName!==String(found.machineName||'') ||
     edited.machineIdentified!==String(found.machineIdentified||'') ||
     edited.type!==String(found.type||'');

   // Safe normal edit: task text/meta only. Never touch merged A:E if identity is unchanged.
   m.ws.getCell(found.row,6).value=edited.part;
   m.ws.getCell(found.row,7).value=edited.maintenance;

   if(identityChanged){
     const identityRows=captureEffectiveIdentityRows(m.ws);
     const currentIdentity=identityRows.find(x=>x.r===found.row);
     if(currentIdentity){
       currentIdentity.section=edited.section;
       currentIdentity.machineNo=edited.machineNo;
       currentIdentity.machineName=edited.machineName;
       currentIdentity.machineIdentified=edited.machineIdentified;
       currentIdentity.type=edited.type;
     }

     const ranges=[];
     if(m.ws._merges){
       for(const key of Object.keys(m.ws._merges)){
         const mm=String(key).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
         if(!mm)continue;
         const c1=colNumber(mm[1]),r1=Number(mm[2]),c2=colNumber(mm[3]),r2=Number(mm[4]);
         if(r2>=pmDataStartRow(m.ws)&&r1<=m.ws.rowCount&&c1<=5&&c2>=1)ranges.push(key);
       }
     }
     for(const rg of ranges)try{m.ws.unMergeCells(rg)}catch(_){}
     materializeIdentityRows(m.ws,identityRows);
     rebuildTaskGroupMerges(m.ws);
   }

   applyProfessionalTaskMeta(m.ws,found.row,found.row+1,edited,Number(b.code));

   // Keep the existing Step86 monthly schedule logic exactly as-is.
   for(let c=16;c<=20;c++){m.ws.getCell(found.row,c).value=null;clearCellFill(m.ws.getCell(found.row,c));}
   for(const w of weeks){m.ws.getCell(found.row,15+w).value=Number(b.code);setPmCodeFill(m.ws.getCell(found.row,15+w),Number(b.code));}

   setMonthlyRowFormulas(m.ws,found.row,found.row+1);
   rebuildMonthlyTaskFormulas(m.ws);
   formatProfessionalMonthlyLayout(m.ws,pmDataStartRow(m.ws),m.ws.rowCount);
   await safeWriteWorkbook(m.wb,m.fullPath);
   return {updated:true,file:m.file,row:found.row};
 }
 if(found)return {updated:false,reason:'Task already exists in selected Monthly Excel',file:m.file};
 if(!weeks.length)return {updated:false,reason:'Task has no planned PM week in selected month',file:m.file};
 const insertAt=findOrderedInsertRow(m.ws,b.machineNo,0,b.taskPosition,b.type),same=tasks.filter(t=>String(t.machineNo||'').trim()===String(b.machineNo||'').trim());
 const template=same.length?same[same.length-1].row:(tasks.filter(t=>t.row<insertAt).slice(-1)[0]?.row||tasks[0]?.row||pmDataStartRow(m.ws)),pair=insertStyledTaskPair(m.ws,insertAt,template);
 writeTaskIdentity(m.ws,pair.planRow,pair.doneRow,b);applyProfessionalTaskMeta(m.ws,pair.planRow,pair.doneRow,b,Number(b.code));
 for(let c=16;c<=20;c++){m.ws.getCell(pair.planRow,c).value=null;m.ws.getCell(pair.doneRow,c).value=0;}
 for(const w of weeks){m.ws.getCell(pair.planRow,15+w).value=Number(b.code);setPmCodeFill(m.ws.getCell(pair.planRow,15+w),Number(b.code));}
 for(let c=21;c<=25;c++){m.ws.getCell(pair.planRow,c).value=null;m.ws.getCell(pair.doneRow,c).value=null;}
 setMonthlyRowFormulas(m.ws,pair.planRow,pair.doneRow);rebuildMonthlyTaskFormulas(m.ws);rebuildTaskGroupMerges(m.ws);formatProfessionalMonthlyLayout(m.ws,pmDataStartRow(m.ws),m.ws.rowCount);formatProfessionalMonthlyLayout(m.ws,pmDataStartRow(m.ws),m.ws.rowCount);await safeWriteWorkbook(m.wb,m.fullPath);
 return {updated:true,file:m.file,row:pair.planRow};
}

app.get('/api/monthly-task-management',requireAuth,requireAdmin,async(req,res)=>{
 try{
const line=String(req.query.line),year=Number(req.query.year),month=Number(req.query.month);await ensureSystemTasksBootstrapped(line,year);const p=await monthlyState122(line,year,month),tasks=monthlyRows122(line,year,month,p);const master={sections:[...new Set(orderedMasterMachines(line).map(m=>m.section))],machines:orderedMasterMachines(line).map(m=>({...m,types:m.type?[m.type]:[],tasks:tasks.filter(t=>same122(t.machineNo,m.machineNo))}))};res.json({line,year,month,file:monthlyFileName(line,year,month),tasks,master,weeks:fridaysInMonth(year,month).map((f,i)=>({index:i+1,label:'W'+(i+1)})),codeInfo:CODE_INFO});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

app.post('/api/monthly-task-management/:line/:year/:month',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year),month=Number(req.params.month),b=req.body||{};await ensureSystemTasksBootstrapped(line,year);const p=await monthlyState122(line,year,month);const before=null;validateTask122(line,b,before);const weeks=[...new Set((b.weekIndexes||[]).map(Number))];if(!weeks.length||weeks.some(w=>w<1||w>fridaysInMonth(year,month).length))throw Error('Select valid PM weeks');const id=before?.taskId||taskLinkId();const task={...before,taskId:id,line,year,machineNo:b.machineNo||before?.machineNo,part:String(b.part??before?.part??''),maintenance:String(b.maintenance??before?.maintenance??''),frequency:String(b.frequency??before?.frequency??''),code:Number(b.code),effectiveCode:Number(b.code),scheduleCodes:undefined,scheduleWeeks:weeks.map(w=>annualColumnForFriday(fridaysInMonth(year,month)[w-1])-15)};if(before){const extra=p.extras.findIndex(t=>t.taskId===id);if(extra>=0)p.extras[extra]=task;else p.overrides[id]=task;}else p.extras.push(task);if(b.applyAnnual){const store=loadSystemTasks(),at=store.tasks.find(t=>t.taskId===id);if(at){for(const k of ['machineNo','part','maintenance'])at[k]=task[k];}else{store.tasks.push({...task,active:true});p.extras=p.extras.filter(t=>t.taskId!==id);p.overrides[id]=task;}saveSystemTasks(store);await rebuildAnnualFromSystemMaster(line,year);}saveMonthlyState122(line,year,month,p);await createMonthlyForLine(line,year,month,{overwrite:true});res.json({ok:true,taskId:id,annualUpdated:!!b.applyAnnual,annualPlanningChanged:false});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

app.put('/api/monthly-task-management/:line/:year/:month/:row',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year),month=Number(req.params.month),b=req.body||{};await ensureSystemTasksBootstrapped(line,year);const p=await monthlyState122(line,year,month);const before=resolveMonthly122(line,year,month,p,req.params.row,b.taskId);validateTask122(line,b,before);const weeks=[...new Set((b.weekIndexes||[]).map(Number))];if(!weeks.length||weeks.some(w=>w<1||w>fridaysInMonth(year,month).length))throw Error('Select valid PM weeks');const id=before?.taskId||taskLinkId();const task={...before,taskId:id,line,year,machineNo:b.machineNo||before?.machineNo,part:String(b.part??before?.part??''),maintenance:String(b.maintenance??before?.maintenance??''),frequency:String(b.frequency??before?.frequency??''),code:Number(b.code),effectiveCode:Number(b.code),scheduleWeeks:weeks.map(w=>annualColumnForFriday(fridaysInMonth(year,month)[w-1])-15)};if(before){const extra=p.extras.findIndex(t=>t.taskId===id);if(extra>=0)p.extras[extra]=task;else p.overrides[id]=task;}else p.extras.push(task);if(b.applyAnnual){const store=loadSystemTasks(),at=store.tasks.find(t=>t.taskId===id);if(at){for(const k of ['machineNo','part','maintenance'])at[k]=task[k];}else{store.tasks.push({...task,active:true});p.extras=p.extras.filter(t=>t.taskId!==id);p.overrides[id]=task;}saveSystemTasks(store);await rebuildAnnualFromSystemMaster(line,year);}saveMonthlyState122(line,year,month,p);await createMonthlyForLine(line,year,month,{overwrite:true});res.json({ok:true,taskId:id,annualUpdated:!!b.applyAnnual,annualPlanningChanged:false});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

app.delete('/api/monthly-task-management/:line/:year/:month/:row',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year),month=Number(req.params.month),p=await monthlyState122(line,year,month),t=resolveMonthly122(line,year,month,p,req.params.row||req.params.monthlyRow,req.body?.taskId);p.excluded.push(t.taskId);p.extras=p.extras.filter(v=>v.taskId!==t.taskId);delete p.overrides[t.taskId];saveMonthlyState122(line,year,month,p);let annualDeleted=false;if(false){const store=loadSystemTasks();annualDeleted=store.tasks.some(v=>v.taskId===t.taskId);store.tasks=store.tasks.filter(v=>v.taskId!==t.taskId);saveSystemTasks(store);await rebuildAnnualFromSystemMaster(line,year);}await createMonthlyForLine(line,year,month,{overwrite:true});res.json({ok:true,annualDeleted,monthlyDeleted:true});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});


app.delete('/api/task-management-both/:line/:year/:month/:monthlyRow',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year),month=Number(req.params.month),p=await monthlyState122(line,year,month),t=resolveMonthly122(line,year,month,p,req.params.row||req.params.monthlyRow,req.body?.taskId);p.excluded.push(t.taskId);p.extras=p.extras.filter(v=>v.taskId!==t.taskId);delete p.overrides[t.taskId];saveMonthlyState122(line,year,month,p);let annualDeleted=false;if(true){const store=loadSystemTasks();annualDeleted=store.tasks.some(v=>v.taskId===t.taskId);store.tasks=store.tasks.filter(v=>v.taskId!==t.taskId);saveSystemTasks(store);await rebuildAnnualFromSystemMaster(line,year);}await createMonthlyForLine(line,year,month,{overwrite:true});res.json({ok:true,annualDeleted,monthlyDeleted:true});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

app.get('/api/task-management-months',requireAuth,requireAdmin,(req,res)=>{
  try{
    const line=String(req.query.line||'Layer Cake');
    const year=Number(req.query.year||new Date().getFullYear());
    res.json({line,year,months:availableMonthlyFiles(line,year).map(x=>({year:x.year,month:x.month,file:x.file}))});
  }catch(e){res.status(500).json({error:e.message});}
});

app.get('/api/task-management',requireAuth,requireAdmin,async(req,res)=>{
 try{
const line=String(req.query.line),year=Number(req.query.year);await ensureSystemTasksBootstrapped(line,year);res.json({line,year,tasks:annualRows122(line,year),codeInfo:CODE_INFO});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

app.put('/api/task-management/:line/:year/:row',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year),b=req.body||{};await ensureSystemTasksBootstrapped(line,year);const ref=resolveAnnual122(line,year,req.params.row,b.taskId),store=loadSystemTasks(),t=store.tasks.find(t=>t.taskId===ref.taskId);validateTask122(line,b,t);for(const k of ['machineNo','part','maintenance','frequency','notes'])if(b[k]!==undefined)t[k]=String(b[k]).trim();if(b.applyCodeToPlannedWeeks){t.effectiveCode=t.code=Number(b.code);delete t.scheduleCodes;}if(b.recalculatePlanning){delete t.scheduleCodes;const date=String(b.firstPmDate||'');if(!date.startsWith(year+'-'))throw Error('Planning date must be inside selected year');const useMachineExpectedHours=b.useMachineExpectedHours!==false;const effectiveHours=useMachineExpectedHours?machineExpectedWeeklyHours(line,b.machineNo||t.machineNo):Number(b.expectedWeeklyHours||0);const calcMethod=String(b.calculationMethod||'calendar').toLowerCase();const calcOpts={...b,expectedWeeklyHours:effectiveHours};const schedule=generateAnnualSchedule(year,date,calcOpts);if(!schedule.weeks.length && calcMethod==='calendar')throw Error('No PM weeks generated');const isHours=calcMethod==='hours'||calcMethod==='plc_hours';Object.assign(t,{firstPmDate:isHours?(schedule.firstDueDate||''):date,runningHoursCounterStartDate:isHours?date:(t.runningHoursCounterStartDate||''),calculationMethod:b.calculationMethod,calendarFrequency:b.calendarFrequency,intervalHours:Number(b.intervalHours||0),expectedWeeklyHours:useMachineExpectedHours?0:effectiveHours,useMachineExpectedHours,scheduleWeeks:schedule.weeks,effectiveCode:schedule.derivedCode||Number(b.code)});}else if(b.useMachineExpectedHours!==undefined){t.useMachineExpectedHours=!!b.useMachineExpectedHours;if(t.useMachineExpectedHours)t.expectedWeeklyHours=0;else if(Number(b.expectedWeeklyHours)>0)t.expectedWeeklyHours=Number(b.expectedWeeklyHours);}saveSystemTasks(store);await rebuildAnnualFromSystemMaster(line,year);let monthlyUpdated=false;if(b.applyCurrentMonth){const mm=Number(b.syncMonthlyMonth),p=await monthlyState122(line,year,mm);if(p.overrides[t.taskId]){for(const k of ['machineNo','part','maintenance','frequency'])p.overrides[t.taskId][k]=t[k];if(b.applyCodeToPlannedWeeks){p.overrides[t.taskId].code=t.code;p.overrides[t.taskId].effectiveCode=t.effectiveCode;delete p.overrides[t.taskId].scheduleCodes;}if(b.recalculatePlanning){p.overrides[t.taskId].scheduleWeeks=t.scheduleWeeks;delete p.overrides[t.taskId].scheduleCodes;}saveMonthlyState122(line,year,mm,p);}await createMonthlyForLine(line,year,mm,{overwrite:true});monthlyUpdated=true;}res.json({ok:true,monthlyUpdated});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});


// ---------- TASK MANAGEMENT ADD / DEACTIVATE ----------
function copyRowStyleAndHeight(ws,sourceRow,targetRow){
  ws.getRow(targetRow).height=ws.getRow(sourceRow).height;
  for(let c=1;c<=ws.columnCount;c++){
    const src=ws.getCell(sourceRow,c), dst=ws.getCell(targetRow,c);
    if(src.style) dst.style=JSON.parse(JSON.stringify(src.style));
    if(src.numFmt) dst.numFmt=src.numFmt;
    if(src.alignment) dst.alignment=JSON.parse(JSON.stringify(src.alignment));
    if(src.border) dst.border=JSON.parse(JSON.stringify(src.border));
    if(src.fill) dst.fill=JSON.parse(JSON.stringify(src.fill));
    if(src.font) dst.font=JSON.parse(JSON.stringify(src.font));
  }
}

function setPmCodeFill(cell,code){
  const colors={1:'92D050',2:'00B0F0',3:'0070C0',4:'275417',5:'002060'};
  cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF'+colors[code]}};
  cell.font={...(cell.font||{}),color:{argb:(code>=3?'FFFFFFFF':'FF000000')}};
}



// ===== STEP 19: Excel safety for inserted rows / shared formulas =====
function normalizeSharedFormulas(ws){
  ws.eachRow({includeEmpty:true}, row=>{
    row.eachCell({includeEmpty:true}, cell=>{
      try{
        const v=cell.value;
        // ExcelJS shared-formula clones can become invalid after row insertion.
        // cell.formula exposes the translated formula for the current cell.
        if(v && typeof v==='object' && (v.sharedFormula || v.formula)){
          const f=cell.formula || v.formula;
          if(f){
            const result=cell.result;
            cell.value={formula:f,result:result};
          }
        }
      }catch(_){}
    });
  });
}

function normalizeWorkbookFormulas(wb){
  for(const ws of wb.worksheets)normalizeSharedFormulas(ws);
}

function excelWriteBackupPath(finalPath){
  const base=path.basename(finalPath);
  let kind='other';
  let rel=base;
  const roots=[
    ['annual',ANNUAL_DIR],
    ['monthly',MONTHLY_WORK_DIR],
    ['monthly_templates',MONTHLY_TEMPLATE_DIR],
    ['system_templates',SYSTEM_TEMPLATE_DIR]
  ];
  for(const [k,root] of roots){
    const relative=path.relative(root,finalPath);
    if(relative && !relative.startsWith('..') && !path.isAbsolute(relative)){
      kind=k; rel=relative; break;
    }
    if(finalPath===root){kind=k;rel=base;break;}
  }
  const relDir=path.dirname(rel)==='.'?'':path.dirname(rel);
  const dir=path.join(BACKUP_DIR,'excel_writes',kind,relDir);
  fs.mkdirSync(dir,{recursive:true});
  return path.join(dir,`${base}.backup_before_write.xlsx`);
}

function migrateLegacyInlineExcelBackups(){
  const roots=[ANNUAL_DIR,MONTHLY_WORK_DIR,MONTHLY_TEMPLATE_DIR,SYSTEM_TEMPLATE_DIR];
  const walk=(dir)=>{
    if(!fs.existsSync(dir))return;
    for(const ent of fs.readdirSync(dir,{withFileTypes:true})){
      const full=path.join(dir,ent.name);
      if(ent.isDirectory()){walk(full);continue;}
      const m=ent.name.match(/^\.(.+\.xlsx)\.backup_before_write\.xlsx$/i);
      if(!m)continue;
      const original=path.join(dir,m[1]);
      const dest=excelWriteBackupPath(original);
      try{
        fs.mkdirSync(path.dirname(dest),{recursive:true});
        if(fs.existsSync(dest))fs.unlinkSync(dest);
        fs.renameSync(full,dest);
      }catch(_){
        try{fs.copyFileSync(full,dest);fs.unlinkSync(full);}catch(__){}
      }
    }
  };
  roots.forEach(walk);
}

migrateLegacyInlineExcelBackups();

async function safeWriteWorkbook(wb,finalPath){
  const dir=path.dirname(finalPath);
  const base=path.basename(finalPath);
  const tmp=path.join(dir,`.${base}.${process.pid}.${Date.now()}.tmp.xlsx`);
  const bak=excelWriteBackupPath(finalPath);

  normalizeWorkbookFormulas(wb);

  try{
    await wb.xlsx.writeFile(tmp);

    const st=fs.statSync(tmp);
    if(!st || st.size<1000)throw new Error('Temporary Excel file is empty or incomplete');

    // Verify that the temporary XLSX can be opened before touching the original.
    const verify=new ExcelJS.Workbook();
    await verify.xlsx.readFile(tmp);

    if(fs.existsSync(finalPath)){
      try{fs.copyFileSync(finalPath,bak);}catch(_){}
    }

    fs.renameSync(tmp,finalPath);

    // Keep only one last-known-good backup during testing.
    return {ok:true,backup:fs.existsSync(bak)?bak:null};
  }catch(e){
    try{if(fs.existsSync(tmp))fs.unlinkSync(tmp);}catch(_){}
    throw e;
  }
}

async function openAnnualForLineSafe(line,year){
  const a=await openAnnualForLine(line,year);
  normalizeWorkbookFormulas(a.wb);
  return a;
}

function naturalMachineKey(v){
  const m=String(v||'').trim().match(/^([^0-9]*)(\d+)(.*)$/);
  if(!m)return {prefix:String(v||'').toUpperCase(),num:999999,tail:''};
  return {prefix:m[1].toUpperCase(),num:Number(m[2]),tail:m[3].toUpperCase()};
}

function compareMachineNo(a,b){
  const A=naturalMachineKey(a),B=naturalMachineKey(b);
  if(A.prefix!==B.prefix)return A.prefix.localeCompare(B.prefix);
  if(A.num!==B.num)return A.num-B.num;
  return A.tail.localeCompare(B.tail);
}

function lastTaskActivityRow(ws){
  const start=pmDataStartRow(ws);
  let last=start-1;
  for(let r=start;r<=ws.rowCount;r++){
    const a=String(safeText(ws,r,8)||'').trim().toLowerCase();
    if(a==='planed'||a==='done')last=r;
  }
  return last;
}

function findOrderedInsertRow(ws,machineNo,insertAfterRow,taskPosition,typeValue){
  const tasks=taskRowsFromAnnual(ws);
  const wanted=String(machineNo||'').trim();
  const wantedType=String(typeValue||'').trim();
  const sameMachine=tasks.filter(t=>String(t.machineNo||'').trim()===wanted);
  const same= wantedType
    ? sameMachine.filter(t=>String(safeText(ws,t.row,5)||'').trim()===wantedType)
    : sameMachine;

  if(taskPosition){
    const p=Math.max(1,Number(taskPosition));
    if(p<=same.length)return same[p-1].row;
    if(same.length)return Math.max(...same.map(t=>t.row))+2;
  }

  if(insertAfterRow){
    const selected=tasks.find(t=>t.row===Number(insertAfterRow));
    if(selected && String(selected.machineNo||'').trim()===wanted)return selected.row+2;
  }

  if(same.length)return Math.max(...same.map(t=>t.row))+2;
  if(sameMachine.length)return Math.max(...sameMachine.map(t=>t.row))+2;

  const machineStarts=[];let seen=new Set();
  for(const t of tasks){
    const no=String(t.machineNo||'').trim();
    if(no&&!seen.has(no)){seen.add(no);machineStarts.push({machineNo:no,row:t.row});}
  }
  const next=machineStarts.find(x=>compareMachineNo(x.machineNo,wanted)>0);
  return next ? next.row : lastTaskActivityRow(ws)+1;
}

function setMonthlyRowFormulas(ws,planRow,doneRow){
  // Hidden helper area AL:AW.
  // AL = Done total, AM = Plan total
  // AN:AR = W1-W5 Plan
  // AS:AW = W1-W5 Done
  ws.getCell(planRow,39).value={formula:`COUNTIF(P${planRow}:T${planRow},">0")`}; // AM
  for(let i=0;i<5;i++){
    const src=String.fromCharCode(80+i); // P..T
    ws.getCell(planRow,40+i).value={formula:`COUNTIF(${src}${planRow},">0")`}; // AN..AR
  }
  ws.getCell(doneRow,38).value={formula:`COUNTIF(P${doneRow}:T${doneRow},">0")`}; // AL
  for(let i=0;i<5;i++){
    const src=String.fromCharCode(80+i);
    ws.getCell(doneRow,45+i).value={formula:`COUNTIF(${src}${doneRow},">0")`}; // AS..AW
  }
}

function rebuildMonthlyTaskFormulas(ws){
  const start=pmDataStartRow(ws);
  let last=start-1;
  for(let r=start;r<=ws.rowCount;r++){
    const a=String(safeText(ws,r,8)||'').trim().toLowerCase();
    if(a==='planed'){
      last=Math.max(last,r);
      const doneRow=r+1;
      setMonthlyRowFormulas(ws,r,doneRow);
      last=Math.max(last,doneRow);
    }
  }

  // Hidden helper totals (AL:AW) remain available for formulas.
  ws.getCell(3,38).value={formula:`SUM(AL${start}:AL${last})`};
  ws.getCell(3,39).value={formula:`SUM(AM${start}:AM${last})`};
  for(let i=0;i<5;i++){
    const pcol=40+i; // AN..AR
    const acol=45+i; // AS..AW
    const pletter=excelCol(pcol), aletter=excelCol(acol);
    ws.getCell(3,pcol).value={formula:`SUM(${pletter}${start}:${pletter}${last})`};
    ws.getCell(3,acol).value={formula:`SUM(${aletter}${start}:${aletter}${last})`};
  }
  // Professional dashboard.
  ws.getCell(3,1).value={formula:`IFERROR(SUMPRODUCT((B${start}:B${last}<>"")/COUNTIFS(B${start}:B${last},B${start}:B${last}&"")),0)`};
  ws.getCell(3,2).value={formula:`SUM(AM${start}:AM${last})`};
  ws.getCell(3,3).value={formula:`SUM(AL${start}:AL${last})`};
  ws.getCell(3,4).value={formula:'MAX(B3-C3,0)'};
  ws.getCell(3,5).value={formula:'IFERROR(C3/B3,0)'};

  // Monthly task classification KPI.
  // Count task definitions (planned rows), not monthly PM occurrences.
  ws.getCell(3,6).value={formula:`COUNTIFS(H${start}:H${last},"planed",L${start}:L${last},"Calendar")`};
  ws.getCell(3,7).value={formula:`COUNTIFS(H${start}:H${last},"planed",L${start}:L${last},"Running Hours")`};
  ws.getCell(3,8).value={formula:`COUNTIFS(H${start}:H${last},"planed",L${start}:L${last},"PLC Hours")`};

  for(let i=0;i<5;i++){
    const p=excelCol(40+i),a=excelCol(45+i),base=10+i*3;
    ws.getCell(3,base).value={formula:`SUM(${p}${start}:${p}${last})`};
    ws.getCell(3,base+1).value={formula:`SUM(${a}${start}:${a}${last})`};
    ws.getCell(3,base+2).value={formula:`IFERROR(${excelCol(base+1)}3/${excelCol(base)}3,0)`};
  }
}


function snapshotAndUnmerge(ws){
  const ranges=[];
  try{
    const merges=ws._merges||{};
    for(const key of Object.keys(merges)){
      const m=merges[key]?.model;
      if(m)ranges.push({top:m.top,left:m.left,bottom:m.bottom,right:m.right});
    }
    for(const r of ranges){
      try{ws.unMergeCells(r.top,r.left,r.bottom,r.right);}catch(_){}
    }
  }catch(_){}
  return ranges;
}

function restoreMergesAfterInsert(ws,ranges,insertAt,count){
  for(const r0 of ranges){
    const r={...r0};
    if(r.top>=insertAt){r.top+=count;r.bottom+=count;}
    else if(r.bottom>=insertAt){r.bottom+=count;}
    try{ws.mergeCells(r.top,r.left,r.bottom,r.right);}catch(_){}
  }
}

function insertStyledTaskPair(ws,insertAt,templatePlan){
  // Preserve each existing task identity BEFORE touching merges/rows.
  // Only planned rows are authoritative; merged slave rows must never become a new machine.
  const existingTasks=(String(ws.name||'').toLowerCase()==='full'?taskRowsFromAnnual(ws):taskRowsFromMonthly(ws))
    .map(t=>({row:t.row,section:t.section||'',machineNo:t.machineNo||'',machineName:t.machineName||'',machineIdentified:t.machineIdentified||'',type:t.type||''}));

  const templateDone=Math.min(templatePlan+1,ws.rowCount);
  const merges=snapshotAndUnmerge(ws);
  ws.spliceRows(insertAt,0,[],[]);

  // After insertion, the old template row moves down when it was at/after insertAt.
  const stylePlan=templatePlan>=insertAt?templatePlan+2:templatePlan;
  const styleDone=templateDone>=insertAt?templateDone+2:templateDone;
  copyRowStyleAndHeight(ws,stylePlan,insertAt);
  copyRowStyleAndHeight(ws,styleDone,insertAt+1);

  for(let c=1;c<=ws.columnCount;c++){
    ws.getCell(insertAt,c).value=null;
    ws.getCell(insertAt+1,c).value=null;
  }

  // Restore non-identity merges only. A:E are rebuilt from actual machine data later.
  const nonIdentity=merges.filter(r=>r.left>5);
  restoreMergesAfterInsert(ws,nonIdentity,insertAt,2);

  // Restore the exact identity of every task at its shifted row.
  for(const t of existingTasks){
    const r=t.row>=insertAt?t.row+2:t.row;
    ws.getCell(r,1).value=t.section||null;
    ws.getCell(r,2).value=t.machineNo||null;
    ws.getCell(r,3).value=t.machineName||null;
    ws.getCell(r,4).value=t.machineIdentified||null;
    ws.getCell(r,5).value=t.type||null;
    // Keep done-row identity empty; rebuildTaskGroupMerges will merge the proper blocks.
    for(let c=1;c<=5;c++)ws.getCell(r+1,c).value=null;
  }
  return {planRow:insertAt,doneRow:insertAt+1};
}


function clearCellFill(cell){
  cell.fill={type:'pattern',pattern:'none'};
}
function professionalFrequencyMeta(b={},code=1){
  const rawMethod=String(b.calculationMethod||'').trim().toLowerCase();
  const rawText=String(b.frequency||b.frequencyText||'').trim();
  const lower=rawText.toLowerCase();

  let mode='Calendar';
  if(rawMethod==='hours'||rawMethod==='running hours'||/\bhr\b|hour/.test(lower))mode='Running Hours';
  if(rawMethod==='plc_hours'||rawMethod==='plc hours')mode='PLC Hours';

  let value=null,text=rawText;
  if(mode==='Calendar'){
    const cf=String(b.calendarFrequency||'').trim().toLowerCase();
    const byCode={1:['Weekly',1],2:['Monthly',1],3:['Quarterly',3],4:['Semi Annual',6],5:['Annual',12]};
    const byName={
      weekly:['Weekly',1],monthly:['Monthly',1],quarterly:['Quarterly',3],
      semiannual:['Semi Annual',6],'semi annual':['Semi Annual',6],annual:['Annual',12]
    };
    const pair=byName[cf]||byCode[Number(code)]||['Weekly',1];
    if(!text)text=pair[0];
    value=pair[1];

    const monthMatch=lower.match(/(\d+)\s*month/);
    if(monthMatch)value=Number(monthMatch[1]);
  }else{
    const n=Number(b.intervalHours||String(rawText).match(/\d+(?:\.\d+)?/)?.[0]||0);
    value=n>0?n:null;
    if(!text&&value)text=`${value} hr`;
  }

  return {
    pmCode:Number(code)||null,
    frequencyValue:value,
    frequencyText:text,
    calculationMode:mode,
    plcTagCounter:String(b.plcTagCounter||b.plcTag||'').trim(),
    firstPmDate:String(b.firstPmDate||'').trim(),
    notes:String(b.notes||'').trim()
  };
}

function applyProfessionalTaskMeta(ws,planRow,doneRow,b,code){
  const m=professionalFrequencyMeta(b,code);
  ws.getCell(planRow,9).value=m.pmCode;
  ws.getCell(planRow,10).value=m.frequencyValue;
  ws.getCell(planRow,11).value=m.frequencyText;
  ws.getCell(planRow,12).value=m.calculationMode;
  ws.getCell(planRow,13).value=m.plcTagCounter||null;
  ws.getCell(planRow,14).value=m.firstPmDate||null;
  ws.getCell(planRow,15).value=m.notes||null;

  for(let c=9;c<=15;c++)ws.getCell(doneRow,c).value=null;
}

// Compatibility name used by older code paths.
// It now follows the Professional schema and never paints I:N as legacy frequency bands.
function applyFrequencyBand(ws,planRow,doneRow,code,b={}){
  applyProfessionalTaskMeta(ws,planRow,doneRow,b,code);
}

function writeTaskIdentity(ws,planRow,doneRow,b){
  ws.getCell(planRow,1).value=String(b.section||'').trim();
  ws.getCell(planRow,2).value=String(b.machineNo||'').trim();
  ws.getCell(planRow,3).value=String(b.machineName||'').trim();
  ws.getCell(planRow,4).value=String(b.machineIdentified||'').trim();
  ws.getCell(planRow,5).value=String(b.type||'').trim();
  ws.getCell(planRow,6).value=String(b.part||'').trim();
  ws.getCell(planRow,7).value=String(b.maintenance||'').trim();
  ws.getCell(planRow,8).value='planed';
  ws.getCell(doneRow,8).value='done';
}

function dateOnly(y,m,d){
  // Noon avoids timezone rollover; Excel number format shows date only.
  return new Date(Number(y),Number(m)-1,Number(d),12,0,0,0);
}



function captureEffectiveIdentityRows(ws){
  const rows=[],start=pmDataStartRow(ws);
  for(let r=start;r<=ws.rowCount;r++){
    const rt=String(safeText(ws,r,8)||'').trim().toLowerCase();
    if(rt!=='planed'&&rt!=='done')continue;
    rows.push({r,rowType:rt,
      section:String(safeText(ws,r,1)||'').trim(),
      machineNo:String(safeText(ws,r,2)||'').trim(),
      machineName:String(safeText(ws,r,3)||'').trim(),
      machineIdentified:String(safeText(ws,r,4)||'').trim(),
      type:String(safeText(ws,r,5)||'').trim()
    });
  }
  const keys=['section','machineNo','machineName','machineIdentified','type'];
  const cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
  for(const x of rows){for(const k of keys){if(x[k])cur[k]=x[k];else if(cur[k])x[k]=cur[k];}}
  for(const k of keys){
    const i=rows.findIndex(x=>x[k]);
    if(i>0){for(let n=0;n<i;n++)rows[n][k]=rows[i][k];}
  }
  return rows;
}
function materializeIdentityRows(ws,rows){
  for(const x of rows||[]){
    ws.getCell(x.r,1).value=x.section||null; ws.getCell(x.r,2).value=x.machineNo||null;
    ws.getCell(x.r,3).value=x.machineName||null; ws.getCell(x.r,4).value=x.machineIdentified||null;
    ws.getCell(x.r,5).value=x.type||null;
  }
}

function rebuildTaskGroupMerges(ws){
  const start=pmDataStartRow(ws),last=Math.max(start,ws.rowCount);
  const rows=captureEffectiveIdentityRows(ws);

  const ranges=[];
  if(ws._merges)for(const key of Object.keys(ws._merges)){
    const m=String(key).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if(!m)continue;
    const c1=colNumber(m[1]),r1=Number(m[2]),c2=colNumber(m[3]),r2=Number(m[4]);
    if(r2>=start&&r1<=last&&c1<=5&&c2>=1)ranges.push(key);
  }
  for(const rg of ranges)try{ws.unMergeCells(rg)}catch(_){}
  materializeIdentityRows(ws,rows);
  if(!rows.length)return;

  let i=0;
  while(i<rows.length){
    let j=i;while(j+1<rows.length&&rows[j+1].section===rows[i].section)j++;
    const r1=rows[i].r,r2=rows[j].r,v=rows[i].section;
    if(v){ws.getCell(r1,1).value=v;for(let k=i+1;k<=j;k++)ws.getCell(rows[k].r,1).value=null;if(r2>r1)try{ws.mergeCells(r1,1,r2,1)}catch(_){}}
    i=j+1;
  }

  i=0;
  while(i<rows.length){
    let j=i;const a=rows[i];
    while(j+1<rows.length){
      const b=rows[j+1];
      if(b.section!==a.section||b.machineNo!==a.machineNo||b.machineName!==a.machineName||b.machineIdentified!==a.machineIdentified)break;
      j++;
    }
    const vals=[a.machineNo,a.machineName,a.machineIdentified],r1=rows[i].r,r2=rows[j].r;
    for(let c=2;c<=4;c++){const v=vals[c-2];if(!v)continue;ws.getCell(r1,c).value=v;for(let k=i+1;k<=j;k++)ws.getCell(rows[k].r,c).value=null;if(r2>r1)try{ws.mergeCells(r1,c,r2,c)}catch(_){}}
    i=j+1;
  }

  i=0;
  while(i<rows.length){
    let j=i;const a=rows[i];
    while(j+1<rows.length){
      const b=rows[j+1];
      if(b.section!==a.section||b.machineNo!==a.machineNo||b.machineName!==a.machineName||b.machineIdentified!==a.machineIdentified||b.type!==a.type)break;
      j++;
    }
    const r1=rows[i].r,r2=rows[j].r,v=a.type;
    if(v){ws.getCell(r1,5).value=v;for(let k=i+1;k<=j;k++)ws.getCell(rows[k].r,5).value=null;if(r2>r1)try{ws.mergeCells(r1,5,r2,5)}catch(_){}}
    i=j+1;
  }

  // Planned/Done pair text is one visual task.
  try{mergeMigratedTaskPairs(ws,start,last)}catch(_){}

  // Keep Annual and Monthly visual formatting consistent after Add/Edit/Delete.
  const name=String(ws.name||'').toLowerCase();
  if(name==='full')formatProfessionalAnnualLayout(ws,start,last);
  else if(name==='sheet1')formatProfessionalMonthlyLayout(ws,start,last);
}

function firstFridayOnOrAfterDate(year,monthIndex,dayOfMonth){
  const lastDay=new Date(year,monthIndex+1,0,12).getDate();
  const targetDay=Math.min(Math.max(1,dayOfMonth),lastDay);
  const target=new Date(year,monthIndex,targetDay,12);
  const fridays=[];
  let d=new Date(year,monthIndex,1,12);
  while(d.getDay()!==5)d.setDate(d.getDate()+1);
  while(d.getMonth()===monthIndex){fridays.push(new Date(d));d.setDate(d.getDate()+7);}
  return fridays.find(x=>x.getDate()>=targetDay)||fridays[fridays.length-1]||null;
}

function annualFridays(year){
  const out=[];
  let d=new Date(year,0,1,12);
  while(d.getDay()!==5)d.setDate(d.getDate()+1);
  while(d.getFullYear()===year){out.push(new Date(d));d.setDate(d.getDate()+7);}
  return out;
}

function nearestAnnualWeekIndex(fridays,date){
  if(!fridays.length)return 0;
  let best=0,dist=Infinity;
  for(let i=0;i<fridays.length;i++){
    const x=Math.abs(fridays[i].getTime()-date.getTime());
    if(x<dist){dist=x;best=i;}
  }
  return best+1;
}

function generateAnnualSchedule(year,firstPmDate,opts={}){
  const [fy,fm,fd]=String(firstPmDate).split('-').map(Number);
  const anchor=dateOnly(fy,fm,fd);
  const fridays=annualFridays(year);
  const method=String(opts.calculationMethod||'calendar').toLowerCase();
  const result=[];

  if(method==='calendar'){
    const freq=String(opts.calendarFrequency||'weekly').toLowerCase();
    if(freq==='weekly'){
      // First PM date belongs to its PM Friday (Friday -> Thursday), then every week.
      const af=new Date(anchor);
      af.setDate(af.getDate()-((af.getDay()-5+7)%7));
      for(let i=0;i<fridays.length;i++)if(fridays[i].getTime()>=af.getTime())result.push(i+1);
      return {weeks:result,derivedCode:1,description:'Weekly'};
    }

    const intervalMonths=freq==='monthly'?1:freq==='quarterly'?3:freq==='semiannual'?6:12;
    // IMPORTANT: recurrence is based on CALENDAR MONTHS, never "every 4 weeks".
    // Preserve the anchor day-of-month as the due-day reference.
    for(let add=0;add<12;add+=intervalMonths){
      const monthIndex=(fm-1)+add;
      const y=year+Math.floor(monthIndex/12);
      const m=((monthIndex%12)+12)%12;
      if(y!==year)continue;
      const dueFriday=firstFridayOnOrAfterDate(year,m,fd);
      if(!dueFriday)continue;
      const wi=fridays.findIndex(x=>x.getTime()===dueFriday.getTime());
      if(wi>=0 && !result.includes(wi+1))result.push(wi+1);
    }
    const derivedCode=freq==='monthly'?2:freq==='quarterly'?3:freq==='semiannual'?4:5;
    return {weeks:result,derivedCode,description:freq};
  }

  if(method==='hours'||method==='plc_hours'){
    const intervalHours=Number(opts.intervalHours||0);
    const expectedWeeklyHours=Number(opts.expectedWeeklyHours||0);
    if(!(intervalHours>0))throw new Error('Enter Interval Hours');
    if(!(expectedWeeklyHours>0))throw new Error('Enter Expected Running Hours per Week');

    // For Running Hours the supplied date is the COUNTER START, not an
    // immediate PM due date.  Never create a false Week-1 PM.
    const calc=generateRunningHoursScheduleForYear(year,fmtIsoDateLocal(anchor),intervalHours,expectedWeeklyHours);
    return calc;
  }

  throw new Error('Invalid Calculation Method');
}

function taskPlanningStatePath(){return path.join(PM_MASTER_DIR,'task_planning_state.json');}
function readTaskPlanningState(){try{return JSON.parse(fs.readFileSync(taskPlanningStatePath(),'utf8'))}catch(_){return {}}}
function loadTaskPlanningMeta(line,year,row){return readTaskPlanningState()[`${line}|${year}|${row}`]||{};}
function writeTaskPlanningState(x){fs.writeFileSync(taskPlanningStatePath(),JSON.stringify(x,null,2),'utf8')}
function saveTaskPlanningMeta(line,year,row,meta){
  const x=readTaskPlanningState();
  x[`${line}|${year}|${row}`]={...meta,updatedAt:new Date().toISOString()};
  writeTaskPlanningState(x);
}
function shiftTaskPlanningMetaAfterInsert(line,year,insertRow){
  const x=readTaskPlanningState(),out={};
  for(const [k,v] of Object.entries(x)){
    const p=k.split('|');
    if(p.length===3&&p[0]===line&&Number(p[1])===year){
      const r=Number(p[2]);out[`${line}|${year}|${r>=insertRow?r+2:r}`]=v;
    }else out[k]=v;
  }
  writeTaskPlanningState(out);
}

app.post('/api/task-management/:line/:year',requireAuth,requireAdmin,async(req,res)=>{
 try{
  const line=decodeURIComponent(req.params.line),year=Number(req.params.year),b=req.body||{};
  const code=Number(b.code),firstPmDate=String(b.firstPmDate||'');
  if(![1,2,3,4,5].includes(code))return res.status(400).json({error:'PM code must be 1 to 5'});
  if(!/^\d{4}-\d{2}-\d{2}$/.test(firstPmDate))return res.status(400).json({error:'Select First PM Date'});
  if(!String(b.machineNo||'').trim()||!String(b.maintenance||'').trim())return res.status(400).json({error:'Machine and Maintenance Task are required'});
  const machine=authoritativeMachineIdentity(line,b.machineNo);
  if(!machine)return res.status(400).json({error:`Machine ${b.machineNo} is not registered in Machine Management`});

  const [fy,fm,fd]=firstPmDate.split('-').map(Number);
  if(dateOnly(fy,fm,fd).getFullYear()!==year)return res.status(400).json({error:`First PM Date must be inside ${year}`});
  const calculationMethod=String(b.calculationMethod||'calendar');
  const calendarFrequency=String(b.calendarFrequency||(code===1?'weekly':code===2?'monthly':code===3?'quarterly':code===4?'semiannual':'annual'));
  const intervalHours=Number(b.intervalHours||0),useMachineExpectedHours=b.useMachineExpectedHours!==false;
  const expectedWeeklyHours=useMachineExpectedHours?machineExpectedWeeklyHours(line,machine.machineNo):Number(b.expectedWeeklyHours||0);
  const schedule=generateAnnualSchedule(year,firstPmDate,{calculationMethod,calendarFrequency,intervalHours,expectedWeeklyHours});
  const isHours=calculationMethod==='hours'||calculationMethod==='plc_hours';
  if(!schedule.weeks.length && !isHours)return res.status(400).json({error:'No annual PM week generated from First PM Date'});
  const effectiveCode=schedule.derivedCode||code;

  const store=await ensureSystemTasksBootstrapped(line,year);
  const taskId=taskLinkId();
  const item={
    taskId,line,year,machineNo:machine.machineNo,
    part:String(b.part||'').trim(),maintenance:String(b.maintenance||'').trim(),frequency:String(b.frequency||calendarFrequency||'').trim(),
    code,effectiveCode,calculationMethod,calendarFrequency,intervalHours,expectedWeeklyHours:useMachineExpectedHours?0:expectedWeeklyHours,useMachineExpectedHours,firstPmDate:isHours?(schedule.firstDueDate||''):firstPmDate,runningHoursCounterStartDate:isHours?firstPmDate:'',
    scheduleWeeks:schedule.weeks,notes:String(b.notes||'').trim(),active:true,
    createdAt:new Date().toISOString(),updatedAt:new Date().toISOString()
  };

  // Optional position within selected machine.
  const sameIndexes=[];
  store.tasks.forEach((t,i)=>{if(String(t.line||'').toLowerCase()===line.toLowerCase()&&Number(t.year)===year&&String(t.machineNo||'').toLowerCase()===machine.machineNo.toLowerCase()&&t.active!==false)sameIndexes.push(i)});
  if(sameIndexes.length)store.tasks.splice(sameIndexes[sameIndexes.length-1]+1,0,item);else store.tasks.push(item);
  saveSystemTasks(store);

  const rebuilt=await rebuildAnnualFromSystemMaster(line,year);

  let monthlyUpdated=false,monthlyReason=null;
  if(b.addCurrentMonthly){
    const syncMonth=Number(b.syncMonthlyMonth||0);
    if(syncMonth>=1&&syncMonth<=12){
      try{
        await createMonthlyForLine(line,year,syncMonth,{overwrite:true});
        monthlyUpdated=true;
      }catch(e){monthlyReason=e.message}
    }
  }
  audit('PM_TASK_ADD_MASTER',req.user,{line,year,taskId,machineNo:machine.machineNo,maintenance:item.maintenance,rebuilt,monthlyUpdated});
  res.json({ok:true,taskId,rebuilt,monthlyUpdated,monthlyReason});
 }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

app.post('/api/task-management/:line/:year/:row/deactivate',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year),t=resolveAnnual122(line,year,req.params.row,req.body?.taskId),store=loadSystemTasks();store.tasks.find(v=>v.taskId===t.taskId).isActive=false;saveSystemTasks(store);res.json({ok:true});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

app.post('/api/task-management/:line/:year/:row/activate',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year),t=resolveAnnual122(line,year,req.params.row,req.body?.taskId),store=loadSystemTasks();store.tasks.find(v=>v.taskId===t.taskId).isActive=true;saveSystemTasks(store);res.json({ok:true});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});



function cloneJsonSafe(v){
  if(v===undefined||v===null)return v;
  try{return JSON.parse(JSON.stringify(v));}catch(_){return v;}
}

function captureMachineBlockFormat(ws,task){
  const tasks=taskRowsFromAnnual(ws);
  const same=tasks.filter(t=>
    t.machineNo===task.machineNo &&
    t.machineName===task.machineName &&
    t.machineIdentified===task.machineIdentified &&
    t.section===task.section
  );
  if(!same.length)return null;

  const firstPlan=Math.min(...same.map(t=>t.row));
  const lastPlan=Math.max(...same.map(t=>t.row));
  const first=firstPlan,last=lastPlan+1;

  const rows=[];
  for(let r=first;r<=last;r++){
    const cells={};
    for(let c=1;c<=15;c++){
      const cell=ws.getCell(r,c);
      cells[c]={
        style:cloneJsonSafe(cell.style),
        numFmt:cell.numFmt,
        alignment:cloneJsonSafe(cell.alignment),
        border:cloneJsonSafe(cell.border),
        fill:cloneJsonSafe(cell.fill),
        font:cloneJsonSafe(cell.font),
        protection:cloneJsonSafe(cell.protection)
      };
    }
    rows.push({offset:r-first,height:ws.getRow(r).height,cells});
  }

  // Capture the visual master styles for A:E from the original machine/type block.
  const masterStyles={};
  for(let c=1;c<=5;c++){
    const cell=ws.getCell(first,c);
    masterStyles[c]={
      style:cloneJsonSafe(cell.style),
      numFmt:cell.numFmt,
      alignment:cloneJsonSafe(cell.alignment),
      border:cloneJsonSafe(cell.border),
      fill:cloneJsonSafe(cell.fill),
      font:cloneJsonSafe(cell.font),
      protection:cloneJsonSafe(cell.protection)
    };
  }

  return {first,last,rows,masterStyles};
}

function applyCellAppearance(cell,snap){
  if(!snap)return;
  if(snap.style)cell.style=cloneJsonSafe(snap.style);
  if(snap.numFmt!==undefined)cell.numFmt=snap.numFmt;
  if(snap.alignment)cell.alignment=cloneJsonSafe(snap.alignment);
  if(snap.border)cell.border=cloneJsonSafe(snap.border);
  if(snap.fill)cell.fill=cloneJsonSafe(snap.fill);
  if(snap.font)cell.font=cloneJsonSafe(snap.font);
  if(snap.protection)cell.protection=cloneJsonSafe(snap.protection);
}

function restoreMachineBlockFormatAfterDelete(ws,task,snap,deletedRow){
  if(!snap)return;
  const tasks=taskRowsFromAnnual(ws);
  const same=tasks.filter(t=>
    t.machineNo===task.machineNo &&
    t.machineName===task.machineName &&
    t.machineIdentified===task.machineIdentified &&
    t.section===task.section
  );
  if(!same.length)return;

  const newFirst=Math.min(...same.map(t=>t.row));
  const newLast=Math.max(...same.map(t=>t.row))+1;

  // Restore original row heights for surviving rows.
  for(let r=newFirst;r<=newLast;r++){
    const originalRow = r>=deletedRow ? r+2 : r;
    const snapRow=snap.rows.find(x=>snap.first+x.offset===originalRow);
    if(snapRow && snapRow.height!==undefined) ws.getRow(r).height=snapRow.height;
  }

  // Restore A:E visual appearance for the whole surviving machine block.
  for(let r=newFirst;r<=newLast;r++){
    for(let c=1;c<=5;c++) applyCellAppearance(ws.getCell(r,c),snap.masterStyles[c]);
  }

  // Restore F:O appearance row-by-row so Parts/Maintenance/frequency borders and sizing stay original.
  for(let r=newFirst;r<=newLast;r++){
    const originalRow = r>=deletedRow ? r+2 : r;
    const snapRow=snap.rows.find(x=>snap.first+x.offset===originalRow);
    if(!snapRow)continue;
    for(let c=6;c<=15;c++) applyCellAppearance(ws.getCell(r,c),snapRow.cells[c]);
  }

  rebuildTaskGroupMerges(ws);
}

function taskRowsForWorksheet(ws){
  return String(ws?.name||'').toLowerCase()==='sheet1'
    ? taskRowsFromMonthly(ws)
    : taskRowsFromAnnual(ws);
}

function captureTaskIdentitiesBeforeDelete(ws){
  // IMPORTANT:
  // Never read A:E directly from a planned row here.
  // A:E can be merged across several task pairs, so continuation rows are
  // physically blank even though they belong to the same machine.
  // taskRowsFromAnnual/taskRowsFromMonthly already resolve the effective
  // machine identity, so use those values as the authoritative snapshot.
  return taskRowsForWorksheet(ws).map(t=>({
    row:t.row,
    section:String(t.section||'').trim(),
    machineNo:String(t.machineNo||'').trim(),
    machineName:String(t.machineName||'').trim(),
    machineIdentified:String(t.machineIdentified||'').trim(),
    type:String(t.type||'').trim()
  }));
}
function restoreTaskIdentitiesAfterDelete(ws,before,deletedRow){
  for(const x of before){
    if(x.row===deletedRow)continue;
    const nr=x.row>deletedRow?x.row-2:x.row;
    if(String(safeText(ws,nr,8)||'').trim().toLowerCase()!=='planed')continue;
    for(const [c,v] of [[1,x.section],[2,x.machineNo],[3,x.machineName],[4,x.machineIdentified],[5,x.type]])
      ws.getCell(nr,c).value=v||null;
    if(String(safeText(ws,nr+1,8)||'').trim().toLowerCase()==='done'){
      for(const [c,v] of [[1,x.section],[2,x.machineNo],[3,x.machineName],[4,x.machineIdentified],[5,x.type]])
        ws.getCell(nr+1,c).value=v||null;
    }
  }
}


function sameTaskIdentity(a,b){
  return String(a.machineNo||'').trim()===String(b.machineNo||'').trim() &&
         String(a.machineName||'').trim()===String(b.machineName||'').trim() &&
         String(a.machineIdentified||'').trim()===String(b.machineIdentified||'').trim() &&
         String(a.type||'').trim()===String(b.type||'').trim() &&
         String(a.part||'').trim()===String(b.part||'').trim() &&
         String(a.maintenance||'').trim()===String(b.maintenance||'').trim();
}


function restoreSurvivingTaskMachinesFromMaster(ws,line,before,deletedRow){
  // A:E in Excel are presentation only. After a delete, reconstruct every surviving
  // task pair from the Machine Management master, never from merged Excel slave cells.
  const survivors=(before||[]).filter(x=>x.row!==deletedRow);
  for(const x of survivors){
    const nr=x.row>deletedRow?x.row-2:x.row;
    if(String(safeText(ws,nr,8)||'').trim().toLowerCase()!=='planed')continue;

    const master=authoritativeMachineIdentity(line,x.machineNo) || {
      section:x.section||'',
      machineNo:x.machineNo||'',
      machineName:x.machineName||'',
      machineIdentified:x.machineIdentified||'',
      type:x.type||''
    };

    for(const r of [nr,nr+1]){
      ws.getCell(r,1).value=master.section||null;
      ws.getCell(r,2).value=master.machineNo||null;
      ws.getCell(r,3).value=master.machineName||null;
      ws.getCell(r,4).value=master.machineIdentified||null;
      ws.getCell(r,5).value=master.type||null;
    }
  }
}

function restoreNonIdentityMergesAfterDelete(ws,merges,deletedRow,count=2){
  // Never restore old A:E merges after deleting rows.
  // They may still point to the deleted merge-master row and erase the surviving identity.
  // A:E will be rebuilt cleanly from Machine Management.
  for(const r0 of (merges||[])){
    if(r0.left<=5)continue;
    const r={...r0};
    if(r.bottom<deletedRow){}
    else if(r.top>=deletedRow+count){r.top-=count;r.bottom-=count;}
    else if(r.top<deletedRow&&r.bottom>=deletedRow+count){r.bottom-=count;}
    else if(r.top<deletedRow&&r.bottom>=deletedRow){r.bottom=Math.max(r.top,r.bottom-count);}
    else continue;
    try{ws.mergeCells(r.top,r.left,r.bottom,r.right);}catch(_){}
  }
}

async function deleteTaskPairFromWorkbook(ws,task,line=''){
  const matches=taskRowsForWorksheet(ws).filter(t=>sameTaskIdentity(t,task));
  if(!matches.length)return {deleted:false};

  // Normally the identity is unique. If duplicated, delete the first exact match only.
  const row=matches[0].row;
  backfillStableTaskLinks(ws);
  const identities=captureTaskIdentitiesBeforeDelete(ws);
  const machineFormat=captureMachineBlockFormat(ws,matches[0]);
  const merges=snapshotAndUnmerge(ws);

  ws.spliceRows(row,2);

  // Restore only F:... merges. A:E machine merges are rebuilt from master data.
  restoreNonIdentityMergesAfterDelete(ws,merges,row,2);
  if(line)rebuildVisibleMachineIdentityFromStableLinks(ws,line);
  else rebuildTaskGroupMerges(ws);
  restoreMachineBlockFormatAfterDelete(ws,matches[0],machineFormat,row);

  // TRUE separation: if the deleted task was the final task for this machine,
  // keep the machine itself as a machine-only row.
  if(line){
    const stillHasTasks=taskRowsForWorksheet(ws).some(t=>
      String(t.machineNo||'').toLowerCase()===String(matches[0].machineNo||'').toLowerCase()
    );
    if(!stillHasTasks){
      const masterMachine=machineFromMaster(line,matches[0].machineNo);
      if(masterMachine)ensureMachinePlaceholderInWorksheet(ws,masterMachine);
    }
  }

  // Monthly workbook needs its formulas/dashboard ranges repaired after row deletion.
  try{rebuildMonthlyTaskFormulas(ws);}catch(_){}
  const last=Math.max(pmDataStartRow(ws),lastLogicalDataRow(ws));
  if(String(ws.name||'').toLowerCase()==='sheet1')formatProfessionalMonthlyLayout(ws,pmDataStartRow(ws),last);
  else formatProfessionalAnnualLayout(ws,pmDataStartRow(ws),last);

  return {deleted:true,row};
}

app.delete('/api/task-management/:line/:year/:row',requireAuth,requireAdmin,async(req,res)=>{
 try{
const {line}=req.params,year=Number(req.params.year);const t=resolveAnnual122(line,year,req.params.row,req.body?.taskId);const store=loadSystemTasks();store.tasks=store.tasks.filter(v=>v.taskId!==t.taskId);saveSystemTasks(store);await rebuildAnnualFromSystemMaster(line,year);res.json({ok:true,annualDeleted:true});
 }catch(e){console.error(e);res.status(e.status||500).json({error:e.message});}
});

// ---------- AUTOMATIC MONTHLY EXCEL CREATOR ----------
// V1.2.1 CLEAN: no built-in production lines. Lines are created only by the user.
const PM_LINES = [];

function fridaysInMonth(year,month){
  const out=[];
  const d=new Date(year,month-1,1,12);
  while(d.getDay()!==5) d.setDate(d.getDate()+1);
  while(d.getMonth()===month-1){
    out.push(new Date(d));
    d.setDate(d.getDate()+7);
  }
  return out;
}
function annualColumnForFriday(friday){
  const first=new Date(friday.getFullYear(),0,1,12);
  while(first.getDay()!==5) first.setDate(first.getDate()+1);
  const idx=Math.round((friday-first)/604800000);
  return 16+idx;
}
function monthShort(month){
  return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][month-1];
}

function availableMonthlyFiles(line,year){
  const out=[];
  for(let month=1;month<=12;month++){
    const fullPath=findExistingMonthlyFile(line,year,month);
    if(fs.existsSync(fullPath)){
      out.push({year,month,file:path.basename(fullPath),fullPath});
    }
  }
  return out;
}

function monthlyLineFolder(line){
  const safe=String(line||'').replace(/[<>:"/\\|?*]+/g,'_').trim()||'Unknown Line';
  return path.join(MONTHLY_WORK_DIR,safe);
}

function monthlyFilePath(line,year,month){
  return path.join(monthlyLineFolder(line),monthlyFileName(line,year,month));
}

function legacyMonthlyFilePath(line,year,month){
  return findExistingMonthlyFile(line,year,month);
}

function findExistingMonthlyFile(line,year,month){
  // Final storage rule: Monthly files live ONLY inside their production-line folder.
  return monthlyFilePath(line,year,month);
}

function listMonthlyExcelFiles(){
  const out=[];
  if(!fs.existsSync(MONTHLY_WORK_DIR))return out;

  function walk(dir){
    for(const ent of fs.readdirSync(dir,{withFileTypes:true})){
      const full=path.join(dir,ent.name);
      if(ent.isDirectory()){
        walk(full);
      }else if(ent.isFile() && /\.xlsx$/i.test(ent.name) && !/^~\$/.test(ent.name)){
        out.push(full);
      }
    }
  }

  walk(MONTHLY_WORK_DIR);
  return out;
}

function organizeExistingMonthlyFilesByLine(){
  fs.mkdirSync(MONTHLY_WORK_DIR,{recursive:true});

  // Ensure configured production-line folders exist.
  const cfg=readJson('lines.json',[]);
  const lines=Array.isArray(cfg) ? cfg : (Array.isArray(cfg.lines)?cfg.lines:[]);
  const created=[];

  for(const item of lines){
    const line=typeof item==='string' ? item : (item.name||item.line||'');
    if(!line)continue;
    const folder=monthlyLineFolder(line);
    if(!fs.existsSync(folder)){
      fs.mkdirSync(folder,{recursive:true});
      created.push(line);
    }
  }

  return {created};
}

function monthlyFileName(line,year,month){
  return `${line} PM_${year}_${String(month).padStart(2,'0')}.xlsx`;
}
async function createMonthlyForLine(line,year,month,{overwrite=false,preserveExecution=true}={}){
  const outFile=monthlyFileName(line,year,month);
  const outPath=monthlyFilePath(line,year,month);
  if(fs.existsSync(outPath) && !overwrite)return {line,file:outFile,status:'exists'};

  await ensureSystemTasksBootstrapped(line,year);

  const machines=orderedMasterMachines(line);
  const period=await monthlyState122(line,year,month);
  const tasks=monthlyTasks122(line,year,month,period);
  period.renderedTasks=tasks;period.machines=machines;saveMonthlyState122(line,year,month,period);
  const fridays=fridaysInMonth(year,month);
  if(fridays.length<4||fridays.length>5)throw new Error(`Unexpected PM week count ${fridays.length}`);

  const executionByTaskId=new Map(Object.entries(preserveExecution?period.execution:{}));

  const built=await buildCanonicalMonthlyTemplate(line,year,month);
  const wb=built.wb,ws=built.ws;
  const startRow=pmDataStartRow(ws);

  const oldDataMerges=[];
  if(ws._merges){
    for(const key of Object.keys(ws._merges)){
      const m=ws._merges[key]?.model;
      if(!m)continue;
      if(m.bottom>=startRow && m.left<=7 && m.right>=1)oldDataMerges.push(key);
    }
  }
  for(const rg of oldDataMerges)try{ws.unMergeCells(rg)}catch(_){}

  const clearLast=Math.max(ws.rowCount,startRow+500);
  for(let r=startRow;r<=clearLast;r++){
    for(let c=1;c<=51;c++)ws.getCell(r,c).value=null;
    ws.getRow(r).height=18;
  }

  ws.getCell(3,25).value=`${line} — ${monthShort(month)} ${year}`;
  try{ws.unMergeCells('P4:T4')}catch(_){}
  try{ws.unMergeCells('U4:Y4')}catch(_){}
  ws.mergeCells('P4:T4');
  ws.getCell(4,16).value='PM WEEK SCHEDULE';
  ws.mergeCells('U4:Y4');
  ws.getCell(4,21).value='Executed By / Execution Date';

  for(let i=0;i<5;i++){
    ws.getCell(5,16+i).value='';
    ws.getCell(5,21+i).value='';
  }
  for(let i=0;i<fridays.length;i++){
    const friday=fridays[i],thursday=new Date(friday);
    thursday.setDate(thursday.getDate()+6);
    const header=`${friday.getDate()}--${thursday.getDate()}`;
    ws.getCell(5,16+i).value=header;
    ws.getCell(5,21+i).value=header;
  }

  const annualWeekToMonthlyIndex=new Map();
  for(let i=0;i<fridays.length;i++){
    const annualCol=annualColumnForFriday(fridays[i]);
    annualWeekToMonthlyIndex.set(annualCol-15,i+1);
  }

  let r=startRow;
  const machineBlocks=[];

  for(const m of machines){
    const machineTasks=tasks.filter(t=>
      String(t.machineNo||'').toLowerCase()===String(m.machineNo||'').toLowerCase()
    );

    const machineStart=r;

    if(!machineTasks.length){
      writeMachinePlaceholder(ws,r,m);
      machineBlocks.push({
        start:r,end:r,
        section:String(m.section||'').trim(),
        machineNo:String(m.machineNo||'').trim(),
        machineName:String(m.machineName||'').trim(),
        machineIdentified:String(m.machineIdentified||'').trim(),
        type:String(m.type||'').trim(),
        taskCount:0
      });
      r+=1;
      continue;
    }

    for(const t of machineTasks){
      const planRow=r,doneRow=r+1;

      for(const rr of [planRow,doneRow]){
        ws.getCell(rr,1).value=String(m.section||'').trim()||null;
        ws.getCell(rr,2).value=String(m.machineNo||'').trim()||null;
        ws.getCell(rr,3).value=String(m.machineName||'').trim()||null;
        ws.getCell(rr,4).value=String(m.machineIdentified||'').trim()||null;
        ws.getCell(rr,5).value=String(m.type||'').trim()||null;
      }

      ws.getCell(planRow,6).value=String(t.part||'').trim()||null;
      ws.getCell(planRow,7).value=String(t.maintenance||'').trim()||null;
      ws.getCell(planRow,8).value='planed';

      ws.getCell(doneRow,6).value=null;
      ws.getCell(doneRow,7).value=null;
      ws.getCell(doneRow,8).value='done';

      applyProfessionalTaskMeta(ws,planRow,doneRow,{
        frequency:t.frequency||'',
        calculationMethod:t.calculationMethod||'Calendar',
        calendarFrequency:t.calendarFrequency||t.frequency||'',
        intervalHours:Number(t.intervalHours||0),
        expectedWeeklyHours:effectiveTaskExpectedWeeklyHours(t),
        firstPmDate:t.firstPmDate||'',
        notes:t.notes||''
      },Number(t.effectiveCode||t.code)||1);

      for(let c=16;c<=20;c++){
        ws.getCell(planRow,c).value=0;
        ws.getCell(doneRow,c).value=0;
        clearCellFill(ws.getCell(planRow,c));
        clearCellFill(ws.getCell(doneRow,c));
      }
      for(let c=21;c<=25;c++){
        ws.getCell(planRow,c).value='';
        ws.getCell(doneRow,c).value='';
      }

      for(const annualWeek of (Array.isArray(t.scheduleWeeks)?t.scheduleWeeks:[])){
        const monthlyIndex=annualWeekToMonthlyIndex.get(Number(annualWeek));
        if(!monthlyIndex)continue;
        const c=15+monthlyIndex;
        const code=Number(t.scheduleCodes?.[annualWeek]||t.effectiveCode||t.code)||1;
        ws.getCell(planRow,c).value=code;
        setPmCodeFill(ws.getCell(planRow,c),code);
      }

      const old=executionByTaskId.get(String(t.taskId||''));
      if(old){
        for(let i=0;i<5;i++){
          ws.getCell(doneRow,16+i).value=Number(old.done?.[i])||0;
          ws.getCell(planRow,21+i).value=old.user?.[i]??'';
          ws.getCell(doneRow,21+i).value=old.date?.[i]?new Date(old.date[i]):'';
        }
      }

      writeStableTaskLink(ws,planRow,doneRow,m.machineNo,t.taskId);

      try{ws.mergeCells(planRow,6,doneRow,6)}catch(_){}
      try{ws.mergeCells(planRow,7,doneRow,7)}catch(_){}

      r+=2;
    }

    machineBlocks.push({
      start:machineStart,end:r-1,
      section:String(m.section||'').trim(),
      machineNo:String(m.machineNo||'').trim(),
      machineName:String(m.machineName||'').trim(),
      machineIdentified:String(m.machineIdentified||'').trim(),
      type:String(m.type||'').trim(),
      taskCount:machineTasks.length
    });
  }

  const last=Math.max(startRow,r-1);

  for(const b of machineBlocks){
    if(b.taskCount<=0)continue;
    for(let c=2;c<=5;c++){
      const value=c===2?b.machineNo:c===3?b.machineName:c===4?b.machineIdentified:b.type;
      if(b.end>b.start){
        try{ws.mergeCells(b.start,c,b.end,c)}catch(_){}
      }
      ws.getCell(b.start,c).value=value||null;
    }
  }

  let i=0;
  while(i<machineBlocks.length){
    let j=i;
    while(j+1<machineBlocks.length &&
          String(machineBlocks[j+1].section||'')===String(machineBlocks[i].section||'')){
      j++;
    }
    const top=machineBlocks[i].start,bottom=machineBlocks[j].end;
    const section=machineBlocks[i].section;
    if(section){
      if(bottom>top){
        try{ws.mergeCells(top,1,bottom,1)}catch(_){}
      }
      ws.getCell(top,1).value=section;
    }
    i=j+1;
  }

  for(const b of machineBlocks){
    if(b.taskCount!==0)continue;
    ws.getCell(b.start,1).value=b.section||null;
    ws.getCell(b.start,2).value=b.machineNo||null;
    ws.getCell(b.start,3).value=b.machineName||null;
    ws.getCell(b.start,4).value=b.machineIdentified||null;
    ws.getCell(b.start,5).value=b.type||null;
    ws.getCell(b.start,8).value='machine';
    writeStableMachineOnlyLink(ws,b.start,b.machineNo);
  }

  rebuildMonthlyTaskFormulas(ws);
  hideMonthlyHelperColumns(ws);
  restoreMonthlyKpiHeaders(ws);
  formatProfessionalMonthlyLayout(ws,startRow,last);
  normalizeMonthlyScheduleDisplay(ws,startRow,last);
  forceExcelRecalculation(wb);

  for(const b of machineBlocks){
    const vals=[
      String(safeText(ws,b.start,1)||'').trim(),
      String(safeText(ws,b.start,2)||'').trim(),
      String(safeText(ws,b.start,3)||'').trim(),
      String(safeText(ws,b.start,4)||'').trim(),
      String(safeText(ws,b.start,5)||'').trim()
    ];
    const expected=[b.section,b.machineNo,b.machineName,b.machineIdentified,b.type].map(x=>String(x||'').trim());
    if(vals.join('|')!==expected.join('|')){
      throw new Error(`Monthly render validation failed for machine ${b.machineNo}: expected ${expected.join(' | ')} but found ${vals.join(' | ')}`);
    }
  }

  fs.mkdirSync(path.dirname(outPath),{recursive:true});
  await safeWriteWorkbook(wb,outPath);

  return {
    line,file:outFile,status:'written',
    machines:machines.length,tasks:tasks.length,lastRow:last,
    mode:'DETERMINISTIC_SYSTEM_MASTER_RENDER'
  };
}

async function ensureMonthlyFiles(year,month,{overwrite=false}={}){
  const results=[];
  for(const line of configuredPmLines()){
    try{ results.push(await createMonthlyForLine(line,year,month,{overwrite})); }
    catch(e){ results.push({line,status:'error',error:e.message}); }
  }
  return results;
}

app.post('/api/monthly/create',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const year=Number(req.body?.year);
    const month=Number(req.body?.month);
    if(year<2020 || year>2100 || month<1 || month>12) return res.status(400).json({error:'Invalid year or month'});
    const results=await ensureMonthlyFiles(year,month,{overwrite:!!req.body?.overwrite});
    if(req.body?.activate){
      const sys=readJson('system.json',{});
      const fsx=fridaysInMonth(year,month)[0];
      sys.activePmDate=`${year}-${String(month).padStart(2,'0')}-${String(fsx.getDate()).padStart(2,'0')}`;
      writeJson('system.json',sys);
      audit('PM_MONTH_ACTIVATED',req.user,{activePmDate:sys.activePmDate});
    }
    res.json({ok:true,year,month,weeks:fridaysInMonth(year,month).length,results});
  }catch(e){res.status(500).json({error:e.message});}
});

app.get('/api/monthly/status',requireAuth,requireAdmin,(req,res)=>{
  const sys=readJson('system.json',{});
  const d=getPmDate();
  res.json({
    activePmDate:sys.activePmDate||'',
    activeYear:d.getFullYear(),
    activeMonth:d.getMonth()+1,
    files:fs.existsSync(MONTHLY_WORK_DIR)?fs.readdirSync(MONTHLY_WORK_DIR).filter(f=>f.endsWith('.xlsx')):[]
  });
});


// ---------- AUTOMATIC WEEKLY / MONTHLY EXCEL EMAIL ----------
const EMAIL_STATE_FILE = 'email_state.json';

function zonedParts(date, timeZone='Asia/Riyadh'){
  const parts = new Intl.DateTimeFormat('en-CA',{
    timeZone, year:'numeric', month:'2-digit', day:'2-digit',
    hour:'2-digit', minute:'2-digit', second:'2-digit',
    hourCycle:'h23', weekday:'short'
  }).formatToParts(date);
  const o={};
  for(const p of parts) if(p.type!=='literal') o[p.type]=p.value;
  return o;
}
function parseHHMM(v, fallback){
  const m=String(v||fallback||'').match(/^(\d{1,2}):(\d{2})$/);
  return m ? {hour:Number(m[1]),minute:Number(m[2])} : {hour:0,minute:0};
}
function ymdLocalDate(year,month,day){
  return new Date(Number(year), Number(month)-1, Number(day), 12,0,0);
}
function fridayStartForDate(d){
  const x=new Date(d.getFullYear(),d.getMonth(),d.getDate(),12);
  while(x.getDay()!==5) x.setDate(x.getDate()-1);
  return x;
}
function pmContextForDate(d){
  const start=fridayStartForDate(d);
  const end=new Date(start); end.setDate(end.getDate()+6);
  return {
    start,end,
    weekHeader:`${start.getDate()}--${end.getDate()}`,
    year:start.getFullYear(),
    month:start.getMonth()+1,
    monthKey:`${start.getFullYear()}_${String(start.getMonth()+1).padStart(2,'0')}`
  };
}
function reportFilesForMonth(year,month){
  const mm=String(month).padStart(2,'0');
  if(!fs.existsSync(MONTHLY_WORK_DIR)) return [];
  return fs.readdirSync(MONTHLY_WORK_DIR)
    .filter(f=>f.toLowerCase().endsWith('.xlsx'))
    .filter(f=>new RegExp(`_${year}_${mm}\\.xlsx$`,'i').test(f))
    .map(f=>({filename:f,path:path.join(MONTHLY_WORK_DIR,f)}));
}
function splitRecipients(v){
  return String(v||'').split(/[;,]/).map(x=>x.trim()).filter(Boolean).join(',');
}
function applyVars(text, vars){
  return String(text||'').replace(/\{\{(company|week|month)\}\}/g,(_,k)=>vars[k]||'');
}
function getSendingAccounts(cfg){
  const enabled=(cfg.accounts||[]).filter(a=>a.enabled!==false && a.smtpHost && a.smtpUser && a.smtpPassword);
  const def=enabled.find(a=>a.isDefault);
  const ordered=def ? [def,...enabled.filter(a=>a!==def)] : enabled;
  return cfg.fallbackEnabled===false ? ordered.slice(0,1) : ordered;
}
async function sendReportEmail(kind, when=new Date(), options={}){
  const emailCfg=readJson('email.json',{accounts:[],fallbackEnabled:true,weekly:{},monthly:{}});
  const mailCfg=emailCfg[kind]||{};
  if(mailCfg.enabled===false && !options.force) throw new Error(`${kind} email is disabled`);

  const accounts=getSendingAccounts(emailCfg);
  if(!accounts.length) throw new Error('No enabled SMTP account is available');

  const ctx=pmContextForDate(when);
  const attachments=reportFilesForMonth(ctx.year,ctx.month);
  if(!attachments.length) throw new Error(`No monthly Excel files found for ${ctx.monthKey} in data/monthly_work`);

  const branding=readJson('branding.json',{});
  const company=branding.companyName || branding.factoryName || 'Smart PM Maintenance';
  const vars={
    company,
    week:`${fmtDate(ctx.start)} - ${fmtDate(ctx.end)}`,
    month:`${String(ctx.month).padStart(2,'0')}/${ctx.year}`
  };

  const to=splitRecipients(mailCfg.to);
  if(!to) throw new Error(`${kind} recipient (To) is empty`);

  const message={
    to,
    cc:splitRecipients(mailCfg.cc)||undefined,
    bcc:splitRecipients(mailCfg.bcc)||undefined,
    subject:applyVars(mailCfg.subject,vars),
    text:applyVars(mailCfg.message,vars),
    attachments
  };

  const errors=[];
  for(const account of accounts){
    try{
      const transporter=buildTransport(account);
      console.log(`[AUTO EMAIL] ${kind} verify: ${account.name||account.smtpUser}`);
      await transporter.verify();
      const info=await transporter.sendMail({
        from:`"${account.fromName||company}" <${account.smtpUser}>`,
        ...message
      });
      audit(kind==='weekly'?'WEEKLY_EXCEL_EMAIL_SENT':'MONTHLY_EXCEL_EMAIL_SENT',
        {name:'System',role:'System'},
        {account:account.name||account.smtpUser,to,files:attachments.map(a=>a.filename),messageId:info.messageId,pmMonth:ctx.monthKey,week:ctx.weekHeader}
      );
      return {ok:true,account:account.name||account.smtpUser,messageId:info.messageId,files:attachments.map(a=>a.filename),context:ctx};
    }catch(e){
      console.error(`[AUTO EMAIL] ${account.name||account.smtpUser} failed:`,e.message);
      errors.push(`${account.name||account.smtpUser}: ${e.message}`);
      audit('EMAIL_ACCOUNT_FAILED',{name:'System',role:'System'},{kind,account:account.name||account.smtpUser,error:e.message});
    }
  }
  throw new Error('All email accounts failed: '+errors.join(' | '));
}

function scheduleKey(kind,ctx){
  return kind==='weekly'
    ? `weekly:${ctx.year}_${String(ctx.month).padStart(2,'0')}:${ctx.weekHeader}`
    : `monthly:${ctx.year}_${String(ctx.month).padStart(2,'0')}`;
}

async function runFridayPrepareForDate(local,{test=false}={}){
  const ctx=pmContextForDate(local);
  const results=await ensureMonthlyFiles(ctx.year,ctx.month);
  const failed=results.filter(x=>x.status==='error');
  if(failed.length){
    audit(test?'TEST_FRIDAY_PREPARE_FAILED':'MONTHLY_AUTO_PREPARE_FAILED',
      {name:test?'Admin Test':'System',role:test?'Admin':'System'},
      {pmMonth:ctx.monthKey,failed});
    throw new Error(failed.map(x=>`${x.line}: ${x.error}`).join(' | '));
  }

  const sys2=readJson('system.json',{});
  sys2.activePmDate=`${ctx.year}-${String(ctx.month).padStart(2,'0')}-${String(ctx.start.getDate()).padStart(2,'0')}`;
  writeJson('system.json',sys2);

  audit(test?'TEST_FRIDAY_PREPARE':'MONTHLY_AUTO_PREPARED',
    {name:test?'Admin Test':'System',role:test?'Admin':'System'},
    {pmMonth:ctx.monthKey,results});

  return {ok:true,context:ctx,results};
}

async function runThursdayCloseForDate(local,{test=false}={}){
  const sys=readJson('system.json',{});
  const ctx=pmContextForDate(local);
  const result={ok:true,context:ctx,weekly:null,monthly:null};

  if(sys.weeklyEmailEnabled!==false){
    result.weekly=await sendReportEmail('weekly',local,{force:test});
  }else{
    result.weekly={skipped:true,reason:'Weekly Email is disabled'};
  }

  // Final PM Thursday: tomorrow is Friday in a new calendar month.
  const tomorrow=new Date(local); tomorrow.setDate(tomorrow.getDate()+1);
  const isMonthEnd=tomorrow.getDay()===5 && tomorrow.getMonth()!==ctx.start.getMonth();
  result.isPmMonthEnd=isMonthEnd;

  if(isMonthEnd){
    if(sys.monthlyEmailEnabled!==false){
      result.monthly=await sendReportEmail('monthly',local,{force:test});
    }else{
      result.monthly={skipped:true,reason:'Monthly Email is disabled'};
    }
  }

  audit(test?'TEST_THURSDAY_CLOSE':'THURSDAY_CLOSE',
    {name:test?'Admin Test':'System',role:test?'Admin':'System'},
    {week:ctx.weekHeader,pmMonth:ctx.monthKey,isPmMonthEnd:isMonthEnd,
     weekly:result.weekly?.skipped?'skipped':'sent',
     monthly:result.monthly?(result.monthly.skipped?'skipped':'sent'):'not month end'});

  return result;
}

async function schedulerTick(){
  try{
    const sys=readJson('system.json',{});
    const tz=sys.timezone||'Asia/Riyadh';
    const p=zonedParts(new Date(),tz);
    const close=parseHHMM(sys.weeklyCloseTime,'23:00');
    const prepare=parseHHMM(sys.newWeekPrepareTime,'00:15');

    // Every Friday at prepare time, ensure that the monthly workbook exists.
    // On the first Friday of a new month this creates the new month; on other Fridays it does nothing.
    if(p.weekday==='Fri' && Number(p.hour)===prepare.hour && Number(p.minute)===prepare.minute){
      const local=ymdLocalDate(p.year,p.month,p.day);
      const ctx=pmContextForDate(local);
      const prepKey=`prepare:${ctx.year}_${String(ctx.month).padStart(2,'0')}`;
      const state=readJson(EMAIL_STATE_FILE,{sent:{}});
      state.sent=state.sent||{};
      if(!state.sent[prepKey]){
        const results=await ensureMonthlyFiles(ctx.year,ctx.month);
        const failed=results.filter(x=>x.status==='error');
        if(!failed.length){
          state.sent[prepKey]={time:new Date().toISOString(),files:results.map(x=>x.file)};
          writeJson(EMAIL_STATE_FILE,state);
          const sys2=readJson('system.json',{});
          sys2.activePmDate=`${ctx.year}-${String(ctx.month).padStart(2,'0')}-${String(ctx.start.getDate()).padStart(2,'0')}`;
          writeJson('system.json',sys2);
          audit('MONTHLY_AUTO_PREPARED',{name:'System',role:'System'},{pmMonth:ctx.monthKey,results});
        }else{
          audit('MONTHLY_AUTO_PREPARE_FAILED',{name:'System',role:'System'},{pmMonth:ctx.monthKey,failed});
        }
      }
    }

    // Thursday at weekly close time.
    if(p.weekday==='Thu' && Number(p.hour)===close.hour && Number(p.minute)===close.minute){
      const local=ymdLocalDate(p.year,p.month,p.day);
      const ctx=pmContextForDate(local);
      const state=readJson(EMAIL_STATE_FILE,{sent:{}});
      state.sent=state.sent||{};

      const wk=scheduleKey('weekly',ctx);
      if(!state.sent[wk]){
        try{
          const r=await sendReportEmail('weekly',local);
          state.sent[wk]={time:new Date().toISOString(),files:r.files};
          writeJson(EMAIL_STATE_FILE,state);
        }catch(e){
          console.error('[SCHEDULER] Weekly email:',e.message);
          audit('WEEKLY_EXCEL_EMAIL_FAILED',{name:'System',role:'System'},{error:e.message,week:ctx.weekHeader});
        }
      }

      // This Thursday is the PM month end when tomorrow's Friday starts a new calendar month.
      const tomorrow=new Date(local); tomorrow.setDate(tomorrow.getDate()+1);
      if(tomorrow.getDay()===5 && tomorrow.getMonth()!==ctx.start.getMonth()){
        const mk=scheduleKey('monthly',ctx);
        const state2=readJson(EMAIL_STATE_FILE,{sent:{}});
        state2.sent=state2.sent||{};
        if(!state2.sent[mk]){
          try{
            const r=await sendReportEmail('monthly',local);
            state2.sent[mk]={time:new Date().toISOString(),files:r.files};
            writeJson(EMAIL_STATE_FILE,state2);
          }catch(e){
            console.error('[SCHEDULER] Monthly email:',e.message);
            audit('MONTHLY_EXCEL_EMAIL_FAILED',{name:'System',role:'System'},{error:e.message,pmMonth:ctx.monthKey});
          }
        }
      }
    }
  }catch(e){ console.error('[SCHEDULER]',e); }
}

app.get('/api/email/status',requireAuth,requireAdmin,(req,res)=>{
  const sys=readJson('system.json',{});
  res.json({
    timezone:sys.timezone||'Asia/Riyadh',
    weeklyCloseTime:sys.weeklyCloseTime||'23:00',
    state:readJson(EMAIL_STATE_FILE,{sent:{}})
  });
});
app.post('/api/email/send-weekly-now',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const when=req.body?.date ? new Date(req.body.date+'T12:00:00') : new Date();
    const r=await sendReportEmail('weekly',when,{force:true});
    res.json(r);
  }catch(e){res.status(500).json({error:e.message});}
});
app.post('/api/email/send-monthly-now',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const when=req.body?.date ? new Date(req.body.date+'T12:00:00') : new Date();
    const r=await sendReportEmail('monthly',when,{force:true});
    res.json(r);
  }catch(e){res.status(500).json({error:e.message});}
});


app.post('/api/system/test-friday-prepare',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const raw=String(req.body?.date||'').trim();
    if(!/^\d{4}-\d{2}-\d{2}$/.test(raw))return res.status(400).json({error:'Select a test date'});
    const [y,m,d]=raw.split('-').map(Number);
    const local=ymdLocalDate(y,m,d);
    if(local.getDay()!==5)return res.status(400).json({error:'Friday Prepare test date must be a Friday'});
    const r=await runFridayPrepareForDate(local,{test:true});
    res.json(r);
  }catch(e){res.status(500).json({error:e.message});}
});

app.post('/api/system/test-thursday-close',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const raw=String(req.body?.date||'').trim();
    if(!/^\d{4}-\d{2}-\d{2}$/.test(raw))return res.status(400).json({error:'Select a test date'});
    const [y,m,d]=raw.split('-').map(Number);
    const local=ymdLocalDate(y,m,d);
    if(local.getDay()!==4)return res.status(400).json({error:'Thursday Close test date must be a Thursday'});
    const r=await runThursdayCloseForDate(local,{test:true});
    res.json(r);
  }catch(e){res.status(500).json({error:e.message});}
});

if(!process.env.PM_TEST_MODE){setInterval(schedulerTick,30000);setTimeout(schedulerTick,3000);}

// ---------- SETTINGS API ----------
app.get('/api/settings/summary', requireAuth, requireAdmin, (req,res)=>{
  res.json({
    users:readJson('users.json',[]).length,
    lines:readJson('lines.json',{lines:[]}).lines.length,
    emailEnabled:(readJson('email.json',{accounts:[]}).accounts||[]).some(a=>a.enabled!==false),
    branding:readJson('branding.json',{}),
    system:readJson('system.json',{})
  });
});

app.get('/api/settings/users', requireAuth, requireAdmin, (req,res)=>{
  res.json(readJson('users.json',[]).map(({password,...u})=>u));
});
app.post('/api/settings/users', requireAuth, requireAdmin, (req,res)=>{
  const users=readJson('users.json',[]);
  const b=req.body||{};
  if(!b.username||!b.password||!b.name||!b.role) return res.status(400).json({error:'Name, username, password and role are required'});
  if(users.some(u=>u.username.toLowerCase()===String(b.username).toLowerCase())) return res.status(409).json({error:'Username already exists'});
  const user={id:Date.now(),username:String(b.username).trim(),password:String(b.password),name:String(b.name).trim(),role:b.role,active:true};
  users.push(user); writeJson('users.json',users); audit('USER_ADD',req.user,{username:user.username,name:user.name,role:user.role});
  res.json({ok:true,user:{id:user.id,username:user.username,name:user.name,role:user.role,active:user.active}});
});
app.put('/api/settings/users/:id', requireAuth, requireAdmin, (req,res)=>{
  const users=readJson('users.json',[]); const i=users.findIndex(u=>String(u.id)===String(req.params.id));
  if(i<0) return res.status(404).json({error:'User not found'});
  const b=req.body||{};
  const nextUsername = b.username!==undefined ? String(b.username).trim() : users[i].username;
  const nextName = b.name!==undefined ? String(b.name).trim() : users[i].name;
  const nextRole = b.role!==undefined ? String(b.role) : users[i].role;
  if(!nextUsername || !nextName || !nextRole) return res.status(400).json({error:'Name, username and role are required'});
  if(users.some((u,idx)=>idx!==i && String(u.username).toLowerCase()===nextUsername.toLowerCase())) return res.status(409).json({error:'Username already exists'});
  users[i].username=nextUsername; users[i].name=nextName; users[i].role=nextRole;
  if(b.active!==undefined) users[i].active=!!b.active;
  if(b.password) users[i].password=String(b.password);
  writeJson('users.json',users); audit('USER_UPDATE',req.user,{id:users[i].id,username:users[i].username,name:users[i].name,role:users[i].role});
  res.json({ok:true,user:{id:users[i].id,username:users[i].username,name:users[i].name,role:users[i].role,active:users[i].active!==false}});
});
app.delete('/api/settings/users/:id', requireAuth, requireAdmin, (req,res)=>{
  let users=readJson('users.json',[]); const target=users.find(u=>String(u.id)===String(req.params.id));
  if(!target) return res.status(404).json({error:'User not found'});
  if(target.username==='admin') return res.status(400).json({error:'Default admin cannot be deleted in this test'});
  users=users.filter(u=>String(u.id)!==String(req.params.id)); writeJson('users.json',users); audit('USER_DELETE',req.user,{username:target.username});
  res.json({ok:true});
});


app.post('/api/settings/branding/logo/:slot', requireAuth, requireAdmin,express.raw({type:['image/png','image/jpeg','image/jpg','application/octet-stream'],limit:'5mb'}),async(req,res)=>{try{const slotNo=Number(req.params.slot);if(!Number.isInteger(slotNo)||slotNo<1||slotNo>3)return res.status(400).json({error:'Logo slot must be 1, 2 or 3'});if(!Buffer.isBuffer(req.body)||req.body.length<32)return res.status(400).json({error:'Select a valid PNG or JPG logo'});const original=String(req.headers['x-file-name']||'logo'),ext=detectImageExtension(req.body,req.headers['content-type'],original);if(!ext)return res.status(400).json({error:'Only PNG or JPG images are supported'});for(const f of fs.readdirSync(BRANDING_DIR))if(new RegExp('^company_logo'+slotNo+'\\.(png|jpg|jpeg)$','i').test(f)){try{fs.unlinkSync(path.join(BRANDING_DIR,f))}catch(_){}}const filename='company_logo'+slotNo+ext,full=path.join(BRANDING_DIR,filename);fs.writeFileSync(full,req.body);const b=currentBranding();while(b.logos.length<3)b.logos.push(normalizeLogoSlot({},b.logos.length));b.logos[slotNo-1]=normalizeLogoSlot({...b.logos[slotNo-1],file:path.relative(__dirname,full).replace(/\\/g,'/'),enabled:true},slotNo-1);if(slotNo===1){b.companyLogo=b.logos[0].file;b.logoMode=b.logos[0].position}writeJson('branding.json',b);try{await refreshCanonicalSystemTemplates()}catch(e){}audit('BRANDING_LOGO_UPLOAD',req.user,{slot:slotNo,file:filename,size:req.body.length});res.json({ok:true,slot:slotNo,file:b.logos[slotNo-1].file,logoUrl:`/api/branding/logo/${slotNo}?v=${Date.now()}`});}catch(e){console.error(e);res.status(500).json({error:e.message})}});
app.post('/api/settings/branding/logo', requireAuth, requireAdmin,express.raw({type:['image/png','image/jpeg','image/jpg','application/octet-stream'],limit:'5mb'}),async(req,res)=>{try{req.params.slot='1';if(!Buffer.isBuffer(req.body)||req.body.length<32)return res.status(400).json({error:'Select a valid PNG or JPG logo'});const ext=detectImageExtension(req.body,req.headers['content-type'],req.headers['x-file-name']);if(!ext)return res.status(400).json({error:'Only PNG or JPG images are supported'});const filename='company_logo1'+ext,full=path.join(BRANDING_DIR,filename);fs.writeFileSync(full,req.body);const b=currentBranding();b.logos[0]=normalizeLogoSlot({...b.logos[0],file:path.relative(__dirname,full).replace(/\\/g,'/'),enabled:true},0);b.companyLogo=b.logos[0].file;b.logoMode=b.logos[0].position;writeJson('branding.json',b);res.json({ok:true,companyLogo:b.companyLogo,logoUrl:`/api/branding/logo/1?v=${Date.now()}`});}catch(e){res.status(500).json({error:e.message})}});

app.get('/api/settings/branding',requireAuth,requireAdmin,(req,res)=>{
  res.json({...currentBranding(),...brandingPublicPayload()});
});
app.put('/api/settings/branding',requireAuth,requireAdmin,async(req,res)=>{try{const before=currentBranding(),body=req.body||{};let logos=Array.isArray(body.logos)?body.logos:before.logos;while(logos.length<3)logos.push({});logos=logos.slice(0,3).map((slot,i)=>normalizeLogoSlot({...before.logos?.[i],...slot,file:String(slot?.file??before.logos?.[i]?.file??'').trim()},i));const next={...before,companyName:String(body.companyName??before.companyName??'').trim(),factoryName:String(body.factoryName??before.factoryName??'').trim(),headerTitle:String(body.headerTitle??before.headerTitle??'').trim(),theme:{primary:String(body.theme?.primary||before.theme?.primary||'#c90000'),header:String(body.theme?.header||before.theme?.header||'#c90000'),background:String(body.theme?.background||before.theme?.background||'#f4f7fb'),card:String(body.theme?.card||before.theme?.card||'#ffffff')},textLayout:normalizeTextLayout(body.textLayout??before.textLayout),headerZones:normalizeHeaderZones(body.headerZones??before.headerZones),excelLogo:normalizeExcelLogo(body.excelLogo??before.excelLogo),logos,leftLogo:String(body.leftLogo??before.leftLogo??'').trim(),rightLogo:String(body.rightLogo??before.rightLogo??'').trim()};next.companyLogo=logos[0]?.file||'';next.logoMode=logos[0]?.position||'left';writeJson('branding.json',next);try{await refreshCanonicalSystemTemplates()}catch(e){}audit('SETTINGS_BRANDING',req.user,{companyName:next.companyName,factoryName:next.factoryName,headerTitle:next.headerTitle,textLayout:next.textLayout,headerZones:next.headerZones,excelLogo:next.excelLogo,logos:next.logos.map(x=>({id:x.id,position:x.position,size:x.size,width:x.width,height:x.height,enabled:x.enabled,hasFile:!!x.file}))});res.json({ok:true,...brandingPublicPayload()});}catch(e){res.status(500).json({error:e.message})}});

app.post('/api/settings/branding/apply-existing',requireAuth,requireAdmin,async(req,res)=>{
  try{
    const results={annual:[],monthly:[]};
    const master=loadPmMaster();
    for(const l of (master.lines||[]).filter(x=>x.enabled!==false)){
      const line=l.name;
      const years=new Set([
        Number(l.createdYear||new Date().getFullYear()),
        ...systemTasksFor(line,Number(l.createdYear||new Date().getFullYear())).map(t=>Number(t.year))
      ].filter(Boolean));

      // Include years from actual Annual files.
      for(const f of annualFilesForLine(line))years.add(Number(f.year));

      for(const year of years){
        try{
          if(fs.existsSync(annualPathFor(line,year))){
            const r=await rebuildAnnualFromSystemMaster(line,year);
            results.annual.push({line,year,ok:true,file:r.file});
          }
        }catch(e){results.annual.push({line,year,ok:false,error:e.message})}

        for(const mf of availableMonthlyFiles(line,year)){
          try{
            const r=await createMonthlyForLine(line,year,mf.month,{overwrite:true,preserveExecution:true});
            results.monthly.push({line,year,month:mf.month,ok:true,file:r.file});
          }catch(e){results.monthly.push({line,year,month:mf.month,ok:false,error:e.message})}
        }
      }
    }

    const failures=[...results.annual,...results.monthly].filter(x=>!x.ok);
    audit('BRANDING_APPLY_EXISTING',req.user,{annual:results.annual.length,monthly:results.monthly.length,failures:failures.length});
    res.status(failures.length?207:200).json({ok:failures.length===0,results,failures});
  }catch(e){console.error(e);res.status(500).json({error:e.message})}
});



// ===== MANAGEMENT ANALYSIS + WEEKLY/MONTHLY REPORTS =====
function reportWeekDates(year,month,weekIndex){
  const fr=fridaysInMonth(Number(year),Number(month));
  const start=fr[Number(weekIndex)-1];
  if(!start)return null;
  const end=new Date(start);end.setDate(end.getDate()+6);
  return {start,end,header:`${start.getDate()}-${end.getDate()}`};
}
async function readReportMonth(line,year,month){
  // Reports/analysis must work for the complete annual plan, not only months
  // already opened for execution. Build a missing monthly view automatically
  // from the annual/system master while preserving any existing execution data.
  const y=Number(year), m=Number(month);
  let full=findExistingMonthlyFile(line,y,m);
  if(!fs.existsSync(full)){
    await createMonthlyForLine(line,y,m,{overwrite:false,preserveExecution:true});
  }
  const x=await openMonthlyForLine(line,y,m);
  const weeks=monthlyWeeks(x.ws), rows=[];
  const systemTaskById=new Map((loadSystemTasks().tasks||[]).map(t=>[String(t.taskId||''),t]));
  let cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
  for(let r=pmDataStartRow(x.ws);r<=x.ws.rowCount;r++){
    if(String(safeText(x.ws,r,8)||'').trim().toLowerCase()!=='planed')continue;
    const vals={section:String(safeText(x.ws,r,1)||'').trim(),machineNo:String(safeText(x.ws,r,2)||'').trim(),machineName:String(safeText(x.ws,r,3)||'').trim(),machineIdentified:String(safeText(x.ws,r,4)||'').trim(),type:String(safeText(x.ws,r,5)||'').trim()};
    if(vals.machineNo&&vals.machineNo!==cur.machineNo)cur={section:'',machineNo:'',machineName:'',machineIdentified:'',type:''};
    for(const k of Object.keys(vals))if(vals[k]&&vals[k]!=='0 0')cur[k]=vals[k];
    const taskId=stableTaskId(x.ws,r), master=systemTaskById.get(String(taskId||''))||{};
    for(let wi=0;wi<weeks.length;wi++){
      const w=weeks[wi],code=Number(x.ws.getCell(r,w.dataCol).value);
      if(![1,2,3,4,5].includes(code))continue;
      const auditCol=w.dataCol+5; // P:T plan, U:Y audit
      const done=Number(x.ws.getCell(r+1,w.dataCol).value)===6;
      const doneBy=done?String(safeText(x.ws,r,auditCol)||'').trim():'';
      const doneDate=done?readDateValue(x.ws.getCell(r+1,auditCol)):'';
      const issue=latestPmIssue(line,taskId,w.header);
      rows.push({line,year:Number(year),month:Number(month),week:wi+1,weekHeader:w.header,taskId,row:r,section:cur.section,machineNo:cur.machineNo,machineName:cur.machineName,machineIdentified:cur.machineIdentified,type:cur.type,part:String(safeText(x.ws,r,6)||'').trim(),maintenance:String(safeText(x.ws,r,7)||'').trim(),arabicMaintenance:String(master.arabicMaintenance||'').trim(),frequency:String(safeText(x.ws,r,11)||'').trim(),code,done,doneBy,doneDate,issue:issue?String(issue.comment||''):''});
    }
  }
  return rows;
}
function reportStatus(row){
  if(row.done)return 'Done';
  const d=reportWeekDates(row.year,row.month,row.week);if(!d)return 'Pending';
  const today=new Date();today.setHours(23,59,59,999);
  return d.end<today?'Overdue':'Pending';
}
app.get('/api/management/analysis',requireAuth,async(req,res)=>{
  try{
    const year=Number(req.query.year)||new Date().getFullYear();
    const lineQ=String(req.query.line||'ALL');
    const lines=lineQ==='ALL'?configuredPmLines():configuredPmLines().filter(x=>x===lineQ);
    const months=[];const tech={};
    for(let m=1;m<=12;m++){
      let rows=[];for(const line of lines){try{rows.push(...await readReportMonth(line,year,m))}catch(_){}}
      const planned=rows.length,done=rows.filter(x=>x.done).length,overdue=rows.filter(x=>reportStatus(x)==='Overdue').length,pending=planned-done-overdue;
      months.push({month:m,planned,done,overdue,pending,adherence:planned?Math.round(done/planned*1000)/10:0});
      for(const r of rows.filter(x=>x.done)){
        const name=r.doneBy||'Unknown'; if(!tech[name])tech[name]={name,total:0,weeks:{}};tech[name].total++;
        const key=`${m}-W${r.week}`;tech[name].weeks[key]=(tech[name].weeks[key]||0)+1;
      }
    }
    const now=new Date(),currentMonth=(year===now.getFullYear()?now.getMonth()+1:12);
    const ytd=months.filter(x=>x.month<=currentMonth).reduce((a,x)=>({planned:a.planned+x.planned,done:a.done+x.done,overdue:a.overdue+x.overdue,pending:a.pending+x.pending}),{planned:0,done:0,overdue:0,pending:0});
    ytd.adherence=ytd.planned?Math.round(ytd.done/ytd.planned*1000)/10:0;
    const annual=months.reduce((a,x)=>({planned:a.planned+x.planned,done:a.done+x.done,overdue:a.overdue+x.overdue,pending:a.pending+x.pending}),{planned:0,done:0,overdue:0,pending:0});
    annual.adherence=annual.planned?Math.round(annual.done/annual.planned*1000)/10:0;
    res.json({year,line:lineQ,months,ytd,annual,technicians:Object.values(tech).sort((a,b)=>b.total-a.total)});
  }catch(e){res.status(500).json({error:e.message})}
});
// Canonical report endpoint: planned occurrences come directly from the same
// pm_tasks.json annual source used by Annual Plan. Monthly workbooks are used
// only to enrich rows with execution/history; they can never make the plan 0.
app.get('/api/annual-plan-report',requireAuth,async(req,res)=>{
  try{
    const year=Number(req.query.year)||new Date().getFullYear();
    const month=Number(req.query.month)||0, week=Number(req.query.week)||0;
    const lineQ=String(req.query.line||'ALL'), statusQ=String(req.query.status||'ALL');
    const lines=lineQ==='ALL'?configuredPmLines():configuredPmLines().filter(x=>String(x).toLowerCase()===lineQ.toLowerCase());
    const monthNames=['January','February','March','April','May','June','July','August','September','October','November','December'];
    const first=new Date(year,0,1,12); while(first.getDay()!==5)first.setDate(first.getDate()+1);
    const execution=new Map();
    const wantedMonths=month?[month]:[1,2,3,4,5,6,7,8,9,10,11,12];
    for(const m of wantedMonths){for(const ln of lines){
      try{for(const r of await readReportMonth(ln,year,m)){
        const d=new Date(year,m-1,1,12);while(d.getDay()!==5)d.setDate(d.getDate()+1);d.setDate(d.getDate()+(Number(r.week)-1)*7);
        const gw=Math.floor((d-first)/604800000)+1;
        execution.set(`${ln.toLowerCase()}|${String(r.taskId||'')}|${gw}`,r);
      }}catch(e){console.warn('[REPORT EXECUTION ENRICH]',ln,year,m,e.message)}
    }}
    let rows=[];
    const all=loadSystemTasks().tasks||[];
    for(const ln of lines){
      const tasks=all.filter(t=>String(t.line||'').toLowerCase()===ln.toLowerCase()&&Number(t.year)===year&&t.active!==false&&t.isActive!==false);
      for(const t of tasks){for(const w0 of (t.scheduleWeeks||[])){
        const gw=Number(w0); if(!gw)continue;
        const d=new Date(first); d.setDate(d.getDate()+(gw-1)*7); const mm=d.getMonth()+1;
        if(month&&mm!==month)continue;
        const fr=fridaysInMonth(year,mm); const lw=Math.max(1,fr.findIndex(x=>x.getDate()===d.getDate()&&x.getMonth()===d.getMonth())+1);
        if(week && week!==lw && week!==gw)continue;
        const ex=execution.get(`${ln.toLowerCase()}|${String(t.taskId||'')}|${gw}`)||{};
        let st=ex.done?'Done':'Pending';
        if(!ex.done){const end=new Date(d);end.setDate(end.getDate()+6);end.setHours(23,59,59,999);if(end<new Date())st='Overdue'}
        rows.push({line:ln,section:t.section||ln,year,month:mm,monthName:monthNames[mm-1],week:lw,globalWeek:gw,weekHeader:'W'+gw,taskId:t.taskId||'',machineNo:t.machineNo||'',machineName:t.machineName||t.machineNo||'',machineIdentified:t.machineIdentified||'',type:t.type||t.maintenanceType||'',part:t.part||'',maintenance:t.maintenance||'',arabicMaintenance:t.arabicMaintenance||'',frequency:t.frequency||'',status:st,done:!!ex.done,doneBy:ex.doneBy||'',doneDate:ex.doneDate||'',issue:ex.issue||''});
      }}
    }
    if(statusQ!=='ALL')rows=rows.filter(r=>r.status===statusQ);
    res.json({year,month,week,line:lineQ,rows});
  }catch(e){console.error(e);res.status(500).json({error:e.message})}
});

app.get('/api/management/report',requireAuth,async(req,res)=>{
  try{
    const year=Number(req.query.year)||new Date().getFullYear(),month=Number(req.query.month)||0,week=Number(req.query.week)||0,lineQ=String(req.query.line||'ALL'),tech=String(req.query.technician||'ALL'),status=String(req.query.status||'ALL');
    const lines=lineQ==='ALL'?configuredPmLines():configuredPmLines().filter(x=>x===lineQ);
    const first=new Date(year,0,1,12);while(first.getDay()!==5)first.setDate(first.getDate()+1);
    const execution=new Map();
    for(const m of (month===0?[1,2,3,4,5,6,7,8,9,10,11,12]:[month])){
      for(const line of lines){
        try{for(const r of await readReportMonth(line,year,m)){
          const key=[line,String(r.taskId||''),m,Number(r.week)||0].join('|');execution.set(key,r);
        }}catch(_){}
      }
    }
    let rows=[];
    for(const line of lines){
      const ts=systemTasksFor(line,year).filter(t=>t&&t.isActive!==false&&t.active!==false);
      for(const t of ts){for(const gw0 of (t.scheduleWeeks||[])){
        const gw=Number(gw0);if(!gw)continue;const d=new Date(first);d.setDate(d.getDate()+(gw-1)*7);
        const mm=d.getMonth()+1;if(month!==0&&mm!==month)continue;
        const fr=fridaysInMonth(year,mm);let lw=fr.findIndex(x=>x.getFullYear()===d.getFullYear()&&x.getMonth()===d.getMonth()&&x.getDate()===d.getDate())+1;if(lw<1)lw=1;
        const ex=execution.get([line,String(t.taskId||''),mm,lw].join('|'))||{};
        rows.push({line,year,month:mm,monthName:['January','February','March','April','May','June','July','August','September','October','November','December'][mm-1],week:lw,globalWeek:gw,weekHeader:'W'+gw,taskId:t.taskId||'',section:t.section||line,machineNo:t.machineNo||'',machineName:t.machineName||t.machineNo||'',machineIdentified:t.machineIdentified||'',type:t.type||t.maintenanceType||'',part:t.part||'',maintenance:t.maintenance||'',arabicMaintenance:t.arabicMaintenance||'',frequency:t.frequency||'',code:Number((t.scheduleCodes||{})[gw]||t.effectiveCode||t.code||1),done:!!ex.done,doneBy:ex.doneBy||'',doneDate:ex.doneDate||'',issue:ex.issue||'',status:ex.status||(ex.done?'Done':'Pending')});
      }}
    }
    if(week)rows=rows.filter(x=>x.week===week||x.globalWeek===week);if(tech!=='ALL')rows=rows.filter(x=>x.doneBy===tech);if(status!=='ALL')rows=rows.filter(x=>x.status===status);
    res.json({year,month,week,line:lineQ,rows});
  }catch(e){res.status(500).json({error:e.message})}
});


for (const key of ['system','lines']) {
  app.get('/api/settings/'+key, requireAuth, requireAdmin, (req,res)=>res.json(readJson(key+'.json',{})));
  app.put('/api/settings/'+key, requireAuth, requireAdmin, (req,res)=>{
    writeJson(key+'.json',req.body||{}); audit('SETTINGS_'+key.toUpperCase(),req.user,req.body||{});
    res.json({ok:true});
  });
}

app.get('/api/settings/audit', requireAuth, requireAdmin, (req,res)=>{
  res.json(readJson('audit.json',[]).slice(0,500));
});

app.post('/api/settings/backup', requireAuth, requireAdmin, (req,res)=>{
  const stamp=new Date().toISOString().replace(/[:.]/g,'-');
  const folder=path.join(BACKUP_DIR,'backup-'+stamp);
  fs.mkdirSync(folder,{recursive:true});
  for(const f of fs.readdirSync(CONFIG_DIR)) fs.copyFileSync(path.join(CONFIG_DIR,f),path.join(folder,f));
  audit('BACKUP_CREATE',req.user,{folder});
  res.json({ok:true,name:path.basename(folder)});
});


const monthlyStorageOrganize=organizeExistingMonthlyFilesByLine();
console.log('[MONTHLY STORAGE]', monthlyStorageOrganize);

(async()=>{
  try{
    const synced=await refreshCanonicalSystemTemplates();
    console.log('[SYSTEM TEMPLATES SYNCED]',synced);
  }catch(e){
    console.error('[SYSTEM TEMPLATE SYNC ERROR]',e.message);
  }

  try{
    const rollover=await autoBuildCurrentYearPlans();
    if(rollover.length)console.log('[YEAR PLANNER AUTO ROLLOVER]',rollover);
  }catch(e){console.error('[YEAR PLANNER AUTO ROLLOVER ERROR]',e.message);}

app.listen(PORT,()=>{
  const w=getWeek(getPmDate());
  console.log('====================================================');
  console.log(' SMART PM SERVER - V1.4.2 LOGIN CLEANUP');
  console.log(` http://localhost:${PORT}`);
  console.log(' Login: admin / 1234   or   tech1 / 1234');
  console.log(` Test: ${TEST_LINE} | ${fmtDate(w.start)} -> ${fmtDate(w.end)}`);
  console.log(' Done = 6 + Username + Execution Date');
  console.log(' Monthly week count is detected automatically: 4 or 5');
  console.log(' Auto email: Thursday close -> Weekly Excel; final PM Thursday -> Monthly Excel');
  console.log('====================================================');
});
})();
