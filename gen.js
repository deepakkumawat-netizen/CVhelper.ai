/**
 * CVhelper.ai — PPTX Generator v6 (Bold Premium Edition)
 * node gen.js <data.json> <output.pptx>
 * LAYOUT_WIDE = 13.33" x 7.5"
 */
"use strict";
const PptxGenJS = require("pptxgenjs");
const fs = require("fs");

const [,, src, out] = process.argv;
if (!src || !out) { console.error("Usage: node gen.js <in.json> <out.pptx>"); process.exit(1); }

const DATA = JSON.parse(fs.readFileSync(src, "utf8"));
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE";

const SW = 13.33;
const SH = 7.5;

// Palette — no # prefix ever
const C = {
  bg:    "050510",
  panel: "0A0A22",
  card:  "0F0F2E",
  card2: "141438",
  line:  "1E1E48",
  white: "FFFFFF",
  offwhite: "E8EDFF",
  silver:"A8B8D0",
  muted: "445070",
  acc:   "6C47FF",
  green: "22D3A0",
  amber: "F5A623",
  cyan:  "18C8EE",
  red:   "F06060",
};
function setAcc(c) { C.acc = (c||"#6C47FF").replace("#",""); }

// ─── BASE HELPERS ──────────────────────────────
function R(sl,x,y,w,h,fill,stroke,pt){
  sl.addShape(pres.shapes.RECTANGLE,{x,y,w,h,
    fill:{color:fill},
    line:stroke?{color:stroke,pt:pt||0.6}:{type:"none"}});
}
function O(sl,x,y,w,h,fill,stroke,pt){
  sl.addShape(pres.shapes.OVAL,{x,y,w,h,
    fill:{color:fill},
    line:stroke?{color:stroke,pt:pt||1}:{type:"none"}});
}
function T(sl,text,x,y,w,h,o){
  sl.addText(text,{x,y,w,h,margin:0,...o});
}
function Tbullets(sl,items,x,y,w,h,fontSize,color){
  if(!items||!items.length) return;
  const arr = items.map((t,i)=>({
    text: t,
    options:{bullet:true, breakLine: i<items.length-1, fontSize, color, paraSpaceAfter:3}
  }));
  sl.addText(arr,{x,y,w,h,margin:6,valign:"top"});
}

// ─── BACKGROUNDS ──────────────────────────────
function bgCover(sl){
  sl.background={color:C.bg};
  // Large gradient glow top-right
  sl.addShape(pres.shapes.OVAL,{x:7.5,y:-3,w:9,h:9,fill:{color:C.acc,transparency:87},line:{type:"none"}});
  // Small cyan glow bottom-left
  sl.addShape(pres.shapes.OVAL,{x:-2,y:5,w:5,h:5,fill:{color:C.cyan,transparency:90},line:{type:"none"}});
  // Bold left accent bar
  R(sl,0,0,0.08,SH,C.acc,null);
}

function bgSection(sl){
  sl.background={color:C.bg};
  // Top colored band
  sl.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:SW,h:2.0,fill:{color:C.acc,transparency:10},line:{type:"none"}});
  sl.addShape(pres.shapes.RECTANGLE,{x:0,y:1.98,w:SW,h:0.05,fill:{color:C.acc},line:{type:"none"}});
  // Right glow
  sl.addShape(pres.shapes.OVAL,{x:10,y:-1.5,w:6,h:6,fill:{color:C.acc,transparency:88},line:{type:"none"}});
}

function bgSplit(sl){
  sl.background={color:C.bg};
  // Left dark panel
  R(sl,0,0,4.8,SH,C.panel,null);
  // Accent divider
  R(sl,4.78,0,0.05,SH,C.acc,null);
}

function bgDefault(sl){
  sl.background={color:C.bg};
  // Subtle glow
  sl.addShape(pres.shapes.OVAL,{x:10,y:-2,w:6,h:6,fill:{color:C.acc,transparency:90},line:{type:"none"}});
  sl.addShape(pres.shapes.OVAL,{x:-1.5,y:5.5,w:4,h:4,fill:{color:C.cyan,transparency:93},line:{type:"none"}});
}

function bgDark(sl){
  sl.background={color:C.bg};
  // Full dark with center glow
  sl.addShape(pres.shapes.OVAL,{x:3,y:0,w:8,h:8,fill:{color:C.acc,transparency:94},line:{type:"none"}});
}

// ─── UI COMPONENTS ──────────────────────────────
function chip(sl,label,x,y){
  const cw = Math.max(label.length*0.098+0.6,2.0);
  R(sl,x,y,cw,0.3,C.acc,C.acc,0.5);
  // Slight transparency overlay
  sl.addShape(pres.shapes.RECTANGLE,{x,y,w:cw,h:0.3,fill:{color:"FFFFFF",transparency:88},line:{type:"none"}});
  T(sl,label,x,y,cw,0.3,{fontSize:8,bold:true,color:C.white,align:"center",valign:"middle",charSpacing:1.5});
}

function pageNum(sl,cur,tot){
  const x=SW-2.0, y=SH-0.42;
  R(sl,x,y,1.9,0.32,C.card2,C.line,0.5);
  T(sl,`${String(cur).padStart(2,"0")}  /  ${String(tot).padStart(2,"0")}`,x,y,1.9,0.32,
    {fontSize:8,color:C.muted,align:"center",valign:"middle"});
}

// Card with optional left accent bar
function card(sl,x,y,w,h,color,leftBar){
  const c=color||C.acc;
  R(sl,x,y,w,h,C.card,c,0.6);
  if(leftBar) R(sl,x,y,0.08,h,c,null);
}

// Bold numbered circle
function numCircle(sl,n,x,y,sz,color){
  O(sl,x,y,sz,sz,color||C.acc,null);
  T(sl,String(n),x,y,sz,sz,{fontSize:sz*14,bold:true,color:C.white,align:"center",valign:"middle"});
}

// Stat card (big value + label)
function statCard(sl,x,y,w,val,lbl,color){
  const c=color||C.acc;
  R(sl,x,y,w,1.1,C.card,c,0.6);
  R(sl,x,y,w,0.05,c,null);
  T(sl,val,x,y+0.05,w,0.6,{fontSize:22,bold:true,color:c,align:"center",valign:"middle",fontFace:"Calibri"});
  T(sl,lbl,x,y+0.7,w,0.32,{fontSize:9.5,color:C.silver,align:"center"});
}

// ─── DIAGRAMS ──────────────────────────────────

// Horizontal flow with arrows
function flowDiagram(sl,steps,x,y,w,h){
  const n=Math.min(steps.length,5);
  if(n<1)return;
  const gap=0.25, bw=(w-gap*(n-1))/n, bh=h*0.65, by=y+h*0.2;
  steps.slice(0,n).forEach((step,i)=>{
    const bx=x+i*(bw+gap);
    const title=typeof step==="object"?step.title:String(step);
    const body=typeof step==="object"&&step.body?step.body:"";
    // Card
    card(sl,bx,by,bw,bh,C.acc);
    R(sl,bx,by,bw,0.06,C.acc,null);
    // Number bubble above
    O(sl,bx+bw/2-0.28,by-0.32,0.56,0.56,C.acc,null);
    T(sl,String(i+1),bx+bw/2-0.28,by-0.32,0.56,0.56,{fontSize:12,bold:true,color:C.white,align:"center",valign:"middle"});
    // Title
    T(sl,title,bx+0.1,by+0.1,bw-0.2,0.42,{fontSize:10.5,bold:true,color:C.white,align:"center",wrap:true});
    // Body
    if(body) T(sl,body,bx+0.1,by+0.56,bw-0.2,bh-0.65,{fontSize:9,color:C.silver,align:"center",valign:"top",wrap:true});
    // Arrow
    if(i<n-1){
      const ax=bx+bw+0.04, ay=by+bh/2;
      R(sl,ax,ay-0.03,0.15,0.06,C.acc,null);
      R(sl,ax+0.1,ay-0.08,0.07,0.16,C.acc,null);
    }
  });
  T(sl,"Process Flow",x,y+h-0.24,w,0.2,{fontSize:8,color:C.muted,align:"center",italic:true});
}

// 2x2 concept grid with real content
function conceptGrid(sl,items,x,y,w,h){
  const gx=0.2,gy=0.2,bw=(w-gx)/2,bh=(h-gy)/2;
  const cols=[C.acc,C.cyan,C.green,C.amber];
  items.slice(0,4).forEach((item,i)=>{
    const col=i%2,row=Math.floor(i/2);
    const bx=x+col*(bw+gx),by=y+row*(bh+gy),c=cols[i];
    const title=typeof item==="object"?item.title:String(item);
    const body=typeof item==="object"&&item.body?item.body:"";
    card(sl,bx,by,bw,bh,c);
    R(sl,bx,by,bw,0.06,c,null);
    // Number
    O(sl,bx+0.14,by+0.14,0.46,0.46,c,null);
    T(sl,String(i+1),bx+0.14,by+0.14,0.46,0.46,{fontSize:11,bold:true,color:C.white,align:"center",valign:"middle"});
    // Title
    T(sl,title,bx+0.7,by+0.14,bw-0.84,0.46,{fontSize:11,bold:true,color:C.white,valign:"middle",wrap:true});
    // Body
    if(body) T(sl,body,bx+0.14,by+0.7,bw-0.28,bh-0.88,{fontSize:10,color:C.silver,valign:"top",wrap:true});
  });
}

// Timeline alternating above/below
function timelineDiagram(sl,events,x,y,w,h){
  const n=Math.min(events.length,5), step=w/n, mid=y+h*0.45;
  R(sl,x,mid-0.03,w,0.06,C.acc,null);
  events.slice(0,n).forEach((ev,i)=>{
    const cx=x+i*step+step/2, above=i%2===0;
    const title=typeof ev==="object"?ev.title:String(ev);
    O(sl,cx-0.18,mid-0.18,0.36,0.36,C.acc,C.white,1.5);
    const cw=Math.min(step*0.88,2.4),ch=0.72,cardX=cx-cw/2;
    const cardY=above?mid-ch-0.58:mid+0.44;
    card(sl,cardX,cardY,cw,ch,C.acc);
    T(sl,title,cardX,cardY,cw,ch,{fontSize:9.5,bold:true,color:C.white,align:"center",valign:"middle",wrap:true});
    const c1=above?cardY+ch:mid+0.18,c2=above?mid-0.18:cardY;
    R(sl,cx-0.03,c1,0.06,Math.abs(c2-c1),C.acc,null);
  });
  T(sl,"Timeline",x,y+h-0.24,w,0.2,{fontSize:8,color:C.muted,align:"center",italic:true});
}

// ─── FULL COURSE ────────────────────────────────
function buildFullCourse(course){
  setAcc(course.color||"#6C47FF");
  const mods=course.modules||[];
  const totalCls=mods.reduce((a,m)=>a+(m.classes||[]).length,0);
  const total=4+mods.length*2+1;
  let sn=0;

  // SLIDE 1: Cover
  sn++;
  const s1=pres.addSlide(); bgCover(s1);
  chip(s1,(course.category||"COURSE").toUpperCase().slice(0,20),0.55,0.25);
  T(s1,course.title||"Course",0.55,0.72,9.2,2.1,{fontSize:40,bold:true,color:C.white,fontFace:"Calibri",wrap:true});
  T(s1,course.subtitle||"",0.55,2.9,8.8,0.55,{fontSize:14,color:C.silver,italic:true});
  [{v:mods.length+" Modules",l:"Curriculum"},{v:totalCls+" Classes",l:"Total Lessons"},{v:course.duration||"Self-paced",l:"Duration"},{v:course.level||"All Levels",l:"Level"}]
    .forEach((st,i)=>{ statCard(s1,0.55+i*3.15,3.6,2.98,st.v,st.l); });
  (course.whatYouLearn||[]).slice(0,4).forEach((o,i)=>{
    card(s1,10.0,0.68+i*1.0,3.1,0.82,C.acc,true);
    T(s1,o,10.22,0.68+i*1.0,2.76,0.82,{fontSize:10,color:C.silver,valign:"middle",wrap:true});
  });
  pageNum(s1,sn,total);

  // SLIDE 2: Outcomes
  sn++;
  const s2=pres.addSlide(); bgSection(s2);
  chip(s2,"LEARNING OUTCOMES",0.55,0.1);
  T(s2,"What You Will Learn",0.55,0.35,12,0.95,{fontSize:34,bold:true,color:C.white,fontFace:"Calibri"});
  (course.whatYouLearn||[]).slice(0,6).forEach((o,i)=>{
    const col=i%3,row=Math.floor(i/3),ox=0.55+col*4.25,oy=1.42+row*1.12;
    card(s2,ox,oy,4.0,0.95,C.acc,true);
    T(s2,"+",ox+0.14,oy,0.4,0.95,{fontSize:16,bold:true,color:C.acc,valign:"middle"});
    T(s2,o,ox+0.6,oy,3.28,0.95,{fontSize:11,color:C.offwhite,valign:"middle",wrap:true});
  });
  pageNum(s2,sn,total);

  // SLIDE 3: Requirements
  sn++;
  const s3=pres.addSlide(); bgDefault(s3);
  chip(s3,"REQUIREMENTS",0.55,0.25);
  T(s3,"Course At a Glance",0.55,0.65,7,0.75,{fontSize:30,bold:true,color:C.white,fontFace:"Calibri"});
  (course.requirements||[]).slice(0,4).forEach((r,i)=>{
    card(s3,0.55,1.58+i*0.82,5.7,0.68,C.acc,true);
    T(s3,String(i+1)+".",0.72,1.58+i*0.82,0.44,0.68,{fontSize:12,bold:true,color:C.acc,valign:"middle"});
    T(s3,r,1.2,1.58+i*0.82,4.92,0.68,{fontSize:11,color:C.offwhite,valign:"middle",wrap:true});
  });
  [{v:mods.length,l:"Modules"},{v:totalCls,l:"Classes"},{v:(course.totalHours||8)+"h",l:"Total Hours"}]
    .forEach((st,i)=>{ statCard(s3,7.0+i*2.12,1.55,2.0,String(st.v),st.l); });
  if(course.description) T(s3,course.description,0.55,5.1,SW-1.1,0.85,{fontSize:11,color:C.silver,italic:true,wrap:true,align:"center"});
  pageNum(s3,sn,total);

  // SLIDE 4: Curriculum — fills full slide height
  sn++;
  const s4=pres.addSlide(); bgSection(s4);
  chip(s4,"CURRICULUM",0.55,0.1);
  T(s4,"Full Course Curriculum",0.55,0.35,12,0.88,{fontSize:32,bold:true,color:C.white,fontFace:"Calibri"});

  // Dynamic row height to fill slide: available = SH - header(1.38) - pagenum(0.5) - padding(0.2)
  const s4AvailH = SH - 1.38 - 0.5 - 0.22;
  const s4Gap    = 0.18;
  const s4RowH   = (s4AvailH - s4Gap*(mods.length-1)) / mods.length;
  const s4Colors = [C.acc,C.amber,C.cyan,C.green];

  mods.forEach((mod,mi)=>{
    const my = 1.38 + mi*(s4RowH+s4Gap);
    const col = s4Colors[mi%4];
    // Card
    R(s4,0.55,my,SW-1.1,s4RowH,C.card,col,0.6);
    R(s4,0.55,my,0.08,s4RowH,col,null);
    R(s4,0.55,my,SW-1.1,0.05,col,null);
    // Number
    O(s4,0.75,my+s4RowH/2-0.32,0.64,0.64,col,null);
    T(s4,String(mi+1).padStart(2,"0"),0.75,my+s4RowH/2-0.32,0.64,0.64,{fontSize:13,bold:true,color:C.white,align:"center",valign:"middle"});
    // Module title
    T(s4,mod.title,1.52,my+0.08,7.2,s4RowH*0.48,{fontSize:14.5,bold:true,color:C.white,valign:"middle",wrap:true});
    // Description
    if(mod.description && s4RowH>0.95){
      T(s4,mod.description,1.52,my+s4RowH*0.52,7.0,s4RowH*0.45,{fontSize:10.5,color:C.silver,valign:"top",wrap:true,italic:true});
    }
    // Right: classes count + duration
    T(s4,(mod.classes||[]).length+" classes",9.4,my,2.2,s4RowH*0.52,{fontSize:11,bold:true,color:col,valign:"middle",align:"right"});
    T(s4,mod.duration||"",9.4,my+s4RowH*0.52,2.2,s4RowH*0.48,{fontSize:9.5,color:C.muted,valign:"middle",align:"right"});
  });
  pageNum(s4,sn,total);

  // Per-module slides
  mods.forEach((mod,mi)=>{
    sn++;
    const sm=pres.addSlide(); bgSplit(sm);
    chip(sm,"MODULE "+(mi+1)+" OF "+mods.length,5.1,0.28);
    // Large number on left
    T(sm,String(mi+1).padStart(2,"0"),0.15,0.5,4.42,2.6,{fontSize:100,bold:true,color:C.acc,align:"center",fontFace:"Calibri",transparency:25});
    // Title + description on right
    T(sm,mod.title,5.1,0.72,8.0,1.2,{fontSize:26,bold:true,color:C.white,fontFace:"Calibri",wrap:true});
    T(sm,mod.description||"",5.1,2.06,8.0,0.65,{fontSize:11.5,color:C.silver,italic:true,wrap:true});

    // Class cards — dynamic height to fill slide
    const smClasses  = mod.classes||[];
    const smAvailH   = SH - 2.85 - 0.45;
    const smGap      = 0.15;
    const smCardH    = Math.min((smAvailH - smGap*(smClasses.length-1)) / Math.max(smClasses.length,1), 1.4);
    const smColors   = [C.acc,C.cyan,C.green];

    smClasses.forEach((cls,ci)=>{
      const cy  = 2.85 + ci*(smCardH+smGap);
      const col = smColors[ci%3];
      R(sm,5.1,cy,8.0,smCardH,C.card,col,0.6);
      R(sm,5.1,cy,8.0,0.05,col,null);
      // Type badge
      const badge = cls.type==="assignment"?"Q":"V";
      R(sm,5.1,cy,0.45,smCardH,col,null);
      T(sm,badge,5.1,cy,0.45,smCardH,{fontSize:9.5,bold:true,color:C.white,align:"center",valign:"middle"});
      // Title — truncate safely
      const titleFontSize = smCardH > 0.9 ? 12 : 11;
      T(sm,cls.title.slice(0,40),5.62,cy,5.1,smCardH*0.55,{fontSize:titleFontSize,bold:true,color:C.offwhite,valign:"middle",wrap:true});
      // Duration
      T(sm,cls.duration||"",10.8,cy,2.1,smCardH*0.55,{fontSize:9.5,color:C.muted,valign:"middle",align:"right"});
      // Topics preview on second line if card is tall enough
      if(smCardH>0.85 && cls.topics && cls.topics.length>0){
        T(sm,cls.topics.slice(0,3).join("  |  "),5.62,cy+smCardH*0.55,6.38,smCardH*0.42,{fontSize:9,color:C.silver,valign:"middle",wrap:true,italic:true});
      }
    });
    pageNum(sm,sn,total);

    sn++;
    const sf=pres.addSlide(); bgSection(sf);
    chip(sf,"MODULE "+(mi+1)+" — CLASS OVERVIEW",0.55,0.1);
    // Title clamped to 1 line max
    const sfTitle = (mod.title+" — Class Overview").slice(0,55);
    T(sf,sfTitle,0.55,0.35,SW-1.1,0.88,{fontSize:28,bold:true,color:C.white,fontFace:"Calibri",wrap:false});

    const sfClasses = mod.classes||[];
    const sfColors  = [C.acc, C.cyan, C.green];
    const sfN       = sfClasses.length;
    const sfGap     = 0.22;
    const sfCardW   = (SW - 1.1 - sfGap*(sfN-1)) / Math.max(sfN,1);
    // Card starts at y=1.42, ends at SH-0.5 = 7.0 — safe boundary
    const sfCardTop = 1.42;
    const sfCardBot = SH - 0.48;
    const sfCardH   = sfCardBot - sfCardTop; // = 5.58"

    sfClasses.forEach((cls,ci)=>{
      const cx  = 0.55 + ci*(sfCardW+sfGap);
      const col = sfColors[ci%3];
      const isAssign = cls.type==="assignment";

      // Card background
      R(sf,cx,sfCardTop,sfCardW,sfCardH,C.card,col,0.7);
      R(sf,cx,sfCardTop,sfCardW,0.06,col,null);

      // Number circle — inside card top center
      const circR = 0.34;
      O(sf,cx+sfCardW/2-circR,sfCardTop+0.12,circR*2,circR*2,col,null);
      T(sf,String(ci+1),cx+sfCardW/2-circR,sfCardTop+0.12,circR*2,circR*2,{fontSize:15,bold:true,color:C.white,align:"center",valign:"middle"});

      // Class title — max 2 lines, font size scales with card width
      const titleFs = sfCardW > 3.5 ? 13 : 11.5;
      T(sf,cls.title,cx+0.12,sfCardTop+0.84,sfCardW-0.24,0.72,{fontSize:titleFs,bold:true,color:C.white,align:"center",valign:"top",wrap:true});

      // Duration badge
      const badgeY = sfCardTop + 1.64;
      R(sf,cx+0.12,badgeY,sfCardW-0.24,0.36,col,null);
      sf.addShape(pres.shapes.RECTANGLE,{x:cx+0.12,y:badgeY,w:sfCardW-0.24,h:0.36,fill:{color:"000000",transparency:55},line:{type:"none"}});
      const badgeTxt = isAssign ? "Assessment  |  "+(cls.duration||"30 min") : (cls.duration||"45 min")+" video class";
      T(sf,badgeTxt,cx+0.12,badgeY,sfCardW-0.24,0.36,{fontSize:9,bold:true,color:C.white,align:"center",valign:"middle"});

      // Divider
      R(sf,cx+0.12,sfCardTop+2.1,sfCardW-0.24,0.04,col,null);

      // Topics — fit within remaining card space
      const topicAreaH = sfCardH - 2.2 - (isAssign ? 0.52 : 0.12);
      const maxTopics  = Math.min((cls.topics||[]).length, 5);
      const topicH     = maxTopics > 0 ? Math.min(topicAreaH/maxTopics, 0.52) : 0.48;
      (cls.topics||[]).slice(0,5).forEach((topic,ti)=>{
        const ty = sfCardTop + 2.18 + ti*topicH;
        O(sf,cx+0.16,ty+topicH/2-0.1,0.18,0.18,col,null);
        T(sf,topic.slice(0,36),cx+0.42,ty,sfCardW-0.56,topicH,{fontSize:Math.min(10,9.5),color:C.offwhite,valign:"middle",wrap:true});
      });

      // MCQ badge for assignments — anchored to bottom of card
      if(isAssign){
        const mqY = sfCardBot - 0.52;
        R(sf,cx+0.12,mqY,sfCardW-0.24,0.44,col,null);
        sf.addShape(pres.shapes.RECTANGLE,{x:cx+0.12,y:mqY,w:sfCardW-0.24,h:0.44,fill:{color:"000000",transparency:50},line:{type:"none"}});
        T(sf,"MCQ Quiz  |  10 Questions  |  Pass: 70%",cx+0.12,mqY,sfCardW-0.24,0.44,{fontSize:9,bold:true,color:C.white,align:"center",valign:"middle"});
      }
    });
    pageNum(sf,sn,total);
  });

  // Closing
  sn++;
  const se=pres.addSlide(); bgCover(se);
  T(se,"Start Learning Today",0.55,1.9,SW-1.1,1.15,{fontSize:50,bold:true,color:C.white,fontFace:"Calibri",align:"center"});
  T(se,course.title||"",0.55,3.15,SW-1.1,0.72,{fontSize:20,color:C.acc,align:"center",bold:true});
  R(se,4.6,3.98,4.1,0.06,C.acc,null);
  T(se,mods.length+" Modules  |  "+totalCls+" Classes  |  "+(course.duration||"Self-paced"),0.55,4.12,SW-1.1,0.52,{fontSize:13,color:C.silver,italic:true,align:"center"});
  R(se,4.5,4.82,4.3,0.78,C.acc,null);
  T(se,"Enroll Now",4.5,4.82,4.3,0.78,{fontSize:16,bold:true,color:C.white,align:"center",valign:"middle"});
  pageNum(se,sn,total);
}

// ─── CLASS PPTX ────────────────────────────────
function buildClassPPTX(D){
  setAcc(D.color||"#6C47FF");
  const courseTitle = D.courseTitle||"Course";
  const moduleTitle = D.moduleTitle||"Module";
  const classTitle  = D.classTitle||"Class";
  const topics      = D.topics||[];
  const kps         = D.keyPoints||[];
  const objectives  = D.objectives||[];
  const summary     = D.summary||"";
  const nextSteps   = D.nextSteps||[];
  const videoScript = D.videoScript||"";
  const realWorld   = D.realWorldExample||"";
  const keyStats    = D.keyStats||[];
  const total = 3 + Math.min(kps.length,6) + 3;
  let sn=0;

  // ── SLIDE 1: COVER ──────────────────────────────────────────
  sn++;
  const s1=pres.addSlide(); bgCover(s1);
  chip(s1,moduleTitle.slice(0,24).toUpperCase(),0.55,0.25);
  T(s1,classTitle,0.55,0.7,9.1,2.1,{fontSize:36,bold:true,color:C.white,fontFace:"Calibri",wrap:true});
  T(s1,courseTitle,0.55,2.92,8.5,0.5,{fontSize:13,color:C.silver,italic:true});
  // Info badge
  R(s1,0.55,3.58,6.2,0.56,C.acc,null);
  sf_trans(s1,0.55,3.58,6.2,0.56,30);
  T(s1,topics.length+" Topics   |   "+(D.duration||"45 min")+"   |   "+(D.level||"Beginner"),0.55,3.58,6.2,0.56,{fontSize:11,color:C.white,align:"center",valign:"middle",bold:true});
  // Right: topics as stacked pills
  topics.slice(0,4).forEach((t,i)=>{
    card(s1,9.85,0.65+i*0.98,3.3,0.78,C.acc,true);
    T(s1,t,10.06,0.65+i*0.98,2.96,0.78,{fontSize:10,color:C.offwhite,valign:"middle",wrap:true});
  });
  // Key stats at bottom if available
  if(keyStats.length>0){
    keyStats.slice(0,3).forEach((st,i)=>{
      R(s1,0.55+i*4.1,4.42,3.85,0.62,C.card,C.acc,0.5);
      T(s1,st,0.72+i*4.1,4.42,3.52,0.62,{fontSize:9,color:C.silver,valign:"middle",wrap:true});
    });
  }
  pageNum(s1,sn,total);

  // ── SLIDE 2: LEARNING OBJECTIVES ────────────────────────────
  sn++;
  const s2=pres.addSlide(); bgSection(s2);
  chip(s2,"LEARNING OBJECTIVES",0.55,0.1);
  T(s2,"By the End of This Class",0.55,0.35,12,0.88,{fontSize:32,bold:true,color:C.white,fontFace:"Calibri"});

  const s2Objs    = objectives.slice(0,4);
  const s2AvailH  = SH - 1.38 - 0.45;   // top header + page num margin
  const s2Gap     = 0.16;
  const s2CardH   = (s2AvailH - s2Gap*(s2Objs.length-1)) / Math.max(s2Objs.length,1);

  s2Objs.forEach((obj,i)=>{
    const oy = 1.38 + i*(s2CardH+s2Gap);
    // Card
    R(s2,0.55,oy,SW-1.1,s2CardH,C.card,C.acc,0.6);
    R(s2,0.55,oy,0.08,s2CardH,C.acc,null);
    // Circle with number — vertically centered in card
    const circY = oy + s2CardH/2 - 0.28;
    O(s2,0.72,circY,0.56,0.56,C.acc,null);
    T(s2,String(i+1),0.72,circY,0.56,0.56,{fontSize:13,bold:true,color:C.white,align:"center",valign:"middle"});
    // Objective text — full height so long text wraps properly
    T(s2,obj,1.42,oy,SW-2.1,s2CardH,{fontSize:13.5,color:C.offwhite,valign:"middle",wrap:true});
  });
  pageNum(s2,sn,total);

  // ── SLIDE 3: TOPICS + CLASS INTRO ───────────────────────────
  sn++;
  const s3=pres.addSlide(); bgDefault(s3);
  chip(s3,"TOPICS COVERED",0.55,0.25);
  T(s3,"What We Cover Today",0.55,0.65,7.5,0.7,{fontSize:28,bold:true,color:C.white,fontFace:"Calibri"});
  // Topics left column
  topics.slice(0,6).forEach((t,i)=>{
    const col=i%2,row=Math.floor(i/2);
    const tx=col===0?0.55:7.02, ty=1.55+row*0.95;
    card(s3,tx,ty,6.12,0.8,C.acc,true);
    T(s3,String(i+1).padStart(2,"0"),tx+0.14,ty,0.52,0.8,{fontSize:13,bold:true,color:C.acc,valign:"middle",align:"center"});
    T(s3,t,tx+0.74,ty,5.24,0.8,{fontSize:12,color:C.offwhite,valign:"middle",wrap:true});
  });
  // If we have a real-world example, show a teaser at bottom
  if(realWorld){
    R(s3,0.55,5.28,SW-1.1,0.92,C.panel,C.amber,0.6);
    T(s3,"Real-World Example:",0.72,5.28,2.2,0.92,{fontSize:9.5,bold:true,color:C.amber,valign:"middle"});
    T(s3,realWorld.slice(0,130),2.98,5.28,SW-3.68,0.92,{fontSize:9.5,color:C.silver,valign:"middle",wrap:true,italic:true});
  }
  pageNum(s3,sn,total);

  // ── KEY CONCEPT SLIDES ───────────────────────────────────────
  kps.slice(0,6).forEach((pt,i)=>{
    sn++;
    const ks=pres.addSlide(); bgDefault(ks);
    const heading = typeof pt==="object" ? pt.heading : `Concept ${i+1}`;
    const body    = typeof pt==="object" ? (pt.body||"") : String(pt);
    const bullets = typeof pt==="object" && Array.isArray(pt.bullets) ? pt.bullets : [];

    chip(ks,"CONCEPT "+(i+1)+" OF "+Math.min(kps.length,6),0.55,0.25);
    T(ks,heading,0.55,0.65,SW-1.1,0.82,{fontSize:26,bold:true,color:C.white,fontFace:"Calibri",wrap:true});
    R(ks,0.55,1.58,SW-1.1,0.05,C.acc,null);

    const layout=i%3;

    if(layout===0){
      // LEFT: full body explanation   RIGHT: bullet points panel
      card(ks,0.55,1.7,7.8,3.42,C.acc,true);
      // Body text — full rich content
      T(ks,body,0.75,1.82,7.45,3.2,{fontSize:13,color:C.offwhite,valign:"top",wrap:true,lineSpacingMultiple:1.35});

      // Right panel — bullets or insight
      if(bullets.length>0){
        card(ks,8.55,1.7,4.65,3.42,C.acc);
        R(ks,8.55,1.7,4.65,0.42,C.acc,null);
        sf_trans(ks,8.55,1.7,4.65,0.42,35);
        T(ks,"Key Points",8.55,1.7,4.65,0.42,{fontSize:11.5,bold:true,color:C.white,align:"center",valign:"middle"});
        R(ks,8.55,2.12,4.65,0.04,C.acc,null);
        // Each bullet as its own row
        bullets.slice(0,5).forEach((b,bi)=>{
          O(ks,8.68,2.22+bi*0.52,0.18,0.18,C.acc,null);
          T(ks,b,8.94,2.18+bi*0.52,4.12,0.44,{fontSize:10.5,color:C.offwhite,valign:"middle",wrap:true});
        });
      } else {
        // Two boxes: insight + topic
        card(ks,8.55,1.7,4.65,1.58,C.amber);
        R(ks,8.55,1.7,4.65,0.04,C.amber,null);
        T(ks,"Key Insight",8.55,1.7,4.65,0.44,{fontSize:11,bold:true,color:C.amber,align:"center",valign:"middle"});
        T(ks,body.split(".")[0]+".",8.68,2.18,4.4,1.0,{fontSize:10.5,color:C.offwhite,italic:true,wrap:true,align:"center"});
        card(ks,8.55,3.38,4.65,1.74,C.cyan);
        R(ks,8.55,3.38,4.65,0.04,C.cyan,null);
        T(ks,"Related Topic",8.55,3.38,4.65,0.44,{fontSize:11,bold:true,color:C.cyan,align:"center",valign:"middle"});
        T(ks,topics[i]||classTitle,8.68,3.86,4.4,1.18,{fontSize:11,color:C.offwhite,italic:true,wrap:true,align:"center"});
      }

    } else if(layout===1){
      // TOP: body paragraph   BOTTOM: numbered bullet cards
      card(ks,0.55,1.7,SW-1.1,1.62,C.acc,true);
      T(ks,body,0.75,1.8,SW-1.5,1.44,{fontSize:13,color:C.offwhite,valign:"top",wrap:true,lineSpacingMultiple:1.35});

      // Bullet cards row — each bullet gets its own card
      const bpts = bullets.length>0 ? bullets.slice(0,4) : topics.slice(0,4);
      const bcw  = (SW-1.1-(bpts.length-1)*0.18)/bpts.length;
      bpts.forEach((b,bi)=>{
        const bx=0.55+bi*(bcw+0.18), by=3.44;
        R(ks,bx,by,bcw,2.72,C.card,[C.acc,C.amber,C.cyan,C.green][bi%4],0.65);
        R(ks,bx,by,bcw,0.06,[C.acc,C.amber,C.cyan,C.green][bi%4],null);
        // Number
        O(ks,bx+bcw/2-0.28,by+0.08,0.56,0.56,[C.acc,C.amber,C.cyan,C.green][bi%4],null);
        T(ks,String(bi+1),bx+bcw/2-0.28,by+0.08,0.56,0.56,{fontSize:13,bold:true,color:C.white,align:"center",valign:"middle"});
        // Bullet text
        T(ks,b,bx+0.12,by+0.74,bcw-0.24,1.88,{fontSize:10.5,color:C.offwhite,align:"center",valign:"top",wrap:true,lineSpacingMultiple:1.2});
      });

    } else {
      // SPLIT: explanation left | numbered points right
      card(ks,0.55,1.7,6.18,3.42,C.acc,true);
      R(ks,0.55,1.7,6.18,0.42,C.acc,null);
      sf_trans(ks,0.55,1.7,6.18,0.42,35);
      T(ks,"Explanation",0.55,1.7,6.18,0.42,{fontSize:11.5,bold:true,color:C.white,align:"center",valign:"middle"});
      R(ks,0.55,2.12,6.18,0.04,C.acc,null);
      T(ks,body,0.72,2.2,5.9,2.84,{fontSize:12.5,color:C.offwhite,valign:"top",wrap:true,lineSpacingMultiple:1.35});

      card(ks,7.03,1.7,6.17,3.42,C.amber);
      R(ks,7.03,1.7,6.17,0.42,C.amber,null);
      sf_trans(ks,7.03,1.7,6.17,0.42,35);
      T(ks,"Quick Points",7.03,1.7,6.17,0.42,{fontSize:11.5,bold:true,color:C.white,align:"center",valign:"middle"});
      R(ks,7.03,2.12,6.17,0.04,C.amber,null);
      const pts = bullets.length>0 ? bullets.slice(0,5) :
        [body.slice(0,110),body.slice(110,220),body.slice(220,330)].filter(p=>p.trim());
      pts.slice(0,5).forEach((p,pi)=>{
        R(ks,7.16,2.22+pi*0.52,5.9,0.46,C.panel,C.amber,0.4);
        O(ks,7.22,2.3+pi*0.52,0.28,0.28,C.amber,null);
        T(ks,String(pi+1),7.22,2.3+pi*0.52,0.28,0.28,{fontSize:8,bold:true,color:C.white,align:"center",valign:"middle"});
        T(ks,p,7.56,2.22+pi*0.52,5.46,0.46,{fontSize:10.5,color:C.offwhite,valign:"middle",wrap:true});
      });
    }
    pageNum(ks,sn,total);
  });

  // ── KEY CONCEPTS REVIEW SLIDE ────────────────────────────────
  sn++;
  const sd=pres.addSlide(); bgDefault(sd);
  chip(sd,"KEY CONCEPTS REVIEW",0.55,0.25);
  T(sd,"All Key Concepts at a Glance",0.55,0.65,12,0.72,{fontSize:28,bold:true,color:C.white,fontFace:"Calibri"});
  R(sd,0.55,1.5,SW-1.1,0.05,C.acc,null);

  const sdColors=[C.acc,C.amber,C.cyan,C.green];
  const sdKps=kps.slice(0,4);
  const sdCols=sdKps.length<=2?sdKps.length:2;
  const sdRows=Math.ceil(sdKps.length/sdCols);
  const sdGx=0.22,sdGy=0.22;
  const sdBw=(SW-1.1-(sdCols-1)*sdGx)/sdCols;
  const sdBh=(5.6-(sdRows-1)*sdGy)/sdRows;

  sdKps.forEach((pt,i)=>{
    const col=i%sdCols, row=Math.floor(i/sdCols);
    const bx=0.55+col*(sdBw+sdGx), by=1.65+row*(sdBh+sdGy);
    const c2=sdColors[i%4];
    const heading=(typeof pt==="object"?pt.heading:`Concept ${i+1}`);
    const body2=(typeof pt==="object"?pt.body:"").slice(0,180);
    const b1=(typeof pt==="object"&&pt.bullets&&pt.bullets[0])?pt.bullets[0]:"";
    const b2=(typeof pt==="object"&&pt.bullets&&pt.bullets[1])?pt.bullets[1]:"";
    // Card
    R(sd,bx,by,sdBw,sdBh,C.card,c2,0.7);
    R(sd,bx,by,sdBw,0.07,c2,null);
    // Number
    O(sd,bx+0.12,by+0.1,0.5,0.5,c2,null);
    T(sd,String(i+1),bx+0.12,by+0.1,0.5,0.5,{fontSize:12,bold:true,color:C.white,align:"center",valign:"middle"});
    // Heading
    T(sd,heading,bx+0.72,by+0.1,sdBw-0.86,0.5,{fontSize:12,bold:true,color:C.white,valign:"middle",wrap:true});
    // Divider
    R(sd,bx+0.12,by+0.68,sdBw-0.24,0.04,c2,null);
    // Body
    T(sd,body2,bx+0.12,by+0.78,sdBw-0.24,sdBh*0.44,{fontSize:10,color:C.offwhite,valign:"top",wrap:true,lineSpacingMultiple:1.15});
    // Bullets
    if(b1){ O(sd,bx+0.16,by+sdBh*0.44+0.92,0.14,0.14,c2,null); T(sd,b1,bx+0.38,by+sdBh*0.44+0.86,sdBw-0.52,0.34,{fontSize:9,color:C.silver,valign:"middle",wrap:true}); }
    if(b2){ O(sd,bx+0.16,by+sdBh*0.44+1.3,0.14,0.14,c2,null);  T(sd,b2,bx+0.38,by+sdBh*0.44+1.24,sdBw-0.52,0.34,{fontSize:9,color:C.silver,valign:"middle",wrap:true}); }
  });
  pageNum(sd,sn,total);

  // ── SUMMARY SLIDE ────────────────────────────────────────────
  sn++;
  const sL=pres.addSlide(); bgSection(sL);
  chip(sL,"CLASS SUMMARY",0.55,0.1);
  T(sL,"What We Covered",0.55,0.35,12,0.9,{fontSize:32,bold:true,color:C.white,fontFace:"Calibri"});
  // Summary paragraph
  T(sL,summary,0.55,1.35,SW-1.1,0.88,{fontSize:12.5,color:C.silver,italic:true,wrap:true});
  // 4 summary cards
  const sumColors=[C.acc,C.cyan,C.green,C.amber];
  kps.slice(0,4).forEach((pt,i)=>{
    const sx=0.55+i*3.2, sy=2.42, cw=3.0, ch=2.15, col=sumColors[i];
    R(sL,sx,sy,cw,ch,C.card,col,0.7);
    R(sL,sx,sy,cw,0.06,col,null);
    O(sL,sx+cw/2-0.3,sy+0.1,0.6,0.6,col,null);
    T(sL,String(i+1),sx+cw/2-0.3,sy+0.1,0.6,0.6,{fontSize:14,bold:true,color:C.white,align:"center",valign:"middle"});
    const h2=(typeof pt==="object"?pt.heading:"Point "+(i+1)).slice(0,30);
    T(sL,h2,sx+0.1,sy+0.8,cw-0.2,0.56,{fontSize:10.5,bold:true,color:C.white,align:"center",valign:"top",wrap:true});
    const b2=(typeof pt==="object"?(pt.body||""):"").slice(0,100);
    T(sL,b2,sx+0.1,sy+1.42,cw-0.2,0.65,{fontSize:9.5,color:C.silver,align:"center",valign:"top",wrap:true});
  });
  pageNum(sL,sn,total);

  // ── NEXT STEPS SLIDE ─────────────────────────────────────────
  sn++;
  const sN=pres.addSlide(); bgDark(sN);
  R(sN,0,0,SW,0.06,C.acc,null);
  T(sN,"Next Steps",0.55,0.72,SW-1.1,1.0,{fontSize:48,bold:true,color:C.white,fontFace:"Calibri",align:"center"});
  T(sN,classTitle,0.55,1.82,SW-1.1,0.48,{fontSize:14,color:C.acc,align:"center",bold:true});
  R(sN,4.5,2.42,4.3,0.05,C.acc,null);

  const ns=nextSteps.length>0?nextSteps.slice(0,4):["Practice all concepts from this class","Complete the MCQ assignment quiz","Download and review the PDF notes","Read ahead for the next class topic"];
  const nsCol=[C.acc,C.cyan,C.green,C.amber];
  ns.forEach((step,i)=>{
    const cx=0.55+i*3.2, cy=2.68, cw=3.0, ch=2.44, col=nsCol[i];
    R(sN,cx,cy,cw,ch,C.card2,col,0.7);
    R(sN,cx,cy,cw,0.06,col,null);
    O(sN,cx+cw/2-0.35,cy+0.1,0.7,0.7,col,null);
    T(sN,["01","02","03","04"][i],cx+cw/2-0.35,cy+0.1,0.7,0.7,{fontSize:16,bold:true,color:C.white,align:"center",valign:"middle"});
    T(sN,step,cx+0.12,cy+0.92,cw-0.24,1.42,{fontSize:11,color:C.offwhite,align:"center",valign:"top",wrap:true,lineSpacingMultiple:1.25});
  });
  pageNum(sN,sn,total);
}

// Helper: semi-transparent overlay on top of a rect
function sf_trans(sl,x,y,w,h,t){
  sl.addShape(pres.shapes.RECTANGLE,{x,y,w,h,fill:{color:"000000",transparency:t},line:{type:"none"}});
}

// ─── MAIN ──────────────────────────────────────
if(DATA.type==="full_course") buildFullCourse(DATA.course||{});
else buildClassPPTX(DATA);

pres.writeFile({fileName:out})
  .then(()=>{ console.log("OK:"+out); })
  .catch(e=>{ console.error("ERR:"+e.message); process.exit(1); });
