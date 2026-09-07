const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625
const F = 'Apple SD Gothic Neo';
const C = { dark:'1F1F1F', orange:'E8622A', white:'FFFFFF', card:'F3F2EF', mid:'6E6E6E', ink:'222222', soft:'FBE9E0', line:'DDDAD4' };
let n = 0;
const TOTAL_HINT = 49;

function base(dark=false){
  const s = pres.addSlide(); n++;
  s.background = { color: dark ? C.dark : C.white };
  s.addText(String(n), { x:9.3, y:5.2, w:0.5, h:0.3, fontFace:F, fontSize:9, color: dark?'8A8A8A':'9A9A9A', align:'right', isTextBox:true, margin:0 });
  return s;
}
function header(s, section, title, dark=false){
  s.addText(section, { x:0.6, y:0.35, w:8, h:0.3, fontFace:F, fontSize:11, color:C.orange, bold:true, isTextBox:true, margin:0 });
  s.addText(title, { x:0.6, y:0.65, w:8.8, h:0.8, fontFace:F, fontSize:28, bold:true, color: dark?C.white:C.ink, isTextBox:true, margin:0, valign:'top' });
}
function circle(s, x, y, txt, d=0.5, fill=C.orange, color=C.white, fs=14){
  s.addShape(pres.ShapeType.ellipse, { x, y, w:d, h:d, fill:{color:fill}, line:{color:fill} });
  s.addText(txt, { x, y, w:d, h:d, fontFace:F, fontSize:fs, bold:true, color, align:'center', valign:'middle', isTextBox:true, margin:0 });
}
function card(s, x, y, w, h, title, body, opt={}){
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill:{color: opt.fill||C.card}, line:{color: opt.fill||C.card}, rectRadius:0.08 });
  let ty=y+0.15;
  if (opt.num){ circle(s, x+0.2, y+0.18, opt.num, 0.4, C.orange, C.white, 12); }
  if (title){ s.addText(title, { x:x+(opt.num?0.7:0.2), y:ty, w:w-(opt.num?0.9:0.4), h:0.45, fontFace:F, fontSize:opt.ts||15, bold:true, color:opt.tc||C.ink, isTextBox:true, margin:0, valign:'middle' }); ty+=0.5; }
  if (body){ s.addText(body, { x:x+0.2, y:ty, w:w-0.4, h:h-(ty-y)-0.15, fontFace:F, fontSize:opt.bs||12.5, color:opt.bc||C.mid, isTextBox:true, margin:0, valign:'top', paraSpaceAfter:4 }); }
}
function statement(section, big, sub, notes){
  const s = base(true);
  s.addText(section, { x:0.6, y:0.5, w:8, h:0.3, fontFace:F, fontSize:11, color:C.orange, bold:true, isTextBox:true, margin:0 });
  s.addShape(pres.ShapeType.ellipse, { x:7.6, y:-0.9, w:3.2, h:3.2, fill:{color:C.orange}, line:{color:C.orange} });
  s.addText(big, { x:0.6, y:1.5, w:7.6, h:2.2, fontFace:F, fontSize:34, bold:true, color:C.white, isTextBox:true, margin:0, valign:'middle' });
  if (sub) s.addText(sub, { x:0.6, y:3.9, w:8.6, h:1.0, fontFace:F, fontSize:15, color:'CFCFCF', isTextBox:true, margin:0, valign:'top' });
  if (notes) s.addNotes(notes);
  return s;
}
function bullets(s, x, y, w, h, items, fs=14, color=C.ink){
  s.addText(items.map((t,i)=>({ text:t, options:{ bullet:true, breakLine:i<items.length-1, paraSpaceAfter:6 } })), { x, y, w, h, fontFace:F, fontSize:fs, color, isTextBox:true, margin:0, valign:'top' });
}
function table(s, x, y, w, rows, colW, opt={}){
  const data = rows.map((r,ri)=> r.map((c,ci)=> ({ text:c, options:{ fontFace:F, fontSize: opt.fs||12, bold: ri===0 || (opt.boldFirstCol && ci===0), color: ri===0? C.white : C.ink, fill:{ color: ri===0? C.orange : (ri%2? C.white: C.card) }, align:'left', valign:'middle', margin:[4,6,4,6] } })));
  s.addTable(data, { x, y, w, colW, border:{ type:'solid', pt:0.5, color:C.line }, rowH: opt.rowH||0.42 });
}

// ============ 열기 ============
{ const s = base(true);
  s.addShape(pres.ShapeType.ellipse, { x:6.9, y:1.2, w:4.2, h:4.2, fill:{color:C.orange}, line:{color:C.orange} });
  s.addText('2026 소상공인 디지털 마케팅 실무교육 · AI를 활용한 SNS 마케팅', { x:0.6, y:0.6, w:7, h:0.3, fontFace:F, fontSize:12, color:C.orange, bold:true, isTextBox:true, margin:0 });
  s.addText('1회차', { x:0.6, y:1.3, w:6, h:0.8, fontFace:F, fontSize:20, color:'CFCFCF', isTextBox:true, margin:0 });
  s.addText('SNS 세팅 +\n우리 비즈니스에 맞는\nSNS 마케팅 전략 수립', { x:0.6, y:1.9, w:6.5, h:2.2, fontFace:F, fontSize:34, bold:true, color:C.white, isTextBox:true, margin:0, valign:'top' });
  s.addText('2026. 9. 7 (월) 09:30 – 17:30', { x:0.6, y:4.5, w:6, h:0.4, fontFace:F, fontSize:14, color:'CFCFCF', isTextBox:true, margin:0 });
  s.addText('오늘 들고 나가는 것\n① 가치×페르소나 시트\n② 내 첫 스킬 (ChatGPT 프로젝트에 쌓인 파일)\n③ 우리를 소개하는 이미지 한 장\n④ 첫 주 게시물 3개', { x:7.3, y:2.1, w:3.4, h:2.4, fontFace:F, fontSize:12.5, bold:false, color:C.white, isTextBox:true, margin:0, valign:'middle', paraSpaceAfter:3 });
  s.addNotes('여러분 대부분 파는 게 있으세요. 오늘은 그걸 바꾸는 날이 아닙니다. 그게 누구의 무엇을 바꾸는지 적는 날이에요. 아직 문 안 여신 분, 정하는 중인 분도 같은 흐름으로 갑니다.');
}
statement('개념 ① 드릴과 구멍', '사람들은 드릴이 아니라,\n벽에 뚫린 구멍을 산다.', '손님이 사는 것은 상품이 아니라, 그 상품이 만드는 변화다.\n오늘 하는 일은 여러분의 드릴 옆에 구멍을 써넣는 것.',
 '드릴을 사는 사람은 드릴이 갖고 싶은 게 아니에요. 벽에 구멍이 필요한 거죠. 손님이 반찬을 사는 것도 반찬이 갖고 싶어서가 아니라 퇴근하고 밥만 하면 되는 저녁이 필요해서예요.');
{ const s = base(); header(s,'개념 ① 드릴과 구멍','반찬가게로 보면');
  card(s,0.6,1.7,3.9,2.6,'드릴 — 파는 것','조미료 안 쓰는 반찬\n매일 아침 새로 만듦\n1인분 소포장',{ts:16,bs:14});
  s.addShape(pres.ShapeType.rightArrow,{x:4.65,y:2.7,w:0.7,h:0.6,fill:{color:C.orange},line:{color:C.orange}});
  card(s,5.5,1.7,3.9,2.6,'구멍 — 손님에게 일어나는 변화','퇴근하고 밥만 하면\n되는 저녁\n\n배달 앱을 덜 켠다',{fill:C.soft,ts:16,bs:14});
  s.addText('"우리는 반찬을 파는 게 아니라, 퇴근하고 밥만 하면 되는 저녁을 판다."', { x:0.6, y:4.55, w:8.8, h:0.5, fontFace:F, fontSize:15, bold:true, color:C.orange, isTextBox:true, margin:0 });
  s.addNotes('아이템은 하나도 안 바꿨어요. 앞세우는 것만 바뀌었습니다. 조미료는 "왜 우리인가"의 근거로 내려갔고, 변화는 "저녁이 해결된다"로 올라갔어요.');
}
{ const s = base(); header(s,'개념 ① 드릴과 구멍','변화와 가치는 다르다');
  table(s,0.6,1.6,8.8,[['','변화 (코어)','가치'],['누가 정하나','사장님이 선언한다','손님이 말한다'],['주어','우리','손님'],['몇 개인가','하나','손님 유형 수만큼'],['틀렸을 때','우리가 뭘 하는 곳인지 다시 생각','다시 쓰면 된다']],[1.8,3.5,3.5],{fs:13,rowH:0.5,boldFirstCol:true});
  s.addText('판정 기준 — 손님을 바꿨을 때 그 문장도 바뀌어야 가치다. 안 바뀌면 변화를 베껴 쓴 것.', { x:0.6, y:4.4, w:8.8, h:0.6, fontFace:F, fontSize:14, bold:true, color:C.orange, isTextBox:true, margin:0 });
  s.addNotes('두 가지를 헷갈리면 안 돼요. 변화는 사장님이 정합니다. 하나예요. 가치는 손님이 말해요. 손님 종류만큼 있어요.');
}
statement('개념 ② AI는 데이터로 답한다', 'AI의 본질은\n무지막지하게 많은 데이터.\n그래서 그냥 물으면\n세상 평균의 답이 나온다.', '"반찬가게 인스타 소개글 써줘" → 어느 반찬가게에나 붙을 수 있는 글.\n그건 우리 글이 아니다. 우리 것으로 만드는 방법은 하나 — 맥락을 설명한다.',
 'AI는 데이터가 무지막지하게 많아요. 세상 모든 반찬가게, 모든 네일샵이 들어 있어요. 그래서 그냥 물으면 세상 평균 답이 나옵니다. 우리 답을 받으려면 맥락을 줘야 해요.');
{ const s = base(); header(s,'개념 ② 맥락은 네 칸으로 · STIC','모르면 캡처해서, 네 줄 붙여 물어본다');
  const items=[['S','상황 Situation','지금 어디서 뭘 하다가 이 화면이 떴나','카카오 파트너센터에서 채널 만드는 중인데 이 화면이 떴어요 (캡처)'],['T','할 일 Task','뭘 해달라는 건가','다음에 어디 눌러야 하는지 알려줘'],['I','의도 Intent','왜, 결국 뭘 하려고','오늘 안에 채널 공개하고 웰컴 메시지까지 넣으려고'],['C','걱정 Concern','뭘 피하고 싶나','잘못 눌러서 결제되거나 삭제될까 봐']];
  items.forEach((it,i)=>{ const x=0.6+i*2.2; card(s,x,1.6,2.05,3.4,null,null,{}); circle(s,x+0.2,1.8,it[0],0.55,C.orange,C.white,18);
    s.addText(it[1],{x:x+0.85,y:1.8,w:1.1,h:0.55,fontFace:F,fontSize:12,bold:true,color:C.ink,isTextBox:true,margin:0,valign:'middle'});
    s.addText(it[2],{x:x+0.2,y:2.5,w:1.7,h:0.7,fontFace:F,fontSize:12,color:C.ink,isTextBox:true,margin:0});
    s.addText(it[3],{x:x+0.2,y:3.3,w:1.7,h:1.5,fontFace:F,fontSize:11,color:C.mid,italic:true,isTextBox:true,margin:0}); });
  s.addNotes('모르는 것은 캡처해서 AI한테 물어보세요. 화면이든 문서든 리뷰든. 네 줄을 붙이면 세상 평균이 아니라 내 상황에 답합니다. 오늘 채널 세팅에서 실제로 이렇게 합니다.');
}
{ const s = base(); header(s,'개념 ② STIC','오늘 오전 인터뷰가 곧 우리 비즈니스의 STIC이다');
  table(s,0.6,1.6,8.8,[['STIC','코어의 어느 칸','파일'],['S 상황','사실 시트 · 손님 한 명 · 재료','03 · 01'],['T 할 일','첫 주 게시물 3개 · 채널 문안','05 · 04'],['I 의도','한 문장 — 우리는 ___를 판다','01'],['C 걱정','말하면 안 되는 것 · 안 보여줄 것 · 안 쓰는 말','03 · 02']],[1.6,5.2,2.0],{fs:13,rowH:0.5,boldFirstCol:true});
  s.addText('그래서 프로젝트 안에서는 할 일(T) 한 줄만 치면 된다.', { x:0.6, y:4.4, w:8.8, h:0.5, fontFace:F, fontSize:14, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'개념 ③ 스킬','스킬 = STIC을 매번 안 쳐도 되게 묶어둔 것');
  s.addText('단골 손님이 오면 "늘 드시던 걸로요" 한마디로 끝난다.\n처음 온 손님에게는 메뉴 설명부터 해야 한다.\n스킬은 AI를 단골로 만드는 방법이다.', { x:0.6, y:1.55, w:8.8, h:1.0, fontFace:F, fontSize:14, color:C.ink, isTextBox:true, margin:0 });
  const parts=[['규칙','이렇게 하고, 이건 하지 마라 (Concern)','프로젝트 「지침」\n00 파일'],['재료','우리 사실 · 우리 손님 · 우리 말투 (Situation · Intent)','프로젝트 「파일」\n01~05'],['반복','다음에 또 같은 방식으로','프로젝트 안 새 대화\nTask 한 줄만']];
  parts.forEach((p,i)=>{ const x=0.6+i*2.95; card(s,x,2.75,2.8,2.3,p[0],p[1]+'\n\nChatGPT에서는 → '+p[2],{num:String(i+1),ts:16,bs:12}); });
  s.addNotes('STIC은 좋은데 매번 네 줄 치는 건 힘들어요. 자기 상황을 글로 설명하는 것 자체가 어려운 분도 있어요. 그럴 때 쓰는 게 스킬이에요.');
}
{ const s = base(); header(s,'개념 ③ 스킬','스킬은 하나, 프로젝트도 하나 — 두 번 만난다');
  card(s,0.6,1.6,4.3,2.5,'오전 — 지침이 질문한다','강사가 넣어드린 지침 + 참조 묶음\n20개 질문을 순서대로 던지고 파일 다섯 개를 만든다\nSTIC을 못 쓰는 분도 답만 하면 STIC이 적힌다\n\n→ 인터뷰 모드',{ts:15,bs:12.5});
  card(s,5.1,1.6,4.3,2.5,'오후 — 파일이 쌓이면 동료가 된다','여러분 답이 파일 01~05로 같은 프로젝트에 올라간다\n같은 지침이 알아서 마케팅 모드로 바뀐다\n이 안에서는 "카드뉴스 써줘" 한 줄이면 된다\n\n→ 마케팅 모드',{fill:C.soft,ts:15,bs:12.5});
  s.addText('STIC을 직접 쓸 수 있으면 쓰고, 어려우면 스킬 안에서 한 줄만 친다. 어느 쪽이든 AI는 우리 맥락 위에서 답한다.', { x:0.6, y:4.3, w:8.8, h:0.4, fontFace:F, fontSize:13.5, bold:true, color:C.orange, isTextBox:true, margin:0 });
  s.addText('2회차 카드뉴스 스킬 → 3회차 이미지 스킬 → 4회차 릴스 스킬이 이 프로젝트 안에 쌓인다.', { x:0.6, y:4.75, w:8.8, h:0.4, fontFace:F, fontSize:12.5, color:C.mid, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'개념 ④ 사실 시트','AI가 지어내지 않게');
  card(s,0.6,1.6,4.3,3.3,'사실 시트가 없으면','AI가 빈칸을 데이터로 채운다. 그게 본질이라 막을 수 없다.\n\n"30년 전통"\n"국내산 100%"\n"망원동 1등"\n\n이미지도 같다 — 없는 간판, 없는 인테리어를 그린다',{ts:15,bs:13});
  card(s,5.1,1.6,4.3,3.3,'사실 시트가 있으면','말해도 되는 사실을 미리 적어둔다 (Situation)\n\n없는 건 [확인 필요]로 표시하게 지침에 넣는다 (Concern)\n\n2·3·4회차 검수가 전부 이 시트 기준',{fill:C.soft,ts:15,bs:13});
}
{ const s = base(); header(s,'열기','오늘의 흐름 — 일곱 블록');
  const blocks=[['1','열기','개념 넷 · GPT 열기'],['2','코어 인터뷰','A~F · 한 문장'],['3','한 달 방향','첫 주 3개 · 촬영 리스트'],['4','내 첫 스킬','파일 5 → 같은 프로젝트'],['5','소개 이미지','한 장 · 검수 5'],['6','채널 세팅','인스타 · 카카오 · 플레이스'],['7','공유와 닫기','한 문장 낭독']];
  blocks.forEach((b,i)=>{ const col=i%4, row=Math.floor(i/4); const x=0.6+col*2.2, y=1.6+row*1.75; card(s,x,y,2.05,1.55,b[1],b[2],{num:b[0],ts:14,bs:11.5}); });
  s.addText('오전 = 블록 1~2\n오후 = 블록 3~7', { x:7.3, y:3.5, w:2.1, h:1.3, fontFace:F, fontSize:12, color:C.mid, isTextBox:true, margin:0, valign:'middle' });
}
{ const s = base(); header(s,'열기','오늘의 규칙 여섯');
  const rules=['모르면 캡처해서 STIC 네 줄로 묻는다','모르겠다고 쳐도 된다. 빈칸은 다음 주 손님한테 물어볼 목록','AI 문장은 꼭 한 단어라도 사장님 말로 고친다','비밀번호 · 인증번호는 AI에게 절대 주지 않는다. 직접 친다','한 달 목표는 끊기지 않는 것. 주 1회','옆 사람이 막히면 "그래서 손님한테 뭐가 달라져요?" 한 번만'];
  rules.forEach((r,i)=>{ const y=1.6+i*0.58; circle(s,0.6,y,String(i+1),0.42,C.orange,C.white,13); s.addText(r,{x:1.2,y:y,w:8.2,h:0.42,fontFace:F,fontSize:14,color:C.ink,isTextBox:true,margin:0,valign:'middle'}); });
}
{ const s = base(true); header(s,'시작 · 10분','프로젝트 만들기 — 오늘 하나만',true);
  const st=['ChatGPT 왼쪽 「프로젝트」 → 「새 프로젝트」 → 이름 「(상호) 마케팅」','「지침」 열기 → 단톡방의 지침 파일 내용 전부 붙여넣기 → 저장','「파일」 → 참조 묶음 파일 한 개 올리기','새 대화에 시작 문구'];
  st.forEach((t,i)=>{ const y=1.55+i*0.62; circle(s,0.6,y+0.06,String(i+1),0.46,C.orange,C.white,14); s.addText(t,{x:1.25,y:y,w:5.2,h:0.58,fontFace:F,fontSize:14,color:C.white,isTextBox:true,margin:0,valign:'middle'}); });
  s.addShape(pres.ShapeType.roundRect,{x:6.7,y:1.55,w:2.7,h:1.15,fill:{color:C.white},line:{color:C.white},rectRadius:0.08});
  s.addText('"비즈니스 GTM 코어\n시작할게요"',{x:6.8,y:1.55,w:2.5,h:1.15,fontFace:F,fontSize:15,bold:true,color:C.ink,align:'center',valign:'middle',isTextBox:true,margin:0});
  s.addShape(pres.ShapeType.roundRect,{x:6.7,y:2.85,w:2.7,h:1.15,fill:{color:C.soft},line:{color:C.soft},rectRadius:0.08});
  s.addText('창업 준비 중이면\n"창업 준비 중인데 시작할게요"',{x:6.8,y:2.85,w:2.5,h:1.15,fontFace:F,fontSize:12.5,bold:true,color:C.ink,align:'center',valign:'middle',isTextBox:true,margin:0});
  s.addText('첫 질문이 뜨면 손 들어주세요. 「프로젝트」 메뉴가 없으면 웹(chatgpt.com)으로.',{x:0.6,y:4.3,w:8.8,h:0.6,fontFace:F,fontSize:13,color:C.orange,bold:true,isTextBox:true,margin:0});
}

// ============ 코어 인터뷰 ============
{ const s = base(); header(s,'블록 2 · 코어 인터뷰','세 상태 — 셋 다 같은 방에서 간다');
  const st=[['팔고 있다','재료 4문항 그대로','단골의 말 · 처음 온 사람의 질문 · 안 오는 이유 · 리뷰'],['팔 것은 정함 · 손님 없음','재료를 「예정 재료」로','왜 이걸 팔기로 했나 · 내가 손님이었을 때 · 써본 사람 말 · 참고하는 곳'],['뭘 팔지 미정','씨앗 찾기 3문항 → 후보 → 합류','하는 일 · 자꾸 묻는 것 · 공짜로 해준 것. 강사가 옆에 앉는다']];
  st.forEach((t,i)=>{ const x=0.6+i*2.95; card(s,x,1.6,2.8,3.1,t[0],t[1]+'\n\n'+t[2],{num:String(i+1),ts:12.5,bs:12,fill:i==0?C.soft:C.card}); });
  s.addText('다른 스킬로 보내지 않는다. 질문만 바뀐다.', { x:0.6, y:4.85, w:8.8, h:0.4, fontFace:F, fontSize:13.5, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 2 · A 재료','재료는 새로 만드는 게 아니라 이미 쌓여 있는 말');
  const q=[['단골이 남한테 우리를 소개할 때 뭐라고 하나요?','손님 언어로 된 가치'],['처음 온 손님이 꼭 묻는 게 있어요?','Before. 묻는 건 불안한 것'],['한 번 오고 다시 안 오는 분들은 왜 안 올까요?','돌아선 사람의 말. 모르는 게 정상'],['리뷰에 반복되는 칭찬이나 불만이 있어요?','실제 문장. 가치 힌트']];
  q.forEach((it,i)=>{ const y=1.6+i*0.82; circle(s,0.6,y+0.1,String(i+1),0.45,C.orange,C.white,13); s.addText(it[0],{x:1.25,y:y,w:5.4,h:0.65,fontFace:F,fontSize:14,bold:true,color:C.ink,isTextBox:true,margin:0,valign:'middle'}); s.addText(it[1],{x:6.8,y:y,w:2.6,h:0.65,fontFace:F,fontSize:12,color:C.mid,isTextBox:true,margin:0,valign:'middle'}); });
  s.addText('[조사] 소개 뒤에 상호로 한 번 검색해요. 이미 올라온 리뷰가 있으면 재료가 됩니다. 5분.', { x:0.6, y:4.95, w:8.8, h:0.35, fontFace:F, fontSize:12, color:C.orange, bold:true, isTextBox:true, margin:0 });
  s.addNotes('세 번째 질문은 대부분 모르세요. 모르는 게 정상이에요. 짐작으로 하나만 치고 넘어가세요. 그 빈칸이 오늘 숙제가 됩니다.');
}
{ const s = base(); header(s,'블록 2 · A 재료','온라인이면 · 손님이 없으면');
  card(s,0.6,1.6,4.3,3.3,'온라인만 판다면','단골의 말 → 두 번 이상 산 분의 후기 · 문의\n처음 온 사람의 질문 → 스토어 Q&A · 톡톡\n안 오는 이유 → 장바구니만 담는 사람 · 반품 사유\n리뷰 → 스토어 리뷰',{ts:15,bs:13});
  card(s,5.1,1.6,4.3,3.3,'손님이 아직 없다면 — 예정 재료','왜 이걸 팔기로 하셨어요?\n사장님이 손님이었을 때 뭐가 불편했어요?\n이미 써본 사람이 뭐라고 했어요? ("없어요"도 답)\n참고하는 가게 손님들은 뭘 좋아해요?',{fill:C.soft,ts:15,bs:13});
}
{ const s = base(); header(s,'블록 2 · B 변화','변화는 관찰이 아니라 선언 — 주어는 우리');
  const steps=['가장 만족한 손님 한 분이 얻어 간 것','Before / After — 한 번만','대안 — 안 사면 대신 뭘 하나','대체 불가 — 우리 쪽 / 사장님 두 칸','한 문장'];
  steps.forEach((t,i)=>{ const x=0.6+i*1.78; circle(s,x+0.45,1.8,String(i+1),0.7,i==4?C.dark:C.orange,C.white,20); s.addText(t,{x:x-0.05,y:2.7,w:1.7,h:1.2,fontFace:F,fontSize:12.5,bold:true,color:C.ink,align:'center',isTextBox:true,margin:0,valign:'top'}); if(i<4) s.addShape(pres.ShapeType.line,{x:x+1.2,y:2.15,w:0.5,h:0,line:{color:C.line,width:1.5}}); });
  s.addText('첫 질문이 "손님"인 이유 — 우리가 뭘 바꾸는지는 우리 머릿속보다 그 손님한테서 이미 일어난 일에 있다.', { x:0.6, y:4.3, w:8.8, h:0.6, fontFace:F, fontSize:13.5, color:C.mid, isTextBox:true, margin:0 });
}
statement('블록 2 · 되묻기 한 문장', '"그건 우리가 하는 일이고,\n그래서 손님 하루에서\n뭐가 달라지나요?"', '메뉴 설명 · 기능 · 공정이 나오면 GPT가 이렇게 되묻는다. 이 질문 하나를 가져가면 된다.');
{ const s = base(); header(s,'블록 2 · 대안','우리 걸 안 사면, 그분은 대신 어떻게 하나요?');
  card(s,0.6,1.6,2.8,2.6,'다른 가게를 간다','더 좋은 걸 만드는 싸움',{ts:15,bs:13});
  card(s,3.55,1.6,2.8,2.6,'집에서 대충 한다','',{ts:15,bs:13});
  card(s,6.5,1.6,2.9,2.6,'그냥 안 한다','움직이게 만드는 싸움\n\n카드뉴스 첫 장이 여기서 나온다',{fill:C.soft,ts:15,bs:13});
  s.addText('경쟁 가게를 묻는 게 아니다. "그냥 안 한다"가 나오면 싸움의 성격이 바뀐다.', { x:0.6, y:4.4, w:8.8, h:0.5, fontFace:F, fontSize:14, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 2 · 대체 불가','왜 우리여야 하나요 — 두 칸으로');
  card(s,0.6,1.6,4.3,2.9,'① 우리 쪽 기준','재료 · 만드는 방식 · 시간 · 위치 · 가격 · 공간 · 배송\n\n"퀄리티가 달라요" → 손님이 그걸 어떻게 알아채나요? 보이는 것으로',{ts:16,bs:13});
  card(s,5.1,1.6,4.3,2.9,'② 사장님 기준 — 확인 가능한 것만','자격 · 년수 · 직접 하는 것 · 전에 했던 일\n\n"정성 · 진심 · 열정" → 손님이 그걸 어떻게 알아채나요?\n안 나오면 빈칸 ☐. 비어도 된다',{fill:C.soft,ts:16,bs:13});
}
{ const s = base(); header(s,'블록 2 · 한 문장','"우리는 ___를 파는 게 아니라, ___를 판다"');
  s.addText('두 칸. 세 칸으로 벌리면 선언이 아니라 설명이 된다. 초안 셋 중 고르고, 꼭 한 단어라도 사장님 말로.', { x:0.6, y:1.5, w:8.8, h:0.5, fontFace:F, fontSize:13, color:C.mid, isTextBox:true, margin:0 });
  table(s,0.6,2.1,8.8,[['업종','한 문장'],['반찬가게','반찬을 파는 게 아니라, 퇴근하고 밥만 하면 되는 저녁을 판다'],['네일샵','네일을 파는 게 아니라, 2주에 한 번 나를 챙기는 시간을 판다'],['온라인몰 (운동복)','운동복을 파는 게 아니라, 사이즈 고민 없이 결제 버튼을 누르는 10초를 판다'],['주문형 떡공방','떡을 파는 게 아니라, 답례 고민이 끝나는 한 통의 카톡을 판다'],['꽃 수업 (창업 예정)','꽃 수업을 파는 게 아니라, 평일 저녁 두 시간 아무 생각 안 하는 시간을 판다']],[2.0,6.8],{fs:12.5,rowH:0.44,boldFirstCol:true});
  s.addNotes('뒷칸이 전부 손님의 하루에서 달라지는 것이에요. 상품 특징이 아니에요.');
}
{ const s = base(); header(s,'블록 2 · D 손님 한 분','"다요"는 답이 아니다');
  card(s,0.6,1.6,4.3,3.1,'이렇게 나오면','"동네 사람 다요"\n"30대 여성"\n"직장인들"',{ts:15,bs:15});
  card(s,5.1,1.6,4.3,3.1,'이렇게 되묻는다','우리가 없으면 제일 아쉬워할 한 분은요?\n\n지난달 오신 분 중 제일 잘 맞았던 한 분\n\n그분은 오기 직전에 뭘 하다가 막혔을까요?',{fill:C.soft,ts:15,bs:13.5});
}
{ const s = base(); header(s,'블록 2 · C 가치 — 그분이 다녀가고 주변에 뭐라고 말할까요','"거기 가보니까 ___" — 그 손님의 말');
  const t=[['손님 입에서 나올 말인가','"신선한 재료" ✕ → "여긴 반찬이 안 남아요" ○'],['손님을 바꾸면 이 문장도 바뀌는가','안 바뀌면 변화를 베껴 쓴 것'],['옆 가게도 똑같이 말할 수 있는가','말할 수 있으면 아직 카테고리 설명']];
  t.forEach((it,i)=>{ const y=1.65+i*0.95; circle(s,0.6,y+0.15,String(i+1),0.5,C.orange,C.white,15); s.addText(it[0],{x:1.3,y:y,w:4.0,h:0.8,fontFace:F,fontSize:15,bold:true,color:C.ink,isTextBox:true,margin:0,valign:'middle'}); s.addText(it[1],{x:5.4,y:y,w:4.0,h:0.8,fontFace:F,fontSize:12.5,color:C.mid,isTextBox:true,margin:0,valign:'middle'}); });
  s.addText('실제로 들은 말, 네이버 리뷰 문장이 제일 좋다. 지어내는 게 아니라 꺼내는 것. 카피가 아니다 — 광고 문장은 2회차.', { x:0.6, y:4.6, w:8.8, h:0.5, fontFace:F, fontSize:13, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 2 · E 사장님','말투는 새로 만들지 않는다 — 지금까지 친 문장이 말투다');
  card(s,0.6,1.6,2.8,3.2,'말투 표본','GPT가 여러분이 실제로 친 문장 세 개를 그대로 보여준다\n\n"그냥 매일 새로 해요"\n"저는 그거 안 해요"',{ts:15,bs:12.5});
  card(s,3.55,1.6,2.8,3.2,'안 쓰는 말','"핫플" · "인생 맛집" · 이모지 남발 · 반말\n\n카드뉴스 · 릴스 자막 전부 이 기준',{ts:15,bs:12.5});
  card(s,6.5,1.6,2.9,3.2,'안 보여줄 것','얼굴 · 주방 · 가격 · 가족 · 매출\n\n적어두면 AI가 그 소재를 제안하지 않는다\n오후 이미지의 규칙이 된다',{fill:C.soft,ts:15,bs:12.5});
}
{ const s = base(); header(s,'블록 2 · F 사실 시트','오늘 유일하게 표를 통째로 — 생각이 아니라 기억이라서');
  bullets(s,0.6,1.6,4.2,3.2,['정식 상호 · 검색되는 이름','주소 · 찾아오는 법 · 주차 (온라인: 배송 · 소요일)','영업시간 · 휴무 (온라인: 주문 마감)','대표 상품 3개와 가격','원산지 · 인증 · 자격 — 증명 가능한 것만','시작한 연도. "거의 30년"은 안 된다','절대 하지 않는 것 — 사실인 것만'],13);
  card(s,5.1,1.6,4.3,3.2,'말하면 안 되는 것 — [조사]로 찾는다','시트를 받으면 우리 업종 규정을 웹에서 찾아요\n예: "천연화장품"은 인증 기준 · "화학성분 없음"은 오인 표현\n\n공통: 최초 · 최고 · 유일 · 1등 · 효능 · 비교 · 시트에 없는 숫자 · 없는 후기',{fill:C.soft,ts:14,bs:12.5});
  s.addNotes('이게 없으면 AI가 30년 전통을 지어냅니다. 식품은 효능, 미용은 재생·탄력, 학원은 합격률, 운동은 감량 보장. 증명 못 하면 안 씁니다.');
}

// ============ 한 달 방향 ============
{ const s = base(); header(s,'블록 3 · 한 달 방향','손님이 있는 곳 — 사전 설문 8명');
  const stats=[['5/8','네이버 검색 · 플레이스로 온다'],['6/8','인스타 비즈니스 계정 이미 있다'],['0/8','카카오채널 있다'],['5/8','최근 한 달 게시물 0개']];
  stats.forEach((st,i)=>{ const x=0.6+i*2.2; card(s,x,1.6,2.05,2.3,null,null,{}); s.addText(st[0],{x:x,y:1.75,w:2.05,h:0.9,fontFace:F,fontSize:36,bold:true,color:C.orange,align:'center',isTextBox:true,margin:0}); s.addText(st[1],{x:x+0.15,y:2.75,w:1.75,h:1.0,fontFace:F,fontSize:12,color:C.ink,align:'center',isTextBox:true,margin:0}); });
  s.addText('플레이스는 새로 안 만든다. 소개글만 오늘 문안으로 바꾼다. 카카오채널은 전원 신규. 빈도는 주 1회.', { x:0.6, y:4.2, w:8.8, h:0.7, fontFace:F, fontSize:14, bold:true, color:C.ink, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 3 · 채널','채널마다 역할이 다르다');
  table(s,0.6,1.6,8.8,[['채널','역할','오늘'],['네이버 플레이스 (이미 있음)','손님이 실제로 찾아오는 길','소개글만 갱신. 10분'],['인스타그램 비즈니스','처음 보는 사람이 발견하는 곳','이름 칸에 지역 + 업종. 프로필 문안'],['카카오톡 채널','한 번 온 사람에게 다시 말 걸기 · 재구매 알림','전원 신규 개설. 웰컴 메시지'],['스레드','사장님 말투로 짧게. 사진 없이 두 줄','선택. 안 올리던 분의 첫 습관']],[2.4,3.4,3.0],{fs:12.5,rowH:0.5,boldFirstCol:true});
  s.addText('채널이 이미 4개 이상이면 — 한 달 손 안 댈 채널 하나를 고른다. 세팅이 아니라 정리.\n[조사] 손님이 실제로 칠 검색어 5개 · 앞으로 석 달의 계기 3개를 찾아 이름 칸과 4주 테마에 넣는다.', { x:0.6, y:4.4, w:8.8, h:0.75, fontFace:F, fontSize:12.5, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 3 · 세 기둥','아침에 적은 3층이 그대로 한 달 콘텐츠가 된다');
  const p=[['내가 파는 것','① 상품 · 공간','사실을 보여주는 사진 · 메뉴 · 공간 · 포장','3회차 이미지'],['사는 사람','② 손님의 변화','후기 · 전후 · "이런 분이 오세요"','2회차 카드뉴스'],['나라는 사람','③ 사장님','얼굴 · 손 · 작업 · 한마디','4회차 릴스']];
  p.forEach((it,i)=>{ const x=0.6+i*2.95; card(s,x,1.6,2.8,3.1,null,null,{fill:i==1?C.soft:C.card}); s.addText(it[0],{x:x+0.2,y:1.75,w:2.4,h:0.35,fontFace:F,fontSize:11,color:C.orange,bold:true,isTextBox:true,margin:0}); s.addText(it[1],{x:x+0.2,y:2.1,w:2.4,h:0.5,fontFace:F,fontSize:17,bold:true,color:C.ink,isTextBox:true,margin:0}); s.addText(it[2],{x:x+0.2,y:2.7,w:2.4,h:1.0,fontFace:F,fontSize:12.5,color:C.mid,isTextBox:true,margin:0}); s.addText(it[3],{x:x+0.2,y:3.9,w:2.4,h:0.5,fontFace:F,fontSize:13,bold:true,color:C.ink,isTextBox:true,margin:0}); });
  s.addText('셋 중 제일 자신 있는 기둥이 첫 주 첫 게시물.', { x:0.6, y:4.85, w:8.8, h:0.4, fontFace:F, fontSize:13.5, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 3 · 첫 주 게시물 3개','방향만 주고 끝내지 않는다 — 오늘 구체적으로 정한다');
  table(s,0.6,1.6,8.8,[['칸','무엇','예 (반찬가게)'],['주제','한 줄','무말랭이 — 조미료 없이 매일 새로'],['찍을 것','폰으로 찍을 사진 · 영상 1~2개. 구체적으로','아침에 무치는 손 + 담긴 통 클로즈업'],['형식','사진 1장 / 사진 3장 / 15초 영상 / 스레드 글','사진 2장'],['첫 줄','사장님 말투로','"오늘도 무말랭이는 아침에 무쳤어요"']],[1.4,3.7,3.7],{fs:12.5,rowH:0.5,boldFirstCol:true});
  s.addText('기둥마다 후보 둘 — 전부 오늘 여러분이 친 말에서. 고르면 네 칸을 채운다. 사진이 어려운 분은 사진 1장 · 스레드 글로만.', { x:0.6, y:4.45, w:8.8, h:0.6, fontFace:F, fontSize:13, color:C.mid, isTextBox:true, margin:0 });
  s.addNotes('설문에서 제일 많이 막힌다고 하신 게 "뭘 올려야 할지"였어요. 그래서 방향만 드리고 끝내지 않습니다. 첫 주 세 개를 지금 정해요.');
}
{ const s = base(); header(s,'블록 3 · 촬영 리스트 · 빈도','이번 주 폰으로 찍어올 10장 · 첫 달은 끊기지 않는 것');
  card(s,0.6,1.6,4.3,3.2,'촬영 리스트 10장','대표 상품 3장\n손 · 재료 · 작업 장면 2장\n간판(입구) 또는 스토어 대표 이미지 1장\n앉는 자리 또는 포장 상태 1장\n첫 주 게시물용 3장\n\n밝을 때 · 정면 · 배경 단순. 잘 찍을 필요 없다',{ts:15,bs:12.5});
  card(s,5.1,1.6,4.3,3.2,'빈도 — 주 1회','게시물 주 1회 + 스레드 두 줄 주 2회\n한 달 게시물 4~6개\n\n지금 한 달에 0개인 분에게 주 3회는 실패하려고 짜는 계획\n\n1주 오늘 정한 3개 → 2주 상품 → 3주 손님의 변화(카드뉴스) → 4주 사장님(릴스)',{fill:C.soft,ts:15,bs:12.5});
  s.addNotes('"사진 만들기가 어렵다"의 절반은 찍을 게 안 정해져서예요. 열 장 목록을 드립니다. 3회차에 AI가 다듬어요.');
}

// ============ 내 첫 스킬 ============
{ const s = base(); header(s,'블록 4 · 내 첫 스킬','파일 다섯 개 = 재료. 규칙은 지침에 이미 있다');
  table(s,0.6,1.6,8.8,[['파일','스킬의 어느 부분','가는 곳'],['(지침 — 오프닝에 붙인 것)','규칙 — 이렇게 하고 이건 하지 마라','「지침」 이미 있음'],['01 가치×페르소나 시트','재료 — 누구에게 무엇을','「파일」'],['02 사장님 보이스','재료 — 어떤 말투로 · 안 보여줄 것','「파일」'],['03 사실 시트','재료 — 말해도 되는 것','「파일」'],['04 채널 세팅','재료 — 채널별 문안 · 소개 이미지','「파일」'],['05 콘텐츠 방향','재료 — 첫 주 3개 · 촬영 리스트 · 2~4회차 출발점','「파일」']],[2.6,4.6,1.6],{fs:12,rowH:0.42,boldFirstCol:true});
  s.addNotes('이제 이 프로젝트가 여러분 첫 스킬이에요. 오전엔 제 지침이 질문했고, 지금 올린 다섯 파일은 여러분 답이에요. 파일이 올라가면 같은 지침이 알아서 마케팅 동료로 바뀝니다.');
}
{ const s = base(); header(s,'블록 4 · 같은 프로젝트에 올리기','세 단계');
  const st=['인터뷰가 만든 01~05 다섯 파일을 내려받는다','지금 이 프로젝트 → 「파일」 → 다섯 개 업로드','새 대화 → 테스트 질문. 올라가면 지침이 스스로 마케팅 모드로 바뀐다'];
  st.forEach((t,i)=>{ const y=1.6+i*0.6; circle(s,0.6,y+0.05,String(i+1),0.45,C.orange,C.white,13); s.addText(t,{x:1.25,y:y,w:5.0,h:0.55,fontFace:F,fontSize:14,color:C.ink,isTextBox:true,margin:0,valign:'middle'}); });
  s.addShape(pres.ShapeType.roundRect,{x:6.5,y:1.6,w:2.9,h:3.0,fill:{color:C.card},line:{color:C.card},rectRadius:0.08});
  s.addText('화면 캡처\n(전날 최신 화면으로)',{x:6.5,y:1.6,w:2.9,h:3.0,fontFace:F,fontSize:12,color:C.mid,align:'center',valign:'middle',isTextBox:true,margin:0});
  s.addText('무료 계정에서 한도에 걸리면 참조 묶음을 지우고 올린다. 인터뷰는 끝났으니 없어도 된다.', { x:0.6, y:4.7, w:8.8, h:0.4, fontFace:F, fontSize:12.5, color:C.mid, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 4 · 테스트','전원 같은 질문 하나');
  s.addShape(pres.ShapeType.roundRect,{x:0.6,y:1.6,w:8.8,h:0.9,fill:{color:C.dark},line:{color:C.dark},rectRadius:0.08});
  s.addText('"우리를 손님 말로 세 줄만 소개해줘"',{x:0.8,y:1.6,w:8.4,h:0.9,fontFace:F,fontSize:20,bold:true,color:C.white,isTextBox:true,margin:0,valign:'middle'});
  table(s,0.6,2.75,8.8,[['결과','뜻','조치'],['시트 안의 사실 · 사장님 말투 · 손님 말','정상','끝'],['"30년 전통" · 시트에 없는 숫자','03이 안 올라감','03 확인'],['여전히 인터뷰 질문을 함','파일 이름이 다르거나 덜 올라감','01~05 파일명 확인'],['광고체 · 이모지 남발','02가 안 읽힘','02 업로드 확인'],['다른 손님에게 말함','01이 안 읽힘','01 업로드 확인']],[3.4,3.2,2.2],{fs:12,rowH:0.4});
}

// ============ 소개 이미지 ============
statement('블록 5 · 단 컷 소개 이미지', '한 장으로\n"우리는 ___를 판다"가\n보이는 이미지', '드릴(상품)만 덩그러니가 아니라 구멍(변화)이 보이는 장면.\n인스타 · 카카오채널 · 스레드에 같은 이미지 — 채널 세 곳이 같은 얼굴이 된다.',
 '오늘 이미지는 한 장이에요. 세 장 고르느라 시간 쓰지 않습니다.');
{ const s = base(); header(s,'블록 5 · 두 길','사진이 있으면 사진이 먼저다');
  card(s,0.6,1.6,4.3,3.1,'A. 내 사진 + AI 정리 (기본)','폰에 상품 · 매장 사진이 있다\n\n사진 1장 올리고 배경 정리 · 밝기 · 여백만\n"없는 건 그리지 마"\n\n없는 것을 그려 넣지 않는다',{fill:C.soft,ts:16,bs:13});
  card(s,5.1,1.6,4.3,3.1,'B. AI 생성','사진이 없다 — 창업 예정 · 온라인몰 초기\n\n프로젝트가 01 · 02 · 03 보고 쓴 프롬프트 3안 → 고르기\n\n시트에 있는 것만 그린다',{ts:16,bs:13});
  s.addText('A가 기본인 이유 — 사실 시트 원칙과 같다. AI가 그린 매장은 손님이 와서 "여기가 거기예요?" 하게 만든다.', { x:0.6, y:4.85, w:8.8, h:0.4, fontFace:F, fontSize:13, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 5 · 프롬프트','구도만 다른 3안 — 고르고, 사장님 말로 한 줄 고친다');
  const p=[['①','상품 클로즈업'],['②','손이 작업하는 장면'],['③','손님이 받는 순간']];
  p.forEach((it,i)=>{ const x=0.6+i*2.95; card(s,x,1.6,2.8,1.5,null,null,{}); circle(s,x+0.2,1.85,it[0],0.5,C.orange,C.white,15); s.addText(it[1],{x:x+0.85,y:1.85,w:1.8,h:0.5,fontFace:F,fontSize:15,bold:true,color:C.ink,isTextBox:true,margin:0,valign:'middle'}); s.addText('영어 프롬프트. 겁내지 않는다 — 우리 파일 보고 GPT가 썼다',{x:x+0.2,y:2.45,w:2.4,h:0.55,fontFace:F,fontSize:11,color:C.mid,isTextBox:true,margin:0}); });
  bullets(s,0.6,3.4,8.8,1.6,['구멍이 보이는 장면으로. 상품만 덩그러니 놓지 않는다','03 사실 시트에 있는 것만. 없는 간판 · 인테리어 · 인증 마크 없음','02 「안 보여줄 것」 없음 · 타사 로고 없음','글자 없음 — 한글은 자주 틀린다. 글자는 내일 캔바에서 · 정방형 1:1 · 자연광 · 배경 단순','고친 한 줄 예 — "우리 포장은 종이봉투예요. 그거 넣어주세요"'],12.5);
}
{ const s = base(); header(s,'블록 5 · 검수 다섯 가지','올리기 전에');
  const ck=['사실 시트에 없는 것이 그려졌나 — 간판 · 인테리어 · 인증 · 메뉴','「안 보여줄 것」이 들어갔나 — 얼굴 · 주방','글자가 있으면 오기 · 깨짐 → 글자 없는 버전으로','타사 로고 · 상표 · 유명 캐릭터','이 한 장만 보고 "구멍"이 보이나 — 드릴만 보이면 구도 ② ③으로'];
  ck.forEach((r,i)=>{ const y=1.6+i*0.62; s.addShape(pres.ShapeType.rect,{x:0.6,y:y+0.08,w:0.38,h:0.38,fill:{color:C.white},line:{color:C.orange,width:2}}); s.addText(r,{x:1.2,y:y,w:8.2,h:0.55,fontFace:F,fontSize:14,color:C.ink,isTextBox:true,margin:0,valign:'middle'}); });
  s.addText('한 장이면 된다.', { x:0.6, y:4.8, w:8.8, h:0.4, fontFace:F, fontSize:14, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 5 · 채널별 쓰임','같은 이미지, 다섯 자리');
  table(s,0.6,1.6,8.8,[['채널','어디에','비고'],['인스타 첫 게시물','첫 주 게시물 ① 또는 ②','캡션은 05의 첫 줄'],['인스타 프로필','원형으로 잘린다 — 가운데에 핵심','얼굴 안 보여주면 상품 클로즈업'],['카카오채널 프로필 · 커버','커버는 가로로 재생성 또는 여백 확장',''],['스레드 첫 글','1:1 그대로','글은 사장님의 고집 한 줄'],['네이버 플레이스','실제 사진만','AI 이미지 안 씀. 플레이스는 사실의 자리']],[2.4,3.4,3.0],{fs:12.5,rowH:0.46,boldFirstCol:true});
}

// ============ 채널 ============
{ const s = base(true); header(s,'블록 6 · 채널 얼굴 맞추기','막히면 캡처 + 네 줄',true);
  const it=[['S','카카오 파트너센터에서 채널 만드는 중인데 이 화면이 떴어요'],['T','다음에 어디 눌러요?'],['I','오늘 안에 채널 공개까지'],['C','잘못 눌러서 결제나 삭제될까 봐']];
  it.forEach((x,i)=>{ const y=1.65+i*0.62; circle(s,0.6,y+0.05,x[0],0.48,C.orange,C.white,15); s.addText(x[1],{x:1.3,y:y,w:8.0,h:0.55,fontFace:F,fontSize:15,color:C.white,isTextBox:true,margin:0,valign:'middle'}); });
  s.addText('네 줄이 귀찮으면 마케팅 프로젝트 안에서 "이 화면에서 다음 어디 눌러요?" 한 줄. 거기엔 우리 상황이 이미 들어 있다.\n비밀번호 · 인증번호 · 서류는 본인만 입력한다.', { x:0.6, y:4.25, w:8.8, h:0.9, fontFace:F, fontSize:12.5, color:'CFCFCF', isTextBox:true, margin:0 });
  s.addNotes('앱 화면은 자주 바뀌어요. 설명과 다른 화면이 나오면 캡처해서 GPT에 올리고 네 줄을 붙이세요. STIC 카드 나눠드립니다.');
}
{ const s = base(); header(s,'블록 6 · 인스타그램','같은 문안, 같은 이미지 — 25분');
  table(s,0.6,1.6,5.4,[['칸','규칙'],['이름','상호 · 지역 · 업종 — 검색에 걸리는 칸'],['소개 (150자)','1줄 변화 · 2줄 손님 말 · 3줄 사실 하나 · 4줄 행동 하나'],['링크','하나. 카카오채널 또는 플레이스'],['프로필 사진','소개 이미지'],['첫 게시물','소개 이미지 + 05의 첫 줄'],['하이라이트','메뉴 · 손님 · 사장님 — 이름만, 비워둔다']],[1.4,4.0],{fs:11.5,rowH:0.42,boldFirstCol:true});
  card(s,6.3,1.6,3.1,3.1,'소개글 예 (반찬가게)','저녁 걱정 없는 퇴근길을 만듭니다\n"여기 반찬은 안 남아요" — 단골 말\n망원동 · 화~토 11:00–20:00 · 월 휴무\n▼ 카카오채널 추가하면 이번 주 반찬을 먼저',{fill:C.soft,ts:13,bs:11.5});
  s.addText('이미 비즈니스 계정인 분이 대부분. 전환 필요한 분만 보조. 전환하면 계정이 공개로 바뀐다.', { x:0.6, y:4.75, w:8.8, h:0.4, fontFace:F, fontSize:12, color:C.mid, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 6 · 카카오톡 채널','전원 신규 — 40분 · 파트너센터 8단계');
  const st=['business.kakao.com 로그인 → 카카오 비즈니스 통합 회원 전환','「새 채널 만들기」','채널 이름(상호) · 검색용 아이디(못 바꾼다) · 카테고리','프로필 → 채널 정보 → 고객센터 연락처 등록 ★ 빠뜨리면 메시지 못 보냄','「채널 공개하기」 ✔ 「검색 허용하기」 ✔','메시지 → 웰컴 메시지 교체','프로필 · 커버 = 소개 이미지 · 채널 추가 QR 내려받기','(선택) 비즈니스 심사 — 서류 준비되면'];
  st.forEach((t,i)=>{ const col=i<4?0:1, row=i%4; const x=0.6+col*4.45, y=1.6+row*0.72; circle(s,x,y+0.1,String(i+1),0.42,C.orange,C.white,12); s.addText(t,{x:x+0.55,y:y,w:3.8,h:0.65,fontFace:F,fontSize:12,color:C.ink,isTextBox:true,margin:0,valign:'middle'}); });
  s.addText('서류 · 비밀번호는 본인이. 가입이 막히면 대부분 카카오 계정 비밀번호 문제 — 본인 폰으로 재설정.', { x:0.6, y:4.65, w:8.8, h:0.4, fontFace:F, fontSize:12, color:C.mid, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 6 · 웰컴 메시지','사장님 말투 2~3줄 · 뭘 보내줄지 하나만 · 팔지 않는다');
  s.addShape(pres.ShapeType.roundRect,{x:0.6,y:1.7,w:5.0,h:2.2,fill:{color:'FEF3C7'},line:{color:'FEF3C7'},rectRadius:0.15});
  s.addText('채널 추가 고맙습니다 🙂\n매주 화요일에 이번 주 반찬 사진만 하나 보내드릴게요.\n필요한 거 있으시면 여기로 편하게 말씀 주세요.',{x:0.9,y:1.85,w:4.4,h:1.9,fontFace:F,fontSize:14,color:C.ink,isTextBox:true,margin:0,valign:'middle',paraSpaceAfter:6});
  card(s,5.9,1.7,3.5,2.2,'재료','02 말투 표본의 어미\n05의 카카오채널 줄 — 1주 웰컴 · 3주 첫 메시지 1회',{ts:14,bs:12.5});
  s.addText('QR은 계산대 · 포장에 붙인다. 3회차 배너 이미지의 재료가 된다.', { x:0.6, y:4.2, w:8.8, h:0.5, fontFace:F, fontSize:13, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 6 · 네이버 플레이스 · 스레드','플레이스 10분 · 스레드는 선택');
  card(s,0.6,1.6,4.3,3.2,'네이버 플레이스 — 소개글만 갱신','스마트플레이스 → 업체 정보 → 소개글\n\n1줄 변화(사장님 말로) · 2줄 손님 말 · 3줄 사실 하나\n대표 키워드: 지역 · 업종 · 대표 상품\n대표 사진은 실제 사진만\n\n온라인몰은 스토어 첫 화면 소개 문장',{ts:15,bs:12.5});
  card(s,5.1,1.6,4.3,3.2,'스레드 (원하는 분만)','앱 설치 → Instagram으로 로그인\n소개 한 줄 — 말투 그대로\n"망원동에서 반찬 만듭니다. 조미료는 안 써요."\n\n첫 글 = 소개 이미지 + 고집 한 줄\n주 2회, 사진 없이 두 줄',{fill:C.soft,ts:15,bs:12.5});
}

// ============ 공유와 닫기 ============
{ const s = base(); header(s,'블록 7 · 공유','한 문장 + 소개 이미지 + 첫 게시물 첫 줄');
  card(s,0.6,1.6,4.3,3.0,'페어 10분','옆 사람에게 한 문장 읽어주기\n→ 상대가 "그래서 손님 하루에서 뭐가 달라져요?" 한 번 묻기\n→ 역할 교대',{ts:16,bs:13.5});
  card(s,5.1,1.6,4.3,3.0,'낭독 — 원하는 분 8명','한 문장 + 이미지 화면에 + 첫 줄\n\n강사가 한 번 되묻는다\n"이 이미지만 보고 구멍이 보이나요?"\n"손님을 바꾸면 이 말도 바뀌나요?"',{fill:C.soft,ts:16,bs:13.5});
  s.addText('소리 내서 읽어보면 세 칸인지 두 칸인지 바로 들린다.', { x:0.6, y:4.8, w:8.8, h:0.4, fontFace:F, fontSize:13, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
statement('블록 7 · 닫기', '오늘 세운 건\n오늘 기준의 답입니다.\n바뀔 겁니다.\n바뀌는 게 정상이에요.', '빈칸이 있어도 괜찮다. 뭐가 비었는지 아는 것 자체가 다음 주에 손님한테 뭘 물어볼지의 목록이다.');
{ const s = base(); header(s,'블록 7 · 닫기','반응이 없을 때 — 셋 중 어디가 틀렸나');
  table(s,0.6,1.6,8.8,[['이런 신호면','틀린 것','하는 일'],['프로필까지 오는데 팔로우 안 하고 나감','사람 (D)','페르소나 교체 → 가치 다시'],['보고 "좋네요" 하는데 안 옴','이유 (C)','가치 재정의'],['와서 보고 "아 이런 데였어요?"','전달','카드뉴스 · 이미지 · 릴스를 바꿈 — 2회차부터']],[3.6,1.8,3.4],{fs:13,rowH:0.55,boldFirstCol:false});
  s.addText('코어가 없으면 이게 전부 "SNS는 우리랑 안 맞네" 하나로 뭉친다. 코어를 먼저 세운 이유.', { x:0.6, y:4.2, w:8.8, h:0.6, fontFace:F, fontSize:14, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 7 · 숙제와 2회차','숙제는 하나');
  card(s,0.6,1.6,4.3,3.0,'촬영 리스트 10장','3회차 전까지. 잘 찍을 필요 없다\n\n(선택) 이번 주 손님 세 분에게\n"어떻게 알고 오셨어요?"',{fill:C.soft,ts:16,bs:13.5});
  card(s,5.1,1.6,4.3,3.0,'2회차 — 내일','마케팅 프로젝트 안에서 시작. 첫 질문 다시 안 한다\n\n05 파일의 「2회차」 줄이 첫 카드뉴스 주제\nBefore에서 시작해 After로. 근거는 사실 시트',{ts:16,bs:13.5});
  s.addText('퇴장 전 — 파일 여섯 개 위치 · 프로젝트 이름 · 소개 이미지 저장 · 촬영 리스트 캡처', { x:0.6, y:4.8, w:8.8, h:0.4, fontFace:F, fontSize:12.5, color:C.mid, isTextBox:true, margin:0 });
}
{ const s = base(); header(s,'블록 7 · 닫기','스킬이 쌓인다');
  const st=[['1회차','비즈니스 GTM 코어','파일 5 · 소개 이미지'],['2회차','카드뉴스 스킬','재사용 카드뉴스'],['3회차','이미지 프롬프트 스킬','용도별 이미지 4종'],['4회차','릴스 기획 스킬','완성 릴스 1편']];
  st.forEach((it,i)=>{ const x=0.6+i*2.2; card(s,x,1.7,2.05,2.4,null,null,{fill:i==0?C.soft:C.card}); circle(s,x+0.2,1.9,String(i+1),0.5,i==0?C.orange:C.dark,C.white,15); s.addText(it[0],{x:x+0.85,y:1.9,w:1.1,h:0.5,fontFace:F,fontSize:12,color:C.mid,isTextBox:true,margin:0,valign:'middle'}); s.addText(it[1],{x:x+0.2,y:2.55,w:1.7,h:0.7,fontFace:F,fontSize:14,bold:true,color:C.ink,isTextBox:true,margin:0}); s.addText(it[2],{x:x+0.2,y:3.3,w:1.7,h:0.7,fontFace:F,fontSize:11.5,color:C.mid,isTextBox:true,margin:0}); if(i<3) s.addShape(pres.ShapeType.rightArrow,{x:x+2.05,y:2.75,w:0.15,h:0.3,fill:{color:C.line},line:{color:C.line}}); });
  s.addText('전부 오늘 만든 마케팅 프로젝트 안에 파일로 들어간다. 4회 끝나면 AI가 우리를 단골처럼 아는 상태.', { x:0.6, y:4.4, w:8.8, h:0.6, fontFace:F, fontSize:14, bold:true, color:C.orange, isTextBox:true, margin:0 });
}
statement('닫기', '집에 가서 모르는 게 나오면\n캡처해서 네 줄 붙여 물어보세요.\n상황 · 할 일 · 의도 · 걱정.', '오늘 배운 것 중에 제일 오래 갈 것. 네 줄이 힘들면 마케팅 프로젝트 안에서 한 줄만.',
 '그게 오늘 배운 것 중에 제일 오래 갈 거예요.');

pres.writeFile({ fileName: 'out.pptx' }).then(()=>console.log('slides:', n));
