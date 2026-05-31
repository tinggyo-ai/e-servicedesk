// IT 통합 관리 시스템 - 기술 스택 PPT 생성기
const PptxGenJS = require('pptxgenjs');
const path = require('path');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches

// ── 컬러 팔레트 ──────────────────────────
const C = {
  navy:       '1e3a5f',
  blue:       '2e86de',
  lightBlue:  '5dade2',
  skyBlue:    'a9cce3',
  white:      'FFFFFF',
  offWhite:   'f4f6f9',
  dark:       '1a252f',
  gray:       '7f8c8d',
  midGray:    'bdc3c7',
  green:      '27ae60',
  orange:     'e67e22',
  purple:     '8e44ad',
};

const FONT = 'Malgun Gothic';

// ── 유틸 ─────────────────────────────────
function addHeader(slide, title, icon) {
  // 상단 네이비 헤더 바
  slide.addShape('rect', { x: 0, y: 0, w: 13.33, h: 1.15, fill: { color: C.navy } });
  // 파란 왼쪽 포인트 바
  slide.addShape('rect', { x: 0, y: 0, w: 0.25, h: 7.5,  fill: { color: C.blue } });
  // 하단 구분선
  slide.addShape('rect', { x: 0.25, y: 1.15, w: 13.08, h: 0.04, fill: { color: C.blue } });
  // 헤더 타이틀
  slide.addText((icon ? icon + '  ' : '') + title, {
    x: 0.5, y: 0.18, w: 12.5, h: 0.78,
    fontSize: 28, bold: true, color: C.white, fontFace: FONT,
  });
}

function addTag(slide, text, x, y, bgColor, textColor) {
  slide.addShape('roundRect', { x, y, w: text.length * 0.13 + 0.4, h: 0.36, fill: { color: bgColor }, rectRadius: 0.05 });
  slide.addText(text, { x, y, w: text.length * 0.13 + 0.4, h: 0.36,
    fontSize: 11, bold: true, color: textColor || C.white, fontFace: FONT, align: 'center' });
}

function addCard(slide, x, y, w, h, fillColor, text, textOpts) {
  slide.addShape('roundRect', { x, y, w, h, fill: { color: fillColor }, rectRadius: 0.08,
    shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 8, offset: 3, angle: 45 } });
  if (text) slide.addText(text, { x: x + 0.15, y, w: w - 0.3, h, ...textOpts });
}

// ══════════════════════════════════════════════════════════════
//  SLIDE 1 — 타이틀
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };

  // 왼쪽 파란 포인트
  s.addShape('rect', { x: 0, y: 0, w: 0.3, h: 7.5, fill: { color: C.blue } });

  // 오른쪽 장식 원들
  s.addShape('ellipse', { x: 10.5, y: -0.8, w: 4, h: 4, fill: { color: '16304f' }, line: { color: '16304f' } });
  s.addShape('ellipse', { x: 11.2, y: 4.5,  w: 2.5, h: 2.5, fill: { color: '16304f' }, line: { color: '16304f' } });

  // 메인 타이틀
  s.addText('IT 통합 관리 시스템', {
    x: 0.7, y: 1.4, w: 10.5, h: 1.4,
    fontSize: 48, bold: true, color: C.white, fontFace: FONT,
  });

  // 서브타이틀
  s.addText('기술 스택 소개', {
    x: 0.7, y: 2.95, w: 8, h: 0.75,
    fontSize: 30, color: C.lightBlue, fontFace: FONT,
  });

  // 구분선
  s.addShape('rect', { x: 0.7, y: 3.85, w: 9.0, h: 0.05, fill: { color: C.blue } });

  // URL / 설명
  s.addText('🌐  it-manager.fly.dev', {
    x: 0.7, y: 4.1, w: 6, h: 0.5,
    fontSize: 17, color: C.skyBlue, fontFace: FONT,
  });
  s.addText('사내 IT 자산·사용자·티켓을 한 곳에서 관리하는 웹 기반 시스템', {
    x: 0.7, y: 4.65, w: 10, h: 0.5,
    fontSize: 15, color: C.midGray, fontFace: FONT,
  });

  // 하단 태그
  const tags = [['Node.js', C.green], ['Express', '34495e'], ['SQLite', C.orange], ['Fly.io', C.blue], ['Docker', '0db7ed']];
  tags.forEach(([t, c], i) => addTag(s, t, 0.7 + i * 1.85, 6.3, c));
}

// ══════════════════════════════════════════════════════════════
//  SLIDE 2 — 주요 기능
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '주요 기능', '⚙');

  const features = [
    { icon: '🖥', title: '자산 관리',       sub: 'PC / 노트북 / 서버\n네트워크 장비 등 등록·수정·삭제', color: C.blue },
    { icon: '👤', title: '사용자 관리',     sub: '사번·부서·직책·이메일\n담당 자산 연결 관리', color: C.purple },
    { icon: '🎫', title: '장애/작업 티켓', sub: '접수·처리·완료 현황\n우선순위·담당자 지정', color: C.orange },
    { icon: '💿', title: '소프트웨어 관리', sub: '라이선스 수량·만료일\n사용 현황 추적', color: '16a085' },
    { icon: '🔧', title: '유지보수 이력',   sub: '장비별 유지보수 기록\n비용·예정일 관리', color: 'c0392b' },
  ];

  features.forEach((f, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 4.2, y = 1.55 + row * 2.8;
    const w = 3.9, h = 2.4;

    // 카드 배경
    s.addShape('roundRect', { x, y, w, h, fill: { color: C.white }, rectRadius: 0.1,
      shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 6, offset: 2, angle: 45 } });
    // 컬러 상단 바
    s.addShape('roundRect', { x, y, w, h: 0.45, fill: { color: f.color }, rectRadius: 0.1 });
    s.addShape('rect',       { x, y: y + 0.2, w, h: 0.25, fill: { color: f.color } });

    // 아이콘 + 타이틀
    s.addText(f.icon + '  ' + f.title, {
      x: x + 0.15, y: y + 0.55, w: w - 0.3, h: 0.55,
      fontSize: 16, bold: true, color: C.dark, fontFace: FONT,
    });
    // 설명
    s.addText(f.sub, {
      x: x + 0.15, y: y + 1.1, w: w - 0.3, h: 1.1,
      fontSize: 13, color: C.gray, fontFace: FONT, valign: 'top',
    });
  });

  // 우측 하단 추가 기능 태그
  s.addText('+ 엑셀 일괄 등록  |  체크박스 일괄 삭제  |  담당자 일괄 지정  |  모바일 반응형', {
    x: 0.5, y: 6.85, w: 12.5, h: 0.45,
    fontSize: 12, color: C.blue, fontFace: FONT, align: 'center', bold: true,
  });
}

// ══════════════════════════════════════════════════════════════
//  SLIDE 3 — Frontend 기술 스택
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };
  addHeader(s, 'Frontend 기술 스택', '🖌');

  const items = [
    { name: 'HTML5 / CSS3',    badge: '마크업·스타일',     desc: '시맨틱 구조와 커스텀 CSS 변수로 테마 관리',              color: 'e44d26' },
    { name: 'Bootstrap 5.3',   badge: 'UI 프레임워크',     desc: '반응형 그리드, 모달, 배지 등 UI 컴포넌트 활용',          color: '7952b3' },
    { name: 'Font Awesome 6.5',badge: '아이콘',             desc: 'CDN 방식으로 1,600+ 벡터 아이콘 사용',                  color: '339af0' },
    { name: 'SheetJS (xlsx)',   badge: '엑셀 파싱',         desc: '브라우저에서 직접 .xlsx/.xls 파일 읽기·날짜 자동 변환',  color: '1d6f42' },
    { name: 'Vanilla JS',      badge: '프레임워크 없음',   desc: 'fetch API로 REST 통신, SPA 방식 섹션 전환 구현',         color: 'f0db4f' },
    { name: 'SPA 구조',        badge: 'Single Page App',   desc: '새로고침 없이 자산·사용자·티켓 화면 전환',               color: C.lightBlue },
  ];

  items.forEach((item, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 6.4, y = 1.4 + row * 1.9;
    const w = 6.0, h = 1.65;

    s.addShape('roundRect', { x, y, w, h, fill: { color: '243447' }, rectRadius: 0.09,
      shadow: { type: 'outer', color: '000000', opacity: 0.3, blur: 5, offset: 2, angle: 45 } });
    // 왼쪽 컬러 바
    s.addShape('roundRect', { x, y, w: 0.18, h, fill: { color: item.color }, rectRadius: 0.09 });
    s.addShape('rect',       { x: x + 0.09, y, w: 0.09, h, fill: { color: item.color } });

    s.addText(item.name, {
      x: x + 0.35, y: y + 0.15, w: w - 1.5, h: 0.45,
      fontSize: 16, bold: true, color: C.white, fontFace: FONT,
    });
    s.addShape('roundRect', { x: x + w - 1.55, y: y + 0.2, w: 1.45, h: 0.32, fill: { color: item.color, transparency: 80 }, rectRadius: 0.04 });
    s.addText(item.badge, {
      x: x + w - 1.55, y: y + 0.2, w: 1.45, h: 0.32,
      fontSize: 10, bold: true, color: item.color, fontFace: FONT, align: 'center',
    });
    s.addText(item.desc, {
      x: x + 0.35, y: y + 0.65, w: w - 0.5, h: 0.75,
      fontSize: 12.5, color: C.midGray, fontFace: FONT,
    });
  });
}

// ══════════════════════════════════════════════════════════════
//  SLIDE 4 — Backend 기술 스택
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.dark };
  addHeader(s, 'Backend 기술 스택', '⚡');

  const items = [
    { name: 'Node.js v16+',       badge: 'Runtime',      desc: '비동기 I/O 기반 서버 런타임, npm 생태계 활용',            color: C.green },
    { name: 'Express.js 4.18',    badge: 'Framework',    desc: 'REST API 라우팅, JSON 미들웨어, 정적 파일 서빙 동시 처리', color: '000000' },
    { name: 'SQLite',             badge: 'Database',     desc: '파일 기반 경량 DB, 별도 DB 서버 불필요',                  color: C.lightBlue },
    { name: 'better-sqlite3',     badge: 'DB Driver',    desc: 'WAL 모드·트랜잭션으로 성능 최적화, Raw SQL 사용',         color: C.orange },
    { name: 'RESTful JSON API',   badge: 'API 방식',     desc: 'GET·POST·PUT·DELETE 메서드, 일괄 처리 Bulk 엔드포인트',   color: C.purple },
    { name: 'ORM 미사용',         badge: '설계 철학',    desc: '가벼운 구조 유지를 위해 Raw SQL 직접 작성',               color: 'c0392b' },
  ];

  items.forEach((item, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 6.4, y = 1.4 + row * 1.9;
    const w = 6.0, h = 1.65;

    s.addShape('roundRect', { x, y, w, h, fill: { color: '243447' }, rectRadius: 0.09,
      shadow: { type: 'outer', color: '000000', opacity: 0.3, blur: 5, offset: 2, angle: 45 } });
    s.addShape('roundRect', { x, y, w: 0.18, h, fill: { color: item.color }, rectRadius: 0.09 });
    s.addShape('rect',       { x: x + 0.09, y, w: 0.09, h, fill: { color: item.color } });

    s.addText(item.name, {
      x: x + 0.35, y: y + 0.15, w: w - 1.5, h: 0.45,
      fontSize: 16, bold: true, color: C.white, fontFace: FONT,
    });
    s.addShape('roundRect', { x: x + w - 1.55, y: y + 0.2, w: 1.45, h: 0.32, fill: { color: item.color, transparency: 80 }, rectRadius: 0.04 });
    s.addText(item.badge, {
      x: x + w - 1.55, y: y + 0.2, w: 1.45, h: 0.32,
      fontSize: 10, bold: true, color: item.color === '000000' ? C.white : item.color, fontFace: FONT, align: 'center',
    });
    s.addText(item.desc, {
      x: x + 0.35, y: y + 0.65, w: w - 0.5, h: 0.75,
      fontSize: 12.5, color: C.midGray, fontFace: FONT,
    });
  });
}

// ══════════════════════════════════════════════════════════════
//  SLIDE 5 — 배포 & 인프라
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.offWhite };
  addHeader(s, '배포 & 인프라', '🚀');

  const items = [
    { icon: '☁',  title: 'Fly.io',               sub: '글로벌 엣지 배포 플랫폼\n전 세계 어디서나 빠른 접속 가능', color: C.blue },
    { icon: '🐳', title: 'Docker',                sub: 'Dockerfile 기반 이미지 빌드\n환경 일관성 보장', color: '0db7ed' },
    { icon: '💾', title: 'Persistent Volume',     sub: 'Fly.io 영구 볼륨에 SQLite 저장\n재배포해도 데이터 유지', color: C.orange },
    { icon: '🔧', title: '환경변수 분리',         sub: 'DB_PATH / PORT 환경변수로\n로컬·클라우드 동일 코드 운영', color: C.purple },
    { icon: '💻', title: '로컬 실행',             sub: 'node app.js\nhttp://localhost:3000 접속', color: C.green },
  ];

  items.forEach((item, i) => {
    const x = 0.4 + i * 2.52, y = 1.5;
    const w = 2.3, h = 4.5;

    s.addShape('roundRect', { x, y, w, h, fill: { color: C.white }, rectRadius: 0.1,
      shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 8, offset: 3, angle: 45 } });
    s.addShape('roundRect', { x, y, w, h: 0.55, fill: { color: item.color }, rectRadius: 0.1 });
    s.addShape('rect',       { x, y: y + 0.28, w, h: 0.28, fill: { color: item.color } });

    s.addText(item.icon, {
      x, y: y + 0.7, w, h: 0.8,
      fontSize: 30, align: 'center', fontFace: FONT,
    });
    s.addText(item.title, {
      x: x + 0.1, y: y + 1.55, w: w - 0.2, h: 0.65,
      fontSize: 14, bold: true, color: C.dark, fontFace: FONT, align: 'center',
    });
    s.addShape('rect', { x: x + 0.4, y: y + 2.25, w: w - 0.8, h: 0.03, fill: { color: item.color } });
    s.addText(item.sub, {
      x: x + 0.1, y: y + 2.35, w: w - 0.2, h: 1.8,
      fontSize: 11.5, color: C.gray, fontFace: FONT, align: 'center', valign: 'top',
    });
  });
}

// ══════════════════════════════════════════════════════════════
//  SLIDE 6 — 아키텍처 요약
// ══════════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  addHeader(s, '아키텍처 요약', '🏗');

  // ── 3단 아키텍처 플로우 ──
  const boxes = [
    { label: 'Client', sub: '브라우저\n(PC / 모바일)', color: C.lightBlue, x: 0.7 },
    { label: 'Server', sub: 'Node.js\nExpress.js', color: C.blue, x: 4.5 },
    { label: 'Database', sub: 'SQLite\nbetter-sqlite3', color: C.orange, x: 8.3 },
  ];

  boxes.forEach(b => {
    s.addShape('roundRect', { x: b.x, y: 1.6, w: 3.3, h: 2.8, fill: { color: b.color, transparency: 87 },
      line: { color: b.color, width: 2 }, rectRadius: 0.12 });
    s.addText(b.label, {
      x: b.x, y: 1.85, w: 3.3, h: 0.6,
      fontSize: 20, bold: true, color: b.color, fontFace: FONT, align: 'center',
    });
    s.addShape('rect', { x: b.x + 0.6, y: 2.5, w: 2.1, h: 0.04, fill: { color: b.color, transparency: 60 } });
    s.addText(b.sub, {
      x: b.x + 0.1, y: 2.65, w: 3.1, h: 1.4,
      fontSize: 13, color: C.skyBlue, fontFace: FONT, align: 'center',
    });
  });

  // 화살표 레이블
  [
    { x: 4.0, label: 'REST API\nHTTP/JSON' },
    { x: 7.8, label: 'SQL Query\nbetter-sqlite3' },
  ].forEach(a => {
    s.addShape('rect', { x: a.x, y: 2.85, w: 0.45, h: 0.04, fill: { color: C.blue } });
    s.addShape('rect', { x: a.x + 0.45, y: 2.75, w: 0.05, h: 0.24, fill: { color: C.blue } });
    s.addShape('rect', { x: a.x + 0.35, y: 2.72, w: 0.15, h: 0.08, fill: { color: C.blue } });
    s.addText(a.label, {
      x: a.x - 0.1, y: 3.0, w: 0.8, h: 0.55,
      fontSize: 9, color: C.skyBlue, fontFace: FONT, align: 'center',
    });
  });

  // ── 하단 특징 4개 ──
  const features = [
    { icon: '📦', text: '단일 서버 구조\n(Monolith)' },
    { icon: '⚡', text: '외부 의존성\n최소화' },
    { icon: '💰', text: '낮은 운영비용\nFly.io 무료 플랜' },
    { icon: '🔄', text: '필요 시 PostgreSQL\n교체 가능' },
  ];

  features.forEach((f, i) => {
    const x = 0.7 + i * 3.0, y = 5.0;
    s.addShape('roundRect', { x, y, w: 2.7, h: 2.0, fill: { color: '16304f' }, rectRadius: 0.08 });
    s.addText(f.icon, { x, y: y + 0.15, w: 2.7, h: 0.65, fontSize: 20, align: 'center', fontFace: FONT });
    s.addText(f.text, { x, y: y + 0.8, w: 2.7, h: 0.95,
      fontSize: 12, color: C.skyBlue, fontFace: FONT, align: 'center' });
  });

  // 우측 파일 구조
  s.addShape('roundRect', { x: 11.0, y: 1.6, w: 2.1, h: 2.8, fill: { color: '16304f' }, rectRadius: 0.08 });
  s.addText('총 코드 구조', { x: 11.0, y: 1.75, w: 2.1, h: 0.4,
    fontSize: 11, bold: true, color: C.lightBlue, fontFace: FONT, align: 'center' });
  s.addShape('rect', { x: 11.2, y: 2.2, w: 1.7, h: 0.03, fill: { color: C.blue, transparency: 60 } });
  [
    { f: '📄 app.js',       d: 'Backend 전체' },
    { f: '🌐 index.html',   d: 'Frontend 전체' },
    { f: '🗄 .db',          d: 'SQLite DB' },
    { f: '🐳 Dockerfile',   d: '배포 설정' },
  ].forEach((r, i) => {
    s.addText(r.f,  { x: 11.05, y: 2.3 + i * 0.48, w: 2.0, h: 0.38, fontSize: 11, color: C.white,    fontFace: FONT });
    s.addText(r.d,  { x: 11.05, y: 2.3 + i * 0.48 + 0.2, w: 2.0, h: 0.28, fontSize: 9,  color: C.gray, fontFace: FONT });
  });
}

// ══════════════════════════════════════════════════════════════
//  저장
// ══════════════════════════════════════════════════════════════
const outPath = path.join('C:\\Users\\Lenovo\\Desktop', 'IT통합관리시스템_기술스택.pptx');
pptx.writeFile({ fileName: outPath })
  .then(() => console.log('✅ 완성:', outPath))
  .catch(e => console.error('오류:', e));
