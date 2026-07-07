import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const artifactToolPath =
  "C:/Users/User/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactToolPath).href);

const root = process.cwd();
const outDir = path.join(root, "outputs");
const previewDir = path.join(outDir, "ppt_preview");
const deckPath = path.join(outDir, "주거복지_내비게이터_1차심사_발표자료.pptx");

const screenshots = {
  dashboard: path.join(outDir, "prototype_ppt_dashboard_crop.png"),
  detail: path.join(outDir, "prototype_ppt_detail_crop.png"),
  map: path.join(outDir, "prototype_ppt_map_crop.png"),
  evidence: path.join(outDir, "prototype_ppt_evidence_crop.png"),
  policy: path.join(outDir, "prototype_ppt_policy_crop.png"),
  workflow: path.join(root, "워크 플로우.png"),
};

const W = 1280;
const H = 720;
const C = {
  bg: "#F5F8FA",
  ink: "#111827",
  muted: "#52616B",
  line: "#D5E0E5",
  panel: "#FFFFFF",
  soft: "#EAF2F4",
  teal: "#0B6B57",
  tealDark: "#08493E",
  blue: "#245B8A",
  amber: "#A15D12",
  rose: "#A43B44",
};

const noLine = { style: "solid", fill: "#00000000", width: 0 };
const lightLine = { style: "solid", fill: C.line, width: 1 };

async function readImageBlob(imagePath) {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function shape(slide, geometry, left, top, width, height, options = {}) {
  return slide.shapes.add({
    geometry,
    position: { left, top, width, height },
    fill: options.fill ?? "#00000000",
    line: options.line ?? noLine,
    borderRadius: options.borderRadius,
    shadow: options.shadow,
  });
}

function text(slide, value, left, top, width, height, options = {}) {
  const box = shape(slide, "rect", left, top, width, height, {
    fill: options.fill ?? "#00000000",
    line: options.line ?? noLine,
  });
  box.text = value;
  box.text.fontSize = options.size ?? 20;
  box.text.color = options.color ?? C.ink;
  box.text.bold = Boolean(options.bold);
  box.text.typeface = options.typeface ?? "Malgun Gothic";
  box.text.alignment = options.align ?? "left";
  box.text.verticalAlignment = options.valign ?? "top";
  box.text.insets = options.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
  return box;
}

function panel(slide, left, top, width, height, options = {}) {
  return shape(slide, "roundRect", left, top, width, height, {
    fill: options.fill ?? C.panel,
    line: options.line ?? lightLine,
    borderRadius: "rounded-lg",
  });
}

function header(slide, eyebrow, title, subtitle) {
  text(slide, eyebrow, 72, 44, 860, 26, { size: 14, bold: true, color: C.teal });
  text(slide, title, 72, 78, 1080, 78, { size: 36, bold: true });
  if (subtitle) {
    text(slide, subtitle, 72, 160, 1040, 44, { size: 18, color: C.muted });
  }
}

function footer(slide, page, source = "") {
  text(slide, String(page).padStart(2, "0"), 72, 662, 44, 22, {
    size: 12,
    bold: true,
    color: C.muted,
  });
  if (source) {
    text(slide, source, 128, 662, 1000, 24, { size: 10, color: C.muted });
  }
}

function noteBand(slide, value) {
  panel(slide, 72, 604, 1136, 42, { fill: "#EAF5F1" });
  text(slide, "발표 메모", 96, 617, 90, 16, {
    size: 11,
    bold: true,
    color: C.teal,
  });
  text(slide, value, 188, 614, 960, 22, {
    size: 15,
    bold: true,
    color: C.tealDark,
  });
}

function bullet(slide, items, left, top, width, lineHeight = 40, options = {}) {
  items.forEach((item, index) => {
    const y = top + index * lineHeight;
    shape(slide, "ellipse", left, y + 9, 9, 9, { fill: options.dot ?? C.teal });
    text(slide, item, left + 22, y, width - 22, lineHeight, {
      size: options.size ?? 18,
      color: options.color ?? C.ink,
      bold: Boolean(options.bold),
    });
  });
}

function chip(slide, value, left, top, width, options = {}) {
  panel(slide, left, top, width, 32, {
    fill: options.fill ?? "#E8F4EF",
    line: { style: "solid", fill: options.fill ?? "#E8F4EF", width: 1 },
  });
  text(slide, value, left + 10, top + 7, width - 20, 18, {
    size: options.size ?? 11,
    bold: true,
    color: options.color ?? C.tealDark,
    align: "center",
  });
}

function speaker(slide, lines) {
  slide.speakerNotes.textFrame.setText(lines.join("\n"));
  slide.speakerNotes.setVisible(true);
}

async function addImage(slide, imagePath, left, top, width, height, alt) {
  slide.images.add({
    blob: await readImageBlob(imagePath),
    contentType: "image/png",
    alt,
    fit: "contain",
    position: { left, top, width, height },
  });
}

const deck = Presentation.create({ slideSize: { width: W, height: H } });

function newSlide() {
  const slide = deck.slides.add();
  slide.background.fill = C.bg;
  return slide;
}

// 1. Title
{
  const slide = newSlide();
  text(slide, "2026 국토교통 서비스 발굴 경연 | 1차 심사", 72, 56, 600, 28, {
    size: 15,
    bold: true,
    color: C.teal,
  });
  text(slide, "주거복지\n내비게이터", 72, 112, 520, 132, {
    size: 52,
    bold: true,
  });
  text(
    slide,
    "흩어진 주거복지 기회를 사용자 조건별 우선순위와 원문 근거로 연결합니다.",
    72,
    270,
    590,
    72,
    { size: 24, color: C.muted },
  );
  panel(slide, 72, 386, 570, 86, { fill: "#EAF5F1" });
  text(
    slide,
    "판정 시스템이 아니라 신청 전 사전 점검과 탐색을 돕는 의사결정 보조 서비스입니다.",
    100,
    414,
    514,
    30,
    { size: 19, bold: true, color: C.tealDark },
  );

  const cards = [
    ["1", "조건 입력", "소득·자산·가구·무주택·지역·통근"],
    ["2", "공고 매칭", "공개 공고와 제도를 같은 기준으로 비교"],
    ["3", "근거 설명", "점수·주의조건·원문 근거·RAG 답변"],
  ];
  cards.forEach(([num, title, body], index) => {
    const y = 132 + index * 140;
    panel(slide, 742, y, 420, 102);
    shape(slide, "ellipse", 764, y + 26, 50, 50, { fill: C.teal });
    text(slide, num, 764, y + 36, 50, 24, {
      size: 21,
      bold: true,
      color: "#FFFFFF",
      align: "center",
    });
    text(slide, title, 842, y + 24, 240, 26, { size: 22, bold: true });
    text(slide, body, 842, y + 58, 290, 28, { size: 15, color: C.muted });
  });
  noteBand(slide, "대체가 아니라 보완, 판정이 아니라 사전 점검이라는 포지션을 먼저 고정합니다.");
  footer(slide, 1);
  speaker(slide, [
    "안녕하세요. 저희는 주거복지 내비게이터를 제안합니다.",
    "이 서비스는 당첨을 예측하거나 자격을 확정하지 않고, 공식 공고를 사용자 조건에 맞게 구조화해 신청 전 우선순위를 안내합니다.",
  ]);
}

// 2. Problem
{
  const slide = newSlide();
  header(
    slide,
    "문제정의",
    "정보가 없는 것이 아니라, 내 조건으로 해석하기 어렵습니다",
    "공식 포털과 공고는 존재하지만 사용자는 여러 사이트와 공고문을 오가며 조건을 직접 비교해야 합니다.",
  );

  panel(slide, 72, 238, 510, 300);
  text(slide, "사용자가 직접 대입해야 하는 조건", 104, 268, 360, 30, {
    size: 22,
    bold: true,
  });
  [
    ["소득", 104, 326, 84],
    ["자산", 202, 326, 84],
    ["무주택", 300, 326, 98],
    ["가구 구성", 412, 326, 112],
    ["희망 지역", 104, 374, 112],
    ["청약통장", 230, 374, 112],
    ["모집기간", 356, 374, 112],
    ["제출서류", 104, 422, 112],
    ["입지·통근", 230, 422, 112],
    ["우선순위", 356, 422, 112],
  ].forEach(([label, x, y, w]) => chip(slide, label, x, y, w));
  text(
    slide,
    "결국 핵심 질문은\n“내가 지금 무엇을 먼저 신청해야 하나?”입니다.",
    104,
    488,
    390,
    48,
    { size: 20, bold: true, color: C.tealDark },
  );

  panel(slide, 640, 238, 568, 300);
  text(slide, "현재 여정에서 막히는 지점", 672, 268, 330, 30, {
    size: 22,
    bold: true,
  });
  bullet(
    slide,
    [
      "공고 검색, 자가진단, 신청, 지도 확인이 기능별로 분리",
      "소득·자산·무주택 기준이 공고문 문장과 표 안에 흩어짐",
      "기준 확인 항목과 주의 조건을 사용자가 직접 판별",
      "기회 누락 또는 조건 불일치 공고 탐색 비용 발생",
    ],
    672,
    326,
    470,
    46,
    { size: 18 },
  );
  noteBand(slide, "문제는 정보 부족이 아니라 조건 해석 비용과 공고별 우선순위 판단입니다.");
  footer(slide, 2);
  speaker(slide, [
    "현재 문제는 정보가 없다는 것이 아닙니다.",
    "마이홈, LH청약플러스, 지자체 포털에는 정보가 많지만 사용자는 본인 조건을 공고문에 직접 대입해야 합니다.",
  ]);
}

// 3. Evidence
{
  const slide = newSlide();
  header(
    slide,
    "근거자료",
    "주거비 부담과 정보 탐색 복잡성은 현재 진행형입니다",
    "공식자료를 주 근거로 두고, 2026년 뉴스 검색 결과는 최근성 보조 근거로 활용합니다.",
  );
  const rows = [
    ["OECD", "주거비 과부담과 affordability를 주요 정책 이슈로 제시"],
    ["마이홈", "공공주택 모집공고, 주거복지서비스, 자가진단, 지역별 검색 제공"],
    ["LH청약플러스", "공고문, 청약연습, 청약신청, 청약자격확인 등 단계별 기능 제공"],
    ["온통청년·서울주거포털", "청년정책 검색, 상담, 지자체 주거지원이 별도 운영"],
  ];
  rows.forEach(([label, body], i) => {
    const y = 242 + i * 70;
    panel(slide, 72, y, 540, 52);
    chip(slide, label, 94, y + 10, 132, { fill: "#E8F4EF" });
    text(slide, body, 248, y + 14, 320, 22, { size: 16, bold: true });
  });

  panel(slide, 676, 238, 532, 300);
  text(slide, "보조 뉴스 근거", 708, 268, 220, 28, {
    size: 22,
    bold: true,
    color: C.amber,
  });
  bullet(
    slide,
    [
      "매일경제: 청년 1분기 주거비 부담 증가 보도",
      "브릿지경제·더팩트: 원스톱 청년주거 지원체계 필요 보도",
      "뉴스핌: 높은 집값으로 독립이 어려운 청년 이슈 보도",
      "지자체·상담소 기사: 정책 정보 제공과 상담형 안내 수요 확인",
    ],
    708,
    320,
    440,
    44,
    { size: 17, dot: C.amber },
  );
  noteBand(slide, "공식 인프라는 존재합니다. 그래서 우리는 대체가 아니라 연결·해석 레이어를 제안합니다.");
  footer(
    slide,
    3,
    "Sources: OECD Affordable Housing, OECD Housing Prices, MyHome, LH Apply Plus, OnTong Youth, Seoul Housing Portal, Google News RSS search.",
  );
  speaker(slide, [
    "주거비 부담은 OECD 자료와 국내 기사에서 계속 다뤄지는 문제입니다.",
    "동시에 마이홈과 LH청약플러스 같은 공식 인프라가 이미 있으므로, 저희는 이를 대체하지 않고 사용자 조건별 해석 레이어를 제안합니다.",
  ]);
}

// 4. User Journey
{
  const slide = newSlide();
  header(
    slide,
    "사용자 여정",
    "여러 사이트 탐색을 한 번의 사전 점검 흐름으로 바꿉니다",
    "청년 1인가구가 공개 공고를 검토하는 상황을 기준으로 발표 데모를 구성했습니다.",
  );
  const nodes = [
    ["1", "공고 검색", "마이홈·LH·지자체 포털 확인"],
    ["2", "공고문 해석", "PDF/HWP/HTML의 조건표 확인"],
    ["3", "내 조건 비교", "소득·자산·무주택·서류 직접 대입"],
    ["4", "입지 확인", "통근·생활SOC·주변 시세 별도 탐색"],
    ["5", "신청 준비", "원문 링크와 제출서류 최종 확인"],
  ];
  nodes.forEach(([num, title, body], i) => {
    const x = 72 + i * 230;
    panel(slide, x, 260, 180, 148);
    shape(slide, "ellipse", x + 66, 282, 48, 48, { fill: i === 2 ? C.amber : C.teal });
    text(slide, num, x + 66, 292, 48, 24, {
      size: 20,
      bold: true,
      color: "#FFFFFF",
      align: "center",
    });
    text(slide, title, x + 18, 350, 144, 24, { size: 19, bold: true, align: "center" });
    text(slide, body, x + 18, 384, 144, 38, { size: 13, color: C.muted, align: "center" });
    if (i < nodes.length - 1) {
      shape(slide, "chevron", x + 186, 310, 34, 38, { fill: C.teal });
    }
  });
  panel(slide, 180, 470, 920, 78, { fill: "#FFF6E8" });
  text(slide, "발표 데모의 목표", 210, 490, 150, 22, {
    size: 13,
    bold: true,
    color: C.amber,
  });
  text(
    slide,
    "조건 입력 → 추천 공고 → 점수 상세 → 지도 입지 → 공고문 근거/RAG → 정책 대시보드",
    210,
    516,
    850,
    26,
    { size: 22, bold: true },
  );
  noteBand(slide, "데모는 사용자가 실제로 막히는 순서 그대로 보여줍니다.");
  footer(slide, 4);
  speaker(slide, [
    "청년 1인가구가 공고를 찾는 상황을 기준으로 데모합니다.",
    "기존에는 여러 단계를 사용자가 직접 오가야 하지만, 저희 화면에서는 사전 점검 흐름으로 묶습니다.",
  ]);
}

// 5. Solution Architecture
{
  const slide = newSlide();
  header(
    slide,
    "해결 구조",
    "AI는 공고문을 구조화하고, 추천은 공개 기준과 규칙으로 설명합니다",
    "환각과 책임 문제를 줄이기 위해 자유 판정보다 원문 근거, 구조화 데이터, 규칙 기반 점수화를 앞세웁니다.",
  );
  panel(slide, 72, 232, 640, 330);
  await addImage(slide, screenshots.workflow, 94, 254, 596, 286, "주거복지 내비게이터 워크플로우");

  panel(slide, 748, 232, 460, 330);
  text(slide, "8인 팀 병렬 구조", 780, 262, 240, 28, { size: 22, bold: true });
  const roles = [
    "데이터 수집",
    "데이터 엔지니어링",
    "AI 공고문 추출",
    "백엔드·점수화 엔진",
    "데이터 분석",
    "지도·시각화",
    "프론트엔드",
    "기획·발표",
  ];
  roles.forEach((role, i) => {
    const x = 780 + (i % 2) * 202;
    const y = 318 + Math.floor(i / 2) * 50;
    chip(slide, role, x, y, 180, { fill: i < 4 ? "#E8F4EF" : "#EAF0F7" });
  });
  text(
    slide,
    "MVP는 공고 수집·조건 입력·기본 매칭부터 완성하고, 점수 설명·지도·RAG·정책 대시보드를 단계적으로 확장합니다.",
    780,
    532,
    360,
    42,
    { size: 15, color: C.muted },
  );
  noteBand(slide, "기능을 한 번에 과장하지 않고 MVP와 확장 기능을 분리해 실현 가능성을 방어합니다.");
  footer(slide, 5, "Application source: participant application workflow image.");
  speaker(slide, [
    "AI는 공고문을 읽기 쉬운 구조로 바꾸는 역할입니다.",
    "추천과 점수는 공개 기준과 규칙 기반으로 설명 가능하게 만들고, 팀은 8개 역할로 병렬 운영합니다.",
  ]);
}

// 6. Prototype dashboard
{
  const slide = newSlide();
  header(
    slide,
    "프로토타입 1",
    "조건 입력과 종합 매칭 결과를 첫 화면에서 함께 보여줍니다",
    "첫 화면에서 서비스의 본질인 조건 기반 공고 매칭과 주의 조건이 바로 보이도록 구성했습니다.",
  );
  panel(slide, 72, 232, 734, 334);
  await addImage(slide, screenshots.dashboard, 92, 252, 694, 294, "프로토타입 대시보드 화면");
  panel(slide, 842, 232, 366, 334);
  text(slide, "심사자가 봐야 할 포인트", 872, 264, 260, 28, { size: 21, bold: true, color: C.teal });
  bullet(
    slide,
    [
      "신청서의 입력 조건이 실제 폼으로 구현",
      "공개 공고와 제도를 같은 기준으로 비교",
      "기준 확인 항목과 주의 조건을 분리",
      "당첨 확률이 아닌 적합도 점수로 표현",
    ],
    872,
    324,
    280,
    48,
    { size: 17 },
  );
  noteBand(slide, "총점보다 중요한 것은 왜 이 공고가 먼저 보이는지 설명하는 구조입니다.");
  footer(slide, 6, "Local prototype screenshot: outputs/prototype_ppt_dashboard_crop.png");
  speaker(slide, [
    "첫 화면은 조건 입력과 추천 결과가 함께 보입니다.",
    "공고 카드에는 출처, 예시 모집기간, 기준 확인 항목, 주의 조건, 원문/신청 링크 버튼이 함께 표시되어 사용자가 다음 행동을 정할 수 있습니다.",
  ]);
}

// 7. Score detail
{
  const slide = newSlide();
  header(
    slide,
    "프로토타입 2",
    "가능/불가능 단정 대신 공고별 기준 점수를 분해합니다",
    "자격, 공고, 입지, 부담 완화, 신청 부담, 준비 용이성을 나눠 보여줘서 점수화의 설명 가능성을 높입니다.",
  );
  panel(slide, 72, 232, 760, 334);
  await addImage(slide, screenshots.detail, 92, 252, 720, 294, "공고별 기준 점수 상세 화면");
  panel(slide, 866, 232, 342, 334);
  text(slide, "방어 포인트", 894, 264, 220, 28, { size: 21, bold: true, color: C.teal });
  bullet(
    slide,
    [
      "점수는 당첨 확률이 아니라 기준 부합도",
      "공고문 기준과 사용자 조건을 항목별 비교",
      "감점·주의 사유를 숨기지 않음",
      "최종 신청은 원문 공고와 기관 심사 기준 확인",
    ],
    894,
    324,
    260,
    48,
    { size: 16 },
  );
  noteBand(slide, "심사자가 점수 임의성을 묻기 전에 기준별 분해 화면으로 먼저 방어합니다.");
  footer(slide, 7, "Local prototype screenshot: outputs/prototype_ppt_detail_crop.png");
  speaker(slide, [
    "점수 상세 화면은 심사에서 중요한 방어 장치입니다.",
    "총점 하나가 아니라 기준별 점수와 감점 사유를 보여줘, 이 서비스가 판정 시스템이 아니라 사전 점검 도구임을 분명히 합니다.",
  ]);
}

// 8. Map / Evidence / Policy
{
  const slide = newSlide();
  header(
    slide,
    "프로토타입 3",
    "지도, 원문 근거, RAG, 정책 대시보드까지 연결합니다",
    "개인 신청 보조에서 끝나지 않고, 익명·집계 지표로 정책 사각지대 탐색까지 확장합니다.",
  );
  const blocks = [
    [screenshots.map, "지도 입지 분석", "통근·생활SOC·주변 시세를 함께 비교"],
    [screenshots.evidence, "공고문 근거 + RAG", "출처 기반 답변으로 AI 환각 위험 통제"],
    [screenshots.policy, "정책 대시보드", "수요는 높지만 공급·신청 연결이 약한 영역 탐색"],
  ];
  for (const [i, [image, title, desc]] of blocks.entries()) {
    const x = 72 + i * 392;
    panel(slide, x, 238, 344, 278);
    await addImage(slide, image, x + 16, 258, 312, 188, title);
    text(slide, title, x + 20, 462, 260, 24, { size: 19, bold: true });
    text(slide, desc, x + 20, 492, 288, 36, { size: 14, color: C.muted });
  }
  noteBand(slide, "개인에게는 신청 전 우선순위, 공공에는 미매칭 패턴 탐색이라는 두 효과를 보여줍니다.");
  footer(slide, 8, "Local prototype screenshots: map, evidence, policy views.");
  speaker(slide, [
    "지도 탭은 생활 가능성을 보여주고, 근거/RAG 탭은 공고문 원문 기반 설명을 제공합니다.",
    "정책 대시보드는 익명 집계로 사각지대 의심 지표를 보여주는 확장 방향입니다.",
  ]);
}

// 9. Effects / Close
{
  const slide = newSlide();
  header(
    slide,
    "기대효과와 마무리",
    "주거 문제 해결을 과장하지 않고, 검증 가능한 효과로 말합니다",
    "탐색 시간 감소, 조건 확인 누락 감소, 상담 전 자가점검 강화, 정책 미매칭 패턴 파악을 목표로 합니다.",
  );
  const effects = [
    ["국민", "공고 탐색 시간 감소\n신청 가능한 공고 누락 감소\n서류·조건 사전 점검"],
    ["공공기관", "공개 공고 데이터 활용도 증가\n반복 민원·중복 탐색 부담 완화\n원문 근거 기반 안내"],
    ["지자체", "지역·계층별 미매칭 패턴 파악\n공급 부족·정보 이탈 의심 영역 탐색\n정책 개선 지표 확보"],
  ];
  effects.forEach(([title, body], i) => {
    const x = 72 + i * 392;
    panel(slide, x, 250, 344, 206);
    text(slide, title, x + 28, 282, 220, 30, {
      size: 25,
      bold: true,
      color: i === 2 ? C.amber : C.teal,
    });
    text(slide, body, x + 28, 334, 280, 82, { size: 17, color: C.ink });
  });
  panel(slide, 172, 492, 936, 66, { fill: "#EAF5F1" });
  text(
    slide,
    "흩어진 주거복지 기회를 사용자 조건별 우선순위와 원문 근거로 연결하겠습니다.",
    210,
    514,
    860,
    26,
    { size: 23, bold: true, color: C.tealDark, align: "center" },
  );
  noteBand(slide, "마지막 문장은 외우기: 흩어진 기회를 조건별 우선순위와 원문 근거로 연결합니다.");
  footer(slide, 9, "Submission schedule: 2026.07.13 materials, 2026.07.15 first review.");
  speaker(slide, [
    "기대효과는 주거 문제를 해결한다는 식으로 과장하지 않습니다.",
    "탐색 시간과 조건 확인 누락을 줄이고, 공공기관에는 미매칭 패턴을 보여주는 것이 목표입니다.",
  ]);
}

await fs.mkdir(outDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });

for (const [index, slide] of deck.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  const png = await deck.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(previewDir, `${stem}.png`), Buffer.from(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(previewDir, `${stem}.layout.json`), await layout.text());
}

const montage = await deck.export({ format: "webp", montage: true, scale: 1 });
await fs.writeFile(path.join(previewDir, "deck-montage.webp"), Buffer.from(await montage.arrayBuffer()));

const pptx = await PresentationFile.exportPptx(deck);
await pptx.save(deckPath);
console.log(deckPath);
