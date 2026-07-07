const programs = [
  {
    title: "대전 청년 매입임대 입주자 모집",
    type: "공공임대",
    source: "LH·마이홈 공고",
    period: "예시 2026.07.15 - 07.29",
    updated: "최종 갱신 예시: 2026.07.07",
    action: "원문 공고 확인",
    base: 91,
    fit: ["무주택", "청년", "소득 기준 확인"],
    caution: ["기본 서류 확인"],
    reason: "희망지역과 직장 접근성이 좋고, 모의 시세 기준 주거비 부담 완화 효과가 큽니다.",
    bars: {
      eligibility: 96,
      notice: 88,
      location: 92,
      saving: 86,
      competition: 67,
      readiness: 74,
    },
  },
  {
    title: "청년 전세임대 수시 지원",
    type: "전세지원",
    source: "LH 청약플러스",
    period: "예시 상시·예산 소진 전",
    updated: "최종 갱신 예시: 2026.07.07",
    action: "신청 링크 확인",
    base: 86,
    fit: ["청년", "무주택", "전세 선택 가능"],
    caution: ["보증금 한도 확인"],
    reason: "예시 점수상 신청 부담은 낮지만, 전세 물건 탐색과 보증 기준 확인이 필요합니다.",
    bars: {
      eligibility: 91,
      notice: 82,
      location: 78,
      saving: 82,
      competition: 84,
      readiness: 62,
    },
  },
  {
    title: "청년 월세 한시 지원",
    type: "주거급여·월세지원",
    source: "마이홈·지자체",
    period: "예시 지자체 접수 일정 확인",
    updated: "최종 갱신 예시: 2026.07.07",
    action: "지자체 안내 확인",
    base: 77,
    fit: ["월세 부담", "청년", "소득 확인 필요"],
    caution: ["가구 소득 재확인"],
    reason: "현재 월 주거비 부담을 낮출 수 있지만, 공고별 소득·거주 요건 확인이 필요합니다.",
    bars: {
      eligibility: 76,
      notice: 73,
      location: 68,
      saving: 80,
      competition: 88,
      readiness: 70,
    },
  },
  {
    title: "청년 특별공급 사전점검",
    type: "특별공급·청약",
    source: "청약홈·LH 공고",
    period: "예시 모집공고별 확인",
    updated: "최종 갱신 예시: 2026.07.07",
    action: "청약 공고 확인",
    base: 71,
    fit: ["청년", "청약통장 확인", "지역 우선순위 검토"],
    caution: ["무주택 기간·납입 조건 확인"],
    reason: "청약통장 조건은 확인 대상이지만, 모집공고별 세부 기준을 별도로 확인해야 합니다.",
    bars: {
      eligibility: 70,
      notice: 72,
      location: 74,
      saving: 66,
      competition: 58,
      readiness: 68,
    },
  },
];

const personaBoost = {
  youth: 0,
  newlywed: -4,
  senior: -8,
  vulnerable: 3,
  general: -6,
};

const regionNames = {
  daejeon: "대전 유성구",
  seoul: "서울 관악구",
  busan: "부산 해운대구",
  gwangju: "광주 북구",
};

const form = document.querySelector("#conditionForm");
const noticeList = document.querySelector("#noticeList");
const detailGrid = document.querySelector("#detailGrid");
const mainScore = document.querySelector("#mainScore");
const topProgram = document.querySelector("#topProgram");
const topReason = document.querySelector("#topReason");
const riskLabel = document.querySelector("#riskLabel");
const riskText = document.querySelector("#riskText");
const summaryText = document.querySelector("#summaryText");
const commuteScore = document.querySelector("#commuteScore");
const savingScore = document.querySelector("#savingScore");
const sourceTitle = document.querySelector("#sourceTitle");

function clamp(value, min = 52, max = 98) {
  return Math.max(min, Math.min(max, value));
}

function readConditions() {
  return {
    persona: document.querySelector("#persona").value,
    income: Number(document.querySelector("#income").value || 0),
    asset: Number(document.querySelector("#asset").value || 0),
    household: document.querySelector("#household").value,
    region: document.querySelector("#region").value,
    commute: document.querySelector("#commute").value,
    rent: Number(document.querySelector("#rent").value || 0),
    account: document.querySelector("#account").value,
    homeless: document.querySelector("#homeless").checked,
    documents: document.querySelector("#documents").checked,
  };
}

function scoreProgram(program, conditions, index) {
  const incomePenalty = conditions.income > 360 ? -9 : conditions.income < 180 ? -2 : 0;
  const assetPenalty = conditions.asset > 28000 ? -8 : 0;
  const homelessPenalty = conditions.homeless ? 0 : -22;
  const documentBoost = conditions.documents ? 5 : -3;
  const commuteBoost = conditions.commute === "near" ? 4 : conditions.commute === "far" ? -6 : 0;
  const accountPenalty = conditions.account === "none" && index === 0 ? -7 : 0;
  const rentBoost = conditions.rent > 70 ? 4 : conditions.rent < 35 ? -3 : 0;
  const total =
    program.base +
    personaBoost[conditions.persona] +
    incomePenalty +
    assetPenalty +
    homelessPenalty +
    documentBoost +
    commuteBoost +
    accountPenalty +
    rentBoost -
    index * 2;

  const bars = Object.fromEntries(
    Object.entries(program.bars).map(([key, value]) => {
      const readinessBoost = key === "readiness" ? documentBoost * 2 : 0;
      const locationBoost = key === "location" ? commuteBoost * 2 : 0;
      const eligibilityPenalty =
        key === "eligibility" ? incomePenalty + assetPenalty + homelessPenalty : 0;
      return [key, clamp(value + readinessBoost + locationBoost + eligibilityPenalty)];
    }),
  );

  return {
    ...program,
    score: clamp(Math.round(total)),
    bars,
    caution: conditions.documents
      ? program.caution.filter((item) => !item.includes("서류"))
      : program.caution,
  };
}

function bar(label, value) {
  return `
    <div class="bar-row">
      <span>${label}</span>
      <span class="track"><i style="width:${value}%"></i></span>
      <span>${value}</span>
    </div>
  `;
}

function scoreRow(label, value, note) {
  return `
    <div class="score-row">
      <span>${label}</span>
      <span class="track"><i style="width:${value}%"></i></span>
      <strong>${value}</strong>
      <small>${note}</small>
    </div>
  `;
}

function renderNotices(data) {
  noticeList.innerHTML = data
    .map(
      (item, index) => `
        <article class="notice-card">
          <div class="rank-score">
            <span>추천 ${index + 1}순위</span>
            <strong>${item.score}</strong>
            <span>/ 100</span>
          </div>
          <div class="card-main">
            <h4>${item.title}</h4>
            <div class="meta">
              <span>${item.type}</span>
              <span>출처: ${item.source}</span>
              <span>모집: ${item.period}</span>
              <span>${item.updated}</span>
            </div>
            <div class="criteria">
              ${item.fit.map((text) => `<span class="ok">기준 확인 · ${text}</span>`).join("")}
              ${item.caution.map((text) => `<span class="warn">주의 · ${text}</span>`).join("")}
            </div>
            <p>${item.reason}</p>
            <button class="ghost-action" type="button">${item.action}</button>
          </div>
          <div class="bar-stack">
            ${bar("자격", item.bars.eligibility)}
            ${bar("입지", item.bars.location)}
            ${bar("부담", item.bars.saving)}
            ${bar("준비", item.bars.readiness)}
          </div>
        </article>
      `,
    )
    .join("");
}

function renderDetails(data) {
  detailGrid.innerHTML = data
    .map(
      (item) => `
        <article class="detail-card">
          <h4>${item.title}</h4>
          ${scoreRow("자격 적합도", item.bars.eligibility, "소득·자산·무주택 조건 비교")}
          ${scoreRow("공고 적합도", item.bars.notice, "모집대상·기간·지역 기준 비교")}
          ${scoreRow("입지 적합도", item.bars.location, "직장/학교 접근성과 생활SOC 반영")}
          ${scoreRow("주거비 부담 완화", item.bars.saving, "모의 시세 기준 예상 부담 완화")}
          ${scoreRow("신청 부담", item.bars.competition, "높을수록 예시 점수상 부담이 낮음")}
          ${scoreRow("준비 용이성", item.bars.readiness, "서류·통장·접수 단계 준비 상태")}
        </article>
      `,
    )
    .join("");
}

function updateSummary(data, conditions) {
  const top = data[0];
  mainScore.textContent = top.score;
  document.documentElement.style.setProperty("--score-angle", `${Math.round(top.score * 3.6)}deg`);
  topProgram.textContent = top.title;
  topReason.textContent = top.reason;
  sourceTitle.textContent = `${top.title} 공고문`;
  riskLabel.textContent = conditions.homeless ? "서류·기준 확인" : "무주택 조건 확인";
  riskText.textContent = conditions.homeless
    ? conditions.documents
      ? "기본 서류가 준비되어 준비 용이성이 상승했습니다."
      : "제출서류를 준비하면 준비 용이성이 상승합니다."
    : "무주택 조건 확인 전에는 추천 점수를 낮게 표시합니다.";
  summaryText.textContent = `${regionNames[conditions.region]} 기준 발표용 예시 공고·제도 4개를 비교했습니다.`;
  commuteScore.textContent =
    conditions.commute === "near" ? "28분" : conditions.commute === "mid" ? "43분" : "61분";
  const saving = Math.max(12, Math.round(conditions.rent * 0.55));
  savingScore.textContent = `월 ${saving}만원 수준`;
}

function recalculate() {
  const conditions = readConditions();
  const scored = programs
    .map((program, index) => scoreProgram(program, conditions, index))
    .sort((a, b) => b.score - a.score);
  updateSummary(scored, conditions);
  renderNotices(scored);
  renderDetails(scored);
}

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    const viewName = tab.dataset.view;
    document.querySelectorAll(".tab").forEach((item) => item.classList.remove("active"));
    document.querySelectorAll(".view").forEach((item) => item.classList.remove("active"));
    tab.classList.add("active");
    document.querySelector(`#view-${viewName}`).classList.add("active");
  });
});

form.addEventListener("submit", (event) => {
  event.preventDefault();
  recalculate();
});

form.addEventListener("change", recalculate);
recalculate();
