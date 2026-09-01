const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
const W = 13.3, H = 7.5;

// ---- palette: steel + safety amber -------------------------------------
const INK = "1C2B36";
const INK2 = "2E4150";
const SLATE = "5A7184";
const MIST = "F0F3F6";
const MIST2 = "E4EAEF";
const LINE = "D3DBE2";
const AMBER = "D9722C";
const AMBER_L = "FAEDE1";
const GREEN = "3B7A5A";
const GREEN_L = "E6F0EA";
const WHITE = "FFFFFF";

const FONT = "Meiryo";
const TODO = (s) => `【要記入：${s}】`;

// ---- helpers -----------------------------------------------------------
function shadow() {
  return { type: "outer", color: "1C2B36", blur: 8, offset: 2, angle: 90, opacity: 0.10 };
}

function lightSlide(kicker, title) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  if (kicker) {
    s.addText(kicker, {
      x: 0.7, y: 0.42, w: 8, h: 0.28, isTextBox: true, margin: 0,
      fontFace: FONT, fontSize: 12, bold: true, color: AMBER, charSpacing: 1,
    });
  }
  s.addText(title, {
    x: 0.7, y: 0.72, w: 11.9, h: 0.75, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 30, bold: true, color: INK, valign: "top",
  });
  return s;
}

function darkSlide(kicker, title) {
  const s = pres.addSlide();
  s.background = { color: INK };
  if (kicker) {
    s.addText(kicker, {
      x: 0.7, y: 0.42, w: 8, h: 0.28, isTextBox: true, margin: 0,
      fontFace: FONT, fontSize: 12, bold: true, color: AMBER, charSpacing: 1,
    });
  }
  s.addText(title, {
    x: 0.7, y: 0.72, w: 11.9, h: 0.75, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 30, bold: true, color: WHITE, valign: "top",
  });
  return s;
}

function card(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.08,
    fill: { color: o.fill || MIST }, line: { color: o.line || LINE, width: 1 },
    shadow: shadow(),
  });
}

function numChip(s, x, y, n, fill, txt) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: 0.46, h: 0.46, fill: { color: fill || INK }, line: { color: fill || INK, width: 1 },
  });
  s.addText(String(n), {
    x, y, w: 0.46, h: 0.46, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 15, bold: true, color: txt || WHITE, align: "center", valign: "middle",
  });
}

function foot(s, text, dark) {
  s.addText(text, {
    x: 0.7, y: 6.82, w: 11.9, h: 0.34, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 11, color: dark ? "9FB0BE" : SLATE, valign: "middle",
  });
}

function bodyText(s, text, o) {
  s.addText(text, {
    isTextBox: true, margin: 0, fontFace: FONT, fontSize: o.size || 13.5,
    color: o.color || INK2, valign: o.valign || "top", lineSpacing: o.lineSpacing || 20,
    x: o.x, y: o.y, w: o.w, h: o.h, bold: o.bold || false, align: o.align || "left",
  });
}

// =======================================================================
// 1. Title
// =======================================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.6, y: -1.6, w: 5.6, h: 5.6, fill: { color: INK2 }, line: { color: INK2, width: 1 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 11.3, y: 4.4, w: 2.6, h: 2.6, fill: { color: AMBER, transparency: 55 }, line: { color: AMBER, width: 1 },
  });
  s.addText("事業ご説明資料", {
    x: 0.9, y: 1.85, w: 8, h: 0.3, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 13, bold: true, color: AMBER, charSpacing: 2,
  });
  s.addText("工場の現場データを\n「その日のうちに」使える形にする", {
    x: 0.9, y: 2.35, w: 9.2, h: 1.85, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 36, bold: true, color: WHITE, lineSpacing: 46,
  });
  s.addText(TODO("サービス名／屋号"), {
    x: 0.9, y: 4.35, w: 8, h: 0.36, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 15, color: AMBER,
  });
  s.addText(
    `${TODO("氏名")}　／　${TODO("2026年◯月◯日")}　／　${TODO("御社名")} 御中`,
    { x: 0.9, y: 5.45, w: 10.5, h: 0.34, isTextBox: true, margin: 0, fontFace: FONT, fontSize: 13, color: "AFC0CD" }
  );
  s.addText("本日のご相談：①ヒアリング先のご紹介　②予備調査を今使うべきかの判断", {
    x: 0.9, y: 5.9, w: 10.5, h: 0.34, isTextBox: true, margin: 0,
    fontFace: FONT, fontSize: 13, bold: true, color: WHITE,
  });
  s.addNotes(
    "冒頭30秒でここを言い切る。\n" +
    "『本日は売り込みではありません。仮説段階で顧客ゼロという現在地をそのままお見せして、穴を指摘いただきたい、そしてヒアリング先をご紹介いただきたい、という2点のお願いです』\n" +
    "相手は現場を知っているコンサル代表。取り繕うと即座に見抜かれるので、最初に現在地を開示する姿勢を出すことが信用の入り口になる。"
  );
}

// =======================================================================
// 2. Summary
// =======================================================================
{
  const s = lightSlide("SUMMARY", "本日お伝えすること");
  const items = [
    ["現在地", "仮説段階・顧客ゼロ。\n課題仮説はあるが、\n現場での検証はこれから。"],
    ["やること", "90日で有償1社を目標に。\n10社ヒアリング → 2社でPoC\n→ 1社と有償契約。"],
    ["お願い", "ヒアリング先のご紹介。\n予備調査は「自分で聞いても\n埋まらない問い」が残った時に。"],
  ];
  const cw = 3.6, gap = 0.45;
  items.forEach((it, i) => {
    const x = 0.7 + i * (cw + gap);
    card(s, { x, y: 1.95, w: cw, h: 2.7 });
    numChip(s, x + 0.35, 2.25, i + 1, i === 2 ? AMBER : INK);
    bodyText(s, it[0], { x: x + 0.95, y: 2.32, w: cw - 1.2, h: 0.4, size: 16, bold: true, color: INK });
    bodyText(s, it[1], { x: x + 0.35, y: 3.0, w: cw - 0.7, h: 1.5, size: 13, lineSpacing: 22 });
  });
  card(s, { x: 0.7, y: 5.0, w: 11.9, h: 1.35, fill: AMBER_L, line: "EBD3BC" });
  bodyText(s, "本日は売り込みではなく、仮説の穴を先に潰すためのご相談です。", {
    x: 1.1, y: 5.3, w: 11.1, h: 0.4, size: 17, bold: true, color: INK,
  });
  bodyText(s, "取り繕った資料より、どこが分かっていないかを開示した資料の方が、有益なご指摘をいただけると考えました。", {
    x: 1.1, y: 5.78, w: 11.1, h: 0.4, size: 12.5, color: INK2,
  });
  s.addNotes(
    "この1枚で全体を握る。以降のスライドは全部この3点の詳細。\n" +
    "「顧客ゼロ」を先に言う。後半で発覚するより、自分から最初に出した方が圧倒的に印象が良い。\n" +
    "相手が途中で質問を始めたら、資料を進めるより会話を優先してよい。紹介をもらうのが目的であって、資料を読み切るのが目的ではない。"
  );
}

// =======================================================================
// 3. Who I am
// =======================================================================
{
  const s = lightSlide("WHO I AM", "自己紹介と、いま実際に動かしているもの");

  card(s, { x: 0.7, y: 1.95, w: 5.75, h: 4.15 });
  bodyText(s, "経歴・現在の稼働", { x: 1.05, y: 2.25, w: 5.05, h: 0.35, size: 16, bold: true, color: INK });
  bodyText(s,
    `${TODO("現職・これまでの職種／年数")}\n\n` +
    `${TODO("製造業との接点があれば（現場経験・取引先など）")}\n\n` +
    `${TODO("本件に週何時間使えるか／資金の目処")}\n\n` +
    `${TODO("なぜ工場向けなのか（動機。必ず聞かれます）")}`,
    { x: 1.05, y: 2.75, w: 5.05, h: 3.1, size: 13, lineSpacing: 22 });

  card(s, { x: 6.85, y: 1.95, w: 5.75, h: 4.15, fill: GREEN_L, line: "C9DED3" });
  bodyText(s, "「作って、動かし続けられる」ことの実証", { x: 7.2, y: 2.25, w: 5.05, h: 0.35, size: 16, bold: true, color: INK });
  bodyText(s, "価格ウォッチ（個人開発・稼働中）", { x: 7.2, y: 2.72, w: 5.05, h: 0.32, size: 13, bold: true, color: GREEN });
  bodyText(s,
    "・ 取得 → 集計 → 公開までを毎日無人で運用\n" +
    "・ 31カテゴリ・880商品の価格を日次で追跡\n" +
    "・ 収集した値動きから「買い時／平常／待ち」を自動判定",
    { x: 7.2, y: 3.15, w: 5.05, h: 1.5, size: 13, lineSpacing: 22 });
  bodyText(s,
    "工場向けサービスに必要なのは、派手な機能ではなく\n「毎日データが溜まり続ける仕組みを止めずに回すこと」。\nその実装と運用は、すでに手元で証明できています。",
    { x: 7.2, y: 4.85, w: 5.05, h: 1.1, size: 12.5, lineSpacing: 20, color: INK2 });

  foot(s, "※ 事業実績ではなく「技術と継続運用の裏付け」として提示。誇張しないこと。");
  s.addNotes(
    "顧客ゼロの状態で信用を得る唯一の材料が『すでに手を動かして動かし続けている物がある』こと。\n" +
    "価格ウォッチは工場向けではないが、"
    + "『毎日自動でデータを取り、集計し、判定を出し、公開まで無人で回している』という点が、現場データSaaSに必要な能力とほぼ同じであることを説明する。\n" +
    "ここは盛らない。『これは工場の実績ではありません』と自分から言う。"
  );
}

// =======================================================================
// 4. Problem hypothesis
// =======================================================================
{
  const s = lightSlide("PROBLEM", "課題仮説：現場のデータが「その場」で消えている");
  const items = [
    ["紙とホワイトボード", "記録は残るが、\n集計されない。\n書いた瞬間から\n誰も見返さない。"],
    ["Excelへの転記", "転記に時間がかかり、\nミスも出る。\n触れるのは\n担当者ひとりだけ。"],
    ["気づくのが遅い", "不良・設備停止・\n段取り遅れが分かるのは\n月次集計が出た後。"],
    ["属人化", "ベテランの勘に依存し、\n数字として\n承継できない。"],
  ];
  const cw = 2.85, gap = 0.31;
  items.forEach((it, i) => {
    const x = 0.7 + i * (cw + gap);
    card(s, { x, y: 2.05, w: cw, h: 3.1 });
    numChip(s, x + 0.32, 2.35, i + 1, INK);
    bodyText(s, it[0], { x: x + 0.32, y: 3.0, w: cw - 0.64, h: 0.7, size: 15, bold: true, color: INK, lineSpacing: 22 });
    bodyText(s, it[1], { x: x + 0.32, y: 3.75, w: cw - 0.64, h: 1.2, size: 12.5, lineSpacing: 20 });
  });
  card(s, { x: 0.7, y: 5.45, w: 11.9, h: 0.95, fill: AMBER_L, line: "EBD3BC" });
  bodyText(s, "以上はすべて仮説です。ヒアリング10社で「実際に困っている順番」を確かめ、外れていれば作るものを変えます。", {
    x: 1.1, y: 5.78, w: 11.1, h: 0.4, size: 13.5, bold: true, color: INK,
  });
  s.addNotes(
    "ここが一番ツッコまれる場所。相手は現場を知っているので『それ、うちの顧客だとこうだよ』が必ず返ってくる。\n" +
    "その返しこそが今日いちばん価値のある情報。反論せず、メモを取り、『4つのうちどれが一番刺さりますか』と聞き返す。\n" +
    "順位を聞き出せたら大成功。それが製品の第一機能になる。"
  );
}

// =======================================================================
// 5. Why now
// =======================================================================
{
  const s = lightSlide("WHY NOW", "なぜ今なのか（3つの追い風・いずれも要検証）");
  const rows = [
    ["人手不足と技能承継", "ベテランの退職で、勘に頼った管理が続けられなくなる。数字にして残す必要が出てきている。"],
    ["原価把握の必要性", "材料費・エネルギー費の変動を価格に反映するには、まず自社の実績値が要る。"],
    ["導入コストの低下", "スマホ・低価格センサ・クラウドで、数百万円だった仕組みが桁を落として作れる。"],
  ];
  rows.forEach((r, i) => {
    const y = 2.0 + i * 1.5;
    card(s, { x: 0.7, y, w: 11.9, h: 1.28 });
    numChip(s, 1.05, y + 0.41, i + 1, AMBER);
    bodyText(s, r[0], { x: 1.75, y: y + 0.24, w: 3.2, h: 0.4, size: 16, bold: true, color: INK });
    bodyText(s, r[1], { x: 5.05, y: y + 0.28, w: 7.15, h: 0.8, size: 13, lineSpacing: 20 });
    s.addShape(pres.ShapeType.roundRect, {
      x: 1.75, y: y + 0.72, w: 0.95, h: 0.3, rectRadius: 0.06,
      fill: { color: AMBER_L }, line: { color: "EBD3BC", width: 1 },
    });
    bodyText(s, "要検証", { x: 1.75, y: y + 0.72, w: 0.95, h: 0.3, size: 10.5, bold: true, color: AMBER, align: "center", valign: "middle" });
  });
  foot(s, "※ 一般論として語られる話であり、自分で確かめた数字ではないことを明示する。");
  s.addNotes(
    "『なぜ今か』は投資家・コンサル双方が必ず確認する論点。ただし借り物の一般論だと見抜かれる。\n" +
    "だからこそ全部に『要検証』を貼っておく。『一般論としてはこう言われていますが、自分で確かめた数字ではありません。ここ、実感と合っていますか』と振ると会話になる。"
  );
}

// =======================================================================
// 6. Solution hypothesis
// =======================================================================
{
  const s = lightSlide("SOLUTION", "解決仮説：現場の作業を1秒も増やさない");
  const steps = [
    ["現場で入力", "スマホで3タップ。\nもしくは今使っている紙を\nそのまま写真で撮るだけ。"],
    ["自動で集計", "転記ゼロ。\n人が Excel を触る工程を\n仕組みから消す。"],
    ["翌朝に届く", "毎朝ダッシュボード。\n異常が出た時だけ通知。\n平常時は何もしない。"],
  ];
  const cw = 3.55, gap = 0.6;
  steps.forEach((st, i) => {
    const x = 0.7 + i * (cw + gap);
    card(s, { x, y: 2.05, w: cw, h: 2.35 });
    numChip(s, x + 0.35, 2.35, i + 1, INK);
    bodyText(s, st[0], { x: x + 0.95, y: 2.42, w: cw - 1.2, h: 0.4, size: 16, bold: true, color: INK });
    bodyText(s, st[1], { x: x + 0.35, y: 3.1, w: cw - 0.7, h: 1.5, size: 13, lineSpacing: 22 });
    if (i < 2) {
      s.addShape(pres.ShapeType.rightArrow, {
        x: x + cw + 0.12, y: 3.25, w: 0.36, h: 0.34,
        fill: { color: AMBER }, line: { color: AMBER, width: 1 },
      });
    }
  });
  card(s, { x: 0.7, y: 5.1, w: 11.9, h: 1.3, fill: MIST2, line: LINE });
  bodyText(s, "設計方針：現場の手間を1つでも増やしたら、その時点で使われなくなる", {
    x: 1.1, y: 5.35, w: 11.1, h: 0.4, size: 16, bold: true, color: INK,
  });
  bodyText(s, `機能ではなく「定着」で勝負する。${TODO("現時点のプロトタイプの状態／触れるデモの有無")}`, {
    x: 1.1, y: 5.82, w: 11.1, h: 0.4, size: 12.5, color: INK2,
  });
  s.addNotes(
    "現場DXが失敗する最大の理由は『現場が入力してくれない』こと。相手も必ずそれを言ってくる。\n" +
    "先回りして『入力を増やさない』を設計方針として掲げておくと、分かっている人だと伝わる。\n" +
    "デモが触れる状態なら、この後スマホで実物を見せるのが最も効く。無ければ『まだ画面だけです』と正直に言う。"
  );
}

// =======================================================================
// 7. Value by role
// =======================================================================
{
  const s = lightSlide("VALUE", "誰の、何が変わるのか");
  const cols = [
    ["現場の作業者", "転記作業がなくなる", "終業後の日報転記に\n毎日◯十分かけている状態から、\n入力その場で完了へ。", INK],
    ["工場長・管理者", "翌日に異常が分かる", "月次集計を待たずに、\n不良・停止・遅れを\n翌朝の画面で把握。", INK],
    ["経営者", "原価と稼働が数字で出る", "勘ではなく実績値で\n価格交渉・設備投資の\n判断ができる。", AMBER],
  ];
  const cw = 3.85, gap = 0.32;
  cols.forEach((c, i) => {
    const x = 0.7 + i * (cw + gap);
    card(s, { x, y: 2.05, w: cw, h: 3.45, fill: i === 2 ? AMBER_L : MIST, line: i === 2 ? "EBD3BC" : LINE });
    bodyText(s, c[0], { x: x + 0.35, y: 2.35, w: cw - 0.7, h: 0.35, size: 13, bold: true, color: c[3] });
    bodyText(s, c[1], { x: x + 0.35, y: 2.85, w: cw - 0.7, h: 0.85, size: 17, bold: true, color: INK, lineSpacing: 26 });
    bodyText(s, c[2], { x: x + 0.35, y: 3.75, w: cw - 0.7, h: 1.25, size: 12.5, lineSpacing: 21 });
  });
  foot(s, "※ 決裁者が誰かはまだ確定していない（工場長／社長／情シス）。ヒアリングで確かめたい論点のひとつ。");
  s.addNotes(
    "3者のうち誰に売るかで、価格も営業導線も全部変わる。まだ決めていないことを正直に言う。\n" +
    "相手は製造業に強いので『この規模ならだいたい社長決裁だよ』といった実務知識が出てくる可能性が高い。それが今日の収穫のひとつ。"
  );
}

// =======================================================================
// 8. Business model
// =======================================================================
{
  const s = lightSlide("MODEL", "ビジネスモデル仮説（数字はすべて仮置き）");

  card(s, { x: 0.7, y: 1.95, w: 5.75, h: 2.35 });
  bodyText(s, "課金の形", { x: 1.05, y: 2.2, w: 5.05, h: 0.35, size: 15, bold: true, color: INK });
  bodyText(s, `初期導入　${TODO("◯◯")}万円`, { x: 1.05, y: 2.72, w: 5.05, h: 0.4, size: 19, bold: true, color: INK });
  bodyText(s, `月額　　　${TODO("◯")}万円／拠点`, { x: 1.05, y: 3.25, w: 5.05, h: 0.4, size: 19, bold: true, color: AMBER });
  bodyText(s, "初期で導入支援の工数を回収し、月額で継続する形を想定。", { x: 1.05, y: 3.8, w: 5.05, h: 0.35, size: 12, color: SLATE });

  card(s, { x: 6.85, y: 1.95, w: 5.75, h: 2.35 });
  bodyText(s, "想定顧客（ICP仮説）", { x: 7.2, y: 2.2, w: 5.05, h: 0.35, size: 15, bold: true, color: INK });
  bodyText(s,
    `・ 従業員 ${TODO("◯◯〜◯◯")}名規模\n` +
    `・ 業種：${TODO("金属加工／食品／樹脂 など")}\n` +
    `・ 地域：${TODO("◯◯県内・車で行ける範囲")}\n` +
    "・ 生産管理システムは未導入、または一部のみ",
    { x: 7.2, y: 2.7, w: 5.05, h: 1.5, size: 13, lineSpacing: 22 });

  card(s, { x: 0.7, y: 4.55, w: 11.9, h: 1.85, fill: AMBER_L, line: "EBD3BC" });
  bodyText(s, "この数字は「決めた価格」ではなく「聞きに行くための仮置き」です", {
    x: 1.1, y: 4.82, w: 11.1, h: 0.4, size: 16, bold: true, color: INK });
  bodyText(s,
    "ヒアリングでは「いくらなら払うか」を直接は聞かず、「今この課題に年間いくら（人件費・不良ロス）かかっているか」を聞き、\n" +
    "そこから逆算して価格を決める方針です。仮置きの数字が現場感覚とズレていれば、その場でご指摘ください。",
    { x: 1.1, y: 5.32, w: 11.1, h: 0.85, size: 12.5, lineSpacing: 21, color: INK2 });
  s.addNotes(
    "価格を聞かれて黙るのが最悪。仮でも数字を出し、『仮置きです』と明示するのが一番強い。\n" +
    "相手が『その規模なら月◯万が上限だよ』と言ってくれたら、それだけで今日は元が取れている。\n" +
    "『いくらなら払いますか』と客に聞いても本音は出ない、という点は相手も同意するはず。逆算アプローチを説明すると、分かっていると評価されやすい。"
  );
}

// =======================================================================
// 9. Current status (honest)
// =======================================================================
{
  const s = lightSlide("STATUS", "現在の状況：できていること／まだできていないこと");

  card(s, { x: 0.7, y: 1.95, w: 5.75, h: 4.0, fill: GREEN_L, line: "C9DED3" });
  bodyText(s, "できていること", { x: 1.05, y: 2.25, w: 5.05, h: 0.4, size: 18, bold: true, color: GREEN });
  bodyText(s,
    "・ 課題仮説と、解くべき順番の整理\n\n" +
    "・ データ収集〜集計〜配信を無人で回す実装力\n　（価格ウォッチで日次稼働中）\n\n" +
    `・ ${TODO("プロトタイプ／画面のどこまでできているか")}\n\n` +
    `・ ${TODO("これまでに話を聞けた工場の件数")}`,
    { x: 1.05, y: 2.8, w: 5.05, h: 3.0, size: 13, lineSpacing: 21 });

  card(s, { x: 6.85, y: 1.95, w: 5.75, h: 4.0, fill: AMBER_L, line: "EBD3BC" });
  bodyText(s, "まだできていないこと", { x: 7.2, y: 2.25, w: 5.05, h: 0.4, size: 18, bold: true, color: AMBER });
  bodyText(s,
    "・ 有償顧客はゼロ。売上実績なし\n\n" +
    "・ 現場ヒアリングが不足。課題仮説は\n　まだ自分の想像の域を出ていない\n\n" +
    "・ 価格の妥当性が未検証\n\n" +
    "・ 導入後の運用負荷（誰が面倒を見るか）が未設計\n\n" +
    "・ 販路がない。今日いちばんの制約はここ",
    { x: 7.2, y: 2.8, w: 5.05, h: 3.0, size: 13, lineSpacing: 21 });

  foot(s, "この2枚を並べて出すことが、本日の資料の主旨です。");
  s.addNotes(
    "この資料の核心。多くの人は左側だけ見せて右側を隠す。隠すと、相手は『何を隠しているか』を探す時間に会話を使ってしまう。\n" +
    "先に全部出せば、残り時間を『どう埋めるか』の相談に使える。コンサル相手にはこちらが圧倒的に有利。\n" +
    "『販路がない』を最後に置いて、次のお願いのスライドへつなげる。"
  );
}

// =======================================================================
// 10. 90-day roadmap
// =======================================================================
{
  const s = lightSlide("PLAN", "今後の方針：90日で「有償1社」に到達する");
  const phases = [
    ["Day 1–30", "現場ヒアリング 10社", "課題の順位と決裁者を特定する", "ゲート：10社中8社が\n同じ課題を挙げること"],
    ["Day 31–60", "無償PoC 2社", "実際に現場で使ってもらう", "ゲート：2週間、\n人に言われず使い続けること"],
    ["Day 61–90", "有償契約 1社", "初期＋月額で契約を結ぶ", "ゲート：値引きなしで\n発注が出ること"],
  ];
  const cw = 3.85, gap = 0.32;
  phases.forEach((p, i) => {
    const x = 0.7 + i * (cw + gap);
    card(s, { x, y: 2.0, w: cw, h: 3.35 });
    bodyText(s, p[0], { x: x + 0.35, y: 2.28, w: cw - 0.7, h: 0.32, size: 12.5, bold: true, color: AMBER });
    bodyText(s, p[1], { x: x + 0.35, y: 2.68, w: cw - 0.7, h: 0.45, size: 18, bold: true, color: INK });
    bodyText(s, p[2], { x: x + 0.35, y: 3.25, w: cw - 0.7, h: 0.6, size: 12.5, lineSpacing: 20 });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.35, y: 3.95, w: cw - 0.7, h: 1.05, rectRadius: 0.06,
      fill: { color: WHITE }, line: { color: LINE, width: 1 },
    });
    bodyText(s, p[3], { x: x + 0.5, y: 4.13, w: cw - 1.0, h: 0.8, size: 11.5, lineSpacing: 18, color: INK2 });
  });
  card(s, { x: 0.7, y: 5.6, w: 11.9, h: 0.85, fill: MIST2, line: LINE });
  bodyText(s, "各ゲートを越えられなければ、粘らずに仮説を捨てて作り直します。期限を切るのはそのためです。", {
    x: 1.1, y: 5.83, w: 11.1, h: 0.4, size: 13.5, bold: true, color: INK });
  s.addNotes(
    "『いつまでに何を』が無い相談は、コンサルから見ると最も動きにくい。日付とゲートを必ず言う。\n" +
    "ゲート条件（8社／2週間／値引きなし）は自分で作った撤退基準。ここを聞かれたら『撤退基準です』と答える。\n" +
    "紹介のお願いは Day 1–30 に直結する、という流れで次のスライドへ。"
  );
}

// =======================================================================
// 11. Five questions
// =======================================================================
{
  const s = lightSlide("OPEN QUESTIONS", "いま埋めたい5つの問い");
  const qs = [
    ["現場は本当に入力してくれるか", "→ 製品設計そのものが決まる"],
    ["誰が決裁するか（工場長／社長／情シス）", "→ 営業導線と資料の宛先が決まる"],
    ["いくらなら払うか（今その課題にいくら払っているか）", "→ 価格が決まる"],
    ["既存の生産管理システムとどう共存するか", "→ 作る範囲（スコープ）が決まる"],
    ["1社目の型は、他社にも横展開できるか", "→ 事業になるかどうかが決まる"],
  ];
  qs.forEach((q, i) => {
    const y = 1.95 + i * 0.92;
    card(s, { x: 0.7, y, w: 11.9, h: 0.78 });
    numChip(s, 0.98, y + 0.16, i + 1, i < 3 ? INK : AMBER);
    bodyText(s, q[0], { x: 1.62, y: y + 0.24, w: 6.5, h: 0.35, size: 14, bold: true, color: INK });
    bodyText(s, q[1], { x: 8.3, y: y + 0.26, w: 3.9, h: 0.35, size: 12.5, color: SLATE });
  });
  foot(s, "1〜3は自分の足で埋めにいきます。4・5は業界構造の話で、独力では時間がかかると考えています。");
  s.addNotes(
    "この5問が、そのまま次のスライドの『予備調査を使うかどうか』の判断材料になる。\n" +
    "5問を先に見せてから『どれを自分でやり、どれを人に頼るか』を相談する流れにすると、判断の筋道が相手に見える。\n" +
    "相手がこの場で4・5に答えてくれるなら、予備調査は要らないかもしれない。それも含めて相談する。"
  );
}

// =======================================================================
// 12. Preliminary research decision frame
// =======================================================================
{
  const s = lightSlide("DECISION", "ご提案いただいた予備調査について（ご相談）");

  card(s, { x: 0.7, y: 1.95, w: 5.75, h: 2.6, fill: GREEN_L, line: "C9DED3" });
  bodyText(s, "自分の足で埋まると考えている", { x: 1.05, y: 2.2, w: 5.05, h: 0.35, size: 15, bold: true, color: GREEN });
  bodyText(s,
    "Q1 現場は入力してくれるか\nQ2 決裁者は誰か\nQ3 いくらなら払うか\n\n→ 10社のヒアリングで見えてくる想定",
    { x: 1.05, y: 2.7, w: 5.05, h: 1.7, size: 13, lineSpacing: 22 });

  card(s, { x: 6.85, y: 1.95, w: 5.75, h: 2.6, fill: AMBER_L, line: "EBD3BC" });
  bodyText(s, "お力をお借りしたい", { x: 7.2, y: 2.2, w: 5.05, h: 0.35, size: 15, bold: true, color: AMBER });
  bodyText(s,
    "Q4 既存システムとの共存の実態\nQ5 横展開できる型があるか\n（＋業界の商習慣・意思決定の流れ）\n\n→ 独力だと時間がかかりすぎる領域",
    { x: 7.2, y: 2.7, w: 5.05, h: 1.7, size: 13, lineSpacing: 22 });

  card(s, { x: 0.7, y: 4.8, w: 11.9, h: 1.6, fill: MIST2, line: LINE });
  bodyText(s, "発注させていただく場合の判断基準", { x: 1.1, y: 5.02, w: 11.1, h: 0.35, size: 15, bold: true, color: INK });
  bodyText(s,
    "「その調査結果で、自分の次の行動が変わるか」——変わらないなら、今は買わない、と決めています。\n" +
    "想定するスコープ・期間・ご予算感については、本日この場でご相談させてください。",
    { x: 1.1, y: 5.45, w: 11.1, h: 0.85, size: 12.5, lineSpacing: 21, color: INK2 });
  s.addNotes(
    "重要：この場で即断しない。『持ち帰って決めます』で構わない。ただし判断基準は先に言い切る。\n" +
    "判断基準を示すと、相手は無理に売り込みにくくなり、かつ『考えて動く人だ』という評価になる。\n" +
    "聞くべきこと：①典型的なスコープと期間 ②費用のレンジ ③成果物の形（報告書か、示唆か、リスト付きか） ④その後の支援の有無と費用。\n" +
    "上限額は面談前に必ず自分で決めておくこと。その場で決めると必ず高い方に流れる。"
  );
}

// =======================================================================
// 13. The ask
// =======================================================================
{
  const s = darkSlide("THE ASK", "本日のお願い");
  const asks = [
    ["ヒアリング先を1〜2社ご紹介ください", `条件：従業員 ${TODO("◯◯")}名規模／${TODO("業種")}／${TODO("地域")}。まずは1社で構いません。`],
    ["ご紹介時は「売り込みではない」と一言添えてください", "「知り合いが工場の現場データの話を聞きたがっている。売り込みではなく、話を聞かせてほしいそうだ」"],
    ["Q4・Q5について、代表のご経験からの所感をください", "既存システムとの共存の実態、横展開の型。この2点は現場を見てきた方にしか分からない領域です。"],
  ];
  asks.forEach((a, i) => {
    const y = 1.9 + i * 1.12;
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.7, y, w: 11.9, h: 0.95, rectRadius: 0.08,
      fill: { color: INK2 }, line: { color: "3E566A", width: 1 },
    });
    numChip(s, 1.0, y + 0.25, i + 1, AMBER);
    bodyText(s, a[0], { x: 1.65, y: y + 0.16, w: 10.6, h: 0.35, size: 15, bold: true, color: WHITE });
    bodyText(s, a[1], { x: 1.65, y: y + 0.55, w: 10.6, h: 0.35, size: 12, color: "B9C8D4" });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.7, y: 5.35, w: 11.9, h: 1.3, rectRadius: 0.08,
    fill: { color: AMBER }, line: { color: AMBER, width: 1 },
  });
  bodyText(s, "ご紹介いただく方に、ご迷惑をかけないためのお約束", { x: 1.1, y: 5.55, w: 11.1, h: 0.35, size: 14, bold: true, color: WHITE });
  bodyText(s,
    "① 売り込みは一切しません（30分のヒアリングのみ）　② 必要ならNDAを締結します　③ 伺った内容は必ずご紹介元にも共有します",
    { x: 1.1, y: 5.98, w: 11.1, h: 0.4, size: 12.5, color: "FDF0E4" });
  s.addNotes(
    "紹介する側にとっての最大のリスクは『変な奴を紹介して自分の信用が落ちること』。ここを先に潰すのが最重要。\n" +
    "3つの約束は必ず口頭でも言う。特に『売り込みはしない』は具体的に『30分のヒアリングだけです』と時間で示す。\n" +
    "『まずは1社で構いません』と敷居を下げる。大きくお願いすると断られ、小さくお願いすると通る。\n" +
    "この場で口約束が出たら、その日のうちに紹介文のたたき台をこちらから送る（相手の手間をゼロにする）。"
  );
}

// =======================================================================
// 14. Risks / appendix
// =======================================================================
{
  const s = lightSlide("APPENDIX", "リスクと、まだ答えを持っていないこと");
  const rows = [
    ["競合", `既存の生産管理システム、現場帳票アプリ、そしてExcel。最大の競合は「今のままで困っていない」という現状維持。差別化は ${TODO("何で差をつけるか")}。`],
    ["技術", "工場内のネットワーク環境（オフライン運用の要否）、既存設備との接続可否。ここは実機を見ないと分からない。"],
    ["事業", "1社目がうまくいっても、それが横展開できる型とは限らない。2社目で作り直しになる可能性を織り込んでいる。"],
    ["自分", `${TODO("使える時間／資金／体制。ひとりで続けられる期間を正直に")}`],
  ];
  rows.forEach((r, i) => {
    const y = 1.95 + i * 1.15;
    card(s, { x: 0.7, y, w: 11.9, h: 1.0 });
    s.addShape(pres.ShapeType.roundRect, {
      x: 1.0, y: y + 0.32, w: 1.05, h: 0.36, rectRadius: 0.06,
      fill: { color: INK }, line: { color: INK, width: 1 },
    });
    bodyText(s, r[0], { x: 1.0, y: y + 0.32, w: 1.05, h: 0.36, size: 12, bold: true, color: WHITE, align: "center", valign: "middle" });
    bodyText(s, r[1], { x: 2.35, y: y + 0.22, w: 9.9, h: 0.7, size: 12.5, lineSpacing: 20 });
  });
  foot(s, "分かっていないことを分かっていると言わない。それが継続的にご協力いただくための最低条件だと考えています。");
  s.addNotes(
    "Appendix として持っておき、聞かれたら開く。時間が余れば自分から開いてもよい。\n" +
    "『最大の競合は現状維持』は製造業では特に真実で、相手も強く同意するはず。ここで会話が弾みやすい。\n" +
    "最後の『自分』の行は書きにくいが、ここを書けるかどうかで本気度の伝わり方が変わる。"
  );
}

pres.writeFile({ fileName: "factory-service-brief.pptx" })
  .then((f) => console.log("wrote", f));
