const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "售前解决方案";
pres.title = "3期10楼仪器研发网络安全优化解决方案";

// ── Color palette ──
const C = {
  darkBg:   "0F172A",
  cardBg:   "1E293B",
  teal:     "0D9488",
  tealLt:   "14B8A6",
  tealAcc:  "5EEAD4",
  white:    "FFFFFF",
  light:    "F1F5F9",
  muted:    "94A3B8",
  text:     "1E293B",
  red:      "EF4444",
  amber:    "F59E0B",
  green:    "10B981",
  blue:     "3B82F6",
  purple:   "8B5CF6",
};

// ── Helpers ──
const makeShadow = () => ({ type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.15 });

function sectionDivider(slide, y) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.06, h: 0.45, fill: { color: C.teal } });
}

function card(slide, x, y, w, h, opts = {}) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: opts.fill || C.cardBg },
    shadow: makeShadow(),
  });
  if (opts.accentTop) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 0.04, fill: { color: opts.accentTop },
    });
  }
}

function tag(slide, x, y, label, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.65, h: 0.28, fill: { color },
  });
  slide.addText(label, {
    x, y, w: 0.65, h: 0.28, fontSize: 8, fontFace: "Arial",
    color: C.white, align: "center", valign: "middle", bold: true,
  });
}

// ── Slide 1: Cover ──
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  // Accent bar top
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });
  // Decorative shape
  s.addShape(pres.shapes.RECTANGLE, { x: 7.2, y: 0, w: 2.8, h: 5.625, fill: { color: C.teal, transparency: 85 } });
  // Title
  s.addText("3期10楼仪器研发网络\n安全优化解决方案", {
    x: 0.8, y: 1.2, w: 8.4, h: 2.2, fontSize: 36, fontFace: "Arial",
    color: C.white, bold: true, lineSpacingMultiple: 1.2,
  });
  // Subtitle
  s.addText("PBR策略路由引流 · 防火墙策略隔离 · 终端安全纵深防护", {
    x: 0.8, y: 3.4, w: 8.4, h: 0.5, fontSize: 14, fontFace: "Arial",
    color: C.tealAcc,
  });
  // Bottom info
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.5, w: 2, h: 0.01, fill: { color: C.teal } });
  s.addText("售前解决方案  |  2026-06  |  V2.0", {
    x: 0.8, y: 4.7, w: 5, h: 0.4, fontSize: 10, fontFace: "Arial",
    color: C.muted,
  });
}

// ── Slide 2: Agenda ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  // Header bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.2, fill: { color: C.darkBg } });
  s.addText("目  录", { x: 0.6, y: 0.25, w: 4, h: 0.7, fontSize: 28, fontFace: "Arial", color: C.white, bold: true });

  const items = [
    ["01", "项目背景与需求", "现状挑战 · 需求目标"],
    ["02", "现有网络架构分析", "拓扑结构 · 问题诊断"],
    ["03", "方案对比与选择", "PBR引流方案 · 可行性分析"],
    ["04", "网络改造详细设计", "VLAN规划 · PBR策略 · 防火墙策略"],
    ["05", "终端安全整合优化", "360杀毒 · Ip_guard DLP · 三层联防"],
    ["06", "实施计划与风险", "6阶段实施 · 风险应对"],
  ];

  items.forEach((item, i) => {
    const y = 1.6 + i * 0.62;
    s.addText(item[0], {
      x: 0.6, y, w: 0.6, h: 0.5, fontSize: 22, fontFace: "Arial",
      color: C.teal, bold: true, valign: "middle",
    });
    s.addText(item[1], {
      x: 1.3, y: y + 0.02, w: 3.5, h: 0.3, fontSize: 14, fontFace: "Arial",
      color: C.text, bold: true,
    });
    s.addText(item[2], {
      x: 1.3, y: y + 0.3, w: 5, h: 0.22, fontSize: 9, fontFace: "Arial",
      color: C.muted,
    });
    if (i < items.length - 1) {
      s.addShape(pres.shapes.LINE, {
        x: 1.3, y: y + 0.58, w: 7.8, h: 0,
        line: { color: "E2E8F0", width: 0.5 },
      });
    }
  });
}

// ── Slide 3: Background & Challenges ──
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("项目背景与需求", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Customer profile card
  card(s, 0.5, 1.3, 4.3, 1.8, { fill: C.white });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.04, fill: { color: C.teal } });
  s.addText("客户现状", { x: 0.8, y: 1.45, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.teal, bold: true });
  s.addText([
    { text: "3期10楼为仪器研发组专用区域", options: { bullet: true, breakLine: true } },
    { text: "含超融合集群、AI应用开发、38+业务系统", options: { bullet: true, breakLine: true } },
    { text: "全网互通（2/3/4期光纤互联），安全边界模糊", options: { bullet: true, breakLine: true } },
    { text: "4期山石防火墙负载低，性能可复用", options: { bullet: true } },
  ], { x: 0.8, y: 1.85, w: 3.8, h: 1.1, fontSize: 9, fontFace: "Arial", color: C.text, paraSpaceAfter: 4 });

  // Core requirements card
  card(s, 5.2, 1.3, 4.3, 1.8, { fill: C.white });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.04, fill: { color: C.tealLt } });
  s.addText("核心需求", { x: 5.5, y: 1.45, w: 3, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.tealLt, bold: true });
  s.addText([
    { text: "10楼研发网络安全隔离，缩小攻击面", options: { bullet: true, breakLine: true } },
    { text: "精细化双向访问控制（研发↔2期/外部）", options: { bullet: true, breakLine: true } },
    { text: "不新增防火墙硬件，复用现有设备", options: { bullet: true, breakLine: true } },
    { text: "保障PLM/MES/打印机等业务不受影响", options: { bullet: true } },
  ], { x: 5.5, y: 1.85, w: 3.8, h: 1.1, fontSize: 9, fontFace: "Arial", color: C.text, paraSpaceAfter: 4 });

  // Bottom insight boxes
  const insights = [
    { t: "零硬件成本", d: "复用4期低负载山石防火墙", c: C.teal },
    { t: "零物理施工", d: "纯网络配置变更，不动机柜", c: C.tealLt },
    { t: "业务零中断", d: "PBR Fail-Open自动回退", c: C.green },
    { t: "纵深防护", d: "网络+终端+DLP三层联防", c: C.blue },
  ];
  insights.forEach((ins, i) => {
    const x = 0.5 + i * 2.35;
    card(s, x, 3.45, 2.1, 1.6, { fill: C.white });
    s.addText(ins.t, { x: x + 0.15, y: 3.6, w: 1.8, h: 0.35, fontSize: 13, fontFace: "Arial", color: ins.c, bold: true });
    s.addText(ins.d, { x: x + 0.15, y: 3.95, w: 1.8, h: 0.9, fontSize: 9, fontFace: "Arial", color: C.text });
  });
}

// ── Slide 4: Current Architecture ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("现有网络架构", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Simplified topology using shapes
  // Internet cloud
  s.addShape(pres.shapes.RECTANGLE, { x: 4, y: 1.2, w: 2, h: 0.4, fill: { color: C.muted } });
  s.addText("Internet / 专线 / ADSL", { x: 4, y: 1.2, w: 2, h: 0.4, fontSize: 7, fontFace: "Arial", color: C.white, align: "center", valign: "middle" });

  // 2期 center
  s.addShape(pres.shapes.RECTANGLE, { x: 3.2, y: 1.9, w: 3.6, h: 1.7, fill: { color: C.darkBg } });
  s.addText("2期机房（核心）\n山石防火墙 | 行为管理 | AD/DNS\n360天擎 | 安全网关 | 态势感知(计划)", {
    x: 3.3, y: 2.0, w: 3.4, h: 1.5, fontSize: 7.5, fontFace: "Arial", color: C.white, align: "center", valign: "middle",
  });

  // Connection lines
  s.addShape(pres.shapes.LINE, { x: 5, y: 1.6, w: 0, h: 0.3, line: { color: C.muted, width: 1 } });

  // 3期 box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.0, w: 4.5, h: 1.4, fill: { color: C.cardBg } });
  s.addText("3期汇聚层\n核心交换机 | WAF | Ip_guard服务端\nVLAN配置 | 策略路由", {
    x: 0.4, y: 4.08, w: 4.3, h: 1.25, fontSize: 7.5, fontFace: "Arial", color: C.white, align: "center", valign: "middle",
  });

  // 4期 box
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 4.0, w: 4.5, h: 1.4, fill: { color: C.cardBg } });
  s.addText("4期（低负载）\n核心交换机 | ★山石防火墙(2期替换)\n策略路由未配", {
    x: 5.3, y: 4.08, w: 4.3, h: 1.25, fontSize: 7.5, fontFace: "Arial", color: C.white, align: "center", valign: "middle",
  });

  // 10楼 box (inside 3期)
  s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 4.6, w: 4.5, h: 0.7, fill: { color: C.teal, transparency: 70 } });
  s.addText("10楼接入层 · 研发PC(360+Ip_guard) · 超融合 · 物理服务器(PLM/MES/监控等)", {
    x: 0.4, y: 4.6, w: 4.3, h: 0.7, fontSize: 7, fontFace: "Arial", color: C.white, align: "center", valign: "middle",
  });

  // Connection lines from 2期 to 3期/4期
  s.addShape(pres.shapes.LINE, { x: 3.4, y: 3.6, w: 0, h: 0.4, line: { color: C.teal, width: 1.5 } });
  s.addShape(pres.shapes.LINE, { x: 6.6, y: 3.6, w: 0, h: 0.4, line: { color: C.teal, width: 1.5 } });
  s.addText("光纤", { x: 2.8, y: 3.72, w: 0.6, h: 0.2, fontSize: 6, fontFace: "Arial", color: C.teal });
  s.addText("光纤", { x: 6.2, y: 3.72, w: 0.6, h: 0.2, fontSize: 6, fontFace: "Arial", color: C.teal });

  // Key equipment list on the right
  card(s, 5.2, 1.2, 4.3, 2.5, { fill: C.light });
  s.addText("关键设备清单", { x: 5.4, y: 1.3, w: 3, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.teal, bold: true });
  const equip = [
    ["2期", "山石防火墙(新) 深信服行为管理 态势感知(计划)"],
    ["3期", "WAF Ip_guard服务端 华为核心交换机"],
    ["4期", "山石防火墙(2期替换退役) 华为核心交换机"],
    ["终端", "360杀毒(全区域) Ip_guard客户端(研发PC)"],
  ];
  equip.forEach((eq, i) => {
    const ey = 1.7 + i * 0.45;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: ey, w: 0.5, h: 0.25, fill: { color: i === 0 ? C.teal : i === 1 ? C.tealLt : i === 2 ? C.amber : C.blue } });
    s.addText(eq[0], { x: 5.4, y: ey, w: 0.5, h: 0.25, fontSize: 7, fontFace: "Arial", color: C.white, align: "center", valign: "middle", bold: true });
    s.addText(eq[1], { x: 6.0, y: ey, w: 3.3, h: 0.25, fontSize: 7.5, fontFace: "Arial", color: C.text, valign: "middle" });
  });
}

// ── Slide 5: Problem Diagnosis ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("现状问题诊断", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  const problems = [
    { icon: "01", title: "安全边界模糊", desc: "研发网与生产网全网互通，一旦被攻击可横向扩散至全区域", color: C.red },
    { icon: "02", title: "访问控制粗放", desc: "仅依赖VLAN做基本隔离，缺乏IP/端口/协议精细化策略控制", color: C.amber },
    { icon: "03", title: "数据泄露隐患", desc: "仪器研发知识产权与生产/办公混网，合规风险大", color: C.red },
    { icon: "04", title: "设备复用不足", desc: "4期山石防火墙性能利用率低，未能发挥策略控制价值", color: C.amber },
    { icon: "05", title: "360杀毒未深度利用", desc: "仅用基础查杀，EDR/漏洞管理/资产清点等高阶能力未启用", color: C.blue },
    { icon: "06", title: "Ip_guard策略粗放", desc: "DLP策略未针对研发场景（源码防泄漏/AI模型保护）精细化", color: C.purple },
  ];

  problems.forEach((p, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 3.1;
    const y = 1.3 + row * 2.0;

    card(s, x, y, 2.85, 1.7, { fill: C.light });
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.85, h: 0.04, fill: { color: p.color } });
    s.addText(p.icon, {
      x: x + 0.15, y: y + 0.2, w: 0.45, h: 0.45, fontSize: 18, fontFace: "Arial",
      color: p.color, bold: true, valign: "middle",
    });
    s.addText(p.title, {
      x: x + 0.65, y: y + 0.2, w: 2, h: 0.35, fontSize: 13, fontFace: "Arial",
      color: C.text, bold: true, valign: "middle",
    });
    s.addText(p.desc, {
      x: x + 0.15, y: y + 0.75, w: 2.55, h: 0.8, fontSize: 8.5, fontFace: "Arial",
      color: C.muted,
    });
  });
}

// ── Slide 6: Solution Comparison ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("方案对比与选择", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Option 1 - Recommended
  card(s, 0.5, 1.3, 4.3, 3.8, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.3, h: 0.05, fill: { color: C.green } });
  tag(s, 5.1, 1.45, "★ 推荐", C.green);
  s.addText("方案一：PBR策略路由引流", {
    x: 0.8, y: 1.5, w: 3.5, h: 0.4, fontSize: 16, fontFace: "Arial", color: C.text, bold: true,
  });
  s.addText("技术可行性 ★★★★★", {
    x: 0.8, y: 1.9, w: 3, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.green, bold: true,
  });

  s.addText([
    { text: "核心思路", options: { bold: true, breakLine: true, fontSize: 11, color: C.teal } },
    { text: "4期防火墙物理位置不动，3期华为汇聚交换机配置PBR，将10楼VLAN流量下一跳指向4期防火墙做策略检查后回注正常路径", options: { breakLine: true, fontSize: 9 } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "核心优势", options: { bold: true, breakLine: true, fontSize: 11, color: C.teal } },
    { text: "零硬件成本 · 零物理施工 · PBR回退<5分钟", options: { bullet: true, breakLine: true, fontSize: 9 } },
    { text: "3期↔4期光纤直连，额外延迟≤1ms", options: { bullet: true, breakLine: true, fontSize: 9 } },
    { text: "不影响4期现有业务，远期可改PBR收回", options: { bullet: true, fontSize: 9 } },
  ], { x: 0.8, y: 2.3, w: 3.8, h: 2.7, fontFace: "Arial", color: C.text, paraSpaceAfter: 3 });

  // Option 2 - Not recommended
  card(s, 5.2, 1.3, 4.3, 3.8, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.3, h: 0.05, fill: { color: C.red } });
  tag(s, 9.8, 1.45, "✗ 不采纳", C.red);
  s.addText("方案二：新增防火墙 + 独立ADSL", {
    x: 5.5, y: 1.5, w: 3.5, h: 0.4, fontSize: 16, fontFace: "Arial", color: C.text, bold: true,
  });
  s.addText("技术可行性 ★★", {
    x: 5.5, y: 1.9, w: 3, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.red, bold: true,
  });
  s.addText([
    { text: "客户反馈", options: { bold: true, breakLine: true, fontSize: 11, color: C.red } },
    { text: "成本高：新购防火墙 + ADSL月租 + 施工", options: { bullet: true, breakLine: true, fontSize: 9 } },
    { text: "技术不可行：独立ADSL后10楼脱离现有网络，但20+套业务系统（PLM/MES等）部署在10楼供全区域使用", options: { bullet: true, breakLine: true, fontSize: 9 } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "远期可重新评估（当服务器迁出10楼后）", options: { italic: true, fontSize: 8, color: C.muted } },
  ], { x: 5.5, y: 2.3, w: 3.8, h: 2.7, fontFace: "Arial", color: C.text, paraSpaceAfter: 3 });
}

// ── Slide 7: PBR Core Concept ──
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });
  s.addText("PBR策略路由引流 — 核心理念", {
    x: 0.6, y: 0.3, w: 8, h: 0.6, fontSize: 26, fontFace: "Arial", color: C.white, bold: true,
  });
  s.addText("\"流量走过去做策略，设备不挪窝\"", {
    x: 0.6, y: 0.9, w: 8, h: 0.4, fontSize: 14, fontFace: "Arial", color: C.tealAcc, italic: true,
  });

  // Flow diagram
  // Before
  s.addText("改造前", { x: 0.5, y: 1.5, w: 1.5, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.muted });
  const beforeBoxes = ["10楼研发PC", "3期汇聚\n交换机", "光纤", "2期核心", "2期系统\nInternet"];
  beforeBoxes.forEach((b, i) => {
    const bx = 0.5 + i * 1.9;
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: 1.85, w: 1.5, h: 0.7, fill: { color: C.cardBg } });
    s.addText(b, { x: bx, y: 1.85, w: 1.5, h: 0.7, fontSize: 8, fontFace: "Arial", color: C.white, align: "center", valign: "middle" });
    if (i < beforeBoxes.length - 1) {
      s.addText("→", { x: bx + 1.5, y: 1.85, w: 0.4, h: 0.7, fontSize: 16, fontFace: "Arial", color: C.muted, align: "center", valign: "middle" });
    }
  });
  s.addText("直通，无策略检查", { x: 0.5, y: 2.65, w: 8.5, h: 0.25, fontSize: 8, fontFace: "Arial", color: C.red, align: "center" });

  // After
  s.addText("改造后", { x: 0.5, y: 3.05, w: 1.5, h: 0.3, fontSize: 10, fontFace: "Arial", color: C.tealAcc });
  const afterBoxes = [
    { t: "10楼研发PC", c: C.cardBg },
    { t: "3期汇聚\nPBR下一跳→4期FW", c: C.teal },
    { t: "4期山石\n防火墙\n策略检查", c: C.tealLt },
    { t: "策略放行\n→回程", c: C.green },
    { t: "2期系统\nInternet", c: C.cardBg },
  ];
  afterBoxes.forEach((b, i) => {
    const ax = 0.5 + i * 1.9;
    s.addShape(pres.shapes.RECTANGLE, { x: ax, y: 3.4, w: 1.5, h: 0.85, fill: { color: b.c } });
    s.addText(b.t, { x: ax, y: 3.4, w: 1.5, h: 0.85, fontSize: 8, fontFace: "Arial", color: C.white, align: "center", valign: "middle" });
    if (i < afterBoxes.length - 1) {
      s.addText("→", { x: ax + 1.5, y: 3.4, w: 0.4, h: 0.85, fontSize: 16, fontFace: "Arial", color: C.tealAcc, align: "center", valign: "middle" });
    }
  });
  s.addText("PBR牵引 → 策略检查 → 放行回注", { x: 0.5, y: 4.35, w: 8.5, h: 0.25, fontSize: 8, fontFace: "Arial", color: C.green, align: "center" });

  // Key metrics
  const metrics = [
    { n: "0", l: "硬件成本（元）", c: C.tealAcc },
    { n: "0", l: "物理施工（天）", c: C.tealAcc },
    { n: "<5min", l: "故障回退时间", c: C.tealAcc },
    { n: "≤1ms", l: "额外延迟", c: C.tealAcc },
  ];
  metrics.forEach((m, i) => {
    const mx = 0.8 + i * 2.3;
    s.addText(m.n, {
      x: mx, y: 4.65, w: 1.8, h: 0.5, fontSize: 24, fontFace: "Arial", color: m.c, bold: true,
      align: "center", valign: "middle",
    });
    s.addText(m.l, {
      x: mx, y: 5.1, w: 1.8, h: 0.3, fontSize: 7.5, fontFace: "Arial", color: C.muted,
      align: "center",
    });
  });
}

// ── Slide 8: VLAN Planning ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("VLAN规划 & 网络改造架构", { x: 0.6, y: 0.18, w: 6, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // VLAN table
  const vlanHeader = [
    { text: "VLAN ID", options: { bold: true, color: C.white, fill: { color: C.teal }, fontSize: 9 } },
    { text: "网段", options: { bold: true, color: C.white, fill: { color: C.teal }, fontSize: 9 } },
    { text: "用途", options: { bold: true, color: C.white, fill: { color: C.teal }, fontSize: 9 } },
    { text: "安全级别", options: { bold: true, color: C.white, fill: { color: C.teal }, fontSize: 9 } },
    { text: "说明", options: { bold: true, color: C.white, fill: { color: C.teal }, fontSize: 9 } },
  ];
  const vlanRows = [
    ["100", "10.0.100.0/24", "研发PC", "高", "PBR→4期防火墙"],
    ["110", "10.0.110.0/24", "超融合集群(内部)", "高", "仅研发访问；PBR→防火墙"],
    ["120", "10.0.120.0/24", "物理服务器", "中", "PLM/MES/打印机；对外端口开放"],
    ["200", "10.0.200.0/24", "管理网段", "最高", "仅IT运维；PBR→防火墙"],
    ["300", "10.0.10.0/24", "3期其他楼层", "低", "已有；ACL隔离10楼VLAN"],
  ];

  s.addTable([vlanHeader, ...vlanRows.map(r => r.map(c => ({ text: c, options: { fontSize: 8, color: C.text } })))], {
    x: 0.5, y: 1.2, w: 9, colW: [0.9, 1.6, 2.0, 1.0, 3.5],
    border: { pt: 0.5, color: "E2E8F0" },
    rowH: [0.35, 0.32, 0.32, 0.32, 0.32, 0.32],
    autoPage: false,
  });

  // 3 traffic paths
  s.addText("流量路径设计", { x: 0.5, y: 3.3, w: 4, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.text, bold: true });

  const paths = [
    { label: "出站", flow: "10楼PC → 3期汇聚[PBR] → 4期FW[策略] → 2期/Internet", color: C.teal },
    { label: "入站", flow: "外部 → 3期汇聚[PBR] → 4期FW[策略] → VLAN 120服务器", color: C.tealLt },
    { label: "本地", flow: "VLAN 100 ↔ VLAN 120（PBR排除，同楼直达）", color: C.green },
  ];
  paths.forEach((p, i) => {
    const py = 3.7 + i * 0.55;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: py, w: 0.7, h: 0.35, fill: { color: p.color } });
    s.addText(p.label, { x: 0.5, y: py, w: 0.7, h: 0.35, fontSize: 9, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" });
    s.addText(p.flow, { x: 1.35, y: py, w: 8.1, h: 0.35, fontSize: 9, fontFace: "Arial", color: C.text, valign: "middle" });
  });
}

// ── Slide 9: PBR Detailed Design ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("PBR策略路由方案 — 核心机制", { x: 0.6, y: 0.18, w: 7, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Left: config logic
  card(s, 0.5, 1.2, 4.5, 2.6, { fill: C.light });
  s.addText("PBR配置逻辑（华为交换机）", { x: 0.7, y: 1.3, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });
  s.addText([
    { text: "出站引流", options: { bold: true, breakLine: true, color: C.teal } },
    { text: "ACL 3010: permit ip 10.0.100.0/24 → 匹配VLAN 100", options: { breakLine: true } },
    { text: "ACL 3010: permit ip 10.0.110.0/24 → 匹配VLAN 110", options: { breakLine: true } },
    { text: "PBR node 10: if-match acl 3010", options: { breakLine: true } },
    { text: "  apply next-hop <4期防火墙Trust IP>", options: { breakLine: true, color: C.teal } },
    { text: "", options: { breakLine: true } },
    { text: "入站引流", options: { bold: true, breakLine: true, color: C.tealLt } },
    { text: "ACL 3020: permit ip dest 10.0.120.0/24 → 匹配VLAN 120", options: { breakLine: true } },
    { text: "PBR node 20: if-match acl 3020", options: { breakLine: true } },
    { text: "  apply next-hop <4期防火墙Trust IP>", options: { color: C.tealLt } },
  ], { x: 0.7, y: 1.6, w: 4.1, h: 2.1, fontSize: 7.5, fontFace: "Courier New", color: C.text, paraSpaceAfter: 1 });

  // Right: PBR exclusion table
  card(s, 5.3, 1.2, 4.3, 2.6, { fill: C.light });
  s.addText("PBR排除项（不引向防火墙）", { x: 5.5, y: 1.3, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });
  const exclusions = [
    ["VLAN间互访", "研发PC↔服务器同楼直达"],
    ["备份流量", "TB级大流量，防火墙无法承载"],
    ["Ip_guard通信", "同楼管控，无需出楼"],
    ["超融合东西向", "VLAN 110内部L2通信"],
  ];
  exclusions.forEach((ex, i) => {
    const ey = 1.7 + i * 0.45;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: ey, w: 3.9, h: 0.35, fill: { color: C.white } });
    s.addText(ex[0], { x: 5.6, y: ey, w: 1.5, h: 0.35, fontSize: 9, fontFace: "Arial", color: C.amber, bold: true, valign: "middle" });
    s.addText(ex[1], { x: 7.1, y: ey, w: 2.2, h: 0.35, fontSize: 8, fontFace: "Arial", color: C.muted, valign: "middle" });
  });

  // Bottom: 3-layer architecture
  s.addText("分层控制架构", { x: 0.5, y: 4.1, w: 4, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.text, bold: true });

  const layers = [
    { label: "第一层", device: "3期华为汇聚交换机", role: "ACL粗粒度隔离 + PBR流量牵引", color: C.blue, y: 4.5 },
    { label: "第二层", device: "4期山石防火墙", role: "L3/L4 ACL + IPS/AV + 会话日志", color: C.teal, y: 4.5 },
    { label: "第三层", device: "2期山石防火墙(出口)", role: "Internet出口策略（已有，不变）", color: C.cardBg, y: 4.5 },
  ];
  layers.forEach((l, i) => {
    const lx = 0.5 + i * 3.15;
    s.addShape(pres.shapes.RECTANGLE, { x: lx, y: 4.5, w: 2.9, h: 0.9, fill: { color: l.color } });
    s.addText(l.label, { x: lx, y: 4.5, w: 2.9, h: 0.3, fontSize: 8, fontFace: "Arial", color: C.white, align: "center", bold: true });
    s.addText(l.device, { x: lx, y: 4.8, w: 2.9, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.white, align: "center" });
    s.addText(l.role, { x: lx, y: 5.1, w: 2.9, h: 0.25, fontSize: 7.5, fontFace: "Arial", color: l.color === C.cardBg ? C.muted : "5EEAD4", align: "center" });
    if (i < 2) {
      s.addText("PBR→", { x: lx + 2.9, y: 4.5, w: 0.25, h: 0.9, fontSize: 9, fontFace: "Arial", color: C.tealAcc, align: "center", valign: "middle" });
    }
  });
}

// ── Slide 10: Firewall Policies ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("防火墙安全策略设计", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Policy overview
  card(s, 0.5, 1.2, 4.3, 1.0, { fill: C.darkBg });
  s.addText("策略总则：默认拒绝 + 白名单放行", {
    x: 0.7, y: 1.3, w: 3.8, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.tealAcc, bold: true,
  });
  s.addText("Any → Any : DENY（记录日志）  ·  逐条开放业务所需最小权限", {
    x: 0.7, y: 1.65, w: 3.8, h: 0.35, fontSize: 9, fontFace: "Arial", color: C.muted,
  });

  // Core policy table
  const pHeader = [
    { text: "#", options: { bold: true, fill: { color: C.teal }, color: C.white, fontSize: 7 } },
    { text: "源", options: { bold: true, fill: { color: C.teal }, color: C.white, fontSize: 7 } },
    { text: "目的", options: { bold: true, fill: { color: C.teal }, color: C.white, fontSize: 7 } },
    { text: "端口", options: { bold: true, fill: { color: C.teal }, color: C.white, fontSize: 7 } },
    { text: "说明", options: { bold: true, fill: { color: C.teal }, color: C.white, fontSize: 7 } },
  ];
  const pRows = [
    ["1", "VLAN 100", "2期DNS", "UDP 53", "DNS解析"],
    ["2", "VLAN 100", "2期AD域控", "88/389/636", "域认证"],
    ["3", "VLAN 100", "2期安全网关", "业务端口", "安全网关"],
    ["4", "VLAN 100", "2期网络AC", "443/8443", "准入控制"],
    ["5", "VLAN 100", "2期监控", "80/443", "监控平台"],
    ["6", "VLAN 100/110", "VLAN 120", "业务端口", "PLM/MES（PBR排除）"],
    ["7", "外部区域", "VLAN 120", "业务端口", "PLM/MES/打印机"],
    ["8", "VLAN 100/110", "Internet", "80/443", "互联网访问"],
    ["9", "IT管理段", "VLAN 200", "SSH/RDP", "运维管理"],
    ["10", "Any", "Any", "Any", "默认拒绝+日志"],
  ];

  s.addTable([pHeader, ...pRows.map(r => r.map(c => ({ text: c, options: { fontSize: 7, color: r[0] === "10" ? C.red : C.text } })))], {
    x: 0.5, y: 2.4, w: 9, colW: [0.3, 1.4, 1.5, 1.2, 4.6],
    border: { pt: 0.3, color: "E2E8F0" },
    rowH: 0.26,
  });

  // Special traffic treatments
  s.addText("特殊流量处理", { x: 0.5, y: 5.1, w: 4, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.text, bold: true });
  const special = ["备份流量: PBR排除，直通2期↔3期", "360杀毒: 经PBR→4期FW放行（TCP 8443/80/443）", "Ip_guard: 同楼本地直达（PBR排除），IT远程管理走PBR→FW"];
  special.forEach((sp, i) => {
    s.addText(sp, { x: 0.5 + (i % 2) * 4.7, y: 5.35, w: 4.5, h: 0.2, fontSize: 7, fontFace: "Arial", color: C.muted });
  });
}

// ── Slide 11: Backup & PBR Failover ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("备份流量优化 & 故障切换", { x: 0.6, y: 0.18, w: 7, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Backup section
  card(s, 0.5, 1.2, 4.3, 2.0, { fill: C.light, accentTop: C.amber });
  s.addText("备份流量 — PBR排除", { x: 0.7, y: 1.35, w: 3.5, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.text, bold: true });
  s.addText([
    { text: "问题", options: { bold: true, breakLine: true, color: C.amber } },
    { text: "2期→3期10楼备份数据量TB级，经PBR→防火墙会拖垮性能", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "方案", options: { bold: true, breakLine: true, color: C.green } },
    { text: "PBR ACL显式排除备份源/目的IP → 不命中PBR → 走正常路由直达", options: { breakLine: true } },
    { text: "备份VLAN 999逻辑隔离，仅备份服务器间可达", options: {} },
  ], { x: 0.7, y: 1.7, w: 3.9, h: 1.4, fontSize: 8.5, fontFace: "Arial", color: C.text, paraSpaceAfter: 3 });

  // Failover section
  card(s, 5.2, 1.2, 4.3, 2.0, { fill: C.light, accentTop: C.green });
  s.addText("PBR故障切换 — Track联动", { x: 5.4, y: 1.35, w: 3.5, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.text, bold: true });
  s.addText([
    { text: "机制", options: { bold: true, breakLine: true, color: C.teal } },
    { text: "PBR关联NQA track检测4期防火墙可达性", options: { breakLine: true } },
    { text: "", options: { breakLine: true } },
    { text: "故障时", options: { bold: true, breakLine: true, color: C.red } },
    { text: "track检测失败 → PBR node自动失效", options: { breakLine: true } },
    { text: "→ 流量回退正常路由表（Fail-Open）", options: { breakLine: true } },
    { text: "→ 10楼恢复未隔离状态，保证业务连续", options: { breakLine: true } },
    { text: "→ 触发告警通知IT运维", options: {} },
  ], { x: 5.4, y: 1.7, w: 3.9, h: 1.4, fontSize: 8.5, fontFace: "Arial", color: C.text, paraSpaceAfter: 3 });

  // Bottom: Implementation order
  s.addText("PBR上线顺序（逐网段渐进，风险可控）", { x: 0.5, y: 3.5, w: 6, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });
  const phases = [
    { step: "Step 1", desc: "VLAN 100(研发PC)试点", color: C.teal },
    { step: "Step 2", desc: "验证→加入VLAN 110", color: C.tealLt },
    { step: "Step 3", desc: "验证→加入VLAN 120入站", color: C.green },
    { step: "Step 4", desc: "全量PBR+防火墙Deny生效", color: C.blue },
  ];
  phases.forEach((ph, i) => {
    const px = 0.5 + i * 2.35;
    s.addShape(pres.shapes.RECTANGLE, { x: px, y: 3.9, w: 2.1, h: 0.7, fill: { color: C.light } });
    s.addShape(pres.shapes.RECTANGLE, { x: px, y: 3.9, w: 0.06, h: 0.7, fill: { color: ph.color } });
    s.addText(ph.step, { x: px + 0.2, y: 3.95, w: 1.7, h: 0.25, fontSize: 8, fontFace: "Arial", color: ph.color, bold: true });
    s.addText(ph.desc, { x: px + 0.2, y: 4.2, w: 1.7, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.text });
    if (i < 3) {
      s.addText("→", { x: px + 2.1, y: 3.9, w: 0.25, h: 0.7, fontSize: 14, fontFace: "Arial", color: C.muted, align: "center", valign: "middle" });
    }
  });

  // Rollback note
  card(s, 0.5, 4.8, 9, 0.65, { fill: C.darkBg });
  s.addText("回退方案：删除PBR配置即可恢复原状，回退时间 < 5分钟  |  防火墙策略初期Log模式观察1-2周 → 确认无误后切换Deny", {
    x: 0.7, y: 4.85, w: 8.6, h: 0.55, fontSize: 9, fontFace: "Arial", color: C.tealAcc, valign: "middle",
  });
}

// ── Slide 12: 3-Layer Defense Architecture ──
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });
  s.addText("终端安全整合 — 三层联防体系", { x: 0.6, y: 0.3, w: 8, h: 0.6, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });
  s.addText("网络边界(PBR+防火墙) → 终端检测(360 EDR) → 数据防泄漏(Ip_guard DLP)", {
    x: 0.6, y: 0.85, w: 8.5, h: 0.4, fontSize: 12, fontFace: "Arial", color: C.tealAcc,
  });

  // 3 columns
  const layers2 = [
    {
      title: "第一层：网络边界", device: "4期山石防火墙", color: C.teal,
      items: ["南北向访问控制", "入侵防御 IPS/AV", "URL过滤", "会话日志 → 态势感知"],
      y: 1.4,
    },
    {
      title: "第二层：终端检测", device: "360天擎 EDR", color: C.blue,
      items: ["病毒/木马实时查杀", "终端行为检测(EDR)", "漏洞扫描与补丁", "资产清点"],
      y: 1.4,
    },
    {
      title: "第三层：数据防泄漏", device: "Ip_guard DLP", color: C.purple,
      items: ["文档透明加密", "外发通道管控", "USB外设管控", "屏幕水印+审计"],
      y: 1.4,
    },
  ];
  layers2.forEach((l, i) => {
    const lx = 0.5 + i * 3.15;
    card(s, lx, 1.4, 2.9, 3.9, { fill: C.cardBg });
    s.addShape(pres.shapes.RECTANGLE, { x: lx, y: 1.4, w: 2.9, h: 0.05, fill: { color: l.color } });
    s.addText(l.title, { x: lx + 0.15, y: 1.55, w: 2.6, h: 0.35, fontSize: 12, fontFace: "Arial", color: l.color, bold: true });
    s.addText(l.device, { x: lx + 0.15, y: 1.9, w: 2.6, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.muted });
    l.items.forEach((item, j) => {
      s.addShape(pres.shapes.RECTANGLE, { x: lx + 0.15, y: 2.3 + j * 0.55, w: 2.6, h: 0.42, fill: { color: C.darkBg } });
      s.addText(item, { x: lx + 0.3, y: 2.3 + j * 0.55, w: 2.3, h: 0.42, fontSize: 9, fontFace: "Arial", color: C.light, valign: "middle" });
    });
    // Arrow between columns
    if (i < 2) {
      s.addText("▼", { x: lx + 2.9, y: 2.5, w: 0.25, h: 1, fontSize: 20, fontFace: "Arial", color: C.tealAcc, align: "center", valign: "middle" });
    }
  });
}

// ── Slide 13: 360 Optimization ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("360杀毒能力优化", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Left: current vs target
  card(s, 0.5, 1.2, 4.3, 2.0, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.3, h: 0.04, fill: { color: C.blue } });
  s.addText("能力启用规划", { x: 0.7, y: 1.35, w: 3, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });
  const capRows = [
    ["病毒查杀", "✅ 已启用", "研发场景定制排除"],
    ["实时防护", "✅ 已启用", "调优敏感度，减少误拦"],
    ["EDR检测响应", "❌ 未启用 → 建议启用", "持续监控异常行为"],
    ["漏洞管理", "❌ 未启用 → 建议启用", "工具链漏洞主动发现"],
    ["资产清点", "❌ 未启用 → 建议启用", "精确软硬件清单"],
    ["微隔离", "❌ 未启用 → 评估", "PC间横向移动防护"],
  ];
  const capTable = capRows.map((r, i) => [
    { text: r[0], options: { fontSize: 8, bold: true, color: C.text } },
    { text: r[1], options: { fontSize: 7, color: r[1].includes("✅") ? C.green : C.amber } },
    { text: r[2], options: { fontSize: 7, color: C.muted } },
  ]);
  s.addTable(capTable, {
    x: 0.7, y: 1.7, w: 3.9, colW: [1.2, 1.5, 1.2],
    border: { pt: 0.3, color: "E2E8F0" },
    rowH: 0.22,
  });

  // Right: R&D customizations
  card(s, 5.2, 1.2, 4.3, 2.0, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.2, w: 4.3, h: 0.04, fill: { color: C.teal } });
  s.addText("研发专属策略优化", { x: 5.4, y: 1.35, w: 3, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });
  s.addText([
    { text: "编译目录白名单", options: { bold: true, breakLine: true } },
    { text: "build/ dist/ target/ __pycache__/ 等扫描排除", options: { breakLine: true } },
    { text: "容器/虚拟化兼容", options: { bold: true, breakLine: true } },
    { text: "Docker镜像/虚拟机磁盘目录 IO优化", options: { breakLine: true } },
    { text: "脚本语言防护增强", options: { bold: true, breakLine: true } },
    { text: "PowerShell/Python/Node.js执行监控", options: { breakLine: true } },
    { text: "AI开发环境防护", options: { bold: true, breakLine: true } },
    { text: "pip/conda安装路径只读扫描 模型文件(.pth/.onnx)排除", options: {} },
  ], { x: 5.4, y: 1.7, w: 3.9, h: 1.4, fontSize: 7.5, fontFace: "Arial", color: C.text, paraSpaceAfter: 2 });

  // Bottom: EDR detection dimensions
  card(s, 0.5, 3.45, 9, 1.75, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.45, w: 9, h: 0.04, fill: { color: C.amber } });
  s.addText("EDR核心监控维度", { x: 0.7, y: 3.6, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });

  const edrDims = [
    { t: "异常进程行为", d: "提权操作监控 · 工具链进程注入检测 · Git/SVN异常外联（代码窃取特征）" },
    { t: "横向移动检测", d: "VLAN间非预期访问 · SMB/RDP/SSH异常连接 · 防火墙拦截+EDR上下文=完整攻击链" },
    { t: "勒索软件防护", d: "批量文件加密/重命名监控 · 超融合集群存储IO突变检测" },
  ];
  edrDims.forEach((dim, i) => {
    const dx = 0.7 + i * 3.0;
    s.addText(dim.t, { x: dx, y: 3.95, w: 2.7, h: 0.25, fontSize: 10, fontFace: "Arial", color: C.amber, bold: true });
    s.addText(dim.d, { x: dx, y: 4.2, w: 2.7, h: 0.8, fontSize: 7.5, fontFace: "Arial", color: C.text });
  });

  // Communication path
  s.addText("360通信路径：研发PC → 3期汇聚[PBR] → 4期FW[放行TCP 8443/80/443] → 2期核心 → 360天擎", {
    x: 0.7, y: 5.3, w: 8.6, h: 0.2, fontSize: 7, fontFace: "Arial", color: C.muted,
  });
}

// ── Slide 14: Ip_guard Optimization ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("Ip_guard DLP能力优化", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  // Data classification
  card(s, 0.5, 1.2, 4.5, 2.1, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 4.5, h: 0.04, fill: { color: C.red } });
  s.addText("数据分类分级保护", { x: 0.7, y: 1.35, w: 3, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });
  const dataClasses = [
    ["绝密", "核心算法源码/专利设计图/AI模型权重", "加密+禁止外发+禁USB+水印+全审计", C.red],
    ["机密", "仪器设计文档/固件代码/测试用例", "加密+外发审批+USB只读+审计", C.amber],
    ["内部", "开发规范/会议纪要/项目周报", "外发审计+敏感词检测", C.blue],
    ["公开", "技术白皮书/产品手册", "无特殊限制", C.green],
  ];
  dataClasses.forEach((dc, i) => {
    const dy = 1.75 + i * 0.35;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: dy, w: 0.55, h: 0.25, fill: { color: dc[3] } });
    s.addText(dc[0], { x: 0.7, y: dy, w: 0.55, h: 0.25, fontSize: 8, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" });
    s.addText(dc[1], { x: 1.35, y: dy, w: 1.7, h: 0.25, fontSize: 7.5, fontFace: "Arial", color: C.text, valign: "middle" });
    s.addText(dc[2], { x: 3.05, y: dy, w: 1.8, h: 0.25, fontSize: 7, fontFace: "Arial", color: C.muted, valign: "middle" });
  });

  // DLP policy details
  card(s, 5.3, 1.2, 4.3, 2.1, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.2, w: 4.3, h: 0.04, fill: { color: C.purple } });
  s.addText("DLP策略精细化", { x: 5.5, y: 1.35, w: 3, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });
  s.addText([
    { text: "文档加密", options: { bold: true, breakLine: true, color: C.purple } },
    { text: "*.c/*.cpp/*.py/*.java/*.go/*.dwg/*.pth 透明加密", options: { breakLine: true } },
    { text: "USB管控", options: { bold: true, breakLine: true, color: C.purple } },
    { text: "默认禁用 → 审批临时开通；调试器/JTAG白名单放行", options: { breakLine: true } },
    { text: "屏幕水印", options: { bold: true, breakLine: true, color: C.purple } },
    { text: "研发PC强制水印（工号+时间戳），AD账号联动追溯", options: { breakLine: true } },
    { text: "外发审批", options: { bold: true, breakLine: true, color: C.purple } },
    { text: "低风险自动放行；高风险 → Leader+IT安全双签审批", options: {} },
  ], { x: 5.5, y: 1.7, w: 3.9, h: 1.5, fontSize: 8, fontFace: "Arial", color: C.text, paraSpaceAfter: 2 });

  // Integration points
  card(s, 0.5, 3.55, 9, 1.6, { fill: C.light });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.55, w: 9, h: 0.04, fill: { color: C.teal } });
  s.addText("与网络隔离的协同设计", { x: 0.7, y: 3.7, w: 5, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.text, bold: true });

  const synergies = [
    { t: "VLAN分层策略", d: "研发PC(VLAN 100)全面DLP；服务器(VLAN 120)仅审计不加密" },
    { t: "PBR通信保障", d: "Ip_guard同楼通信PBR排除直达；IT管理台跨区走PBR→FW放行" },
    { t: "远程场景协同", d: "VPN/零信任接入自动收紧DLP(剪贴板/截屏/外发)；远期联动aTrust" },
  ];
  synergies.forEach((sy, i) => {
    const sx = 0.7 + i * 3.0;
    s.addShape(pres.shapes.RECTANGLE, { x: sx, y: 4.05, w: 2.7, h: 0.9, fill: { color: C.white } });
    s.addText(sy.t, { x: sx + 0.1, y: 4.1, w: 2.5, h: 0.25, fontSize: 9, fontFace: "Arial", color: C.teal, bold: true });
    s.addText(sy.d, { x: sx + 0.1, y: 4.4, w: 2.5, h: 0.45, fontSize: 7.5, fontFace: "Arial", color: C.text });
  });
}

// ── Slide 15: Implementation Plan ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("实施计划", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  const stages = [
    { n: "阶段一", t: "调研", time: "1周", risk: "低", color: C.teal,
      desc: "盘点IP/端口/访问关系；评估防火墙规格；梳理360和Ip_guard现状" },
    { n: "阶段二", t: "方案细化", time: "1周", risk: "低", color: C.tealLt,
      desc: "完善VLAN/PBR ACL/防火墙策略表；输出360和Ip_guard优化文档" },
    { n: "阶段三", t: "测试验证", time: "2-3天", risk: "中", color: C.amber,
      desc: "PBR试点IP测试；验证业务可达性、360在线率、Ip_guard管控通信" },
    { n: "阶段四", t: "部署实施", time: "1个窗口日", risk: "高", color: C.red,
      desc: "PBR配置上线→防火墙策略导入→VLAN切割→逐网段启用；同步360/Ip_guard策略" },
    { n: "阶段五", t: "监控优化", time: "2周", risk: "中", color: C.amber,
      desc: "PBR流量监控；策略日志审计；360 EDR调优；Ip_guard外发审批上线" },
    { n: "阶段六", t: "持续运营", time: "持续", risk: "低", color: C.green,
      desc: "EDR+SIEM联动；DLP全量加密；PBR策略定期审计；漏洞常态化" },
  ];

  stages.forEach((st, i) => {
    const sy = 1.2 + i * 0.7;
    // Timeline bar
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: sy + 0.15, w: 0.25, h: 0.25, fill: { color: st.color } });
    if (i < stages.length - 1) {
      s.addShape(pres.shapes.LINE, { x: 0.62, y: sy + 0.4, w: 0, h: 0.3, line: { color: C.muted, width: 1 } });
    }
    // Content
    s.addText(`${st.n}: ${st.t}`, { x: 1.0, y: sy, w: 2.5, h: 0.25, fontSize: 11, fontFace: "Arial", color: st.color, bold: true });
    s.addText(st.desc, { x: 1.0, y: sy + 0.3, w: 5.5, h: 0.25, fontSize: 8.5, fontFace: "Arial", color: C.text });
    // Time & risk tags
    tag(s, 7.3, sy + 0.05, st.time, C.cardBg);
    s.addText(st.time, { x: 7.3, y: sy + 0.05, w: 0.8, h: 0.25, fontSize: 8, fontFace: "Arial", color: C.text, align: "center", valign: "middle" });
    tag(s, 8.2, sy + 0.05, `风险:${st.risk}`, st.color);
    s.addText(`风险:${st.risk}`, { x: 8.2, y: sy + 0.05, w: 0.7, h: 0.25, fontSize: 8, fontFace: "Arial", color: C.white, align: "center", valign: "middle" });
  });

  // Key principles
  s.addText("核心原则", { x: 0.5, y: 5.3, w: 2, h: 0.2, fontSize: 7, fontFace: "Arial", color: C.muted });
  s.addText("业务窗口期变更 · PBR逐网段上线 · 初期防火墙Log模式 · PBR回退<5分钟 · 360在线率即时验证", {
    x: 2.5, y: 5.3, w: 7, h: 0.2, fontSize: 7, fontFace: "Arial", color: C.muted,
  });
}

// ── Slide 16: Risk Assessment ──
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.darkBg } });
  s.addText("风险评估与应对", { x: 0.6, y: 0.18, w: 5, h: 0.65, fontSize: 26, fontFace: "Arial", color: C.white, bold: true });

  const risks = [
    ["PBR下一跳不可达", "低", "高", "NQA track自动检测，故障时PBR失效Fallback"],
    ["PBR配置错误/流量黑洞", "中", "高", "逐网段上线；少量IP先验证；配置前备份"],
    ["防火墙性能不足", "中", "高", "实施前评估吞吐量；PBR排除备份大流量；逐网段监控"],
    ["策略遗漏致业务中断", "高", "高", "初期Log模式观察1-2周；逐条梳理访问关系表"],
    ["外部用户无法访问10楼服务器", "中", "中", "入站PBR策略+VLAN 120重点验证；初期宽松后收紧"],
    ["360客户端经PBR后断连", "中", "中", "第一时间验证在线率；保留云端直连备用通道"],
    ["EDR误报阻断开发工具", "高", "中", "初期仅告警不阻断；建立编译工具链白名单"],
    ["交换机TCAM表项不足", "低", "中", "实施前检查TCAM使用率；PBR ACL精简设计"],
  ];

  const rHeader = [
    { text: "风险", options: { bold: true, fill: { color: C.darkBg }, color: C.white, fontSize: 8 } },
    { text: "概率", options: { bold: true, fill: { color: C.darkBg }, color: C.white, fontSize: 8 } },
    { text: "影响", options: { bold: true, fill: { color: C.darkBg }, color: C.white, fontSize: 8 } },
    { text: "应对措施", options: { bold: true, fill: { color: C.darkBg }, color: C.white, fontSize: 8 } },
  ];
  const rRows = risks.map(r => [
    { text: r[0], options: { fontSize: 8, color: C.text } },
    { text: r[1], options: { fontSize: 8, color: r[1] === "高" ? C.red : r[1] === "中" ? C.amber : C.green, bold: true, align: "center" } },
    { text: r[2], options: { fontSize: 8, color: r[2] === "高" ? C.red : r[2] === "中" ? C.amber : C.green, bold: true, align: "center" } },
    { text: r[3], options: { fontSize: 7.5, color: C.muted } },
  ]);

  s.addTable([rHeader, ...rRows], {
    x: 0.5, y: 1.2, w: 9, colW: [2.3, 0.7, 0.7, 5.3],
    border: { pt: 0.3, color: "E2E8F0" },
    rowH: [0.3, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38],
  });
}

// ── Slide 17: Solution Highlights ──
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });
  s.addText("方案核心亮点", { x: 0.6, y: 0.3, w: 8, h: 0.6, fontSize: 28, fontFace: "Arial", color: C.white, bold: true });

  const highlights = [
    { n: "01", t: "PBR策略路由引流", d: "4期防火墙物理位置不动，3期华为汇聚PBR将10楼流量牵引至4期防火墙做策略检查", sub: "创新点：零物理施工，纯配置变更，回退<5分钟" },
    { n: "02", t: "分层纵深防护", d: "第一层ACL+PBR粗粒度隔离 → 第二层防火墙L3/L4精粒度控制+IPS/AV → 第三层2期出口防护", sub: "架构亮点：三层递进，PBR为管道串联各层" },
    { n: "03", t: "零硬件成本复用", d: "4期已启用但业务量小，山石防火墙性能利用率低；将空闲性能转化为10楼安全策略执行点", sub: "经济性：不新增任何硬件，仅配置变更" },
    { n: "04", t: "终端安全纵深协同", d: "360天擎EDR+Ip_guard DLP与网络层防护形成三层联防；同楼通信PBR排除，跨区通信经防火墙放行", sub: "协同：网络-终端-DLP全链路覆盖" },
    { n: "05", t: "智能故障切换", d: "PBR关联NQA track检测防火墙可达性，故障时自动Fail-Open保证业务连续性，同时触发告警", sub: "韧性：不因安全加固引入新的单点故障" },
  ];

  highlights.forEach((h, i) => {
    const hy = 1.2 + i * 0.85;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: hy, w: 9, h: 0.72, fill: { color: C.cardBg } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: hy, w: 0.06, h: 0.72, fill: { color: C.teal } });
    s.addText(h.n, { x: 0.75, y: hy + 0.05, w: 0.5, h: 0.62, fontSize: 22, fontFace: "Arial", color: C.tealAcc, bold: true, valign: "middle" });
    s.addText(h.t, { x: 1.4, y: hy + 0.05, w: 3.5, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.white, bold: true, valign: "bottom" });
    s.addText(h.d, { x: 1.4, y: hy + 0.38, w: 7.8, h: 0.3, fontSize: 8.5, fontFace: "Arial", color: C.muted, valign: "top" });
    // Sub-tag
    if (h.sub) {
      s.addText(h.sub, { x: 5.2, y: hy + 0.08, w: 4.1, h: 0.25, fontSize: 7, fontFace: "Arial", color: C.tealAcc });
    }
  });
}

// ── Slide 18: Next Steps ──
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal } });

  s.addText("下一步行动建议", { x: 0.6, y: 0.5, w: 8, h: 0.7, fontSize: 28, fontFace: "Arial", color: C.white, bold: true });

  const nextSteps = [
    { n: "1", t: "方案评审", d: "组织IT运维/研发组/安全团队评审技术路线" },
    { n: "2", t: "现场调研", d: "盘点IP/端口/访问关系/防火墙规格/交换机PBR能力" },
    { n: "3", t: "方案细化", d: "输出完整VLAN规划/PBR ACL/防火墙策略配置文档" },
    { n: "4", t: "测试验证", d: "非业务窗口搭建PBR试点，逐网段验证业务可达性" },
    { n: "5", t: "正式部署", d: "窗口期上线PBR+防火墙策略+360/Ip_guard优化" },
  ];

  nextSteps.forEach((ns, i) => {
    const ny = 1.6 + i * 0.7;
    s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: ny, w: 0.5, h: 0.5, fill: { color: C.teal } });
    s.addText(ns.n, { x: 1.5, y: ny, w: 0.5, h: 0.5, fontSize: 20, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle" });
    s.addText(ns.t, { x: 2.3, y: ny, w: 2, h: 0.5, fontSize: 15, fontFace: "Arial", color: C.white, bold: true, valign: "middle" });
    s.addText(ns.d, { x: 4.5, y: ny, w: 4.5, h: 0.5, fontSize: 10, fontFace: "Arial", color: C.muted, valign: "middle" });
    if (i < nextSteps.length - 1) {
      s.addShape(pres.shapes.LINE, { x: 1.75, y: ny + 0.5, w: 0, h: 0.2, line: { color: C.teal, width: 2 } });
    }
  });

  // Bottom info
  s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 5.2, w: 7, h: 0.01, fill: { color: C.teal } });
  s.addText("3期10楼仪器研发网络安全优化解决方案  |  V2.0  |  2026-06", {
    x: 1.5, y: 5.3, w: 7, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.muted, align: "center",
  });
}

// ── Output ──
pres.writeFile({ fileName: "/Users/linwei/Documents/My project/project02/pro_xinchanye/3期10楼网络安全优化解决方案_V2.0.pptx" })
  .then(() => console.log("PPT saved successfully!"))
  .catch(err => console.error("Error:", err));
