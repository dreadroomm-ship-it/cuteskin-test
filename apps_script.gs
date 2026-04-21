// ============================================================
//  CuteSkin Test Handler — Google Apps Script
//  Принимает POST с результатами теста, шлёт письмо
//  с красивым HTML-отчётом + PDF-вложением на RECIPIENT.
//
//  После правки: Deploy → Manage deployments → Edit →
//  Version: "New version" → Deploy. Иначе /exec вернёт старую версию.
// ============================================================

const RECIPIENT = 'dreadroomm@gmail.com';

const BRAND = {
  pink:        '#d9466f',
  pinkDeep:    '#b02551',
  pinkMid:     '#ec7aa0',
  pinkSoft:    '#fcd4e0',
  pinkMist:    '#fef0f5',
  ok:          '#2e7d32',
  okSoft:      '#e8f5e9',
  warn:        '#ef6c00',
  warnSoft:    '#fff3e0',
  bad:         '#c62828',
  badSoft:     '#ffebee',
  ink:         '#1a1a1a',
  inkSoft:     '#4a4a4a',
  line:        '#f0dbe4',
  paper:       '#ffffff',
  canvas:      '#fff7fa'
};

// ------------ Entry points ------------

function doGet() {
  return out({ ok: true, service: 'CuteSkin Test Handler' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const pdf = buildPdf(data);
    const html = buildHtmlEmail(data);
    const subject = formatSubject(data);
    GmailApp.sendEmail(RECIPIENT, subject, plainFallback(data), {
      htmlBody: html,
      attachments: [pdf],
      name: 'CuteSkin Test'
    });
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err && err.stack || err) });
  }
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------ Helpers ------------

function formatSubject(d) {
  const name = (d.name || 'Без имени').trim();
  return 'CuteSkin Test — ' + name + ' — ' + d.score + ' (' + d.percent + '%)';
}

function verdict(percent) {
  if (percent >= 80) return { label: 'Отлично',   color: BRAND.ok,   bg: BRAND.okSoft,   emoji: '✨' };
  if (percent >= 50) return { label: 'Средне',    color: BRAND.warn, bg: BRAND.warnSoft, emoji: '⚖' };
  return                      { label: 'Слабо',     color: BRAND.bad,  bg: BRAND.badSoft,  emoji: '⚠' };
}

function formatDate(iso) {
  try {
    const d = new Date(iso);
    const pad = n => (n < 10 ? '0' + n : '' + n);
    return pad(d.getDate()) + '.' + pad(d.getMonth() + 1) + '.' + d.getFullYear() +
           ' · ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  } catch (e) { return iso; }
}

function formatDuration(totalSec) {
  totalSec = Math.max(0, Math.round(Number(totalSec) || 0));
  const m = Math.floor(totalSec / 60);
  const s = totalSec % 60;
  if (m === 0) return s + 'с';
  return m + 'м ' + (s < 10 ? '0' + s : s) + 'с';
}

function integrityFlags(a) {
  const bits = [];
  if (typeof a.time_sec === 'number') {
    if (a.time_sec <= 3 && a.type !== 'free') bits.push('быстрый ответ');
    if (a.time_sec <= 8 && a.type === 'free') bits.push('быстрый ввод');
  }
  if (a.paste_attempts > 0) bits.push('вставка×' + a.paste_attempts);
  if (a.copy_attempts > 0) bits.push('копирование×' + a.copy_attempts);
  return bits;
}

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function plainFallback(d) {
  const v = verdict(d.percent);
  return 'CuteSkin Test\n\n' +
    'Кандидат: ' + (d.name || '—') + '\n' +
    'Дата: ' + formatDate(d.date) + '\n' +
    'Результат: ' + d.score + ' (' + d.percent + '%) — ' + v.label + '\n\n' +
    'Полный отчёт — в приложенном PDF и в HTML-версии письма.';
}

// ============================================================
//   HTML EMAIL
// ============================================================

function buildHtmlEmail(d) {
  const v = verdict(d.percent);
  const sections = (d.sections || []).map(sectionRowHtml).join('');
  const answers = (d.answers || []).map(answerCardHtml).join('');
  const shots = Number(d.screenshot_attempts || 0);
  const alertLines = [];
  if (shots > 0) alertLines.push('пытался сделать скриншот ' + shots + ' раз' + (shots === 1 ? '' : (shots < 5 ? 'а' : '')));
  const alertBanner = alertLines.length ? `
    <tr><td style="background:${BRAND.bad};color:#fff;padding:14px 32px;font-size:14px;font-weight:700;letter-spacing:0.5px;line-height:1.5;">
      ⚠ Кандидат ${alertLines.join(' · ')} во время прохождения теста
    </td></tr>` : '';

  return `<!doctype html>
<html lang="ru">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:${BRAND.pinkMist};font-family:-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:${BRAND.ink};">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:${BRAND.pinkMist};">
    <tr><td align="center" style="padding:28px 12px;">

      <!-- Hero -->
      <table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0" style="max-width:640px;width:100%;border-collapse:separate;">
        <tr><td style="background:linear-gradient(135deg,${BRAND.pinkDeep} 0%,${BRAND.pink} 60%,${BRAND.pinkMid} 100%);border-radius:20px 20px 0 0;padding:28px 32px;color:#fff;">
          <div style="font-size:13px;letter-spacing:3px;text-transform:uppercase;opacity:0.85;">CuteSkin · Project Manager Test</div>
          <div style="font-size:28px;font-weight:800;margin-top:6px;line-height:1.1;">Отчёт о прохождении</div>
          <div style="font-size:14px;margin-top:10px;opacity:0.9;">${esc(formatDate(d.date))}</div>
          <div style="font-size:14px;margin-top:10px;line-height:1.9;">
            <span style="display:inline-block;background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:999px;margin-right:6px;">⏱ ${esc(formatDuration(d.duration_sec))}</span>
            <span style="display:inline-block;background:${shots > 0 ? '#fff' : 'rgba(255,255,255,0.2)'};color:${shots > 0 ? BRAND.bad : '#fff'};padding:4px 12px;border-radius:999px;font-weight:${shots > 0 ? 800 : 500};">
              ${shots > 0 ? '⚠ Скриншотов: ' + shots : '✓ Без скриншотов'}
            </span>
          </div>
        </td></tr>
        ${alertBanner}

        <!-- Body -->
        <tr><td style="background:${BRAND.paper};padding:28px 32px;border-left:1px solid ${BRAND.line};border-right:1px solid ${BRAND.line};">

          <!-- Candidate + Score side-by-side -->
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td valign="top" width="58%" style="padding-right:12px;">
                <div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;color:${BRAND.inkSoft};">Кандидат</div>
                <div style="font-size:22px;font-weight:800;margin-top:6px;color:${BRAND.ink};line-height:1.2;">${esc(d.name || '—')}</div>
              </td>
              <td valign="top" width="42%" align="right">
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="background:${v.bg};border-radius:16px;border:1px solid ${v.color}22;">
                  <tr><td style="padding:14px 20px;text-align:center;">
                    <div style="font-size:11px;letter-spacing:2px;text-transform:uppercase;color:${v.color};font-weight:700;">${v.emoji} ${esc(v.label)}</div>
                    <div style="font-size:30px;font-weight:800;color:${v.color};line-height:1;margin-top:6px;">${esc(d.percent)}%</div>
                    <div style="font-size:13px;color:${BRAND.inkSoft};margin-top:4px;">${esc(d.score)}</div>
                  </td></tr>
                </table>
              </td>
            </tr>
          </table>

          <!-- Progress bar -->
          <div style="margin:22px 0 6px;">
            <div style="height:12px;background:${BRAND.pinkSoft};border-radius:999px;overflow:hidden;">
              <div style="width:${Math.max(2, Math.min(100, d.percent))}%;height:12px;background:linear-gradient(90deg,${BRAND.pinkMid},${BRAND.pinkDeep});border-radius:999px;"></div>
            </div>
          </div>

          <!-- Sections -->
          <div style="margin-top:28px;font-size:11px;letter-spacing:2px;text-transform:uppercase;color:${BRAND.inkSoft};font-weight:700;">Разбивка по разделам</div>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;border-collapse:separate;border-spacing:0 6px;">
            ${sections}
          </table>

          <!-- Answers -->
          <div style="margin-top:30px;font-size:11px;letter-spacing:2px;text-transform:uppercase;color:${BRAND.inkSoft};font-weight:700;">Ответы кандидата</div>
          <div style="margin-top:12px;">
            ${answers}
          </div>

        </td></tr>

        <!-- Footer -->
        <tr><td style="background:${BRAND.paper};border-radius:0 0 20px 20px;padding:20px 32px 26px;border-left:1px solid ${BRAND.line};border-right:1px solid ${BRAND.line};border-bottom:1px solid ${BRAND.line};">
          <div style="font-size:12px;color:${BRAND.inkSoft};line-height:1.5;">
            Полная версия с дословными формулировками вариантов ответа — в PDF-приложении.
            <br>Открытые вопросы автоматически не оцениваются, проверяются вручную.
          </div>
        </td></tr>
      </table>

    </td></tr>
  </table>
</body>
</html>`;
}

function sectionRowHtml(s) {
  const pct = s.total ? Math.round((s.correct / s.total) * 100) : 0;
  const barColor = s.free ? BRAND.pinkMid : (pct >= 80 ? BRAND.ok : pct >= 50 ? BRAND.warn : BRAND.bad);
  const scoreText = s.free ? 'свободный ответ' : (s.correct + ' / ' + s.total);
  const pctText = s.free ? '—' : (pct + '%');
  const barWidth = s.free ? 100 : Math.max(4, pct);
  return `
    <tr>
      <td style="background:${BRAND.canvas};border-radius:12px;padding:12px 16px;border:1px solid ${BRAND.line};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td valign="middle" style="font-size:14px;font-weight:700;color:${BRAND.ink};">${esc(s.name)}</td>
            <td valign="middle" align="right" style="font-size:13px;color:${BRAND.inkSoft};white-space:nowrap;">
              <span style="color:${BRAND.ink};font-weight:700;">${esc(pctText)}</span>
              &nbsp;·&nbsp;${esc(scoreText)}
            </td>
          </tr>
          <tr><td colspan="2" style="padding-top:8px;">
            <div style="height:6px;background:${BRAND.pinkSoft};border-radius:999px;overflow:hidden;">
              <div style="width:${barWidth}%;height:6px;background:${barColor};border-radius:999px;"></div>
            </div>
          </td></tr>
        </table>
      </td>
    </tr>`;
}

function answerCardHtml(a) {
  const meta = answerMetaHtml(a);
  if (a.type === 'free') {
    const text = (a.free_text || '').trim();
    const hasText = !!text;
    return `
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:10px;">
        <tr><td style="background:${BRAND.pinkMist};border-radius:14px;padding:14px 16px;border:1px solid ${BRAND.line};border-left:4px solid ${BRAND.pink};">
          <div style="font-size:11px;color:${BRAND.pinkDeep};font-weight:700;letter-spacing:1.5px;text-transform:uppercase;">#${a.num} · ${esc(a.section)} · свободный ответ</div>
          <div style="font-size:15px;font-weight:700;color:${BRAND.ink};margin-top:6px;line-height:1.35;">${esc(a.q)}</div>
          <div style="font-size:14px;color:${hasText ? BRAND.ink : BRAND.inkSoft};margin-top:10px;white-space:pre-wrap;line-height:1.5;${hasText ? '' : 'font-style:italic;'}">
            ${hasText ? esc(text) : 'Ответ не введён'}
          </div>
          ${meta}
        </td></tr>
      </table>`;
  }

  const answered = !!a.answered;
  const correct = !!a.is_correct;
  const color = !answered ? BRAND.inkSoft : (correct ? BRAND.ok : BRAND.bad);
  const bg    = !answered ? BRAND.canvas : (correct ? BRAND.okSoft : BRAND.badSoft);
  const label = !answered ? 'не отвечено' : (correct ? '✓ верно' : '✕ неверно');
  const typeLabel = a.type === 'multi' ? 'мульти' : 'одиночный';

  return `
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:10px;">
      <tr><td style="background:${BRAND.paper};border-radius:14px;padding:14px 16px;border:1px solid ${BRAND.line};border-left:4px solid ${color};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td style="font-size:11px;color:${BRAND.inkSoft};font-weight:700;letter-spacing:1.5px;text-transform:uppercase;">#${a.num} · ${esc(a.section)} · ${esc(typeLabel)}</td>
            <td align="right" style="font-size:11px;font-weight:800;color:${color};letter-spacing:1px;text-transform:uppercase;white-space:nowrap;">${esc(label)}</td>
          </tr>
        </table>
        <div style="font-size:15px;font-weight:700;color:${BRAND.ink};margin-top:8px;line-height:1.35;">${esc(a.q)}</div>
        <div style="margin-top:10px;padding:10px 12px;background:${bg};border-radius:10px;">
          <div style="font-size:11px;color:${BRAND.inkSoft};font-weight:700;letter-spacing:1px;text-transform:uppercase;">Ответ кандидата</div>
          <div style="font-size:14px;color:${color};font-weight:600;margin-top:4px;line-height:1.4;">${esc(a.picked_text || '—')}</div>
        </div>
        ${correct ? '' : `
          <div style="margin-top:6px;padding:10px 12px;background:${BRAND.okSoft};border-radius:10px;">
            <div style="font-size:11px;color:${BRAND.inkSoft};font-weight:700;letter-spacing:1px;text-transform:uppercase;">Правильный ответ</div>
            <div style="font-size:14px;color:${BRAND.ok};font-weight:600;margin-top:4px;line-height:1.4;">${esc(a.correct_text || '—')}</div>
          </div>`}
        ${meta}
      </td></tr>
    </table>`;
}

function answerMetaHtml(a) {
  const flags = integrityFlags(a);
  const timeTxt = 'Время: ' + formatDuration(a.time_sec || 0);
  const flagTxt = flags.length
    ? ' · <span style="color:' + BRAND.bad + ';font-weight:700;">⚑ ' + esc(flags.join(' · ')) + '</span>'
    : '';
  return '<div style="margin-top:10px;font-size:11px;color:' + BRAND.inkSoft + ';letter-spacing:0.5px;">' +
         esc(timeTxt) + flagTxt + '</div>';
}

// ============================================================
//   PDF (DocumentApp)
// ============================================================

function buildPdf(data) {
  const safeName = (data.name || 'Без имени').replace(/[\\/:*?"<>|]/g, ' ').trim();
  const doc = DocumentApp.create('CuteSkin Test — ' + safeName);
  const body = doc.getBody();
  body.setMarginTop(50).setMarginBottom(50).setMarginLeft(60).setMarginRight(60);

  const v = verdict(data.percent);

  // ---- Hero banner: table initialized with first-line content (avoids setText) ----
  const hero = body.appendTable([['CUTESKIN  ·  PROJECT MANAGER TEST']]);
  // Now it's safe to drop the initial empty paragraph (hero table is the new last child)
  const first = body.getChild(0);
  if (first && first.getType() === DocumentApp.ElementType.PARAGRAPH && first.asParagraph().getText() === '') {
    first.asParagraph().removeFromParent();
  }
  styleBand(hero, BRAND.pink);
  const heroCell = hero.getCell(0, 0);
  heroCell.setPaddingTop(18).setPaddingBottom(18).setPaddingLeft(20).setPaddingRight(20);
  styleFirstPara(heroCell, { color: '#ffe4ef', size: 9, bold: true });
  appendStyledPara(heroCell, 'Отчёт о прохождении', { color: '#ffffff', size: 22, bold: true });
  appendStyledPara(heroCell, formatDate(data.date), { color: '#ffe4ef', size: 10 });
  const pdfShots = Number(data.screenshot_attempts || 0);
  appendStyledPara(heroCell,
    'Длительность: ' + formatDuration(data.duration_sec),
    { color: '#ffe4ef', size: 10 });
  appendStyledPara(heroCell,
    (pdfShots > 0 ? '⚠ ПОПЫТОК СКРИНШОТА: ' + pdfShots : '✓ Без попыток скриншота'),
    { color: pdfShots > 0 ? '#ffffff' : '#ffe4ef', size: 10, bold: pdfShots > 0 });

  // ---- Candidate + Score row: init with first-line content in each cell ----
  body.appendParagraph(' ').editAsText().setFontSize(6);
  const info = body.appendTable([[
    'КАНДИДАТ',
    (v.emoji + ' ' + v.label).toUpperCase()
  ]]);
  info.setBorderWidth(0);
  const leftCell = info.getCell(0, 0);
  leftCell.setBackgroundColor(BRAND.canvas).setPaddingTop(14).setPaddingBottom(14).setPaddingLeft(16).setPaddingRight(16);
  styleFirstPara(leftCell, { color: BRAND.inkSoft, size: 9, bold: true });
  appendStyledPara(leftCell, data.name || '—', { color: BRAND.ink, size: 18, bold: true });

  const rightCell = info.getCell(0, 1);
  rightCell.setBackgroundColor(v.bg).setPaddingTop(14).setPaddingBottom(14).setPaddingLeft(16).setPaddingRight(16);
  styleFirstPara(rightCell, { color: v.color, size: 9, bold: true, align: DocumentApp.HorizontalAlignment.RIGHT });
  appendStyledPara(rightCell, data.percent + '%', { color: v.color, size: 26, bold: true, align: DocumentApp.HorizontalAlignment.RIGHT });
  appendStyledPara(rightCell, data.score, { color: BRAND.inkSoft, size: 11, align: DocumentApp.HorizontalAlignment.RIGHT });

  // ---- Sections header ----
  body.appendParagraph(' ').editAsText().setFontSize(6);
  appendKicker(body, 'РАЗБИВКА ПО РАЗДЕЛАМ');

  // Sections table: name | score | pct
  const secs = data.sections || [];
  const secRows = [['Раздел','Баллы','%']].concat(secs.map(s => {
    const pct = s.total ? Math.round((s.correct / s.total) * 100) : 0;
    return [
      s.name,
      s.free ? 'свободный ответ' : (s.correct + ' / ' + s.total),
      s.free ? '—' : (pct + '%')
    ];
  }));
  const secTable = body.appendTable(secRows);
  styleDataTable(secTable);
  // Color pct cells
  for (let i = 1; i < secRows.length; i++) {
    const s = secs[i - 1];
    const pct = s.total ? Math.round((s.correct / s.total) * 100) : 0;
    const col = s.free ? BRAND.pinkDeep : (pct >= 80 ? BRAND.ok : pct >= 50 ? BRAND.warn : BRAND.bad);
    const cell = secTable.getCell(i, 2);
    cell.editAsText().setForegroundColor(col).setBold(true);
  }

  // ---- Answers ----
  body.appendParagraph(' ').editAsText().setFontSize(6);
  appendKicker(body, 'ОТВЕТЫ КАНДИДАТА');

  (data.answers || []).forEach(a => {
    body.appendParagraph(' ').editAsText().setFontSize(4);
    appendAnswerCard(body, a);
  });

  // Save and convert
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  const pdf = file.getAs('application/pdf').copyBlob();
  pdf.setName('CuteSkin-Test-' + safeName + '.pdf');
  file.setTrashed(true);
  return pdf;
}

function appendKicker(body, text) {
  const p = body.appendParagraph(text);
  p.editAsText().setForegroundColor(BRAND.pinkDeep).setFontSize(10).setBold(true);
  return p;
}

function styleBand(table, color) {
  table.setBorderWidth(0);
  table.getCell(0, 0).setBackgroundColor(color);
}

function styleDataTable(table) {
  const rows = table.getNumRows();
  table.setBorderWidth(0);
  // Header
  for (let c = 0; c < table.getRow(0).getNumCells(); c++) {
    const cell = table.getCell(0, c);
    cell.setBackgroundColor(BRAND.pinkSoft).setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
    cell.editAsText().setForegroundColor(BRAND.pinkDeep).setFontSize(9).setBold(true);
  }
  // Body rows
  for (let r = 1; r < rows; r++) {
    for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
      const cell = table.getCell(r, c);
      cell.setBackgroundColor(r % 2 ? BRAND.canvas : BRAND.paper)
          .setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
      cell.editAsText().setForegroundColor(BRAND.ink).setFontSize(10);
    }
  }
}

function appendAnswerCard(body, a) {
  const isFree = a.type === 'free';
  const answered = !!a.answered;
  const correct = !!a.is_correct;

  let accent, bg;
  if (isFree) { accent = BRAND.pink; bg = BRAND.pinkMist; }
  else if (!answered) { accent = BRAND.inkSoft; bg = BRAND.canvas; }
  else if (correct) { accent = BRAND.ok; bg = BRAND.okSoft; }
  else { accent = BRAND.bad; bg = BRAND.badSoft; }

  const typeLabel = isFree ? 'свободный ответ' : (a.type === 'multi' ? 'мульти' : 'одиночный');
  const statusLabel = isFree
    ? ((a.free_text || '').trim() ? 'ОТПРАВЛЕНО' : 'ПУСТО')
    : (!answered ? 'НЕ ОТВЕЧЕНО' : (correct ? '✓ ВЕРНО' : '✕ НЕВЕРНО'));

  const kickerText = '#' + a.num + '  ·  ' + (a.section || '') + '  ·  ' + typeLabel + '   —   ' + statusLabel;

  // Two-column table: thin accent stripe ('·' invisible at size 1) + content (initialized with kicker text)
  const card = body.appendTable([['\u00A0', kickerText]]);
  card.setBorderWidth(0);
  const stripe = card.getCell(0, 0);
  stripe.setBackgroundColor(accent).setWidth(4).setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  stripe.getChild(0).asParagraph().editAsText().setFontSize(1);
  const content = card.getCell(0, 1);
  content.setBackgroundColor(bg).setPaddingTop(12).setPaddingBottom(12).setPaddingLeft(14).setPaddingRight(14);
  styleFirstPara(content, { color: accent, size: 9, bold: true });
  appendStyledPara(content, a.q || '—', { color: BRAND.ink, size: 12, bold: true });

  if (isFree) {
    const text = (a.free_text || '').trim();
    appendStyledPara(content, text || 'Ответ не введён', {
      color: text ? BRAND.ink : BRAND.inkSoft,
      size: 11,
      italic: !text
    });
  } else {
    appendStyledPara(content, 'Ответ кандидата: ' + (a.picked_text || '—'), { color: accent, size: 11 });
    if (!correct) {
      appendStyledPara(content, 'Правильный ответ: ' + (a.correct_text || '—'), { color: BRAND.ok, size: 11 });
    }
  }

  const flags = integrityFlags(a);
  const metaLine = 'Время: ' + formatDuration(a.time_sec || 0) +
                   (flags.length ? '   ⚑ ' + flags.join(' · ') : '');
  appendStyledPara(content, metaLine, {
    color: flags.length ? BRAND.bad : BRAND.inkSoft,
    size: 9,
    bold: flags.length > 0
  });
}

// ------------ Paragraph helpers ------------

function applyTextStyle(paragraph, opts) {
  opts = opts || {};
  const t = paragraph.editAsText();
  if (opts.color) t.setForegroundColor(opts.color);
  if (opts.size) t.setFontSize(opts.size);
  if (opts.bold != null) t.setBold(!!opts.bold);
  if (opts.italic != null) t.setItalic(!!opts.italic);
  if (opts.align) paragraph.setAlignment(opts.align);
  return paragraph;
}

// Style the first paragraph of a table cell (cell must already have text in it)
function styleFirstPara(cell, opts) {
  return applyTextStyle(cell.getChild(0).asParagraph(), opts);
}

// Append a new paragraph to a cell and style it. Guarantees non-empty text.
function appendStyledPara(cell, text, opts) {
  const safe = (text == null || text === '') ? ' ' : text;
  return applyTextStyle(cell.appendParagraph(safe), opts);
}
