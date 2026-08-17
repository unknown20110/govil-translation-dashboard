/**
 * govil-extractor.js
 * -------------------
 * חילוץ תוכן מדפי gov.il לקובץ Word, לצורך העלאה לסוכן התרגום.
 * נטען דינמית ע"י bookmarklet קטן (ראו bookmarklet.txt).
 *
 * תומך בשני "משפחות" תבניות:
 *   1. Angular (guide + info)  -> div[id^="htmlContent_"], ואופציונלית טאבים ב-nav[aria-label*="Guide chapters"]
 *   2. Service (תבנית ישנה)    -> div.ServiceContainer
 *
 * מבוסס על ניתוח DOM בפועל של שלוש דוגמאות (08/2026):
 *   - /he/service/car_licence_renewal              (Service)
 *   - /he/pages/drivers-car-license-fee-boards      (Angular, עם טאבים)
 *   - /he/pages/topics_appointments_schedul_video   (Angular, בלי טאבים)
 *
 * הערה חשובה: קובץ ה-Word שנוצר משתמש בטכניקת altchunks (html-docx-js).
 * זה נפתח כמו שצריך ב-Microsoft Word, אבל לא ב-LibreOffice/Google Docs/Pages ישנים.
 * מספיק לצורך העלאה לסוכן התרגום, אבל שווה לדעת.
 */
(function () {
  'use strict';

  const HTML_DOCX_CDN = 'https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.js';

  // ---------- עזרים כלליים ----------

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function escapeHtml(str) {
    return (str || '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }

  function sanitizeFilename(name) {
    return (name || 'govil-content').replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 80) || 'govil-content';
  }

  function loadScriptOnce(src) {
    return new Promise((resolve, reject) => {
      if (window.htmlDocx) return resolve();
      const existing = document.querySelector(`script[src="${src}"]`);
      if (existing) {
        existing.addEventListener('load', () => resolve());
        existing.addEventListener('error', reject);
        return;
      }
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('נכשלה טעינת ספריית html-docx-js מה-CDN'));
      document.head.appendChild(s);
    });
  }

  // ---------- Toast UI פשוט (מצב התקדמות) ----------

  function toast(msg) {
    let el = document.getElementById('__govil_extract_toast');
    if (!el) {
      el = document.createElement('div');
      el.id = '__govil_extract_toast';
      el.style.cssText =
        'position:fixed;bottom:20px;left:20px;z-index:2147483647;background:#1a73e8;color:#fff;' +
        'padding:10px 16px;border-radius:8px;font-family:Arial,sans-serif;font-size:14px;' +
        'box-shadow:0 2px 10px rgba(0,0,0,.35);direction:rtl;max-width:320px;';
      document.body.appendChild(el);
    }
    el.textContent = msg;
    clearTimeout(el._hideTimer);
    el._hideTimer = setTimeout(() => el.remove(), 4000);
    return el;
  }

  // ---------- זיהוי סוג עמוד ----------

  function detectPageType() {
    if (document.querySelector('div.ServiceContainer')) return 'service';
    if (document.querySelector('div[id^="htmlContent_"]')) return 'angular';
    return 'unknown';
  }

  // ---------- ניקוי HTML שחולץ (הסרת כפתורים/אייקונים/תמונות לא רלוונטיים לתרגום) ----------

  function cleanHtml(html) {
    const wrapper = document.createElement('div');
    wrapper.innerHTML = html || '';

    // הסרת אלמנטים לא רלוונטיים לתרגום, כולל widget השיתוף (שתפו:) של gov.il
    // הערה: לא מסירים h2[name="ServiceDescription"] כאן לפי selector עיוור —
    // בחלק מהדפים (למשל car_licence_renewal) הוא ריק לגמרי (זבל תבנית),
    // אבל בדפים אחרים (למשל digital-tachograph-card) ה-CMS שם בו את תקציר
    // השירות בפועל. ההסרה הגנרית של כותרות ריקות למטה מספיקה ובטוחה יותר —
    // מסירה אותו רק כשהוא באמת ריק, בלי לאבד תוכן אמיתי (נתפס ותוקן 9.8.2026).
    wrapper
      .querySelectorAll(
        'script, style, button, svg, img, iframe, nav, [aria-hidden="true"], .hidden-print, [name="AllSocialURLs"]'
      )
      .forEach((el) => el.remove());

    // שאריות ידועות של widget השיתוף שלא נתפסות ע"י ההסרות הכלליות (label בודד "שתפו:")
    wrapper.querySelectorAll('div, span, p').forEach((el) => {
      const t = el.textContent.trim();
      if (t === 'שתפו:' || t === 'שתפו') el.remove();
    });

    // הסרת קישורים ריקים (icon-only, למשל שאריות של כפתורי שיתוף) + ה-li/ul/ol שנשארים ריקים אחריהם
    let changed = true;
    while (changed) {
      changed = false;
      wrapper.querySelectorAll('a').forEach((a) => {
        if (!a.textContent.trim()) {
          a.remove();
          changed = true;
        }
      });
      wrapper.querySelectorAll('li, ul, ol, div, span').forEach((el) => {
        if (!el.textContent.trim() && el.children.length === 0) {
          el.remove();
          changed = true;
        }
      });
    }

    // הסרת כותרות ריקות (קיימות לפעמים כבר בדף המקור, למשל h2[name="ServiceDescription"])
    wrapper.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach((h) => {
      if (!h.textContent.trim()) h.remove();
    });

    // הסרת attributes שגורמים לרעש (class/style עם hash-ים של Angular, onclick וכו')
    wrapper.querySelectorAll('*').forEach((el) => {
      el.removeAttribute('class');
      el.removeAttribute('style');
      el.removeAttribute('id');
      el.removeAttribute('onclick');
      el.removeAttribute('target');
    });

    return wrapper.innerHTML.trim();
  }

  // בעמודי Service, הכותרת (h1) כבר נמצאת בתוך אזור התוכן עצמו (בשונה מעמודי Angular,
  // שם היא חיה מחוץ ל-htmlContent). בודקים את זה כדי לא להזריק כותרת כפולה.
  function contentAlreadyStartsWithTitle(rawHtml, title) {
    const wrapper = document.createElement('div');
    wrapper.innerHTML = rawHtml || '';
    const firstHeading = wrapper.querySelector('h1, h2, h3');
    if (!firstHeading) return false;
    const normalize = (s) => (s || '').replace(/\s+/g, ' ').trim();
    return normalize(firstHeading.textContent) === normalize(title);
  }

  // [17.8.2026] פותחת אקורדיונים מקופלים בתוך container (בלולאה, לכיסוי
  // אקורדיונים מקוננים). נשמר כרשת ביטחון כללית לעמודים שבהם תוכן מקופל
  // באמת lazy — אבל ראו ההערה למטה: זה *לא* היה שורש הבאג ב-subsidy-faq.
  async function expandAllCollapsedSections(container) {
    if (!container) return 0;
    let totalOpened = 0;
    for (let round = 0; round < 5; round++) {
      const collapsed = Array.from(container.querySelectorAll('[aria-expanded="false"]'));
      if (collapsed.length === 0) break;
      collapsed.forEach((el) => el.click());
      totalOpened += collapsed.length;
      await sleep(200); // מרווח לרינדור התוכן שנחשף
    }
    return totalOpened;
  }

  // [17.8.2026, גרסה 2 — אחרי בדיקת קובץ MHTML אמיתי של הדף החי] האבחון
  // המדויק: בלוק "שאלות ותשובות" (id="faqs_N") הוא *לא* צאצא של
  // htmlContent_N בכלל — הוא אח שלו, div נפרד תחת אותו הורה (לצד
  // h3#content_title). לכן:
  //   1. חיפוש [aria-expanded="false"] בתוך htmlContent עצמו (גרסה 1 של
  //      התיקון) תמיד מצא אפס תוצאות — לא הייתה טעות במנגנון עצמו, רק
  //      בקונטיינר שחיפשנו בו.
  //   2. תוכן התשובות כבר קיים ב-DOM גם כשהפריט מקופל (רק גובה מוסתר
  //      ב-CSS, לא lazy-render אמיתי) — אומת על כמה שאלות שונות. לא היינו
  //      חייבים בכלל ללחוץ, רק לאסוף את הבלוק הנכון.
  //   3. כותרת כל שאלה עטופה ב-<button aria-expanded=...>, ו-cleanHtml
  //      הכללית (למטה) כבר מוחקת כל תגית button — בלי טיפול נפרד זה מוחק
  //      גם את טקסט השאלה עצמו, לא רק את מנגנון הקיפול.
  function findAdjacentFaqBlock(contentEl) {
    if (!contentEl || !contentEl.parentElement) return null;
    return contentEl.parentElement.querySelector('[id^="faqs_"]');
  }

  function unwrapFaqQuestionButtons(container) {
    if (!container) return;
    container.querySelectorAll('button[aria-expanded]').forEach((btn) => {
      while (btn.firstChild) btn.parentNode.insertBefore(btn.firstChild, btn);
      btn.remove();
    });
  }

  // ---------- חילוץ: תבנית Angular (guide + info) ----------

  async function waitForContentSwap(prevHtml, timeoutMs = 4000, intervalMs = 100) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const el = document.querySelector('div[id^="htmlContent_"]');
      if (el && el.innerHTML !== prevHtml) return el;
      await sleep(intervalMs);
    }
    return document.querySelector('div[id^="htmlContent_"]'); // best effort
  }

  async function extractAngular() {
    const sections = [];
    const tabsNav = document.querySelector(
      'nav[aria-label*="Guide chapters"], nav[aria-label*="פרקי המדריך"]'
    );
    const tabButtons = tabsNav ? Array.from(tabsNav.querySelectorAll('button[id^="sideNav_"]')) : [];

    if (tabButtons.length === 0) {
      // עמוד מידע ללא טאבים - פרק יחיד
      const contentEl = document.querySelector('div[id^="htmlContent_"]');
      const titleEl =
        document.getElementById('content_title') ||
        document.getElementById('contentPageHeadTitle') ||
        document.querySelector('h1');
      await expandAllCollapsedSections(contentEl);
      const faqEl = findAdjacentFaqBlock(contentEl); // [17.8.2026]
      await expandAllCollapsedSections(faqEl);
      unwrapFaqQuestionButtons(faqEl);
      sections.push({
        title: titleEl ? titleEl.textContent.trim() : document.title,
        html: (contentEl ? contentEl.innerHTML : '') + (faqEl ? faqEl.innerHTML : ''),
      });
      return sections;
    }

    // עמוד מדריך עם כמה טאבים - לעבור על כולם
    for (let i = 0; i < tabButtons.length; i++) {
      const btn = tabButtons[i];
      const label = (btn.querySelector('span') || btn).textContent.trim();
      toast(`סורק לשונית ${i + 1} מתוך ${tabButtons.length}: ${label}`);

      const before = document.querySelector('div[id^="htmlContent_"]');
      const prevHtml = before ? before.innerHTML : null;

      btn.click();
      const contentEl = await waitForContentSwap(prevHtml);
      await sleep(150); // מרווח ביטחון קטן לרינדור מלא
      await expandAllCollapsedSections(contentEl);
      const faqEl = findAdjacentFaqBlock(contentEl); // [17.8.2026] — כאן נתפס הבאג בפועל
      await expandAllCollapsedSections(faqEl);
      unwrapFaqQuestionButtons(faqEl);

      const titleEl = document.getElementById('content_title');
      sections.push({
        title: (titleEl ? titleEl.textContent.trim() : '') || label,
        html: (contentEl ? contentEl.innerHTML : '') + (faqEl ? faqEl.innerHTML : ''),
      });
    }
    return sections;
  }

  // ---------- חילוץ: תבנית Service (ישנה) ----------

  async function extractService() {
    const contentEl = document.querySelector('div.ServiceContainer');
    const titleEl = document.querySelector('.PageTitle') || document.querySelector('h1');
    await expandAllCollapsedSections(contentEl);
    const faqEl = findAdjacentFaqBlock(contentEl); // [17.8.2026] הגנה זהה, לא נבדק על תבנית Service בפועל
    await expandAllCollapsedSections(faqEl);
    unwrapFaqQuestionButtons(faqEl);
    return [
      {
        title: titleEl ? titleEl.textContent.trim() : document.title,
        html: (contentEl ? contentEl.innerHTML : '') + (faqEl ? faqEl.innerHTML : ''),
      },
    ];
  }

  // ---------- בניית ה-Word והורדה ----------

  function buildAndDownloadDocx(pageTitle, sections) {
    const bodyParts = sections.map((sec, idx) => {
      const pageBreak = idx > 0 ? '<br clear="all" style="page-break-before:always">' : '';
      const skipOwnTitle = contentAlreadyStartsWithTitle(sec.html, sec.title);
      const titleHtml = skipOwnTitle ? '' : `<h2>${escapeHtml(sec.title)}</h2>\n`;
      return `${pageBreak}${titleHtml}${cleanHtml(sec.html)}`;
    });

    const fullHtml = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>${escapeHtml(pageTitle)}</title></head>
<body dir="rtl" style="direction:rtl; text-align:right; font-family:Arial, sans-serif;">
${bodyParts.join('\n')}
</body>
</html>`;

    const blob = window.htmlDocx.asBlob(fullHtml);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${sanitizeFilename(pageTitle)}.docx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  // ---------- ראשי ----------

  async function main() {
    try {
      toast('טוען כלי חילוץ...');
      await loadScriptOnce(HTML_DOCX_CDN);

      const pageType = detectPageType();
      let sections = [];

      if (pageType === 'service') {
        sections = await extractService();
      } else if (pageType === 'angular') {
        sections = await extractAngular();
      } else {
        alert('לא זוהה מבנה עמוד מוכר (לא Service ולא Angular guide/info). אולי זה סוג עמוד חדש שצריך למפות.');
        return;
      }

      const emptyCount = sections.filter((s) => !s.html || !s.html.trim()).length;
      if (emptyCount > 0) {
        console.warn(`govil-extractor: ${emptyCount} מתוך ${sections.length} פרקים חזרו ריקים`);
      }

      const pageTitle = (document.querySelector('h1') || {}).textContent?.trim() || document.title;
      buildAndDownloadDocx(pageTitle, sections);
      toast(`הקובץ הורד בהצלחה ✓ (${sections.length} פרקים)`);
    } catch (err) {
      console.error('govil-extractor error:', err);
      alert('שגיאה בחילוץ התוכן: ' + err.message);
    }
  }

  main();
})();
