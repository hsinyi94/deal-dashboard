(() => {
  const LOCALE_KEY = 'deal_dashboard_locale';
  const JA = 'ja';
  const ZH = 'zh-Hant';
  const translatableFields = ['name', 'sub', 'typeLabel', 'deadlineText', 'eventText', 'linkLabel', 'earlyBird'];

  const jaDeals = {
    pd4: { name: 'ポイントフェスティバル #4（春半ばセール）', sub: 'Qualtrics で申請', typeLabel: 'ポイント', linkLabel: '今すぐ申請' },
    monthly3: { name: '第3回月次セール — Manual BD', sub: '2回募集｜1%以上のポイント設定が必要', deadlines: ['第1回', '第2回'], linkLabel: '今すぐ申請', earlyBird: '💡 第1回に申請すると、プロモーション料金は最大6件分' },
    'monthly-points': { name: '4月月次セール（ポイントキャンペーン）', sub: 'セラーセントラル / Qualtrics', typeLabel: 'ポイント', linkLabel: '今すぐ申請' },
    pd5: { name: 'ポイントフェスティバル #5', sub: 'セラーセントラル / Qualtrics', typeLabel: 'ポイント', linkLabel: '今すぐ申請' },
    'mde4-points-deal': { name: '第4回月次セール（MDE#4）— Points DEAL', sub: 'Qualtrics で申請｜無料｜5%～50%のポイント設定', typeLabel: 'ポイント', linkLabel: '今すぐ申請', earlyBird: '💡 ポイントバッジは5/27 0:00（JST）から有効｜セルフサービスは Points DEAL ポータルを参照' },
    monthly4: { name: '第4回月次セール（MDE#4）— Manual BD', sub: '2回募集｜1%以上のポイント設定が必要', deadlines: ['第1回', '第2回'], linkLabel: '今すぐ申請', earlyBird: '💡 第1回に申請すると、プロモーション料金は最大6件分（大幅節約）' },
    hde2: { name: '初夏セール（HDE#2）— Manual BD', sub: '2回募集｜1%以上のポイント設定が必要', deadlines: ['第1回', '第2回'], linkLabel: '今すぐ申請', earlyBird: '💡 第1回に申請すると、プロモーション料金は最大6件分（大幅節約）' },
    'points-festival-7': { name: 'ポイントフェスティバル #7（夏季）— Points DEAL', sub: 'セラーセントラルからセルフ申請｜無料｜100% BO はポイント設定が必要', typeLabel: 'ポイント', deadlines: ['ポイント設定期限'], linkLabel: 'セラーセントラルで申請', earlyBird: '💡 Qualtrics は7/13に終了しましたが、セラーセントラルから追加申請できます（変更・取消不可）' },
    hde3: { name: '夏日 Sale（HDE#3）— Manual BD', sub: '1回募集｜生活応援夏日 Sale＋夏ギフト Sale｜1%以上のポイント設定が必要', deadlines: ['申請期限'], linkLabel: '今すぐ申請', earlyBird: '💡 キャンペーン全体でプロモーション料金は最大5件分｜ポイント設定期限 7/23' },
    'points-festival-8': { name: 'ポイントフェスティバル #8（盛夏）— Points DEAL', sub: 'ポイントセントラルからセルフ申請｜無料｜ポイント2倍 8/8～8/9、8/15～8/16', typeLabel: 'ポイント', deadlineText: '招待メール：7/29締切（時刻記載なし）<br>ポイントセントラルから随時申請可能', linkLabel: 'ポイントセントラルで申請', earlyBird: '💡 推奨ポイント率10%以上｜招待リンクから送信後は変更・取消不可' },
    'sl-mde5-september-dotd': { name: 'ファッションカテゴリー MDE#5＋9月 DOTD', sub: 'Fashion ASIN 限定｜FDE#7、MDE#5、FBS#7、FDE#8', eventText: 'FDE#7：8/21 09:00～8/27 23:59<br>MDE#5：8/28 00:00～9/3 23:59<br>FBS#7：9/5 09:00～9/14 23:59<br>FDE#8：9/18 09:00～9/28 23:59', deadlines: ['申請期限'], linkLabel: 'DOTD に申請', earlyBird: '💡 FDE#7 は MDE#5 との連続実施のみ対応｜MDE ASIN は1%以上のポイント設定が必要' },
    hde4: { name: '生活応援 夏祭り Sale（HDE#4）— Manual BD', sub: '2回募集｜Manual BD｜Hardline カテゴリー特設ページ', deadlines: ['第1回', '第2回'], linkLabel: '今すぐ申請', earlyBird: '💡 第1回はプロモーション料金が最大6件分｜第2回以降は最大10件分' },
    mde5: { name: '第5回 Amazon 月次セール（MDE#5）— Manual BD', sub: '2回募集｜Manual BD', deadlines: ['第1回', '第2回'], linkLabel: '今すぐ申請', earlyBird: '💡 第1回はプロモーション料金が最大6件分｜第2回以降は最大10件分' },
    'b2b-brand-standard-2026': { name: '2026 Amazon.co.jp 法人向けブランド特集（通常枠）', sub: '申請不要｜5%以上の法人価格割引を設定すると自動掲載', typeLabel: 'B2B', deadlines: ['キャンペーン期間中'], deadlineText: 'キャンペーン期間中', linkLabel: '設定ガイド', earlyBird: '💡 通常枠は毎日更新されます。早めに設定して露出機会を増やしましょう' },
    'prime-manual-bd': { name: 'プライムデー — Manual BD', sub: '3回募集｜前払費用 ¥600＋売上の1.0%｜1%以上のポイント設定が必要', deadlines: ['第1回', '第2回', '第3回'], linkLabel: '今すぐ申請', earlyBird: '💡 第1回に申請すると、プロモーション料金は最大6件分（大幅節約）' },
    'prime-bd': { name: 'プライム会員限定 Zお買い得', sub: 'PE-BD｜前払費用 ¥600＋売上の1.0%', linkLabel: 'セラーセントラル' },
    'prime-ld': { name: 'プライム会員限定タイムセール', sub: 'PE-LD', linkLabel: 'セラーセントラル' },
    'prime-ped': { name: 'プライム会員限定割引', sub: 'PED｜前払費用 ¥600＋売上の1.0%', deadlines: ['開始12時間前'], deadlineText: '開始12時間前', linkLabel: 'セラーセントラル' },
    'prime-points': { name: 'ポイントキャンペーン（Points Deal）', sub: '無料｜推奨ポイント率10%', typeLabel: 'ポイント', deadlines: ['キャンペーン期間中'], deadlineText: 'キャンペーン期間中', linkLabel: 'セラーセントラル' },
    'prime-coupon': { name: 'クーポン', sub: '作成料 ¥150/枚＋クーポン経由売上の1.5%', typeLabel: 'クーポン', deadlines: ['キャンペーン期間中'], deadlineText: 'キャンペーン期間中', linkLabel: 'セラーセントラル' },
    'pbdd-pe-bd': { name: 'プライム会員限定 Zお買い得（PE-BD）', sub: 'プライム感謝祭｜タイムセールページ掲載枠', eventText: '2026年10月<br>詳細日程は公式発表待ち', linkLabel: 'セラーセントラル' },
    'pbdd-ped': { name: 'プライム会員限定割引（PED）', sub: 'プライム感謝祭｜商品ページのプライム限定価格', eventText: '2026年10月<br>詳細日程は公式発表待ち', linkLabel: 'セラーセントラル' },
    'pbdd-manual-bd': { name: 'プライム感謝祭 — Manual BD', sub: 'Manual BD｜申請方法は公式発表待ち', eventText: '2026年10月<br>詳細日程は公式発表待ち', linkLabel: '更新待ち' },
    'pbdd-coupon': { name: 'プライム感謝祭 — クーポン', sub: '緑色バッジで露出｜申請方法は公式発表待ち', typeLabel: 'クーポン', eventText: '2026年10月<br>詳細日程は公式発表待ち', linkLabel: '更新待ち' },
    'pbdd-points-deal': { name: 'プライム感謝祭 — Points DEAL', sub: 'ポイントキャンペーン｜申請受付前（公式発表待ち）｜受付開始後：Points Central → Promotions', typeLabel: 'ポイント', eventText: '2026年10月<br>詳細日程は公式発表待ち', linkLabel: '入口（受付前）' }
  };

  const jaText = [
    ['目前沒有進行中或開放提報的活動', '現在、開催中または申請受付中のキャンペーンはありません'],
    ['目前沒有尚未開放提報的活動', '現在、申請受付前のキャンペーンはありません'],
    ['目前沒有已結束的活動', '終了したキャンペーンはありません'],
    ['目前沒有近期提醒', '直近のリマインダーはありません'],
    ['Prime Day 早鳥優惠截止 — Z劃算/秒殺省 ¥500', 'プライムデー早期申請特典期限 — Zお買い得/タイムセールが ¥500割引'],
    ['Prime 感謝特賣參加商品需送達 FBA', 'プライム感謝祭の対象商品はFBAへの納品が必要'],
    ['建議海運常規補貨、緊急用標準空運', '通常補充は船便、緊急時は標準航空便を推奨'],
    ['Prime 感謝特賣 FBA 入倉截止', 'プライム感謝祭 FBA納品期限'],
    ['Prime Day FBA 入倉截止', 'プライムデー FBA納品期限'],
    ['Prime 感謝特賣 PBDD（預計10月）', 'プライム感謝祭 PBDD（10月予定）'],
    ['2026 Prime 會員日（預計7月）', '2026 プライムデー（7月予定）'],
    ['尚未開放提報', '申請受付前'],
    ['進行中 / 開放提報', '開催中 / 申請受付中'],
    ['已結束活動', '終了したキャンペーン'],
    ['活動進行中', '開催中'],
    ['開放提報中', '申請受付中'],
    ['開放提報', '申請受付中'],
    ['即將到來', '近日開催'],
    ['提報連結', '申請リンク'],
    ['提報截止', '申請期限'],
    ['活動期間', '開催期間'],
    ['活動名稱', 'キャンペーン名'],
    ['活動日曆', 'キャンペーンカレンダー'],
    ['Deal 類型', 'Dealタイプ'],
    ['近期提醒', '直近のリマインダー'],
    ['FBA 入倉截止', 'FBA納品期限'],
    ['活動開始', 'キャンペーン開始'],
    ['提報截止', '申請期限'],
    ['已截止', '受付終了'],
    ['已結束', '終了'],
    ['待官方公告', '公式発表待ち'],
    ['待公告', '発表待ち'],
    ['待發布', '発表待ち'],
    ['待更新', '更新待ち'],
    ['活動期間內', 'キャンペーン期間中'],
    ['剩餘 ', '残り'],
    ['今天 ', '本日 '],
    [' 天', '日'],
    ['活動期間', '開催期間'],
    ['Deal 提報管理', 'Deal 申請管理'],
    ['Prime 會員日', 'プライムデー'],
    ['Prime 感謝特賣', 'プライム感謝祭'],
    ['前提交', 'までに申請']
  ];

  const originalFmt = fmt;
  const originalFmtShort = fmtShort;
  let currentLocale = ZH;

  function rememberChinese() {
    deals.forEach(deal => {
      if (deal._zh) return;
      deal._zh = {};
      translatableFields.forEach(field => {
        deal._zh[field] = { exists: Object.prototype.hasOwnProperty.call(deal, field), value: deal[field] };
      });
      deal._zh.deadlines = deal.deadlines.map(deadline => deadline.label);
    });
  }

  function localizeDeals(locale) {
    deals.forEach(deal => {
      const translated = jaDeals[deal.id] || {};
      translatableFields.forEach(field => {
        if (locale === JA && Object.prototype.hasOwnProperty.call(translated, field)) {
          deal[field] = translated[field];
        } else if (deal._zh[field].exists) {
          deal[field] = deal._zh[field].value;
        } else {
          delete deal[field];
        }
      });
      deal.deadlines.forEach((deadline, index) => {
        deadline.label = locale === JA && translated.deadlines
          ? (translated.deadlines[index] || '')
          : deal._zh.deadlines[index];
      });
    });
  }

  function setDateFormat(locale) {
    if (locale === JA) {
      const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
      fmt = d => {
        const dt = p(d);
        return `${dt.getMonth() + 1}/${dt.getDate()}（${weekdays[dt.getDay()]}）${String(dt.getHours()).padStart(2, '0')}:${String(dt.getMinutes()).padStart(2, '0')}`;
      };
      fmtShort = d => {
        const dt = p(d);
        return `${dt.getMonth() + 1}/${dt.getDate()}（${weekdays[dt.getDay()]}）`;
      };
    } else {
      fmt = originalFmt;
      fmtShort = originalFmtShort;
    }
  }

  function setStaticUi(locale) {
    const ja = locale === JA;
    document.querySelector('.navbar .title').textContent = ja ? 'Deal 申請管理' : 'Deal 提報管理';
    document.querySelector('[data-tab="active"]').innerHTML = `${ja ? '📋 開催中 / 申請受付中' : '📋 進行中 / 開放提報'} <span class="count" id="activeCount">0</span>`;
    const upcomingTab = document.querySelector('[data-tab="upcoming"]');
    if (upcomingTab) upcomingTab.innerHTML = `${ja ? '🗓️ 申請受付前' : '🗓️ 尚未開放提報'} <span class="count" id="upcomingCount">0</span>`;
    document.querySelector('[data-tab="ended"]').innerHTML = `${ja ? '📁 終了したキャンペーン' : '📁 已結束活動'} <span class="count" id="endedCount">0</span>`;

    const activeHeaders = ja ? ['✓', 'キャンペーン名', 'Dealタイプ', '開催期間', '申請期限', 'ステータス', '申請リンク'] : ['✓', '活動名稱', 'Deal 類型', '活動期間', '提報截止', '狀態', '提報連結'];
    ['#tab-active', '#tab-upcoming'].forEach(selector => {
      document.querySelectorAll(`${selector} th`).forEach((th, index) => { th.textContent = activeHeaders[index]; });
    });
    const endedHeaders = ja ? ['キャンペーン名', 'Dealタイプ', '開催期間', '申請期限', 'ステータス'] : ['活動名稱', 'Deal 類型', '活動期間', '提報截止', '狀態'];
    document.querySelectorAll('#tab-ended th').forEach((th, index) => { th.textContent = endedHeaders[index]; });

    document.getElementById('calHeader').textContent = ja ? '📅 キャンペーンカレンダー' : '📅 活動日曆';
    const headers = document.querySelectorAll('.panel-header');
    if (headers[1]) headers[1].textContent = ja ? '🔔 直近のリマインダー' : '🔔 近期提醒';
    const legend = document.querySelectorAll('.calendar > div:last-child > span');
    if (legend[0]?.lastChild) legend[0].lastChild.nodeValue = ja ? ' 開催期間' : ' 活動期間';
    if (legend[1]?.lastChild) legend[1].lastChild.nodeValue = ja ? ' 申請期限' : ' 提報截止';
  }

  function translateRenderedText() {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(node => {
      const parent = node.parentElement;
      if (!parent || ['SCRIPT', 'STYLE'].includes(parent.tagName) || parent.closest('.language-switch')) return;
      let value = node.nodeValue;
      jaText.forEach(([source, target]) => { value = value.split(source).join(target); });
      node.nodeValue = value;
    });
  }

  function translateCalendarDays() {
    if (currentLocale !== JA) return;
    const days = ['日', '月', '火', '水', '木', '金', '土'];
    document.querySelectorAll('#calGrid .day-header').forEach((element, index) => { element.textContent = days[index]; });
  }

  function addLanguageSwitch() {
    const font = document.createElement('link');
    font.rel = 'stylesheet';
    font.href = 'https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&display=swap';
    document.head.appendChild(font);

    const style = document.createElement('style');
    style.textContent = `
      html[lang="ja"] body { font-family:'Noto Sans JP','Noto Sans TC',sans-serif; }
      .language-switch { margin-left:auto; display:flex; align-items:center; gap:4px; padding:3px; border:1px solid #4B5563; border-radius:8px; background:#111827; }
      .language-switch button { border:0; border-radius:5px; padding:6px 10px; color:#D1D5DB; background:transparent; font:600 12px/1.2 inherit; cursor:pointer; }
      .language-switch button:hover { color:#FFF; background:#374151; }
      .language-switch button.active { color:#111827; background:#FF9900; }
      @media (max-width:600px) { .navbar { padding:12px 16px; flex-wrap:wrap; } .navbar .subtitle { order:3; width:100%; } .language-switch { margin-left:auto; } }
    `;
    document.head.appendChild(style);

    const control = document.createElement('div');
    control.className = 'language-switch';
    control.setAttribute('role', 'group');
    control.setAttribute('aria-label', '語言 / Language');
    control.innerHTML = '<button type="button" data-locale="zh-Hant">繁中</button><button type="button" data-locale="ja">日本語</button>';
    document.querySelector('.navbar').appendChild(control);
    control.querySelectorAll('button').forEach(button => button.addEventListener('click', () => applyLocale(button.dataset.locale, true)));
  }

  function updateUrl(locale) {
    try {
      const url = new URL(window.location.href);
      url.searchParams.set('lang', locale === JA ? 'ja' : ZH);
      history.replaceState(null, '', url);
    } catch (_) { /* Local file previews may block history updates. */ }
  }

  function applyLocale(locale, updateAddress = false) {
    currentLocale = locale === JA ? JA : ZH;
    document.documentElement.lang = currentLocale;
    document.title = currentLocale === JA
      ? 'Deal 申請管理ダッシュボード | Amazon Global Selling Taiwan'
      : 'Deal 提報管理儀表板 | Amazon Global Selling Taiwan';
    try { localStorage.setItem(LOCALE_KEY, currentLocale); } catch (_) { /* Storage can be disabled. */ }
    if (updateAddress) updateUrl(currentLocale);

    localizeDeals(currentLocale);
    setDateFormat(currentLocale);
    setStaticUi(currentLocale);
    renderTables();
    renderReminders();
    renderCalendar();
    if (currentLocale === JA) translateRenderedText();
    translateCalendarDays();

    document.querySelectorAll('.language-switch button').forEach(button => {
      const selected = button.dataset.locale === currentLocale;
      button.classList.toggle('active', selected);
      button.setAttribute('aria-pressed', String(selected));
    });
  }

  rememberChinese();
  addLanguageSwitch();
  let initialLocale = ZH;
  try {
    const requested = new URLSearchParams(window.location.search).get('lang');
    const stored = localStorage.getItem(LOCALE_KEY);
    if (requested === JA || requested === 'ja-JP') initialLocale = JA;
    else if (requested === ZH || requested === 'zh-TW') initialLocale = ZH;
    else if (stored === JA) initialLocale = JA;
  } catch (_) { /* Use Traditional Chinese by default. */ }
  applyLocale(initialLocale);

  ['calPrev', 'calNext'].forEach(id => {
    document.getElementById(id).addEventListener('click', () => translateCalendarDays());
  });
})();