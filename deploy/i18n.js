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
