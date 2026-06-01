/* Cloudflare Pages Function — POST /ocr
 * 差旅工具的掃描圖 OCR 代理：金鑰只存在伺服器端（Pages 機密變數），絕不進前端。
 * 流程：收 base64 圖 → 伺服器端加 GEMINI_KEY 呼叫 Gemini pro → 背景寄 Email 通知 → 回傳 rows
 *
 * 需要在 Cloudflare Pages → Settings → Environment variables 設定（機密）：
 *   GEMINI_KEY   你的付費 Gemini API key
 *   RESEND_KEY   Resend API key（寄信用）
 *   RESEND_TO    收通知的 email（Resend 免網域時須為你註冊 Resend 的那個 email）
 *   RESEND_FROM  (選填) 寄件者，免網域預設 onboarding@resend.dev
 */
export async function onRequestPost(context) {
    const { request, env } = context;
    const json = (obj, status = 200) =>
        new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });

    try {
        const body = await request.json().catch(() => ({}));
        const { imageB64, mime, context: ctx, file } = body;
        if (!imageB64) return json({ error: '缺少圖片' }, 400);
        if (!env.GEMINI_KEY) return json({ error: '伺服器尚未設定 GEMINI_KEY' }, 500);

        const model = env.GEMINI_MODEL || 'gemini-2.5-pro';
        const prompt =
            '這是台南市政府衛生局的差旅費/補助掃描清單表。逐列辨識，回傳純 JSON 陣列，' +
            '每列物件鍵：office(衛生所/單位全名，如「佳里區衛生所」)、name(領受人/姓名)、amount(金額數字)。' +
            '姓名務必逐字辨識、不可漏字或縮寫；若該表沒有姓名欄就回空陣列。' +
            '不要「合計/總計/以下空白/承辦人」等列。不要輸出身分證字號。只回純 JSON 陣列。';
        const gReq = {
            contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mime || 'image/jpeg', data: imageB64 } }] }],
            generationConfig: { responseMimeType: 'application/json', temperature: 0 },
        };
        const gRes = await fetch(
            'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + encodeURIComponent(env.GEMINI_KEY),
            { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(gReq) }
        );
        if (!gRes.ok) {
            const t = await gRes.text();
            return json({ error: 'Gemini ' + gRes.status, detail: t.slice(0, 200) }, 502);
        }
        const gJson = await gRes.json();
        const txt = (gJson.candidates && gJson.candidates[0] && gJson.candidates[0].content
            && gJson.candidates[0].content.parts[0] && gJson.candidates[0].content.parts[0].text) || '[]';
        let rows;
        try { rows = JSON.parse(txt); } catch (e) { rows = []; }
        if (!Array.isArray(rows)) rows = rows.data || rows.rows || [];

        /* 背景寄通知（用 waitUntil，不拖慢回傳；失敗不影響辨識結果） */
        if (env.RESEND_KEY && env.RESEND_TO) {
            const ip = request.headers.get('cf-connecting-ip') || '?';
            const ua = request.headers.get('user-agent') || '?';
            const payload = {
                from: env.RESEND_FROM || 'onboarding@resend.dev',
                to: env.RESEND_TO,
                subject: '差旅工具 OCR 使用通知｜' + (ctx || '(未標示衛生所)'),
                text:
                    '差旅明細整理系統的掃描圖 OCR 被使用：\n' +
                    '時間(UTC)：' + new Date().toISOString() + '\n' +
                    '衛生所/情境：' + (ctx || '(未標示)') + '\n' +
                    '檔案：' + (file || '-') + '\n' +
                    '辨識列數：' + rows.length + '\n' +
                    '來源 IP：' + ip + '\n' +
                    'User-Agent：' + ua,
            };
            context.waitUntil(
                fetch('https://api.resend.com/emails', {
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + env.RESEND_KEY, 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload),
                }).catch(() => {})
            );
        }
        return json({ rows, model });
    } catch (e) {
        return json({ error: e.message || String(e) }, 500);
    }
}
