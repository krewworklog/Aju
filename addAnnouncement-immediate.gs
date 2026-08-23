/* ═══════════════════════════════════════════════════════════════
   ██  REPLACE the existing addAnnouncement() in the LIVE v22 code
   ██  with this one. Do NOT paste Desktop/Aju/Code.gs — that file
   ██  is v21 and has no announcements code at all.
   ██
   ██  Change: announcements go live the moment they are posted,
   ██  instead of at 10:00 on the next working day. They still
   ██  expire 24 hours later.
   ══════════════════════════════════════════════════════════════ */

function addAnnouncement(ss, params) {
  const now = new Date();

  // Live immediately; still expires exactly 24 hours later.
  const showFrom  = now;
  const showUntil = new Date(now.getTime() + 24 * 60 * 60 * 1000);

  const obj = {
    id: 'ann' + Date.now(),
    title: String(params.title || '').trim(),
    body: String(params.body || '').trim(),
    postedById: params.postedById || '',
    postedByName: params.postedByName || '',
    postedAt: now.toISOString(),
    showFrom: showFrom.toISOString(),
    showUntil: showUntil.toISOString()
  };

  // Written through the project's own header-safe helper, so the column
  // order in the sheet is whatever row 1 says — not assumed here.
  addRow(ss, 'announcements', obj);

  return { ok: true, showFrom: obj.showFrom, showUntil: obj.showUntil };
}
