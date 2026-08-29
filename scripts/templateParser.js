/* ─────────────────────────────────────────
   templateParser.js
   Splits the uploaded template into
   individual email objects and replaces
   all placeholders with real values.

   Template block structure (one block per vendor).
   Vendor sections are always demarcated by a
   "# 1 — Vendor Name" style checklist header — this
   is fixed by the standardized Prep Day Confirmation
   Template format used company-wide, so it's detected
   automatically rather than asked for as user input:
     TO:         →  Recipient list — may be separated by
                     semicolons, commas, and/or line breaks
     Subject:    →  Email subject line
     CC:         →  Template CC list (e.g. [COLEADER_EMAILS]) —
                     merged with the user-entered CC field
     Attachments: →  Comma/semicolon-separated list of files this
                     vendor needs (e.g. "Rooming List") — the preview
                     prompts for each one and blocks sending until
                     it's been attached, so it can't be forgotten.
     EMAIL BODY: →  Everything below this line becomes the
                     email body. Blocks with no TO: line are
                     treated as non-email front matter (cover
                     page, vendor checklist, data-entry table)
                     and are dropped rather than rendered as
                     a phantom email.

   Placeholders in body:

[PREP_STAFF_NAME]
[NUM_GUESTS]
[NUM_LEADERS]
[NUM_GUESTS+1]
[NUM_GUESTS+2]
[NUM_GUESTS+3]
[NUM_ROOMS_HOTEL1]
[NUM_ROOMS_HOTEL2]
[NUM_ROOMS_HOTEL3]
[NUM_ROOMS_HOTEL4]
[DIETARY_RESTRICTIONS]
[SPECIAL_ROOM_REQUESTS]
[LEADER_1_NAME]
[LEADER_1_PHONE]
[LEADER_2_NAME]
[LEADER_2_PHONE]
[COLEADER_EMAILS]
[D1_DATE]
[D2_DATE]
[D3_DATE]
[D4_DATE]
[D5_DATE]
[D6_DATE]
───────────────────────────────────────── */

// ── Applies all substitutions to a single string ──
function applySubstitutions(text) {
  const dateMap = buildDateMap();          // from dateUtils.js — { D1: "July 10", D2: "July 11", ... }

  // Safely read each field — returns '' if the element doesn't exist
  const val = id => document.getElementById(id)?.value.trim() ?? '';
  const prepDay = val('d1-date');
  const dietary = val('dietary-input');
  const guests  = Number(val('numGuest'));
  const adults  = Number(val('numAdult'));
  const kids    = val('numKid');
  const kidAge    = val('kidAges');
  const leaders = val('numLeader');
  const specialRequest = val('specialRequest');
  const prepperName = val('prepperName');
  const coleaderEmails = val('cc-input');

  // Replace [D1_DATE]…[D6_DATE] with their formatted dates.
  // dateMap keys are bare "D1".."D6" — the real placeholders in the
  // template are wrapped like "[D1_DATE]", so build the bracketed
  // form here rather than replacing the bare substring (which would
  // never match "[D1_DATE]" and could corrupt unrelated text).
  for (const [dayKey, formattedDate] of Object.entries(dateMap)) {
    text = text.replaceAll(`[${dayKey}_DATE]`, formattedDate);
  }

  // Replace dietary-restrictions placeholder
  text = text.replaceAll('[DIETARY_RESTRICTIONS]', dietary || '[No dietary restrictions provided]');
  // Replace guest count placeholder
  text = text.replaceAll('[NUM_GUESTS]', guests);
  text = text.replaceAll('[NUM_GUESTS+1]', guests+1);
  text = text.replaceAll('[NUM_GUESTS+2]', guests+2);
  text = text.replaceAll('[NUM_GUESTS+3]', guests+3);
  text = text.replaceAll('[NUM_GUESTS+4]', guests+4);
  // Replace adult count placeholder
  text = text.replaceAll('[NUM_ADULTS]', adults);
  // Replace kid count placeholder
  text = text.replaceAll('[NUM_KIDS]', kids);
  text = text.replaceAll('[KIDS_AGES]', kidAge);
  // Replace leader count placeholder
  text = text.replaceAll('[NUM_LEADERS]', leaders);
  // Replace each leader's name/phone placeholder — the "Number of
  // Leaders" dropdown tops out at 5, so that's the max we substitute.
  for (let i = 1; i <= 5; i++) {
    text = text.replaceAll(`[LEADER_${i}_NAME]`, val(`leader${i}Name`));
    text = text.replaceAll(`[LEADER_${i}_PHONE]`, val(`leader${i}Num`));
  }

  // Replace each hotel's room-count placeholder — the "Number of
  // Hotels" dropdown tops out at 5, so that's the max we substitute.
  for (let i = 1; i <= 5; i++) {
    text = text.replaceAll(`[NUM_ROOMS_HOTEL${i}]`, val(`hotel${i}Rooms`));
  }
  text = text.replaceAll('[SPECIAL_ROOM_REQUESTS]', specialRequest);

  text = text.replaceAll('[PREP_DATE]', prepDay);
  text = text.replaceAll('[PREP_STAFF_NAME]', prepperName);
  text = text.replaceAll('[COLEADER_EMAILS]', coleaderEmails);


  return text;
}

// ── Finds any placeholder-shaped tokens left in a string after
// substitution (e.g. a typo'd [D2DATE] or [NUM_GUEST] in the source
// doc that doesn't match a known placeholder). Used to warn prep
// staff before a vendor ever sees raw template syntax. ──
function findUnresolvedPlaceholders(text) {
  const matches = text.match(/\[[A-Z0-9_+]+\]/g) || [];
  return [...new Set(matches)];
}

// Vendor checklist headers always look like "# 1 — Vendor Name" (a "#"
// immediately followed by optional whitespace and a digit). This is the
// fixed structural boundary between vendor sections in every confirm
// template — matching on the digit avoids false splits on unrelated "#"
// characters elsewhere in the doc (e.g. a "#" table column header, or
// "# Rooms at First Hotel" in the data-entry table, neither of which is
// followed by a digit).
const VENDOR_HEADER_SPLIT = /#(?=\s*\d)/;

// ── Splits the full template into one object per email block ──
function parseEmails(rawTemplate) {
  // Split on each vendor checklist header, drop any empty blocks
  const allBlocks = rawTemplate
    .split(VENDOR_HEADER_SPLIT)
    .map(block => block.trim())
    .filter(block => block.length > 0);

  // Some vendor sections are intentionally non-email tasks (e.g. a
  // "WhatsApp BODY:" contact instead of "TO:" / "EMAIL BODY:"), and any
  // leftover front matter won't have a TO: line either — only blocks
  // with a real "TO:" line represent a vendor email.
  const blocks = allBlocks.filter(block => /(^|\n)\s*TO:\s*\S/i.test(block));

  const skipped = allBlocks.length - blocks.length;
  if (skipped > 0 && typeof addLog === 'function') {
    addLog('info', `Skipped ${skipped} non-email block(s) with no TO: line (front matter, or a non-email task like a WhatsApp/Call contact).`);
  }

  return blocks.map((block, index) => parseEmailBlock(block, index + 1));
}

// ── Parses a single email block into a structured object ──
//
// Template structure observed:
//   - TO:, Subject:, and CC: appear ABOVE the EMAIL BODY: marker
//   - TO: recipients can span multiple lines and may be separated by
//     semicolons, commas, and/or plain whitespace (real vendor lists
//     are inconsistent about this)
//   - EMAIL BODY: marks the start of the actual email body
//   - Blocks with no EMAIL BODY: tag use the whole block as the body
//
function parseEmailBlock(block, emailNumber) {

  // ── Step 1: Extract TO: from the full block (before any slicing).
  // TO: is uppercase and recipients may continue on the next line(s)
  // until a blank line or another tag is encountered.
  let recipients = [];
  const toMatch = block.match(/^TO:\s*([\s\S]*?)(?=\n\s*\n|\n[A-Z][a-zA-Z]*:|$)/mi);
  if (toMatch) {
    // Recipients may be separated by semicolons, commas, newlines, or
    // just whitespace (some vendor rows are missing a semicolon between
    // addresses) — split on any run of those and keep only real addresses.
    recipients = toMatch[1]
      .split(/[;,\s]+/)
      .map(e => e.trim())
      .filter(e => e.includes('@'));  // only keep actual email addresses
  }

  // ── Step 2: Extract Subject: from the full block (also above EMAIL BODY:).
  const subjectMatch = block.match(/^Subject:\s*(.+)$/mi);
  const subject = subjectMatch ? subjectMatch[1].trim() : '';

  // ── Step 3: Extract CC: from the full block.
  const ccMatch = block.match(/^CC:\s*(.+)$/mi);
  const templateCC = ccMatch
    ? ccMatch[1].split(/[;,]+/).map(e => e.trim()).filter(e => e.includes('@'))
    : [];
  const userCC = document.getElementById('cc-input').value
    .split(/[;,]+/).map(e => e.trim()).filter(Boolean);
  const allCC = [...new Set([...templateCC, ...userCC])];  // deduplicate

  // ── Step 3b: Extract Attachments: — a comma/semicolon-separated list
  // of files this vendor needs (e.g. "Attachments: Rooming List"). Each
  // one becomes a slot the preview prompts for and blocks sending on
  // until a file is actually attached. "None"/"none" means this vendor
  // needs nothing attached, so no prompt should appear at all.
  const attachmentsMatch = block.match(/^Attachments:\s*(.+)$/mi);
  const attachments = attachmentsMatch
    ? applySubstitutions(attachmentsMatch[1].trim())
        .split(/[;,]+/)
        .map(label => label.trim())
        .filter(Boolean)
        .filter(label => label.toLowerCase() !== 'none')
        .map(label => ({ label, file: null }))
    : [];

  // ── Step 4: Extract body — everything after the 'EMAIL BODY:' line.
  // If no EMAIL BODY: tag exists, use the whole block as the body.
  let body = block;
  const emailTagMatch = block.match(/^\s*(?:email\s+)?body\s*:.*$/im);
  if (emailTagMatch) {
    const afterEmailLine = block.indexOf(emailTagMatch[0]) + emailTagMatch[0].length;
    body = block.slice(afterEmailLine).trim();
  }

  // Apply substitutions to both body and subject
  body = applySubstitutions(body);
  const resolvedSubject = applySubstitutions(subject);

  // Flag anything that still looks like a placeholder so it never
  // silently goes out to a vendor (e.g. a typo'd [D2DATE] in the doc).
  const unresolvedPlaceholders = [...new Set([
    ...findUnresolvedPlaceholders(body),
    ...findUnresolvedPlaceholders(resolvedSubject),
  ])];

  return {
    emailNumber,
    recipients,
    cc: allCC,
    subject: resolvedSubject,
    body,
    unresolvedPlaceholders,
    attachments,
  };
}
