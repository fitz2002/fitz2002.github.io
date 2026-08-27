/* ─────────────────────────────────────────
   app.js
   Top-level orchestration.
   Wires up UI event listeners and runs
   the main "process & send" workflow.

   Depends on (loaded before this file):
     fileHandler.js    — templateText, file upload
     dateUtils.js      — buildDateMap, buildDayGrid
     templateParser.js — parseEmails, applySubstitutions
     emailSender.js    — authenticateOutlook, sendEmail
───────────────────────────────────────── */

// ── Leader Name/Number fields, built dynamically from the
// "Number of Leaders" dropdown so there's always exactly one
// Name/Number pair per leader (ids: leader1Name/leader1Num …) ──
const LEADER_PLACEHOLDER_NAMES = ['Michael Jordan', 'LeBron James', 'Tom Hale', 'Kobe Bryant', 'Larry Bird'];

function buildLeaderFields() {
  const count = parseInt(document.getElementById('numLeader').value) || 0;
  const container = document.getElementById('leader-fields');

  // Preserve any values already entered when the count changes
  const prevValues = {};
  container.querySelectorAll('textarea').forEach(el => { prevValues[el.id] = el.value; });
  container.innerHTML = '';

  for (let i = 1; i <= count; i++) {
    container.appendChild(createLeaderFieldPair(i, prevValues));
  }
}

function createLeaderFieldPair(leaderNumber, prevValues) {
  const row = document.createElement('div');
  row.className = 'two-col';

  const nameField = document.createElement('div');
  nameField.className = 'field';
  const nameLabel = document.createElement('label');
  nameLabel.innerHTML = `Leader ${leaderNumber} Name <span class="required-mark">*</span>`;
  const nameInput = document.createElement('textarea');
  nameInput.id = `leader${leaderNumber}Name`;
  nameInput.rows = 1;
  nameInput.required = true;
  nameInput.placeholder = LEADER_PLACEHOLDER_NAMES[leaderNumber - 1] || 'Leader name';
  nameInput.value = prevValues[nameInput.id] || '';
  nameField.appendChild(nameLabel);
  nameField.appendChild(nameInput);

  const numField = document.createElement('div');
  numField.className = 'field';
  const numLabel = document.createElement('label');
  numLabel.innerHTML = `Leader ${leaderNumber} Number <span class="required-mark">*</span>`;
  const numInput = document.createElement('textarea');
  numInput.id = `leader${leaderNumber}Num`;
  numInput.rows = 1;
  numInput.required = true;
  numInput.placeholder = '(xxx) xxx-xxxx';
  numInput.value = prevValues[numInput.id] || '';
  numField.appendChild(numLabel);
  numField.appendChild(numInput);

  row.appendChild(nameField);
  row.appendChild(numField);
  return row;
}

// ── Hotel room-count fields, built dynamically from the
// "Number of Hotels" dropdown (ids: hotel1Rooms, hotel2Rooms, …) ──
// Column count per hotel count, chosen so rows stay visually balanced:
// 1→1, 2→2, 3→3 (all in one row), 4→2 (2x2), 5→3 (3 top, 2 below)
const HOTEL_GRID_COLS = { 1: 1, 2: 2, 3: 3, 4: 2, 5: 3 };

function buildHotelFields() {
  const count = parseInt(document.getElementById('numHotels').value) || 0;
  const container = document.getElementById('hotel-fields');

  // Preserve any values already entered when the count changes
  const prevValues = {};
  container.querySelectorAll('textarea').forEach(el => { prevValues[el.id] = el.value; });
  container.innerHTML = '';

  const cols = HOTEL_GRID_COLS[count] || Math.min(count, 3) || 1;
  container.className = `field-grid cols-${cols}`;

  for (let i = 1; i <= count; i++) {
    container.appendChild(createHotelField(i, prevValues));
  }
}

function createHotelField(hotelNumber, prevValues) {
  const field = document.createElement('div');
  field.className = 'field';
  const label = document.createElement('label');
  label.innerHTML = `Number of Rooms Hotel ${hotelNumber} <span class="required-mark">*</span>`;
  const input = document.createElement('textarea');
  input.id = `hotel${hotelNumber}Rooms`;
  input.rows = 1;
  input.required = true;
  input.placeholder = '...';
  input.value = prevValues[input.id] || '';
  field.appendChild(label);
  field.appendChild(input);
  return field;
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('numLeader').addEventListener('change', buildLeaderFields);
  buildLeaderFields();

  document.getElementById('numHotels').addEventListener('change', buildHotelFields);
  buildHotelFields();

  // ── "Kids Trip?" toggle — Number of Adults / Number of Kids / Kids
  // Ages stay hidden until this is checked (unchecked by default) ──
  const kidsToggle = document.getElementById('kids-trip-toggle');
  const kidsFields = document.getElementById('kids-fields');
  const syncKidsFields = () => kidsFields.classList.toggle('visible', kidsToggle.checked);
  kidsToggle.addEventListener('change', syncKidsFields);
  syncKidsFields();
});

// ── Required-field validation ──
// Every field marked with the `required` attribute (Prep Staff Name,
// Number of Guests, Co-leader emails, plus each currently-rendered
// Leader Name/Number and Hotel Rooms field) must be filled in before
// emails can be processed & sent. Uses one delegated listener so newly
// created leader/hotel fields are covered automatically.
document.addEventListener('input', (e) => {
  if (e.target.required && e.target.value.trim()) {
    e.target.classList.remove('field-invalid');
  }
});

function validateRequiredFields() {
  const invalid = [];
  document.querySelectorAll('[required]').forEach(el => {
    const empty = !el.value.trim();
    el.classList.toggle('field-invalid', empty);
    if (empty) invalid.push(el);
  });
  return invalid;
}

// ── Returns the list of blocking issues for an email in its current
// state (empty array once everything's resolved, or if it's been
// removed from the batch). Shared by the initial parse-time log pass
// and the Edit/Delete/Undo handlers so the log can report exactly
// what changed as issues are fixed. ──
function getEmailIssues(email) {
  if (email.deleted) return [];
  const issues = [];
  if (!email.recipients.length) issues.push('no To: recipients found');
  if (email.unresolvedPlaceholders.length) {
    issues.push(`unresolved placeholder(s) ${email.unresolvedPlaceholders.join(', ')}`);
  }
  const missingAttachments = (email.attachments || []).filter(a => !a.file);
  if (missingAttachments.length) {
    issues.push(`missing attachment(s): ${missingAttachments.map(a => a.label).join(', ')}`);
  }
  return issues;
}

// ── Main workflow: parse template → preview → send ──
async function processAndSend() {
  clearUI();

  if (!templateText) {
    addLog('error', 'No template loaded. Please upload a template file first.');
    return;
  }

  const invalidFields = validateRequiredFields();
  if (invalidFields.length) {
    addLog('error', `Please fill in all required fields (marked *) before sending — ${invalidFields.length} field(s) still empty.`);
    invalidFields[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
    invalidFields[0].focus();
    return;
  }

  const sendBtn = document.getElementById('send-btn');
  sendBtn.disabled = true;
  sendBtn.innerHTML = '<span class="btn-icon">⏳</span> Processing…';

  try {
    const token = document.getElementById('access-token')?.value.trim() || '';
    if (!token) {
      addLog('warn', 'No access token — will preview emails but NOT send. Authenticate in Step 6 to send.');
    }

    addLog('info', 'Parsing template…');
    const emails = parseEmails(templateText);
    addLog('info', `Found ${emails.length} email(s) in template.`);

    emails.forEach(email => {
      getEmailIssues(email).forEach(issue => {
        addLog('warn', `Email ${email.emailNumber}: ${issue}.`);
      });
    });

    const statusEls = renderEmailCards(emails);

    if (!token) {
      addLog('warn', 'Preview complete. No emails sent (not authenticated).');
      return;
    }

    sendBtn.innerHTML = '<span class="btn-icon">⏳</span> Sending…';

    let successCount = 0;
    let failCount    = 0;

    for (let i = 0; i < emails.length; i++) {
      const email = emails[i];

      if (email.deleted) {
        addLog('info', `Email ${email.emailNumber} was removed — skipping.`);
        continue;
      }

      if (!email.recipients.length) {
        addLog('warn', `Email ${email.emailNumber}: No To: recipients found — skipped.`);
        setEmailStatus(statusEls[i], 'failed', 'Skipped');
        failCount++;
        continue;
      }

      if (email.unresolvedPlaceholders.length) {
        addLog('error', `Email ${email.emailNumber}: unresolved placeholder(s) ${email.unresolvedPlaceholders.join(', ')} — skipped, not sent to vendor.`);
        setEmailStatus(statusEls[i], 'failed', 'Blocked');
        failCount++;
        continue;
      }

      const missingAttachments = (email.attachments || []).filter(a => !a.file);
      if (missingAttachments.length) {
        addLog('error', `Email ${email.emailNumber}: missing attachment(s) ${missingAttachments.map(a => a.label).join(', ')} — skipped, not sent to vendor.`);
        setEmailStatus(statusEls[i], 'failed', 'Blocked');
        failCount++;
        continue;
      }

      addLog('info', `Sending email ${email.emailNumber} → ${email.recipients.join(', ')}…`);

      try {
        await sendEmail(email.recipients, email.cc, email.subject, email.body, token, email.attachments);
        addLog('success', `✓ Email ${email.emailNumber} sent to ${email.recipients.join(', ')}`);
        setEmailStatus(statusEls[i], 'sent', 'Sent');
        successCount++;
      } catch (err) {
        addLog('error', `✗ Email ${email.emailNumber} failed: ${err.message}`);
        setEmailStatus(statusEls[i], 'failed', 'Failed');
        failCount++;
      }

      await delay(400);
    }

    const allOk = failCount === 0;
    addLog(
      allOk ? 'success' : 'warn',
      `Done. ${successCount} sent, ${failCount} failed.`
    );

  } catch (err) {
    addLog('error', `Unexpected error: ${err.message}`);
    console.error(err);
  } finally {
    sendBtn.disabled = false;
    sendBtn.innerHTML = '<span class="btn-icon">✉️</span> Process &amp; Send All Emails';
  }
}

// ── Builds the "To: ... CC: ... Subject: ..." summary markup,
// including a warning line when unresolved placeholder syntax
// (e.g. a typo'd [D2DATE]) is still present in the body/subject ──
function buildToAddrHTML(email) {
  const warning = email.unresolvedPlaceholders.length
    ? `<br><span style="color:#c0392b;font-weight:600;">⚠ Unresolved placeholder(s): ${email.unresolvedPlaceholders.join(', ')} — will not be sent</span>`
    : '';
  return `To: ${email.recipients.join('; ') || '(no recipients found)'}` +
    (email.cc.length ? ` &nbsp;|&nbsp; CC: ${email.cc.join('; ')}` : '') +
    `<br><span class="subject-preview">Subject: ${email.subject || '(no subject)'}</span>` +
    warning;
}

// ── Builds the full header row markup for one email card ──
function buildHeadSummaryHTML(email) {
  return `
      <span class="to-addr">
        ${buildToAddrHTML(email)}
      </span>
      <span class="email-num">Email ${email.emailNumber}</span>
    `;
}

// ── Builds the "attachments needed" prompt for one email card, with a
// file input per item from its "Attachments:" line (e.g. "Rooming
// List"). Sending is blocked (see getEmailIssues) until every slot has
// a file, so a required attachment can't get forgotten. ──
function buildAttachmentsSection(email) {
  const section = document.createElement('div');
  section.className = 'attachments-prompt';

  const title = document.createElement('div');
  title.className = 'attachments-prompt-title';
  title.textContent = '📎 Attachments needed';
  section.appendChild(title);

  email.attachments.forEach(attachment => {
    const slot = document.createElement('div');
    slot.className = 'attachment-slot';

    const label = document.createElement('label');
    label.textContent = attachment.label;

    const fileInput = document.createElement('input');
    fileInput.type = 'file';

    const status = document.createElement('span');
    status.className = 'attachment-status';

    const refreshSlotStatus = () => {
      slot.classList.toggle('attached', !!attachment.file);
      status.textContent = attachment.file ? `✓ ${attachment.file.name}` : 'Not attached';
    };
    refreshSlotStatus();

    fileInput.addEventListener('change', () => {
      const issuesBefore = getEmailIssues(email);
      attachment.file = fileInput.files[0] || null;
      refreshSlotStatus();

      // Log exactly what changed, same as the Edit/Save flow, so the
      // send log reflects attachments as they're actually provided.
      const issuesAfter = getEmailIssues(email);
      issuesBefore
        .filter(issue => !issuesAfter.includes(issue))
        .forEach(issue => addLog('success', `Email ${email.emailNumber}: resolved — ${issue}.`));
      issuesAfter
        .filter(issue => !issuesBefore.includes(issue))
        .forEach(issue => addLog('warn', `Email ${email.emailNumber}: ${issue}.`));
    });

    slot.appendChild(label);
    slot.appendChild(fileInput);
    slot.appendChild(status);
    section.appendChild(slot);
  });

  return section;
}

// ── Renders an email preview card for each parsed email ──
// Returns an array of status <span> elements (one per email)
function renderEmailCards(emails) {
  const previewWrap  = document.getElementById('emails-preview');
  const emailsList   = document.getElementById('emails-list');
  const previewLabel = document.getElementById('emails-preview-label');

  previewWrap.style.display = 'block';
  previewLabel.textContent  = `${emails.length} Parsed Email${emails.length !== 1 ? 's' : ''}`;

  const statusEls = [];

  emails.forEach((email, i) => {
    // ── Status badge ──
    const statusSpan = document.createElement('span');
    statusSpan.className = 'email-send-status status-pending';
    statusSpan.textContent = 'Pending';
    statusEls.push(statusSpan);

    // ── Action buttons (Edit / Delete) ──
    const actions = document.createElement('div');
    actions.className = 'email-card-actions';

    const editBtn   = document.createElement('button');
    editBtn.className   = 'btn-edit';
    editBtn.textContent = '✏️ Edit';

    const deleteBtn   = document.createElement('button');
    deleteBtn.className   = 'btn-delete';
    deleteBtn.textContent = '🗑 Delete';

    actions.appendChild(editBtn);
    actions.appendChild(deleteBtn);

    // ── Header row ──
    const head = document.createElement('div');
    head.className = 'email-card-head';
    head.innerHTML = buildHeadSummaryHTML(email);
    head.appendChild(statusSpan);
    head.appendChild(actions);

    // ── Body (read-only display) ──
    const bodyDiv = document.createElement('div');
    bodyDiv.className = 'email-card-body';
    bodyDiv.textContent = email.body;

    // ── Assemble card ──
    const card = document.createElement('div');
    card.className = 'email-card';
    card.appendChild(head);
    if (email.attachments && email.attachments.length) {
      card.appendChild(buildAttachmentsSection(email));
    }
    card.appendChild(bodyDiv);
    emailsList.appendChild(card);

    // ── Helper: rebuilds the read-only header summary line ──
    function refreshHeadSummary() {
      head.querySelector('.to-addr').innerHTML = buildToAddrHTML(email);
    }

    // ── Helper: rebuilds the read-only body display ──
    function refreshBodyDisplay() {
      bodyDiv.className = 'email-card-body';
      bodyDiv.innerHTML = '';
      bodyDiv.textContent = email.body;
    }

    // ── Edit button logic ──
    editBtn.addEventListener('click', () => {
      if (card.classList.contains('deleted')) return;

      // Snapshot of issues as of opening the editor, so Save can log
      // exactly what this edit resolved or introduced.
      const issuesBeforeEdit = getEmailIssues(email);

      // Build the edit form
      const form = document.createElement('div');
      form.className = 'email-edit-form';

      // Subject field
      const subjectLabel = document.createElement('label');
      subjectLabel.className = 'edit-field-label';
      subjectLabel.textContent = 'Subject';
      const subjectInput = document.createElement('input');
      subjectInput.type  = 'text';
      subjectInput.value = email.subject;

      // To field
      const toLabel = document.createElement('label');
      toLabel.className = 'edit-field-label';
      toLabel.textContent = 'To  (semicolon-separated)';
      const toInput = document.createElement('input');
      toInput.type  = 'text';
      toInput.value = email.recipients.join('; ');

      // CC field
      const ccLabel = document.createElement('label');
      ccLabel.className = 'edit-field-label';
      ccLabel.textContent = 'CC  (semicolon-separated)';
      const ccInput = document.createElement('input');
      ccInput.type  = 'text';
      ccInput.value = email.cc.join('; ');

      // Body textarea
      const bodyLabel = document.createElement('label');
      bodyLabel.className = 'edit-field-label';
      bodyLabel.textContent = 'Body';
      const bodyTextarea = document.createElement('textarea');
      bodyTextarea.value = email.body;

      form.appendChild(subjectLabel);
      form.appendChild(subjectInput);
      form.appendChild(toLabel);
      form.appendChild(toInput);
      form.appendChild(ccLabel);
      form.appendChild(ccInput);
      form.appendChild(bodyLabel);
      form.appendChild(bodyTextarea);

      // Swap body div content for the form
      bodyDiv.className = 'email-card-body editing';
      bodyDiv.innerHTML = '';
      bodyDiv.appendChild(form);
      card.classList.add('editing-mode');
      subjectInput.focus();

      // Swap action buttons to Save / Cancel
      const saveBtn = document.createElement('button');
      saveBtn.className   = 'btn-save';
      saveBtn.textContent = '💾 Save';

      const cancelBtn = document.createElement('button');
      cancelBtn.className   = 'btn-cancel';
      cancelBtn.textContent = '✕ Cancel';

      actions.innerHTML = '';
      actions.appendChild(saveBtn);
      actions.appendChild(cancelBtn);

      // ── Save ──
      saveBtn.addEventListener('click', () => {
        email.subject    = subjectInput.value.trim();
        email.recipients = toInput.value.split(';').map(e => e.trim()).filter(Boolean);
        email.cc         = ccInput.value.split(';').map(e => e.trim()).filter(Boolean);
        email.body       = bodyTextarea.value;
        // Re-check for leftover placeholder syntax since manual edits
        // could introduce or resolve one.
        email.unresolvedPlaceholders = [...new Set([
          ...findUnresolvedPlaceholders(email.body),
          ...findUnresolvedPlaceholders(email.subject),
        ])];

        // Log exactly what this edit changed, so the send log reflects
        // issues as they're resolved instead of staying stuck at the
        // state from the initial parse.
        const issuesAfterEdit = getEmailIssues(email);
        issuesBeforeEdit
          .filter(issue => !issuesAfterEdit.includes(issue))
          .forEach(issue => addLog('success', `Email ${email.emailNumber}: resolved — ${issue}.`));
        issuesAfterEdit
          .filter(issue => !issuesBeforeEdit.includes(issue))
          .forEach(issue => addLog('warn', `Email ${email.emailNumber}: ${issue}.`));

        refreshHeadSummary();
        refreshBodyDisplay();
        card.classList.remove('editing-mode');

        actions.innerHTML = '';
        actions.appendChild(editBtn);
        actions.appendChild(deleteBtn);
      });

      // ── Cancel ──
      cancelBtn.addEventListener('click', () => {
        refreshBodyDisplay();
        card.classList.remove('editing-mode');

        actions.innerHTML = '';
        actions.appendChild(editBtn);
        actions.appendChild(deleteBtn);
      });
    });

    // ── Delete button logic ──
    deleteBtn.addEventListener('click', () => {
      const issuesAtDelete = getEmailIssues(email);
      email.deleted = true;
      card.classList.add('deleted');
      statusSpan.className   = 'email-send-status status-failed';
      statusSpan.textContent = 'Removed';

      addLog('info', issuesAtDelete.length
        ? `Email ${email.emailNumber} removed from batch (no longer blocked by: ${issuesAtDelete.join('; ')}).`
        : `Email ${email.emailNumber} removed from batch.`);

      // Replace action buttons with an Undo option
      actions.innerHTML = '';
      const undoBtn   = document.createElement('button');
      undoBtn.className   = 'btn-edit';   // reuse teal style
      undoBtn.textContent = '↩ Undo';
      actions.appendChild(undoBtn);

      undoBtn.addEventListener('click', () => {
        email.deleted = false;
        card.classList.remove('deleted');
        statusSpan.className   = 'email-send-status status-pending';
        statusSpan.textContent = 'Pending';

        addLog('info', `Email ${email.emailNumber} restored to batch.`);
        // Re-surface any issues that still apply now that it's back in.
        getEmailIssues(email).forEach(issue => addLog('warn', `Email ${email.emailNumber}: ${issue}.`));

        actions.innerHTML = '';
        actions.appendChild(editBtn);
        actions.appendChild(deleteBtn);
      });
    });
  });

  return statusEls;
}

// ── Updates one email card's status badge ──
function setEmailStatus(spanEl, type, label) {
  spanEl.className = `email-send-status status-${type}`;
  spanEl.textContent = label;
}

// ── Appends a line to the send log ──
function addLog(type, message) {
  const logSection = document.getElementById('log-section');
  logSection.classList.add('visible');

  const entry = document.createElement('div');
  entry.className = `log-entry ${type}`;
  entry.innerHTML = `<span class="log-dot"></span><span>${message}</span>`;

  document.getElementById('log-list').appendChild(entry);
  entry.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// ── Resets the log and email preview before each run ──
function clearUI() {
  document.getElementById('log-list').innerHTML    = '';
  document.getElementById('emails-list').innerHTML = '';
  document.getElementById('log-section').classList.remove('visible');
  document.getElementById('emails-preview').style.display = 'none';
}

// ── Simple promise-based delay ──
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
