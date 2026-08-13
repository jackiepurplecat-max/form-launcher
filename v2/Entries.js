/**
 * v2 — entry creation.
 *
 * NOT YET DEPLOYED. See Config.js.
 *
 * There is exactly one way an entry is born: createEntry(). The custom form,
 * Siri and OCR all call it, and differ only in where the values came from.
 *
 * v1 had two paths, because Google Forms wrote the row itself and left the
 * trigger to finalise it. That split is gone with Forms — nothing writes a row
 * behind this module's back, so there is no "already exists" case to handle
 * and no onFormSubmit trigger to keep in step.
 *
 * initializeEntry() is the single place bookkeeping columns are set, documents
 * are named and filed, the registry is taught and creation mail is sent.
 */

/* ============================== Filenames ================================= */

/**
 * Reduce free text to something safe and consistent inside a filename:
 * accents flattened, each word capitalised so the result stays readable once
 * the spaces go, then everything non-alphanumeric removed.
 *
 *   "Hospital da Luz"  -> "HospitalDaLuz"
 *   "Farmácia Sá"      -> "FarmaciaSa"
 *   "José & Cia, Lda." -> "JoseCiaLda"
 */
function slugForFilename(text) {
  return (text || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // strip combining accents exposed by NFD
    .split(/\s+/)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join('')
    .replace(/[^A-Za-z0-9]/g, '');
}

/**
 * Base filename for a document: YYMMDD_Counterparty_Amount_<document>
 * The state suffix chain is appended later by applyFileState().
 */
function buildBaseFilename(sheet, cols, row, fileCol) {
  const dateValue = readCell(sheet, cols, row, COMMON.date);
  const stamp = dateValue
    ? Utilities.formatDate(new Date(dateValue), Session.getScriptTimeZone(), 'yyMMdd')
    : 'nodate';

  const who = slugForFilename(readCell(sheet, cols, row, COMMON.counterparty)) || 'unknown';
  const amount = formatAmountForFilename(readCell(sheet, cols, row, COMMON.amount));

  return [stamp, who, amount, fileCol.suffix].filter(part => part !== '').join('_');
}

/* ============================== Validation ================================ */

/**
 * Fields an entry should have. Returns a list of missing labels rather than
 * throwing, so the caller decides whether that is fatal.
 */
function missingFields(section, sheet, cols, row) {
  const missing = [];

  const core = [
    { header: COMMON.date, label: 'Date' },
    { header: COMMON.amount, label: 'Amount' },
    { header: COMMON.counterparty, label: counterpartyLabel(section) }
  ];
  // Required unless the config says otherwise. Work's Expense Reason and
  // Health's Patient are required; Income's Reason declares required: false,
  // because it prefills from the registry and is genuinely optional - treating
  // it as required made every Income entry report itself incomplete.
  if (section.category && section.category.required !== false) {
    core.push({ header: section.category.header, label: section.category.label });
  }

  core.forEach(field => {
    const value = readCell(sheet, cols, row, field.header);
    if (value === '' || value === null || value === undefined) missing.push(field.label);
  });

  section.extraFields.filter(f => f.required).forEach(field => {
    const value = readCell(sheet, cols, row, field.header);
    if (value === '' || value === null || value === undefined) missing.push(field.label);
  });

  return missing;
}

/** attached / awaiting / none required, from what actually arrived. */
function receiptStateFor(section, sheet, cols, row) {
  if (!section.fileColumns.length) return RECEIPT_STATE.notRequired;

  const present = section.fileColumns.filter(fileCol => {
    const value = readCell(sheet, cols, row, fileCol.header);
    return value !== '' && value !== null && value !== undefined;
  });

  if (!present.length) return RECEIPT_STATE.awaiting;
  return present.length === section.fileColumns.length
    ? RECEIPT_STATE.attached
    : RECEIPT_STATE.awaiting;
}

/* =============================== Registry ================================= */

/**
 * Teach the registry about this entry's counterparty.
 *
 * This is what makes the registry self-populating: nothing is entered up
 * front, every entry contributes, so it is current by construction rather
 * than by maintenance. Without this call the Suppliers sheet would simply
 * stay empty forever.
 *
 * What gets learned is per-section, because "usually the same" is not true
 * everywhere - see registryTypeField / registryNifField in Config.js. Work
 * learns that Uber is a Taxi but never learns the Expense Reason, since the
 * same supplier serves many trips.
 *
 * Never fatal. A registry write failing must not cost you the entry, so this
 * reports rather than throws.
 *
 * Note: the registry is shared across sections, so a counterparty that
 * appears in two of them with different type fields would see its stored type
 * cleared on the conflict. In practice they do not overlap; if one ever does,
 * the fix is a per-section registry rather than a smarter merge.
 */
function learnCounterparty(section, sheet, cols, row) {
  try {
    const name = readCell(sheet, cols, row, COMMON.counterparty);
    if (!name) return null;

    const details = {};
    if (section.registryTypeField) {
      details.type = readCell(sheet, cols, row, section.registryTypeField);
    }
    if (section.registryNifField) {
      details.nif = readCell(sheet, cols, row, section.registryNifField);
    }

    return recordSupplier(name, details);

  } catch (error) {
    return { ok: false, error: error.toString() };
  }
}

/* ============================ Creation email ============================== */

/**
 * Mail sent when an entry is created — currently IVA only.
 *
 * Deliberately tied to creation rather than to a status change, so no
 * transition has a side effect beyond moving files, and re-selecting a state
 * can never re-send it.
 *
 * Only sent when the claim is actually sendable. A partial entry - Siri caught
 * the core and the receipt is still to come - would otherwise mail a claim
 * missing its Número and NIF, with no attachment, and nothing would ever send
 * the real one.
 *
 * TODO: the completion path must call this again once the receipt lands, and
 * needs a "Claim Emailed" marker column before it does, or editing a row twice
 * mails the claim twice. Until that exists, a deferred claim is sent by hand.
 */
function sendCreationEmail(section, sheet, cols, row, missing) {
  const spec = section.emailOnCreate;
  if (!spec) return null;

  // Sent once, ever. Checked first so no amount of re-running any caller can
  // produce a second claim.
  const alreadySent = readCell(sheet, cols, row, CLAIM_EMAILED_COLUMN);
  if (alreadySent) {
    return { ok: false, skipped: true, reason: `already emailed ${alreadySent}` };
  }

  const incomplete = missing || missingFields(section, sheet, cols, row);
  if (incomplete.length) {
    return { ok: false, deferred: true, reason: `incomplete: missing ${incomplete.join(', ')}` };
  }

  if (spec.attachReceipt &&
      readCell(sheet, cols, row, COMMON.receiptState) !== RECEIPT_STATE.attached) {
    return { ok: false, deferred: true, reason: 'receipt not attached yet' };
  }

  try {
    const recipient = PropertiesService.getScriptProperties()
      .getProperty(spec.recipientProperty);
    if (!recipient) {
      return { ok: false, error: `${spec.recipientProperty} not set in Script Properties` };
    }

    const who = readCell(sheet, cols, row, COMMON.counterparty);
    const amount = readCell(sheet, cols, row, COMMON.amount);
    const currency = readCell(sheet, cols, row, COMMON.currency);

    // The category goes in the subject where there is one, because it is what
    // the claim is filed under: v1's work subject led with the trip, and
    // "Uber 12.50" without it does not tell you which trip to bill.
    const category = section.category
      ? (readCell(sheet, cols, row, section.category.header) || '').toString().trim()
      : '';
    const subject =
      `${section.label}: ${category ? category + ' - ' : ''}${who} ${amount} ${currency}`.trim();

    const lines = [`${section.label} entry created.`, ''];
    [COMMON.date, COMMON.counterparty, COMMON.amount, COMMON.currency].forEach(header => {
      lines.push(`${header}: ${readCell(sheet, cols, row, header)}`);
    });
    section.extraFields.forEach(field => {
      lines.push(`${field.label}: ${readCell(sheet, cols, row, field.header)}`);
    });

    const options = {};
    if (spec.attachReceipt) {
      const attachments = [];
      section.fileColumns.forEach(fileCol => {
        const fileId = extractFileId(readCell(sheet, cols, row, fileCol.header));
        if (fileId) attachments.push(DriveApp.getFileById(fileId).getBlob());
      });

      // Receipt State says a URL is present, not that it resolves. A dead link
      // would otherwise mail a claim with nothing attached and report ok.
      if (!attachments.length) {
        return { ok: false, deferred: true, reason: 'no attachable document found' };
      }
      options.attachments = attachments;
    }

    // MailApp, not GmailApp. Both send as you, but GmailApp asks for
    // https://mail.google.com/ - full read and write of the whole mailbox -
    // whereas MailApp needs only script.send_mail. This code has no business
    // being able to read your email (principle 6, least privilege).
    MailApp.sendEmail(recipient, subject, lines.join('\n'), options);

    // Stamped only after the send succeeded, so a failure leaves the claim
    // genuinely unsent and re-sendable rather than silently marked done
    writeCell(sheet, cols, row, CLAIM_EMAILED_COLUMN, new Date());

    return { ok: true, recipient: recipient, subject: subject };

  } catch (error) {
    // Reported, never fatal: the entry exists and must not be lost because
    // mail failed.
    return { ok: false, error: error.toString() };
  }
}

/**
 * Send a claim whose document arrived after the entry was made.
 *
 * This is the other half of "email when the file is uploaded". An entry created
 * from Siri has no receipt yet, so its claim defers; when the receipt is added
 * this fires it. Re-running the same gate rather than a second implementation of
 * it, so late claims cannot behave differently from prompt ones.
 *
 * Safe to call repeatedly and on any row: the Claim Emailed stamp stops a second
 * send, and a still-incomplete row simply defers again. Nothing here decides
 * whether to send - sendCreationEmail does, exactly as at creation.
 *
 * The completion step and the edit path both call this. It is also runnable by
 * hand for a row whose receipt you attached in the sheet.
 */
function sendPendingClaim(sectionKey, sheetRow) {
  const section = getSection(sectionKey);
  if (!section.emailOnCreate) {
    return { ok: false, error: `${sectionKey} does not send a claim email` };
  }

  const sheet = getSheet(section);
  const row = resolveDataRow(sheet, sheetRow);
  const cols = resolveColumns(sheet);

  // Recomputed from the sheet, never taken from a caller
  writeCell(sheet, cols, row, COMMON.receiptState, receiptStateFor(section, sheet, cols, row));

  const result = sendCreationEmail(section, sheet, cols, row, null);
  Logger.log(`${section.sheet} row ${row}: pending claim -> ${JSON.stringify(result)}`);
  return result;
}

/* ========================== More-info request ============================= */

/** Script Property holding the address that "more info needed" mail goes to. */
const COMPLETION_RECIPIENT_PROPERTY = 'COMPLETION_EMAIL_RECIPIENT';

/** Script Property holding the /exec URL of the versioned deployment. */
const WEB_APP_URL_PROPERTY = 'WEB_APP_URL';

/**
 * The deployed web app's URL, or '' if it cannot be established.
 *
 * The Script Property wins, because it can name the PINNED deployment — the one
 * the phone and every mailed link should reach. `getService().getUrl()` is only a
 * fallback: it needs no setup, but it returns whichever endpoint the running
 * context belongs to, and anything run from the editor belongs to /dev, which
 * only opens for accounts that can edit the script. Wrapped because the manifest
 * pins its scopes and no URL lookup should ever be the reason mail stops going.
 */
function webAppUrl() {
  const configured = (PropertiesService.getScriptProperties()
    .getProperty(WEB_APP_URL_PROPERTY) || '').toString().trim();
  if (configured) return mailableUrl(configured);

  try {
    return mailableUrl((ScriptApp.getService().getUrl() || '').toString());
  } catch (error) {
    Logger.log(`Could not establish the web app URL: ${error}`);
    return '';
  }
}

/**
 * Refuse a /dev URL, however it was arrived at.
 *
 * `/dev` serves HEAD and opens only for accounts that can EDIT the script. A
 * mailed link is opened later, on whatever device is to hand, signed in as
 * whichever Google account happens to be that browser's default — so a /dev link
 * is not merely suboptimal, it is guaranteed to fail there. And it fails at
 * Drive's layer, before doGet, so it reads as "Sorry, unable to open the file at
 * this time" rather than as anything to do with this project. That is a bad half
 * hour, and it happened.
 *
 * Returning '' sends the caller to its fallback, which is the spreadsheet row: a
 * worse destination that actually opens beats a better one that cannot.
 */
function mailableUrl(url) {
  const clean = (url || '').toString().trim();
  if (!clean) return '';
  if (/\/dev(\?|#|$)/.test(clean)) {
    Logger.log(
      `Refusing to mail a /dev link (${clean}) - it only opens for accounts that ` +
      `can edit the script. Set ${WEB_APP_URL_PROPERTY} to the /exec URL of the ` +
      `versioned deployment.`
    );
    return '';
  }
  return clean;
}

/**
 * Where the completion mail sends you: the form, open on that entry.
 *
 * The sheet row was fine for correcting a word and wrong for what this mail is
 * usually about — a missing document, a date, a category value. The sheet offers
 * no upload, no date picker and no dropdown; the form offers all three and
 * already edits in place, so only this URL had to change.
 *
 * The entry is named by its TIMESTAMP, not its row number. Archiving a row shifts
 * every row beneath it up by one, and a completion mail is precisely the thing
 * that sits in an inbox for days — a row number would then open a DIFFERENT
 * entry, prefilled and looking entirely plausible. Both sides read the timestamp
 * back out of the cell that was just written, so they compare the same stored
 * value rather than two formattings of it.
 *
 * Falls back to the spreadsheet row when no web app URL is known: a reminder that
 * lands somewhere useful beats one that does not land.
 *
 * Carries the accounts problem with it — /exec takes the browser's default
 * account and cannot be pointed at one by URL, so on a device where the v2
 * account is not the default this link opens as the wrong person. That is the
 * same unresolved constraint as the phone, not something this can fix.
 */
/**
 * Where to go and find the document, for the completion mail.
 *
 * Returns null when there is nothing useful to say — no medium recorded, or no
 * document outstanding — because a line that appears on every reminder saying
 * something vague is a line you stop reading, and this one has to still be worth
 * reading on the day it matters.
 *
 * The staging folder is where Genius Scan and saved mail attachments land. It is
 * linked rather than described, and the link carries `authuser` for the usual
 * reason: several accounts are signed in on every device, and without it the
 * default account answers and the failure reads as a missing folder.
 */
function documentLocationHint(section, sheet, cols, row, needsDocument) {
  if (!needsDocument) return null;

  // The medium refines step one; it does not gate it. A document is still
  // wanted whether or not anyone recorded where it came from, so this returns
  // a hint with no sentence rather than nothing at all.
  const medium = sectionReceiptMedium(section);
  const value = medium
    ? (readCell(sheet, cols, row, medium.header) || '').toString().trim()
    : '';

  const where = {};
  where[RECEIPT_MEDIUM.electronic] = 'It was electronic — save it out of your mail.';
  where[RECEIPT_MEDIUM.physical] = 'It was on paper — scan it with Genius Scan.';
  where[RECEIPT_MEDIUM.both] = 'There was paper and an email — either will do.';

  const sentence = where[value] || null;

  const folderId = PropertiesService.getScriptProperties()
    .getProperty(STAGING_FOLDER_PROPERTY);
  if (!folderId) return { sentence: sentence, folderUrl: null };

  // The effective user is the account the script runs as, which is also who
  // this mail goes to — so it is the right account to open the folder as.
  let viewer = '';
  try {
    viewer = (Session.getEffectiveUser().getEmail() || '').toString();
  } catch (error) {
    viewer = '';
  }

  const url = `https://drive.google.com/drive/folders/${folderId}`;
  return {
    sentence: sentence,
    folderUrl: viewer ? `${url}?authuser=${encodeURIComponent(viewer)}` : url
  };
}

function completionLink(section, sheet, cols, row) {
  const sheetLink = `${getSpreadsheet().getUrl()}` +
    `#gid=${sheet.getSheetId()}&range=A${row}`;

  const base = webAppUrl();
  const key = sectionKeyOf(section);
  const stamp = readCell(sheet, cols, row, COMMON.timestamp);
  if (!base || !key || !(stamp instanceof Date)) return sheetLink;

  return `${base}?section=${encodeURIComponent(key)}&t=${stamp.getTime()}`;
}

/**
 * Tell you an entry arrived incomplete.
 *
 * This is the safety net the whole partial-entry design rests on: a row is
 * written even when it is missing fields, so something has to say so or the gap
 * is only ever visible by scrolling the sheet. A mishearing then costs one tap
 * rather than a lost capture.
 *
 * Fires when required fields are blank OR a document is still awaited - both
 * mean "come back to this", which is the only thing the mail is for.
 *
 * Goes to its own address, not to the IVA claims address: this is a note to
 * yourself, and it must never land in front of whoever processes claims.
 *
 * The link points at the row in the spreadsheet. Once the web form exists this
 * becomes the completion link that opens the form on that row - the thing
 * Google Forms could never do - and only the URL built here has to change.
 *
 * Never fatal. The entry exists; failing to send a reminder must not undo it.
 */
function sendCompletionRequest(section, sheet, cols, row, missing, receiptState) {
  const needsDocument = receiptState === RECEIPT_STATE.awaiting;
  if (!missing.length && !needsDocument) return null;

  try {
    const recipient = PropertiesService.getScriptProperties()
      .getProperty(COMPLETION_RECIPIENT_PROPERTY);
    if (!recipient) {
      return { ok: false, error: `${COMPLETION_RECIPIENT_PROPERTY} not set in Script Properties` };
    }

    const outstanding = missing.slice();
    if (needsDocument) {
      section.fileColumns.forEach(fileCol => {
        const value = readCell(sheet, cols, row, fileCol.header);
        if (value === '' || value === null || value === undefined) outstanding.push(fileCol.label);
      });
    }

    const who = readCell(sheet, cols, row, COMMON.counterparty) || 'unknown';
    const subject = `${section.label}: more info needed - ${who}`;

    const link = completionLink(section, sheet, cols, row);

    const captured = [];
    [COMMON.date, COMMON.counterparty, COMMON.amount, COMMON.currency, COMMON.notes]
      .forEach(header => {
        const value = readCell(sheet, cols, row, header);
        if (value !== '' && value !== null && value !== undefined) {
          captured.push({ label: header, value: value });
        }
      });

    const hint = documentLocationHint(section, sheet, cols, row, needsDocument);

    /*
     * TWO NUMBERED STEPS when a document is outstanding, because it is two
     * separate jobs done in two different apps, and the second cannot be
     * started until the first is finished. Written as one list of things
     * needed, the "get the receipt into Drive" part reads as a note rather
     * than as work, and you open the form, find you have nothing to attach,
     * and close it again.
     *
     * When only fields are blank there is no first job, so there are no steps
     * — just the link. Numbering a single step is a form of noise.
     */
    const lines = [
      `A ${section.label} entry was created but is not finished.`,
      '',
      'Still needed:'
    ];
    outstanding.forEach(label => lines.push(`  - ${label}`));
    lines.push('');

    if (needsDocument) {
      lines.push(`Step 1 — get the document into ${STAGING_FOLDER_NAME}`);
      if (hint && hint.sentence) lines.push(`  ${hint.sentence}`);
      lines.push('  Save it there from your mail, or scan straight to it.');
      if (hint && hint.folderUrl) lines.push(`  ${hint.folderUrl}`);
      lines.push('', 'Step 2 — finish the entry');
      lines.push('  Open the form and choose the document from the list.');
    }

    lines.push('', 'What was captured:');
    captured.forEach(item => lines.push(`  ${item.label}: ${item.value}`));
    lines.push('', `Finish it here: ${link}`);

    /*
     * Sent as HTML with a plain-text alternative, rather than text alone.
     *
     * The link is ~130 characters and plain-text mail wraps at about 78, which
     * leaves the receiving client guessing where a URL ends. As an href it is one
     * unbreakable attribute instead. A bare long URL in an otherwise plain message
     * is also a fair imitation of spam, and this mail was landing in junk - a
     * multipart message with real structure is at least a less suspicious shape,
     * though no message can guarantee its own delivery to an inbox.
     *
     * Every interpolated value goes through escapeHtml. These come from sheet
     * cells, which is the same untrusted text safeCellValue neutralises on the way
     * in: a supplier called "Smith & Sons <Lda>" would otherwise eat the rest of
     * the line, and that is the mild version.
     */
    const html = [
      `<div style="font:15px/1.5 -apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif">`,
      `<p>A ${escapeHtml(section.label)} entry was created but is not finished.</p>`,
      `<p><strong>Still needed</strong></p>`,
      `<ul>${outstanding.map(label => `<li>${escapeHtml(label)}</li>`).join('')}</ul>`,

      // Same two steps as the plain-text alternative. The two bodies must say
      // the same thing: which one a client renders is not ours to choose.
      needsDocument
        ? `<p><strong>Step 1 — get the document into ${escapeHtml(STAGING_FOLDER_NAME)}</strong><br>` +
          (hint && hint.sentence ? `${escapeHtml(hint.sentence)}<br>` : '') +
          `Save it there from your mail, or scan straight to it.` +
          (hint && hint.folderUrl
            ? `<br><a href="${escapeHtml(hint.folderUrl)}">Open the ${escapeHtml(STAGING_FOLDER_NAME)} folder</a>`
            : '') +
          `</p><p><strong>Step 2 — finish the entry</strong><br>` +
          `Open the form and choose the document from the list.</p>`
        : '',

      `<p><strong>What was captured</strong></p>`,
      `<ul>${captured.map(item =>
        `<li>${escapeHtml(item.label)}: ${escapeHtml(item.value)}</li>`).join('')}</ul>`,
      `<p><a href="${escapeHtml(link)}" style="display:inline-block;padding:10px 16px;`,
      `background:#1a73e8;color:#fff;border-radius:8px;text-decoration:none;`,
      `font-weight:600">${needsDocument ? 'Step 2 — finish this entry' : 'Finish this entry'}</a></p>`,
      `</div>`
    ].join('');

    MailApp.sendEmail(recipient, subject, lines.join('\n'), { htmlBody: html });
    return { ok: true, recipient: recipient, outstanding: outstanding, link: link };

  } catch (error) {
    return { ok: false, error: error.toString() };
  }
}

/* ============================ Initialisation ============================== */

/**
 * Fill in bookkeeping, name and file the documents, send any creation mail.
 * Shared by every intake path.
 */
function initializeEntry(section, sheet, row, source) {
  const cols = resolveColumns(sheet);

  if (!readCell(sheet, cols, row, COMMON.timestamp)) {
    writeCell(sheet, cols, row, COMMON.timestamp, new Date());
  }
  writeCell(sheet, cols, row, COMMON.source, source || 'manual');

  // Initial state. Its date is taken from the entry when supplied, because a
  // state like Invoiced is normally backdated; today is only a fallback.
  const first = section.states[0];
  writeCell(sheet, cols, row, COMMON.status, first.name);
  if (first.dateColumn && !readCell(sheet, cols, row, first.dateColumn)) {
    writeCell(sheet, cols, row, first.dateColumn, today());
  }

  const receiptState = receiptStateFor(section, sheet, cols, row);
  writeCell(sheet, cols, row, COMMON.receiptState, receiptState);

  const documents = nameAndFileDocuments(section, sheet, cols, row, 0);
  const renames = documents.renames;
  const files = documents.files;
  const warnings = missingFields(section, sheet, cols, row);
  const registry = learnCounterparty(section, sheet, cols, row);
  const email = sendCreationEmail(section, sheet, cols, row, warnings);
  const completion = sendCompletionRequest(section, sheet, cols, row, warnings, receiptState);

  if (warnings.length) {
    Logger.log(`${section.sheet} row ${row}: incomplete — missing ${warnings.join(', ')}`);
  }

  return {
    ok: true,
    section: section.sheet,
    row: row,
    state: first.name,
    receiptState: receiptState,
    warnings: warnings,
    renames: renames,
    files: files,
    fileErrors: files.filter(f => !f.ok).concat(renames.filter(r => !r.ok)),
    registry: registry,
    email: email,
    completionRequest: completion
  };
}

/**
 * Give every document its name from the row, then file it for a given state.
 *
 * The base name is rebuilt from the row's CURRENT values, so this is also what
 * makes an edit correct: change the date, the counterparty or the amount and
 * the filenames follow. applyFileState then appends the state suffix chain and
 * moves the file, so the whole naming rule lives in one place whether the row
 * was just created, just edited, just changed state, or just had its supplier
 * renamed underneath it.
 *
 * folderName is passed straight through to applyFileState; see the note there.
 * It is only ever set when repairing an archived row.
 */
function nameAndFileDocuments(section, sheet, cols, row, targetIndex, folderName) {
  const renames = [];
  section.fileColumns.forEach(fileCol => {
    const fileId = extractFileId(readCell(sheet, cols, row, fileCol.header));
    if (!fileId) return;
    try {
      const file = DriveApp.getFileById(fileId);
      const ext = splitExtension(file.getName()).ext;
      file.setName(`${buildBaseFilename(sheet, cols, row, fileCol)}${ext}`);
      renames.push({ column: fileCol.header, ok: true, name: file.getName() });
    } catch (error) {
      renames.push({ column: fileCol.header, ok: false, error: error.toString() });
    }
  });

  return {
    renames: renames,
    files: applyFileState(section, sheet, cols, row, targetIndex, folderName)
  };
}

/* ================================= Intake ================================= */

/**
 * Append a new entry. THE only way a row is born.
 *
 * The custom form, Siri and OCR are all callers of this — they differ in where
 * the field values came from, and in nothing else. Adding a fourth intake
 * means writing a caller, not touching anything here.
 *
 * A row is written even when it is incomplete, deliberately: a partial entry
 * you can see and finish beats a capture that failed. Incompleteness comes
 * back as ok:false with the missing labels, so the caller can send you to the
 * completion step — but the row exists either way, and a caller must not
 * retry on ok:false or it will duplicate the entry.
 *
 * @param {string} sectionKey key into SECTIONS
 * @param {Object} fields     keyed by COLUMN HEADER, matching the sheet
 * @param {string} source     'form' | 'siri' | 'ocr' | 'manual'
 */
function createEntry(sectionKey, fields, source) {
  const section = getSection(sectionKey);
  const sheet = getSheet(section);
  const cols = resolveColumns(sheet);

  const supplied = fields || {};
  const headers = Object.keys(supplied);

  // Every header checked BEFORE anything is written. Rejecting one part way
  // through the loop would leave a half-written row with no status and no
  // bookkeeping - and that row is then the one getLastRow() reports, so the
  // next entry would land on top of it.
  headers.forEach(header => {
    if (!cols[header]) throw new Error(`Unknown column "${header}" for ${sectionKey}`);
  });

  const width = sheet.getLastColumn();

  // Appending is the only moment two callers can collide, so the lock covers
  // exactly that and nothing slow. Written as one range rather than a setValue
  // per field, so the row is never briefly half-present.
  const row = withLock(() => {
    const target = sheet.getLastRow() + 1;
    const values = new Array(width).fill('');

    // safeCellValue matters here above all: this is the one path that accepts
    // field values from outside, so a counterparty of "=IMPORTXML(...)" must be
    // stored as text rather than executed.
    headers.forEach(header => {
      values[cols[header] - 1] = safeCellValue(supplied[header]);
    });

    sheet.getRange(target, 1, 1, width).setValues([values]);

    // Committed before the lock is released, or the next caller's getLastRow()
    // still returns the old value and picks the same row.
    SpreadsheetApp.flush();
    return target;
  });

  const result = initializeEntry(section, sheet, row, source || 'manual');

  if (result.warnings.length) {
    result.ok = false;
    result.error = `Missing required: ${result.warnings.join(', ')}`;
  }
  return result;
}
