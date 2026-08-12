/**
 * HelpfulForms v2 — the Siri intake shim.
 *
 * THIS IS THE WHOLE PROJECT. There is deliberately nothing else in it.
 *
 * WHY IT EXISTS AT ALL. Siri needs an endpoint an anonymous caller can reach,
 * and `webapp.access` is a property of the PROJECT, not of the deployment.
 * Setting it to ANYONE_ANONYMOUS on the main project would open the web UI to
 * the internet — and worse than that, anonymous access blanks
 * Session.getActiveUser() for EVERYONE INCLUDING THE OWNER, so `uiAccessCheck()`
 * would begin refusing me and no check inside `doGet` could tell the two cases
 * apart. A second project is the only version of this that stays safe.
 *
 * WHY IT IS EMPTY. Everything the endpoint does — the key check, the field
 * whitelist, the registry lookup, the row — lives in `v2/Siri.js`, in the main
 * project, reached from here as a library under the symbol `HF`. That keeps ONE
 * copy of the code and ONE set of Script Properties, and it means the entire
 * endpoint is covered by `npm run v2:test`, which runs the real source in node.
 * The delegation below is the only part no test can reach, which is why there
 * is no logic in it to get wrong.
 *
 * So: do not add anything to this file. A check written here is a check the
 * harness cannot see. If you need behaviour, it goes in `v2/Siri.js`.
 *
 * HOW THE LIBRARY IS PINNED. `v2-siri/appsscript.json` is GENERATED — it holds
 * the main project's script id, which is why it is gitignored like every other
 * clasp identifier in this repo. Regenerate it with:
 *
 *   npm run v2:siri:manifest
 *
 * It is generated in development mode, so this shim runs whatever `npm run
 * v2:push` last put on HEAD. There is no version to bump and nothing to go
 * quietly stale.
 */

function doPost(e) {
  return HF.siriHandlePost(e);
}

/**
 * A GET is not part of the protocol, but this URL is anonymous and will be
 * visited — by a link checker, by a mistyped tap, eventually by a crawler.
 * Answering flatly is better than an Apps Script error page, which would
 * confirm the project exists and show its name.
 */
function doGet() {
  return ContentService
    .createTextOutput('Not available.')
    .setMimeType(ContentService.MimeType.TEXT);
}
