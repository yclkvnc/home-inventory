# Copilot instructions — home-inventory

## Project constraints

- No build step: plain HTML, CSS and vanilla JavaScript. No npm, no bundler, no
  CDN, no CSS framework.
- `styles.css` drives everything through custom properties. `[data-theme="dark"]`
  redefines values only — never duplicate rules.
- Data lives in one Excel table on OneDrive, reached through Microsoft Graph.
  Column headers are read at runtime; do not hardcode new columns.
- Keep `app.js` in the existing ES5-flavoured style (`var`, `function`, no
  arrow functions or template literals).

## Screenshots on pull requests

Any change that alters the UI should be verified locally with Playwright and
documented with screenshots in the pull request description.

Prefer attaching the images to the pull request directly, if a tool for that is
available to you. If it is not, say so explicitly rather than falling back
silently.

Only if no attachment path exists, use the two-commit route:

1. Commit the captures to the working branch under `docs/screenshots/` and push.
2. Embed them in the description with raw URLs pinned to that commit's full SHA:
   `https://raw.githubusercontent.com/yclkvnc/home-inventory/<full-sha>/docs/screenshots/<file>`
3. Push a second commit deleting `docs/screenshots/`, so the final diff contains
   only the source files the change touches. The images stay reachable through
   the pinned SHA.
4. Note in the description that the images are hosted from a commit outside the
   diff, and that the links will eventually break once that commit is
   unreachable.

Do not push to any branch other than the pull request's own — the sandbox is
refused with 403.

Keep the files small: JPEG at a sensible quality, one viewport per file.

### Order of writes

Write the description **once, last, and in full**. Compose the complete final
text — summary, notes, and image markdown together — and publish it as a single
update after the screenshot commits are pushed.

Never publish a description first and refine it later. The description has no
visible history: a second update silently discards whatever the first one
contained, and the images disappear with it.

If a further update is genuinely unavoidable, read the current description back
first and carry its image markdown forward verbatim into the new text. Never
replace images with a pointer to an earlier version or an earlier commit.

### Before you finish

All three must hold. Check them in this order, and fix and re-check rather than
reporting a failure:

1. Re-read the published pull request description as it currently stands on
   GitHub — not the draft you composed — and confirm every screenshot is present
   in it.
2. Request each image URL from that published description and confirm it
   returns 200.
3. Confirm the final diff contains no image files.

Say explicitly in your closing summary that you performed check 1, and how many
images you found. A UI change whose published description carries no images is
unfinished work, regardless of the state of the code.

### Layout

Group the captures under headings by scenario, pair light and dark theme in a
two-column table, and wrap each scenario in a `<details>` block — leave the
first one `open`.

Include a "before" column, captured against the base branch, whenever the view
existed beforehand.

State which browser engine produced them, and note anything a still image cannot
show — transitions, hover states, or `prefers-reduced-motion` behaviour.
