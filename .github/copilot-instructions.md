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

Verify every URL returns 200, and confirm the final diff contains no image
files, before finishing.

The description you publish with the images must be the final one. If you update
the description again afterwards, carry the same image markdown forward — never
replace it with a reference to earlier versions, since the description has no
visible history and the images would simply disappear.

Keep the files small: JPEG at a sensible quality, one viewport per file.

### Layout

Group the captures under headings by scenario, pair light and dark theme in a
two-column table, and wrap each scenario in a `<details>` block — leave the
first one `open`.

Include a "before" column, captured against the base branch, whenever the view
existed beforehand.

State which browser engine produced them, and note anything a still image cannot
show — transitions, hover states, or `prefers-reduced-motion` behaviour.
