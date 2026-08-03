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
documented with screenshots on the pull request page.

The agent sandbox cannot upload attachments to GitHub, so the images live on a
dedicated branch instead:

1. Push the captures to the `pr-screenshots` branch, inside a folder named after
   the pull request number — for example `39/card-grid-light-after.jpg`. The
   branch shares no history with `main`, holds nothing but these folders, and is
   never merged or deleted; removing it would break the image links in older
   pull requests.
2. Embed them in the PR description as inline markdown images pointing at that
   branch:
   `https://raw.githubusercontent.com/yclkvnc/home-inventory/pr-screenshots/<pr-number>/<file>`
3. Never commit image files to the working branch — its diff must contain only
   the source files the change touches.
4. Keep the files small: JPEG at a sensible quality, one viewport per file.

### Before and after

For every scenario, capture the same view twice: once with the base branch
checked out and once with the change applied. Present the pair side by side in a
markdown table so the difference is obvious:

| Before | After |
|---|---|
| ![before](…) | ![after](…) |

Group the scenarios under headings, pair light and dark theme, and wrap long
sets in collapsible `<details>` blocks. Skip the "before" shot only when the
view did not exist previously.

Alongside the images, state in the description which browser engine produced
them, and note anything a still image cannot show — transitions, hover states,
or `prefers-reduced-motion` behaviour.
