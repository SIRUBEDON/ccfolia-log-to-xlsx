Security regression tests
=========================

Run with Node.js 24: `npm ci --ignore-scripts` followed by `npm test`.

Tests exercise the application DOM, file imports and exports using jsdom.
All test resource requests are answered locally; no production site is modified.
These tests do not replace visual checks in a real browser or spreadsheet application.
