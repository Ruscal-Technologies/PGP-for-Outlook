import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests/**/*.test.js'],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html', 'lcov'],
      // Scoped to the modules this suite actually targets (web/js/pgp/* +
      // wkd.js) — the four Office.js UI entry points have no tests yet (see
      // CLAUDE.md's "Known gaps"), so including them here would just produce
      // a misleadingly low number rather than a meaningful signal.
      include: ['web/js/pgp/**/*.js', 'web/js/wkd.js'],
    },
  },
});
