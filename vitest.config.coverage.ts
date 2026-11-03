import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

/**
 * Coverage runs against `src/` instead of the built package, so the numbers describe the
 * source. The normal `npm test` suite intentionally imports the built package: a packaging
 * mistake should fail tests, not slip through.
 */
export default defineConfig({
  resolve: {
    alias: {
      'excel-lite': fileURLToPath(new URL('./src/index.ts', import.meta.url)),
    },
  },
  test: {
    include: ['test/**/*.test.ts'],
    environment: 'node',
    testTimeout: 30_000,
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      exclude: ['src/cli.ts', 'src/index.ts', 'src/types.ts'],
      reporter: ['text', 'lcov', 'json-summary'],
      thresholds: {
        lines: 80,
        functions: 80,
        branches: 75,
      },
    },
  },
});
