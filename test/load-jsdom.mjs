import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import fs from 'node:fs';
import path from 'node:path';

const here = path.dirname(fileURLToPath(import.meta.url));

/**
 * Resolve jsdom for the test suites only.
 * Order: JSDOM_PATH, the repo's node_modules (npm ci in CI), then the
 * historical /tmp/dom install used by this sandbox. Never a runtime import.
 */
export async function loadJsdom() {
  const env = process.env.JSDOM_PATH;
  const roots = [
    env,
    path.join(here, '../node_modules/jsdom'),
    '/tmp/dom/node_modules/jsdom'
  ].filter(Boolean);

  for (const root of roots) {
    const file = root.endsWith('.js') ? root : path.join(root, 'lib/api.js');
    if (fs.existsSync(file)) return import(pathToFileURL(file).href);
  }

  try {
    const require = createRequire(path.join(here, '../package.json'));
    return import(pathToFileURL(require.resolve('jsdom')).href);
  } catch {
    /* try a bare specifier next */
  }

  try {
    return await import('jsdom');
  } catch (err) {
    const error = new Error(
      'jsdom is required for the UI and OOXML suites. Run `npm install` (devDependency only) or set JSDOM_PATH. It is not a runtime dependency.'
    );
    error.cause = err;
    throw error;
  }
}
