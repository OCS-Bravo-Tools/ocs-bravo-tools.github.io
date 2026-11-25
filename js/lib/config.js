/**
 * The shared collector for cadre mode: Collector v2, one Google Sheet that the maintainer
 * sets up once. Its address lives only here, never in a link, so a new deployment or a
 * hand-off changes the site and nothing else. Set from the maintainer's setup (SETUP.md §1).
 *
 * Tests on this computer only (localhost, 127.0.0.1 or [::1]) can point the pages at the
 * mock with ?collector=http://127.0.0.1:PORT/ in the page address. Nowhere else is it read.
 * A page opened here without it talks to the real sheet.
 */
export const COLLECTOR = 'https://script.google.com/macros/s/AKfycbyH7lo-WKCXrGa-Husvyk8ffftJ4329LCA-0lk9gQb6CXl3H6sNtJR-jq6hGoHN1o1UxQ/exec';

const LOCAL = /^(localhost|127\.0\.0\.1|\[::1\])$/;

/** The collector this page talks to: the built-in one, or the test mock on this computer. */
export function collectorUrl() {
  if (typeof location !== 'undefined' && LOCAL.test(location.hostname)) {
    const q = new URLSearchParams(location.search).get('collector');
    if (q && /^http:\/\/(127\.0\.0\.1|localhost):\d+\/?$/.test(q)) return q;
  }
  return COLLECTOR;
}

/** The query part the links carry on a test page (so phones in tests reach the mock), else ''. */
export function testQuery() {
  if (typeof location === 'undefined' || !LOCAL.test(location.hostname)) return '';
  const q = new URLSearchParams(location.search).get('collector');
  return q && collectorUrl() === q ? '?collector=' + encodeURIComponent(q) : '';
}
