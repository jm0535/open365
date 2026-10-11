/* Tiny pub/sub used to fan store + UI events out to whichever view is mounted. */
const map = new Map();
export const bus = {
  on(type, fn) {
    if (!map.has(type)) map.set(type, new Set());
    map.get(type).add(fn);
    return () => map.get(type)?.delete(fn);
  },
  emit(type, detail) {
    for (const fn of map.get(type) || []) { try { fn(detail); } catch (e) { console.warn(e); } }
    for (const fn of map.get('*') || []) { try { fn({ type, detail }); } catch (e) { console.warn(e); } }
  }
};
