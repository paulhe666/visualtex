// Convert validated numerical paint only. Never accept an arbitrary CSS value.
export function rgbHex(value: unknown): string {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Invalid RGB paint');
  const object = value as Record<string, unknown>;
  if (Object.keys(object).sort().join(',') !== 'b,g,r') throw new Error('Unexpected RGB paint fields');
  const components = [object.r, object.g, object.b];
  if (!components.every(n => typeof n === 'number' && Number.isInteger(n) && n >= 0 && n <= 255)) {
    throw new Error('RGB paint channels must be integers between 0 and 255');
  }
  return '#' + components.map(n => (n as number).toString(16).padStart(2, '0')).join('');
}
