window.VMCRYPTO = (function () {
  const b64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
  async function keyFor(pw, saltB64, iter) {
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: b64(saltB64), iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
  }
  async function open(url, pw) {
    const r = await fetch(url, { cache: 'no-cache' });
    if (!r.ok) throw new Error('fetch');
    const env = await r.json();
    const key = await keyFor(pw, env.salt, env.iter);
    const buf = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: b64(env.iv) }, key, b64(env.ct));
    return JSON.parse(new TextDecoder().decode(buf));
  }
  return { open };
})();
