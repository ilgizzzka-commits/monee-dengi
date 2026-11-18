'use strict';

/* Деньги Monee — service worker: чтобы приложение открывалось без интернета.

   ВАЖНО: при КАЖДОЙ выкладке повышай VERSION (v1 → v2 → v3 …).
   Иначе телефон так и будет показывать старую версию из своего запаса.

   Что сохраняется на телефоне: только «оболочка» приложения — страница,
   оформление, код, иконки и шрифт. Цифры из таблицы и всё, что идёт
   через Google (вход, Диск), НЕ сохраняются никогда. */

const VERSION = 'v1';
const CACHE = 'monee-dengi-' + VERSION;
const FONTS = 'monee-dengi-fonts';

const SHELL = [
  './',
  'index.html',
  'style.css',
  'app.js',
  'manifest.webmanifest',
  'icon.svg',
  'icon-192.png',
  'icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(SHELL.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys
        .filter((k) => k.startsWith('monee-dengi-') && k !== CACHE && k !== FONTS)
        .map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Своё: сначала из запаса (быстро и без сети), иначе из сети.
  if (url.origin === self.location.origin) {
    if (req.mode === 'navigate') {
      event.respondWith(
        caches.match('index.html', { cacheName: CACHE })
          .then((hit) => hit || fetch(req))
      );
      return;
    }
    event.respondWith(
      caches.match(req, { cacheName: CACHE, ignoreSearch: true })
        .then((hit) => hit || fetch(req))
    );
    return;
  }

  // Шрифт Onest с Google Fonts: берём из запаса, в фоне обновляем.
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      caches.open(FONTS).then((c) => c.match(req).then((hit) => {
        const net = fetch(req)
          .then((res) => { if (res.ok || res.type === 'opaque') c.put(req, res.clone()); return res; })
          .catch(() => hit);
        return hit || net;
      }))
    );
    return;
  }

  // Всё остальное (вход Google, Google Диск с цифрами) — только напрямую, без запаса.
});
