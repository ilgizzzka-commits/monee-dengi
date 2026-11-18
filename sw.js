'use strict';

/* Деньги Monee — service worker: чтобы приложение открывалось без интернета.

   ВАЖНО: при КАЖДОЙ выкладке повышай VERSION (v1 → v2 → v3 …).
   Иначе телефон так и будет показывать старую версию из своего запаса.

   Что этот файл сохраняет на телефоне: только «оболочку» приложения —
   страницу, оформление, код, иконки и шрифт. Ответы Google (вход, Диск)
   он не сохраняет никогда. Последние посчитанные цифры хранит само
   приложение (app.js, localStorage), «Выйти» их стирает. */

const VERSION = 'v2';
const CACHE = 'monee-dengi-' + VERSION;
const FONTS = 'monee-dengi-fonts';
const DATA = 'monee-dengi-data';

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
        .filter((k) => k.startsWith('monee-dengi-') && k !== CACHE && k !== FONTS && k !== DATA)
        .map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // data.json (курсы ЦБ и календарь, открытые данные): сначала из сети, чтобы были свежие;
  // без интернета — последняя сохранённая копия.
  if (url.origin === self.location.origin && url.pathname.endsWith('/data.json')) {
    event.respondWith(
      fetch(req).then((res) => {
        if (res.ok) { const copy = res.clone(); caches.open(DATA).then((c) => c.put('data.json', copy)); }
        return res;
      }).catch(() => caches.open(DATA).then((c) => c.match('data.json')).then((hit) => hit || Response.error()))
    );
    return;
  }

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
