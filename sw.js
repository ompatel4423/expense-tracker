const CACHE_NAME = 'expense-tracker-v4';
const STATIC_ASSETS = [
  '/',
  '/app_logo.jpg',
  '/manifest.json'
];

// Safari / WebKit throws "error response served by service worker has redirections"
// if event.respondWith receives a Response with redirected: true.
// This sanitizer generates a clean new Response without redirect flags.
async function cleanRedirectedResponse(response) {
  if (!response || !response.redirected) {
    return response;
  }
  const body = await response.blob();
  return new Response(body, {
    status: response.status,
    statusText: response.statusText,
    headers: response.headers
  });
}

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      for (const asset of STATIC_ASSETS) {
        try {
          const res = await fetch(asset, { redirect: 'follow' });
          const clean = await cleanRedirectedResponse(res);
          if (clean && clean.ok) {
            await cache.put(asset, clean);
          }
        } catch (e) {}
      }
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  const url = event.request.url;

  // Let Google Apps Script API calls bypass the service worker completely
  if (url.includes('script.google.com') || url.includes('googleusercontent.com')) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then(async (cachedResponse) => {
      if (cachedResponse) {
        return cleanRedirectedResponse(cachedResponse);
      }

      try {
        const networkResponse = await fetch(event.request);
        const cleanResponse = await cleanRedirectedResponse(networkResponse);

        // Cache static CDN files like GSAP and fonts
        if (cleanResponse && cleanResponse.ok && (url.includes('cdnjs.cloudflare.com') || url.includes('fonts.gstatic.com'))) {
          const clone = cleanResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, clone);
          });
        }

        return cleanResponse;
      } catch (err) {
        if (event.request.mode === 'navigate') {
          const fallback = await caches.match('/');
          if (fallback) {
            return cleanRedirectedResponse(fallback);
          }
        }
        throw err;
      }
    })
  );
});
