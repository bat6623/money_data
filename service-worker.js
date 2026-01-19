const CACHE_NAME = 'asset-stats-v1';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './us.html',
    './styles.css',
    './script.js',
    './icon.svg',
    'https://cdn.jsdelivr.net/npm/remixicon@3.5.0/fonts/remixicon.css',
    'https://cdn.jsdelivr.net/npm/chart.js',
    'https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0',
    'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'
];

// Install Event: Cache core assets
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('Opened cache');
                return cache.addAll(ASSETS_TO_CACHE);
            })
    );
});

// Activate Event: Clean up old caches
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cacheName) => {
                    if (cacheName !== CACHE_NAME) {
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
});

// Fetch Event: Stale-While-Revalidate Strategy
self.addEventListener('fetch', (event) => {
    // Skip cross-origin requests like Google Sheets CSV if needed, 
    // but Stale-While-Revalidate is generally safe for CORS if opaque response is okay.
    // However, for API data (CSV), we might prefer Network First.

    const url = new URL(event.request.url);

    // Dynamic CSV Data -> Network First (Don't cache or simple fallback)
    if (url.href.includes('docs.google.com') || url.href.includes('output=csv')) {
        event.respondWith(
            fetch(event.request).catch(() => {
                // Return empty content or cached fallback if available?
                // For now, simple network only for data to avoid stale data confusion on re-open
                return new Response('', { status: 408, statusText: 'Offline' });
            })
        );
        return;
    }

    // Static Assets -> Stale-While-Revalidate
    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            const fetchPromise = fetch(event.request).then((networkResponse) => {
                // Update cache
                if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
                    const responseToCache = networkResponse.clone();
                    caches.open(CACHE_NAME).then((cache) => {
                        cache.put(event.request, responseToCache);
                    });
                }
                return networkResponse;
            });

            // Return cached response immediately if available, else wait for network
            return cachedResponse || fetchPromise;
        })
    );
});
