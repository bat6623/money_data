const CACHE_NAME = 'asset-stats-v3';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './us.html',
    './styles.css',
    './script.js',
    './icon.svg',
    './manifest.json',
    'https://cdn.jsdelivr.net/npm/remixicon@3.5.0/fonts/remixicon.css',
    'https://cdn.jsdelivr.net/npm/chart.js',
    'https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0',
    'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'
];

// Install Event: Cache core assets
self.addEventListener('install', (event) => {
    self.skipWaiting(); // Force activation
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
        }).then(() => self.clients.claim()) // Take control immediately
    );
});

// Fetch Event: Stale-While-Revalidate Strategy
self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);

    // Ignore Chrome Extensions and non-http(s) schemes
    if (!url.protocol.startsWith('http')) {
        return;
    }

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
