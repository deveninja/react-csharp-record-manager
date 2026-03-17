const apiCache = new Map();

export async function fetchJsonWithCache(
  url,
  { ttlMs = 30_000, forceRefresh = false } = {},
) {
  const now = Date.now();
  const cached = apiCache.get(url);

  if (!forceRefresh && cached) {
    if (cached.data && cached.expiry > now) {
      return cached.data;
    }

    if (cached.promise) {
      return cached.promise;
    }
  }

  const requestPromise = fetch(url)
    .then((response) => {
      if (!response.ok) {
        throw new Error(`Request failed for ${url}.`);
      }

      return response.json();
    })
    .then((data) => {
      apiCache.set(url, { data, expiry: Date.now() + ttlMs, promise: null });
      return data;
    })
    .catch((error) => {
      apiCache.delete(url);
      throw error;
    });

  apiCache.set(url, {
    data: cached?.data ?? null,
    expiry: cached?.expiry ?? 0,
    promise: requestPromise,
  });

  return requestPromise;
}

export function invalidateCache(url) {
  apiCache.delete(url);
}
