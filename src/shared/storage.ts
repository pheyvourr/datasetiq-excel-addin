// Type declaration for Office.js global
declare const OfficeRuntime: any;

const STORAGE_KEY = 'DATASETIQ_API_KEY';
const FAVORITES_KEY = 'DATASETIQ_FAVORITES';
const RECENT_KEY = 'DATASETIQ_RECENT';

export interface StoredKey {
  key: string | null;
  supported: boolean;
}

export async function getStoredApiKey(): Promise<StoredKey> {
  // Try OfficeRuntime first (Excel environment)
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    try {
      const key = await OfficeRuntime.storage.getItem(STORAGE_KEY);
      return { key: key ?? null, supported: true };
    } catch (_err) {
      return { key: null, supported: false };
    }
  }
  
  // Fallback to localStorage for testing/development
  try {
    const key = localStorage.getItem(STORAGE_KEY);
    return { key: key ?? null, supported: true };
  } catch (_err) {
    return { key: null, supported: false };
  }
}

export async function setStoredApiKey(key: string): Promise<void> {
  // Try OfficeRuntime first (Excel environment)
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    await OfficeRuntime.storage.setItem(STORAGE_KEY, key);
    return;
  }
  
  // Fallback to localStorage for testing/development
  try {
    localStorage.setItem(STORAGE_KEY, key);
  } catch (_err) {
    throw new Error('Storage not available');
  }
}

export async function clearStoredApiKey(): Promise<void> {
  // Try OfficeRuntime first (Excel environment)
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    try {
      await OfficeRuntime.storage.removeItem(STORAGE_KEY);
    } catch (_err) {
      // ignore
    }
    return;
  }
  
  // Fallback to localStorage for testing/development
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch (_err) {
    // ignore
  }
}

// Favorites management
export async function getFavorites(): Promise<string[]> {
  // Try OfficeRuntime first (Excel environment)
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    try {
      const data = await OfficeRuntime.storage.getItem(FAVORITES_KEY);
      return data ? JSON.parse(data) : [];
    } catch (_err) {
      return [];
    }
  }
  
  // Fallback to localStorage for testing/development
  try {
    const data = localStorage.getItem(FAVORITES_KEY);
    return data ? JSON.parse(data) : [];
  } catch (_err) {
    return [];
  }
}

export async function addFavorite(seriesId: string): Promise<void> {
  const favorites = await getFavorites();
  if (!favorites.includes(seriesId)) {
    favorites.unshift(seriesId);
    const updated = JSON.stringify(favorites.slice(0, 50));
    
    // Try OfficeRuntime first
    if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(FAVORITES_KEY, updated);
    } else {
      localStorage.setItem(FAVORITES_KEY, updated);
    }
  }
}

export async function removeFavorite(seriesId: string): Promise<void> {
  const favorites = await getFavorites();
  const filtered = favorites.filter(id => id !== seriesId);
  const updated = JSON.stringify(filtered);
  
  // Try OfficeRuntime first
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    await OfficeRuntime.storage.setItem(FAVORITES_KEY, updated);
  } else {
    localStorage.setItem(FAVORITES_KEY, updated);
  }
}

// Recent series management
export async function getRecent(): Promise<string[]> {
  // Try OfficeRuntime first (Excel environment)
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    try {
      const data = await OfficeRuntime.storage.getItem(RECENT_KEY);
      return data ? JSON.parse(data) : [];
    } catch (_err) {
      return [];
    }
  }
  
  // Fallback to localStorage for testing/development
  try {
    const data = localStorage.getItem(RECENT_KEY);
    return data ? JSON.parse(data) : [];
  } catch (_err) {
    return [];
  }
}

export async function addRecent(seriesId: string): Promise<void> {
  const recent = await getRecent();
  const filtered = recent.filter(id => id !== seriesId);
  filtered.unshift(seriesId);
  const updated = JSON.stringify(filtered.slice(0, 20));
  
  // Try OfficeRuntime first
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    await OfficeRuntime.storage.setItem(RECENT_KEY, updated);
  } else {
    localStorage.setItem(RECENT_KEY, updated);
  }
}
