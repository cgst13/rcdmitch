import type { RCDReport, AccountCode, Signatory, RPTCollectionItem } from '../types/rcd';

// Service to handle Google Sheets interactions via Google Apps Script Web App

export interface User {
  email: string;
  name: string;
  role: string;
}

// Service to handle Google Sheets interactions via Node.js Express Backend or Google Apps Script Web App
export const getApiUrl = (): string => {
  const customUrl = localStorage.getItem('rcd_backend_url');
  if (customUrl && customUrl.trim()) {
    return customUrl.trim();
  }

  const envUrl = import.meta.env.VITE_API_URL || import.meta.env.VITE_GOOGLE_SCRIPT_URL || '';
  if (envUrl && envUrl.trim()) {
    return envUrl.trim();
  }

  // If hosted on HTTPS (like GitHub Pages), default to empty or HTTPS GAS URL if available
  if (typeof window !== 'undefined' && window.location.protocol === 'https:') {
    return '';
  }

  if (typeof window !== 'undefined' && window.location.hostname && window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
    return `http://${window.location.hostname}:5000`;
  }
  return 'http://localhost:5000';
};

export const setApiUrl = (url: string): void => {
  if (url && url.trim()) {
    localStorage.setItem('rcd_backend_url', url.trim());
  } else {
    localStorage.removeItem('rcd_backend_url');
  }
};

const clearLocalDatabaseCache = () => {
  localStorage.removeItem('rcd_reports');
  localStorage.removeItem('account_codes');
  localStorage.removeItem('signatories');
  localStorage.removeItem('rpt_collections');
  localStorage.removeItem('collections');
};

/**
 * Utility to extract Google Sheet ID from a URL or raw ID
 */
export const extractSpreadsheetId = (urlOrId: string): string | null => {
  if (!urlOrId || !urlOrId.trim()) return null;
  const trimmed = urlOrId.trim();
  if (/^[a-zA-Z0-9-_]{25,100}$/.test(trimmed)) {
    return trimmed;
  }
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return null;
};

export interface SpreadsheetConfig {
  spreadsheetId: string;
  spreadsheetUrl: string;
}

/**
 * Retrieves the current Google Sheet database configuration.
 */
export const getSpreadsheetConfig = async (): Promise<SpreadsheetConfig> => {
  const result = await callApi('getSpreadsheet');
  if (result && result.result === 'success' && result.spreadsheetId) {
    localStorage.setItem('rcd_spreadsheet_id', result.spreadsheetId);
    if (result.spreadsheetUrl) {
      localStorage.setItem('rcd_spreadsheet_url', result.spreadsheetUrl);
    }
    return {
      spreadsheetId: result.spreadsheetId,
      spreadsheetUrl: result.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${result.spreadsheetId}/edit`
    };
  }

  // Fallback to localStorage or default
  const storedId = localStorage.getItem('rcd_spreadsheet_id') || '1kZh86P60Meu3YhKy7neFdlqfH6AJjuHSUfGqmvYvfVo';
  const storedUrl = localStorage.getItem('rcd_spreadsheet_url') || `https://docs.google.com/spreadsheets/d/${storedId}/edit`;
  return {
    spreadsheetId: storedId,
    spreadsheetUrl: storedUrl
  };
};

/**
 * Updates the Google Sheet database link/ID in the system.
 */
export const updateSpreadsheetConfig = async (urlOrId: string): Promise<{ success: boolean; message: string; config?: SpreadsheetConfig }> => {
  const extractedId = extractSpreadsheetId(urlOrId);
  if (!extractedId) {
    return { success: false, message: 'Invalid Google Sheet URL or ID format. Please paste a full Google Sheet link or a valid Sheet ID.' };
  }

  // Clear stale cached data from previous database
  clearLocalDatabaseCache();

  const result = await callApi('updateSpreadsheet', { spreadsheetUrl: urlOrId, spreadsheetId: extractedId });
  
  if (result && result.result === 'success') {
    const finalId = result.spreadsheetId || extractedId;
    const finalUrl = result.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${finalId}/edit`;
    
    localStorage.setItem('rcd_spreadsheet_id', finalId);
    localStorage.setItem('rcd_spreadsheet_url', finalUrl);
    
    return {
      success: true,
      message: result.message || 'Google Sheet database link updated and verified successfully!',
      config: {
        spreadsheetId: finalId,
        spreadsheetUrl: finalUrl
      }
    };
  }

  // Fallback: If backend fails or offline, save to localStorage
  const fallbackUrl = `https://docs.google.com/spreadsheets/d/${extractedId}/edit`;
  localStorage.setItem('rcd_spreadsheet_id', extractedId);
  localStorage.setItem('rcd_spreadsheet_url', fallbackUrl);

  return {
    success: true,
    message: result?.message ? `Notice: ${result.message}` : 'Google Sheet link updated in local client settings.',
    config: {
      spreadsheetId: extractedId,
      spreadsheetUrl: fallbackUrl
    }
  };
};

/**
 * Helper to call Backend API (Node.js Express or Google Apps Script Web App)
 * All requests are POST with { action, ...payload }
 */
const callApi = async (action: string, payload: Record<string, any> = {}) => {
  const targetUrl = getApiUrl();

  // If page is HTTPS and targetUrl is unencrypted HTTP, browser will block (Mixed Content)
  if (typeof window !== 'undefined' && window.location.protocol === 'https:' && targetUrl.startsWith('http:')) {
    console.warn(`Mixed Content Prevention (${action}): Unable to fetch HTTP endpoint "${targetUrl}" from HTTPS site.`);
    return {
      result: 'error',
      message: 'Mixed Content Error: GitHub Pages runs on HTTPS and blocks connection to unencrypted http:// backends. Please configure your HTTPS Google Apps Script Web App URL in Settings.'
    };
  }

  if (!targetUrl) {
    if (typeof window !== 'undefined' && window.location.protocol === 'https:') {
      return {
        result: 'error',
        message: 'No HTTPS API URL configured. Please set your Google Apps Script Web App URL in Settings.'
      };
    }
    return null;
  }

  const isGas = targetUrl.includes('script.google.com');

  try {
    const options: RequestInit = {
      method: 'POST',
      headers: isGas
        ? { 'Content-Type': 'text/plain;charset=utf-8' }
        : { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action, ...payload }),
    };

    const response = await fetch(targetUrl, options);
    const data = await response.json();
    return data;
  } catch (error: any) {
    console.error(`API Error (${action}):`, error);
    return null;
  }
};


/**
 * Submits an RCD report to the backend.
 */
export const submitRCDReport = async (report: RCDReport): Promise<boolean> => {
  console.log('Submitting Report Service:', report);

  const result = await callApi('submitReport', { report });
  
  if (result && result.result === 'success') {
    return true;
  }
  
  console.warn('Backend submission failed, saving to local storage as backup');
  
  // Fallback: Save to localStorage
  try {
    const existingReports = JSON.parse(localStorage.getItem('rcd_reports') || '[]');
    const newReports = [report, ...existingReports];
    localStorage.setItem('rcd_reports', JSON.stringify(newReports));
    return true; 
  } catch (e) {
    console.error('Failed to save to localStorage', e);
    return false;
  }
};

/**
 * Retrieves recent reports from backend.
 */
export const getRecentReports = async (): Promise<RCDReport[]> => {
  const result = await callApi('getReports');
  
  if (result && result.result === 'success') {
    return result.reports;
  }

  // Fallback: localStorage
  console.warn('Backend fetch failed, loading from local storage');
  try {
    return JSON.parse(localStorage.getItem('rcd_reports') || '[]');
  } catch {
    return [];
  }
};

/**
 * Logs in by checking credentials against the backend.
 */
export const loginWithGoogleSheet = async (email: string, password: string): Promise<User | null> => {
  console.log(`Attempting login for ${email}`);
  
  const result = await callApi('login', { email, password });
  
  if (result && result.result === 'success') {
    return {
      email: result.email,
      name: result.name,
      role: result.role
    };
  }

  return null;
};

/**
 * Account Code Management
 */

export const getAccountCodes = async (): Promise<AccountCode[]> => {
  const result = await callApi('getAccountCodes');
  
  if (result && result.result === 'success') {
    return result.accountCodes;
  }

  // Fallback: LocalStorage
  const stored = localStorage.getItem('account_codes');
  if (stored) {
    return JSON.parse(stored);
  }
  
  // Default Mock Data (if nothing in storage and backend fails)
  return [
    { id: 1, mainCategory: 'General Fund', subCategory: 'Tax Revenue', code: '1-01-01-010' },
    { id: 2, mainCategory: 'General Fund', subCategory: 'Tax Revenue', code: '4-01-01-010' },
    { id: 3, mainCategory: 'Special Education Fund', subCategory: 'Tax Revenue', code: '4-01-02-020' },
    { id: 4, mainCategory: 'General Fund', subCategory: 'Business Income', code: '4-01-02-010' },
    { id: 5, mainCategory: 'General Fund', subCategory: 'Service Income', code: '4-01-03-010' },
    { id: 6, mainCategory: 'Trust Fund', subCategory: 'Inter-Agency', code: '2-02-01-010' },
  ];
};

export const saveAccountCode = async (code: AccountCode): Promise<boolean> => {
  const result = await callApi('saveAccountCode', { accountCode: code });
  
  if (result && result.result === 'success') {
    return true;
  }

  // Fallback: LocalStorage
  console.warn('Backend save failed, saving to local storage');
  const currentCodes = await getAccountCodes();
  const index = currentCodes.findIndex(c => c.id === code.id);
  
  let newCodes;
  if (index >= 0) {
    newCodes = [...currentCodes];
    newCodes[index] = code;
  } else {
    newCodes = [...currentCodes, code];
  }
  
  localStorage.setItem('account_codes', JSON.stringify(newCodes));
  return true;
};

export const deleteAccountCode = async (id: number): Promise<boolean> => {
  const result = await callApi('deleteAccountCode', { id });
  
  if (result && result.result === 'success') {
    return true;
  }

  // Fallback: LocalStorage
  console.warn('Backend delete failed, removing from local storage');
  const currentCodes = await getAccountCodes();
  const newCodes = currentCodes.filter(c => c.id !== id);
  localStorage.setItem('account_codes', JSON.stringify(newCodes));
  return true;
};

/**
 * Signatory Management
 */

export const getSignatories = async (): Promise<Signatory[]> => {
  const result = await callApi('getSignatories');
  
  if (result && result.result === 'success') {
    return result.signatories;
  }

  // Fallback: LocalStorage
  const stored = localStorage.getItem('signatories');
  if (stored) {
    return JSON.parse(stored);
  }
  
  // Default Mock Data
  return [
    { id: 1, fullName: 'CHRISTIAN S. TOLENTINO', position: 'RCC I', department: 'Treasury' },
    { id: 2, fullName: 'LEON F. PAZ, JR.', position: 'Chief, Accounting Department/Unit', department: 'Accounting' },
    { id: 3, fullName: 'SISTINE F. ATILLANO', position: 'SAA I', department: 'Treasury' },
    { id: 4, fullName: 'MENARD A. HERRERA', position: 'Municipal Treasurer', department: 'Treasury' },
  ];
};

export const saveSignatory = async (signatory: Signatory): Promise<boolean> => {
  const result = await callApi('saveSignatory', { signatory });
  
  if (result && result.result === 'success') {
    return true;
  }

  // Fallback: LocalStorage
  console.warn('Backend save failed, saving to local storage');
  const currentSignatories = await getSignatories();
  const index = currentSignatories.findIndex(s => s.id === signatory.id);
  
  let newSignatories;
  if (index >= 0) {
    newSignatories = [...currentSignatories];
    newSignatories[index] = signatory;
  } else {
    newSignatories = [...currentSignatories, signatory];
  }
  
  localStorage.setItem('signatories', JSON.stringify(newSignatories));
  return true;
};

export const deleteSignatory = async (id: number): Promise<boolean> => {
  const result = await callApi('deleteSignatory', { id });
  
  if (result && result.result === 'success') {
    return true;
  }

  // Fallback: LocalStorage
  console.warn('Backend delete failed, removing from local storage');
  const currentSignatories = await getSignatories();
  const newSignatories = currentSignatories.filter(s => s.id !== id);
  localStorage.setItem('signatories', JSON.stringify(newSignatories));
  return true;
};


export interface CollectionEntry {
  id: number;
  afNo: string;
  orNo: string;
  payor: string;
  subCategory: string;
  mainCategory: string;
  accountCode: string;
  amount: number;
  date: string;
  remarks: string;
}

export const getCollectionEntries = async (): Promise<CollectionEntry[]> => {
  const result = await callApi('getCollections');
  
  if (result && result.result === 'success') {
    // Sanitize data: Ensure string fields are strings
    return result.entries.map((e: any) => ({
      ...e,
      afNo: String(e.afNo || ''),
      orNo: String(e.orNo || ''),
      payor: String(e.payor || ''),
      subCategory: String(e.subCategory || ''),
      mainCategory: String(e.mainCategory || ''),
      accountCode: String(e.accountCode || ''),
      remarks: String(e.remarks || ''),
      date: String(e.date || '')
    }));
  }
  return [];
};

export const saveCollectionEntry = async (entry: CollectionEntry): Promise<boolean> => {
  const result = await callApi('saveCollection', { entry });
  
  if (result && result.result === 'success') {
    return true;
  }
  return false;
};

export interface CollectionHeader {
  afNo: string;
  orNo: string;
  payor: string;
  date: string;
  remarks: string;
}

export interface CollectionCharge {
  subCategory: string;
  mainCategory: string;
  accountCode: string;
  amount: number;
}

export const saveCollectionEntryBulk = async (
  header: CollectionHeader,
  charges: CollectionCharge[]
): Promise<boolean> => {
  const result = await callApi('saveCollectionBulk', { header, charges });
  if (result && result.result === 'success') {
    return true;
  }
  return false;
};

export const updateCollectionEntry = async (entry: CollectionEntry): Promise<boolean> => {
  const result = await callApi('updateCollection', { entry });
  if (result && result.result === 'success') {
    return true;
  }
  return false;
};

export const deleteCollectionEntry = async (id: number): Promise<boolean> => {
  const result = await callApi('deleteCollection', { id });
  if (result && result.result === 'success') {
    return true;
  }
  return false;
};

/**
 * RPT Collection Management
 */

export const getRPTCollections = async (): Promise<RPTCollectionItem[]> => {
  const result = await callApi('getRPTCollections');
  
  if (result && result.result === 'success') {
    // Sanitize data: Ensure string fields are strings
    return result.collections.map((c: any) => ({
      ...c,
      af56Id: String(c.af56Id || ''),
      orNumber: String(c.orNumber || ''),
      payor: String(c.payor || ''),
      barangay: String(c.barangay || ''),
      landName: String(c.landName || ''),
      tdNumber: String(c.tdNumber || ''),
      yearsPaid: String(c.yearsPaid || ''),
      remarks: String(c.remarks || ''),
      date: String(c.date || '')
    }));
  }

  // Fallback: LocalStorage
  const stored = localStorage.getItem('rpt_collections');
  if (stored) {
    return JSON.parse(stored);
  }
  
  return [];
};

export const saveRPTCollection = async (collection: RPTCollectionItem): Promise<boolean> => {
  const result = await callApi('saveRPTCollection', { collection });
  
  if (result && result.result === 'success') {
    return true;
  }

  // Fallback: LocalStorage
  console.warn('Backend save failed, saving to local storage');
  const currentCollections = await getRPTCollections();
  const index = currentCollections.findIndex(c => c.id === collection.id);
  
  let newCollections;
  if (index >= 0) {
    newCollections = [...currentCollections];
    newCollections[index] = collection;
  } else {
    newCollections = [...currentCollections, collection];
  }
  
  localStorage.setItem('rpt_collections', JSON.stringify(newCollections));
  return true;
};

export const deleteRPTCollection = async (id: number): Promise<boolean> => {
  const result = await callApi('deleteRPTCollection', { id });
  
  if (result && result.result === 'success') {
    return true;
  }

  // Fallback: LocalStorage
  console.warn('Backend delete failed, removing from local storage');
  const currentCollections = await getRPTCollections();
  const newCollections = currentCollections.filter(c => c.id !== id);
  localStorage.setItem('rpt_collections', JSON.stringify(newCollections));
  return true;
};
