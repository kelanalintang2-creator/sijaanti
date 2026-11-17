/**
 * ============================================================================
 * SiJaAnti PRO - Konfigurasi Client Frontend (config.js)
 * Standar: /gas-pro-api, /github-pages-deploy-guide
 * ============================================================================
 * PENTING:
 * 1. Ganti 'GAS_URL' di bawah dengan URL deployment Web App Google Apps Script
 *    yang Anda peroleh setelah melakukan Deploy > New Deployment (Who has access: Anyone).
 * 2. URL wajib berakhiran '/exec'.
 * ============================================================================
 */

const APP_CONFIG = {
  APP_NAME: 'SiJaAnti PRO',
  APP_VERSION: '2.0.0-PRO',
  SCHOOL_NAME: 'SD Negeri Pilangrejo 1',

  // ⚠️ GANTI URL INI DENGAN URL WEB APP DEPLOYMENT GAS ANDA (berakhiran /exec)
  GAS_URL: 'https://script.google.com/macros/s/AKfycbzFpwhOARFtT4e-vg8Kj5Us_GxYDmf0oz2p6DCVEHG_PGJQq0r8LyqzQY4pS_9T4bnCdA/exec',

  // Konfigurasi Cache Lokal SWR
  STORAGE_KEY_DATA: 'SIJAANTI_PRO_DATA_V2',
  STORAGE_KEY_AUTH: 'SIJAANTI_PRO_AUTH_V2',
  CACHE_TTL_MS: 30 * 60 * 1000 // 30 Menit
};
