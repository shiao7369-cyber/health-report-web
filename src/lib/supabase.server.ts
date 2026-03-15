import { createClient } from '@supabase/supabase-js';

/**
 * 伺服器端 Supabase client（使用 service_role key）
 * 只在 API Routes / SSR 頁面中使用，絕對不要在前端 JS 中匯入此檔案
 */
export function createServerClient() {
  const url = import.meta.env.SUPABASE_URL;
  const key = import.meta.env.SUPABASE_SERVICE_ROLE_KEY;

  if (!url || !key) {
    throw new Error('Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY environment variables');
  }

  return createClient(url, key, {
    auth: {
      // 伺服器端不需要自動更新 session
      autoRefreshToken: false,
      persistSession: false,
    },
  });
}
