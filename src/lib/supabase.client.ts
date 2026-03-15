import { createClient } from '@supabase/supabase-js';

/**
 * 前端 Supabase client（使用 anon key）
 * 可在 React 元件中安全使用，適合 Supabase Auth、即時訂閱
 */
export const supabase = createClient(
  import.meta.env.PUBLIC_SUPABASE_URL,
  import.meta.env.PUBLIC_SUPABASE_ANON_KEY
);
