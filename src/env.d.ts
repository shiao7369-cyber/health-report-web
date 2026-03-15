/// <reference path="../.astro/types.d.ts" />

interface ImportMetaEnv {
  // 伺服器端（無 PUBLIC_ 前綴，不會打包進前端 JS）
  readonly SUPABASE_URL: string;
  readonly SUPABASE_SERVICE_ROLE_KEY: string;

  // 前端可見（有 PUBLIC_ 前綴）
  readonly PUBLIC_SUPABASE_URL: string;
  readonly PUBLIC_SUPABASE_ANON_KEY: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}