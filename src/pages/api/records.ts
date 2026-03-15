import type { APIRoute } from 'astro';
import { createServerClient } from '../../lib/supabase.server';

// 此 API Route 為 SSR（不預先渲染），讓 service_role key 保持在伺服器端
export const prerender = false;

/**
 * GET /api/records
 * 查詢健檢紀錄
 *
 * Query params:
 *   limit    預設 50，最大 200
 *   offset   預設 0
 *   name     依姓名模糊搜尋
 *   date     依體檢日（YYYY-MM-DD）精確搜尋
 *   health_id 依健檢號碼精確搜尋
 */
export const GET: APIRoute = async ({ request }) => {
  try {
    const url       = new URL(request.url);
    const limit     = Math.min(Number(url.searchParams.get('limit')  ?? 50), 200);
    const offset    = Number(url.searchParams.get('offset') ?? 0);
    const name      = url.searchParams.get('name')      ?? '';
    const date      = url.searchParams.get('date')      ?? '';
    const healthId  = url.searchParams.get('health_id') ?? '';

    const supabase = createServerClient();
    let query = supabase
      .from('health_records')
      .select('*', { count: 'exact' })
      .range(offset, offset + limit - 1)
      .order('exam_date', { ascending: false });

    if (name)     query = query.ilike('name',      `%${name}%`);
    if (date)     query = query.eq('exam_date',    date);
    if (healthId) query = query.eq('health_id',    healthId);

    const { data, error, count } = await query;
    if (error) throw error;

    return new Response(JSON.stringify({ data, count }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    return new Response(JSON.stringify({ error: message }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
    });
  }
};

/**
 * POST /api/records
 * 寫入健檢紀錄
 *
 * Body: 單筆 { name, exam_date, ... } 或陣列 [{ ... }, ...]
 * 欄位對應 supabase/migrations/20260315000000_create_health_records.sql
 */
export const POST: APIRoute = async ({ request }) => {
  try {
    const body    = await request.json();
    const records = Array.isArray(body) ? body : [body];

    if (records.some((r: unknown) => typeof r !== 'object' || r === null || !('name' in r))) {
      return new Response(JSON.stringify({ error: 'Each record must have at least a "name" field' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
      });
    }

    const supabase = createServerClient();
    const { data, error } = await supabase
      .from('health_records')
      .insert(records)
      .select();

    if (error) throw error;

    return new Response(JSON.stringify({ data }), {
      status: 201,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    return new Response(JSON.stringify({ error: message }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
    });
  }
};
