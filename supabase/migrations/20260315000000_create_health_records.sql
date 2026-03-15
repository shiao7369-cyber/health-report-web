-- 健檢紀錄資料表
-- 欄位對應 src/lib/report-logic.ts 的 COL_MAP

create table if not exists health_records (
  -- 系統欄位
  id           uuid primary key default gen_random_uuid(),
  created_at   timestamptz not null default now(),

  -- 基本資料
  exam_date    date,            -- 體檢日
  health_id    text,            -- 健檢號碼
  name         text not null,   -- 姓名
  gender       text,            -- 性別
  age          text,            -- 年齡
  birthday     text,            -- 生日

  -- 慢性病史
  hypertension     text,        -- 高血壓
  diabetes         text,        -- 糖尿病
  hyperlipidemia   text,        -- 高血脂症
  heart_disease    text,        -- 心臟病
  kidney_disease   text,        -- 腎臟病
  stroke           text,        -- 腦中風

  -- 生活習慣
  smoking      text,            -- 吸菸
  drinking     text,            -- 喝酒
  betel_nut    text,            -- 嚼檳榔
  exercise     text,            -- 運動

  -- 憂鬱篩檢
  depression1  text,
  depression2  text,

  -- 身體測量
  height       text,            -- 身高 (cm)
  weight       text,            -- 體重 (kg)
  waist        text,            -- 腰圍 (cm)
  bmi          text,            -- 身體質量指數
  pulse        text,            -- 脈搏
  sbp          text,            -- 收縮壓
  dbp          text,            -- 舒張壓

  -- 尿液
  urine_protein text,           -- 蛋白質

  -- 血液生化
  cholesterol  text,            -- 膽固醇
  triglyceride text,            -- 三酸甘油脂
  got          text,            -- ASR(GOT)
  gpt          text,            -- ALT(GPT)
  bun          text,            -- 尿素氮
  creatinine   text,            -- 肌酐酸
  glucose      text,            -- 血糖
  uric_acid    text,            -- 尿酸
  hdl          text,            -- 高密度膽固醇
  ldl          text,            -- 低密度膽固醇
  egfr         text,            -- 腎絲球過濾率

  -- 肝炎檢查
  hbsag        text,            -- B型肝炎表面抗原
  hcv          text,            -- C型肝炎抗體
  prev_hep     text,            -- 曾於成健B,C肝檢查

  -- 衛教建議
  counsel_quit_smoke    text,
  counsel_quit_alcohol  text,
  counsel_quit_betel    text,
  counsel_accident      text,
  counsel_oral          text,
  counsel_weight        text,
  counsel_diet          text,
  counsel_exercise_old  text,
  counsel_maintain_weight text,
  counsel_healthy_diet  text,
  counsel_exercise      text,
  counsel_healthy_meal  text,
  counsel_chronic       text,
  counsel_chronic2      text,
  counsel_physical      text,
  counsel_report        text,
  counsel_kidney        text,

  -- 慢性疾病風險值
  risk_cad     text,            -- 冠心病
  risk_dm      text,            -- 糖尿病
  risk_htn     text,            -- 高血壓
  risk_stroke  text,            -- 腦中風
  risk_cv      text,            -- 血管不良事件
  kidney_stage text             -- 腎功能檢查期別
);

-- 常用查詢索引
create index on health_records (exam_date desc);
create index on health_records (name);
create index on health_records (health_id);

-- Row Level Security（預設關閉，需要時再開）
-- alter table health_records enable row level security;

comment on table health_records is '成人健康檢查紀錄，對應 Excel 健檢資料匯入';
