-- SQL Script to create the configuration storage tables in Supabase for SGC App
-- Run this script in the SQL Editor on your Supabase Dashboard

-- 1. Create the summary configurations storage table
CREATE TABLE IF NOT EXISTS sgc_summary_configs (
    id bigint primary key generated always as identity,
    created_at timestamp with time zone default timezone('utc'::text, now()) not null,
    name text not null,
    project text,
    giao_table jsonb default '[]'::jsonb,
    nhan_table jsonb default '[]'::jsonb
);

-- 2. Add description comment for the table and fields
COMMENT ON TABLE sgc_summary_configs IS 'Lưu trữ các cấu hình tổng hợp dự án SGC';
COMMENT ON COLUMN sgc_summary_configs.name IS 'Tên hiển thị của cấu hình';
COMMENT ON COLUMN sgc_summary_configs.project IS 'Dự án liên kết';
COMMENT ON COLUMN sgc_summary_configs.giao_table IS 'Mảng lưu các đơn vị giao nhận từ bảng Đơn Giao (Giảm trừ, Bỏ qua, Tính toán)';
COMMENT ON COLUMN sgc_summary_configs.nhan_table IS 'Mảng lưu các đơn vị giao nhận từ bảng Đơn Nhận (Giảm trừ, Bỏ qua, Tính toán)';

-- 3. Enable Row Level Security (RLS) to keep the data safe
ALTER TABLE sgc_summary_configs ENABLE ROW LEVEL SECURITY;

-- 4. Create RLS Policies to allow standard operations
-- Note: Adjust policy rules to fit your access control specifications if needed.
CREATE POLICY "Allow public read access" 
ON sgc_summary_configs 
FOR SELECT 
USING (true);

CREATE POLICY "Allow public insert access" 
ON sgc_summary_configs 
FOR INSERT 
WITH CHECK (true);

CREATE POLICY "Allow public update access" 
ON sgc_summary_configs 
FOR UPDATE 
USING (true);

CREATE POLICY "Allow public delete access" 
ON sgc_summary_configs 
FOR DELETE 
USING (true);
