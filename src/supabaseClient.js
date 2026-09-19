import { createClient } from '@supabase/supabase-js'

// Try finding in localStorage first (browser fallback), then in environment variables or default hardcoded credentials
const storedUrl = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_url') : null
const storedKey = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_key') : null

const DEFAULT_SUPABASE_URL = 'https://ojknvcprfwcrnydkspqp.supabase.co'
const DEFAULT_SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9qa252Y3ByZndjcm55ZGtzcHFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM2NDM5MDIsImV4cCI6MjA5OTIxOTkwMn0.UpuOPw6WyXQQv-2zYzRdExjg1MzzQvYRdRcTM-wHLUY'

export const supabaseUrl = storedUrl || import.meta.env.VITE_SUPABASE_URL || DEFAULT_SUPABASE_URL
export const supabaseAnonKey = storedKey || import.meta.env.VITE_SUPABASE_ANON_KEY || DEFAULT_SUPABASE_ANON_KEY

export const isSupabaseConfigured = !!(supabaseUrl && supabaseAnonKey)

export const supabase = isSupabaseConfigured 
  ? createClient(supabaseUrl, supabaseAnonKey)
  : null
