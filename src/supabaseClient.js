import { createClient } from '@supabase/supabase-js'

// Try finding in localStorage first (browser fallback), then in environment variables
const storedUrl = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_url') : null
const storedKey = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_key') : null

export const supabaseUrl = storedUrl || import.meta.env.VITE_SUPABASE_URL || ''
export const supabaseAnonKey = storedKey || import.meta.env.VITE_SUPABASE_ANON_KEY || ''

export const isSupabaseConfigured = !!(supabaseUrl && supabaseAnonKey)

export const supabase = isSupabaseConfigured 
  ? createClient(supabaseUrl, supabaseAnonKey)
  : null
