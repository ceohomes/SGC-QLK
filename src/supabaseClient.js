import { createClient } from '@supabase/supabase-js'

// Try finding in localStorage first (browser fallback), then in environment variables or default hardcoded credentials
const storedUrl = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_url') : null
const storedKey = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_key') : null

const DEFAULT_SUPABASE_URL = 'https://luhsnaqlajbwkftrsbeg.supabase.co'
const DEFAULT_SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx1aHNuYXFsYWpid2tmdHJzYmVnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMDk3MjIsImV4cCI6MjA5NTU4NTcyMn0.NqqOG1KkzceGquzudBcPqOSsX1BhB24U_jmew0Mqsc4'

export const supabaseUrl = storedUrl || import.meta.env.VITE_SUPABASE_URL || DEFAULT_SUPABASE_URL
export const supabaseAnonKey = storedKey || import.meta.env.VITE_SUPABASE_ANON_KEY || DEFAULT_SUPABASE_ANON_KEY

export const isSupabaseConfigured = !!(supabaseUrl && supabaseAnonKey)

export const supabase = isSupabaseConfigured 
  ? createClient(supabaseUrl, supabaseAnonKey)
  : null
