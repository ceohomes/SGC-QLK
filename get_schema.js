import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://luhsnaqlajbwkftrsbeg.supabase.co'
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx1aHNuYXFsYWpid2tmdHJzYmVnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMDk3MjIsImV4cCI6MjA5NTU4NTcyMn0.NqqOG1KkzceGquzudBcPqOSsX1BhB24U_jmew0Mqsc4'

async function tryOptions() {
  try {
    const res = await fetch(`${supabaseUrl}/rest/v1/don_giao`, {
      method: 'OPTIONS',
      headers: {
        'apikey': supabaseAnonKey
      }
    })
    console.log('OPTIONS status:', res.status, res.statusText)
    console.log('OPTIONS Headers:', [...res.headers.entries()])
    const txt = await res.text()
    console.log('OPTIONS Body text:', txt)
  } catch (e) {
    console.error('Error:', e)
  }
}

tryOptions()
