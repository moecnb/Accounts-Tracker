import { createClient } from '@supabase/supabase-js';

export const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

export interface WipStatusRow {
  company_number: string;
  client_name: string;
  is_wip: boolean;
  marked_at: string | null;
  marked_by: string | null;
}
