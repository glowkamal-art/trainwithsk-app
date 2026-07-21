import { createClient } from "@supabase/supabase-js";

// These two values are safe to keep in the front-end code —
// the "anon" key is designed to be public. Never put your database
// password or "service_role" key here.
const SUPABASE_URL = "https://syatbzihqcaxcjgmiwwi.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InN5YXRiemlocWNheGNqZ21pd3dpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ2MTIxMDEsImV4cCI6MjEwMDE4ODEwMX0.9fQ-HjBmKQe0RBRRq2CD0s6xSv568-SB-XiQMtzMN7Q";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
