# Our Budget

A private, shared household budget app (installable on iPhone) built with plain HTML/JS,
hosted on GitHub Pages, with data in Supabase protected by Row Level Security.

- Only the two household accounts can sign in; sign-ups are disabled.
- `config.js` holds the Supabase **publishable** key, which is designed to be public.
  All data access is enforced by Row Level Security in the database.
- Receipts are read on the device (Tesseract.js / pdf.js); nothing is sent to third parties.
