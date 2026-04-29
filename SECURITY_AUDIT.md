# Security Audit

Date: 2026-04-29

## Current Posture

This is a browser-only React/Vite app that stores task data in Supabase through the public anon key. React escapes rendered user content by default, and the app does not use `dangerouslySetInnerHTML`.

## Fixes Applied

- Supabase reads are treated as the source of truth, preventing stale device caches from resurrecting deleted tasks.
- Deleted task IDs are written to Supabase metadata as tombstones so stale clients cannot re-upload deleted records.
- Uploaded files are capped at 8 MB and extracted text is capped at 120,000 characters to reduce browser memory and storage risk.
- Local development unregisters stale service workers so old cached app code does not keep running.
- The service worker uses network-first reads and avoids caching cross-origin Supabase responses.
- A `no-referrer` policy was added to avoid leaking local URLs or query strings via the referrer header.

## Critical Finding

The current Supabase schema grants anonymous read, insert, update, and delete access to all task, output, and metadata rows. This is functional for a private prototype, but it is not secure for sensitive personal, client, or business data if the app is reachable by anyone else.

## Required Production Fix

Before using this with sensitive data, add authentication and replace the anon-wide RLS policies with user-scoped policies. Each row should include a `user_id uuid not null` tied to `auth.uid()`, and policies should only allow users to read and write their own rows.

## Residual Risk

Without user authentication, any browser client with the Supabase URL and anon key can access the shared dataset according to the current RLS policies. The frontend cannot fully secure this on its own.
