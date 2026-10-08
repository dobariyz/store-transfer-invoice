-- Run this in Supabase SQL Editor after creating staff Auth accounts.
-- RLS stays enabled. Only signed-in staff accounts receive access.

alter table public.products enable row level security;
alter table public.transfers enable row level security;
alter table public.transfer_items enable row level security;

revoke all on table public.products, public.transfers, public.transfer_items from anon;
grant select, insert, update, delete on table public.products, public.transfers, public.transfer_items to authenticated;
grant usage, select on all sequences in schema public to authenticated;

drop policy if exists "staff can access shared products" on public.products;
create policy "staff can access shared products"
  on public.products for all to authenticated
  using (true) with check (true);

drop policy if exists "staff can access shared transfers" on public.transfers;
create policy "staff can access shared transfers"
  on public.transfers for all to authenticated
  using (true) with check (true);

drop policy if exists "staff can access shared transfer items" on public.transfer_items;
create policy "staff can access shared transfer items"
  on public.transfer_items for all to authenticated
  using (true) with check (true);
