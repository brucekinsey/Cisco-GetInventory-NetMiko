import pynetbox
from collections import defaultdict
from datetime import datetime

NB_URL = "http://nova.suwannee.k12.fl.us:7200"
TOKEN = ""  # <-- don't paste real tokens

DRY_RUN = False
MOVE_PREFIXES = False   # leave False unless you know you need it
DELETE_DUPES = True    # set True only after you're satisfied with merge-only

BATCH_SIZE = 200        # tune if needed (100-500 usually fine)

nb = pynetbox.api(NB_URL, token=TOKEN)

# Ensure manual nb.http_session requests carry auth (bulk PATCH uses this)
nb.http_session.headers.update({
    "Authorization": f"Token {TOKEN}",
    "Content-Type": "application/json",
    "Accept": "application/json",
})

print("Auth header present:", "Authorization" in nb.http_session.headers)
print("Authorization header:", nb.http_session.headers.get("Authorization", "")[:20] + "...")

# -------------------------
# Helpers (normalize fields)
# -------------------------

def vlan_id_list(tagged_vlans):
    """Return tagged VLAN IDs as list[int] whether elements are ints or objects."""
    if not tagged_vlans:
        return []
    out = []
    for v in tagged_vlans:
        out.append(v if isinstance(v, int) else v.id)
    return out

def vlan_ref_id(v):
    """Normalize VLAN reference to int ID (None / int / object)."""
    if v is None:
        return None
    return v if isinstance(v, int) else getattr(v, "id", None)

def scope_key(v):
    """Defines VLAN uniqueness scope; only merge VLANs in the same scope."""
    site_id = v.site.id if v.site else None
    group_id = v.group.id if v.group else None
    tenant_id = v.tenant.id if v.tenant else None
    status = v.status.value if hasattr(v.status, "value") else v.status
    return (v.vid, site_id, group_id, tenant_id, status)

# -------------------------
# Bulk patch primitives
# -------------------------

def bulk_patch(endpoint_path, payloads, batch_size=BATCH_SIZE):
    if not payloads:
        return 0

    url = nb.base_url.rstrip("/") + endpoint_path
    total = 0

    for start in range(0, len(payloads), batch_size):
        batch = payloads[start:start + batch_size]
        if DRY_RUN:
            total += len(batch)
            continue

        r = nb.http_session.patch(url, json=batch)

        if not r.ok:
            print("\nPATCH FAILED")
            print("URL:", url)
            print("Status:", r.status_code)
            try:
                print("JSON:", r.json())
            except Exception:
                print("TEXT:", r.text)
            # Also print headers if you suspect a proxy
            # print("Resp headers:", dict(r.headers))
            r.raise_for_status()

        total += len(batch)

    return total

# -------------------------
# Index interfaces once
# -------------------------

def build_interface_index(all_ifaces):
    """
    Build:
      untagged_index[vlan_id] -> list of interface objects
      tagged_index[vlan_id]   -> list of interface objects
    """
    untagged_index = defaultdict(list)
    tagged_index = defaultdict(list)

    for i in all_ifaces:
        u = vlan_ref_id(i.untagged_vlan)
        if u is not None:
            untagged_index[u].append(i)

        for vid in vlan_id_list(i.tagged_vlans):
            tagged_index[vid].append(i)

    return untagged_index, tagged_index

# -------------------------
# Build patches for a VLAN merge
# -------------------------

def build_dcim_interface_patches(from_vlan_id, to_vlan_id, untagged_index, tagged_index):
    """
    Returns list of patch dicts for dcim/interfaces bulk PATCH.
    Each dict includes id and only fields that need changing.
    """
    patches_by_id = {}

    # untagged changes
    for i in untagged_index.get(from_vlan_id, []):
        patches_by_id[i.id] = {"id": i.id, "untagged_vlan": to_vlan_id}

    # tagged changes
    for i in tagged_index.get(from_vlan_id, []):
        tagged_ids = vlan_id_list(i.tagged_vlans)
        if from_vlan_id not in tagged_ids:
            continue
        new_ids = [to_vlan_id if x == from_vlan_id else x for x in tagged_ids]
        new_ids = sorted(set(new_ids))
        d = patches_by_id.setdefault(i.id, {"id": i.id})
        d["tagged_vlans"] = new_ids

    return list(patches_by_id.values())

def repoint_prefixes_bulk(from_vlan_id, to_vlan_id):
    """
    Optional: bulk patch prefixes that reference VLAN.
    If you never use prefix.vlan, leave MOVE_PREFIXES=False.
    """
    prefixes = list(nb.ipam.prefixes.filter(vlan_id=from_vlan_id))
    if not prefixes:
        return 0

    payloads = [{"id": p.id, "vlan": to_vlan_id} for p in prefixes]
    # NetBox prefixes endpoint:
    count = bulk_patch("/ipam/prefixes/", payloads, batch_size=BATCH_SIZE)
    return count

# -------------------------
# Verification (scan-based, reliable)
# -------------------------

def verify_vlan_not_used(vlan_id, all_ifaces):
    untagged_refs = 0
    tagged_refs = 0
    for i in all_ifaces:
        if vlan_ref_id(i.untagged_vlan) == vlan_id:
            untagged_refs += 1
        if vlan_id in vlan_id_list(i.tagged_vlans):
            tagged_refs += 1
    return untagged_refs, tagged_refs

# -------------------------
# VLAN discovery
# -------------------------

def fetch_all_vlans():
    vlans = list(nb.ipam.vlans.all())
    print("Fetched VLANs:", len(vlans))
    return vlans

def find_duplicate_vlan_groups(vlans):
    buckets = defaultdict(list)
    for v in vlans:
        buckets[scope_key(v)].append(v)

    dupes = []
    for key, items in buckets.items():
        if len(items) > 1:
            dupes.append((key, sorted(items, key=lambda x: x.id)))
    return dupes

def delete_vlan(vlan_id):
    v = nb.ipam.vlans.get(vlan_id)
    print(f"  [delete] VLAN id={v.id} vid={v.vid} name={v.name}")
    if not DRY_RUN and DELETE_DUPES:
        v.delete()

# -------------------------
# Merge one duplicate group
# -------------------------

def merge_group(key, vlan_list):
    vid, site_id, group_id, tenant_id, status = key
    canonical = vlan_list[0]
    dupes = vlan_list[1:]

    print("\n")
    print(datetime.now().strftime("%H:%M:%S"))
    print("=== DUPLICATE VLAN GROUP ===")
    print(f"VID={vid} site={site_id} group={group_id} tenant={tenant_id} status={status}")
    print("Canonical VLAN ID:", canonical.id, "name:", canonical.name)
    print("Dupes:", [(d.id, d.name) for d in dupes])

    # Fetch interfaces once per group and build index
    all_ifaces = list(nb.dcim.interfaces.all())
    untagged_index, tagged_index = build_interface_index(all_ifaces)

    # Drain dupes into canonical
    for d in dupes:
        if MOVE_PREFIXES:
            pref_count = repoint_prefixes_bulk(d.id, canonical.id)
            print(f"  [prefixes] {d.id} -> {canonical.id}: patched={pref_count}")

        patches = build_dcim_interface_patches(d.id, canonical.id, untagged_index, tagged_index)
        print(f"  [interfaces] {d.id} -> {canonical.id}: patches={len(patches)}")

        patched = bulk_patch("/dcim/interfaces/", patches, batch_size=BATCH_SIZE)
        if DRY_RUN:
            print(f"    (dry-run) would patch {patched} interfaces")
        else:
            print(f"    patched {patched} interfaces")

        # Refresh interface index after applying changes so subsequent dupes are accurate
        if not DRY_RUN:
            all_ifaces = list(nb.dcim.interfaces.all())
            untagged_index, tagged_index = build_interface_index(all_ifaces)

    # Verify + delete only when requested
    if not DELETE_DUPES:
        return

    all_ifaces_verify = list(nb.dcim.interfaces.all())
    for d in dupes:
        u, t = verify_vlan_not_used(d.id, all_ifaces_verify)
        print(f"  [verify] vlan={d.id}: untagged={u} tagged={t}")
        if u == 0 and t == 0:
            delete_vlan(d.id)
        else:
            print(f"  [skip delete] VLAN {d.id} still referenced; investigate.")

# -------------------------
# MAIN
# -------------------------

vlans = fetch_all_vlans()
dupe_groups = find_duplicate_vlan_groups(vlans)

if not dupe_groups:
    print("No duplicate VLAN groups found.")
    raise SystemExit(0)

print("Duplicate groups found:", len(dupe_groups))

for key, vlan_list in dupe_groups:
    merge_group(key, vlan_list)

print("\nDONE.")
print("Suggested run order:")
print("  1) DRY_RUN=True,  DELETE_DUPES=False  (no changes)")
print("  2) DRY_RUN=False, DELETE_DUPES=False  (merge only)")
print("  3) DRY_RUN=False, DELETE_DUPES=True   (verify + delete)")