import urllib.request, json, http.cookiejar

jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
BASE = 'http://127.0.0.1:5000'

def post(path, body, expect_error=False):
    req = urllib.request.Request(BASE+path, data=json.dumps(body).encode(),
                                 headers={'Content-Type': 'application/json'})
    try:
        r = opener.open(req)
        return r.status, json.loads(r.read())
    except urllib.error.HTTPError as e:
        if expect_error:
            return e.code, json.loads(e.read())
        raise

def get(path):
    r = opener.open(urllib.request.Request(BASE+path))
    return json.loads(r.read())

def delete(path):
    req = urllib.request.Request(BASE+path, method='DELETE')
    r = opener.open(req)
    return r.status, json.loads(r.read())

PASS = "PASS"
FAIL = "FAIL"

# 1. Auth
s, d = post('/api/auth', {'pin': 'ast2026'})
print(f"1. Auth:         {PASS if d.get('role')=='editor' else FAIL} -- {d}")

# 2. Use item_name (auto-create) instead of item_id
s, d = post('/api/transactions', {
    'item_name': 'HySBatt v2.1 Pack',
    'type': 'IN', 'quantity': 5,
    'reference': 'TEST-001',
    'notes': 'Test entry',
    'logged_by': 'TestBot',
    'subitems': [
        {'name': 'Battery Unit 1', 'notes': 'Serial A001'},
        {'name': 'Battery Unit 2', 'notes': 'Serial A002'},
        {'name': 'Battery Unit 3', 'notes': ''},
    ]
})
print(f"2. Log IN (item_name + 3 sub-items): {PASS if s==200 else FAIL} -- {d}")

# 3. Fetch transactions and check subitems
txs = get('/api/transactions')
found = next((t for t in txs if t.get('reference')=='TEST-001'), None)
sub_ok = found and len(found.get('subitems',[])) == 3
print(f"3. Sub-items returned in history: {PASS if sub_ok else FAIL} -- {len(found.get('subitems',[])) if found else 'tx not found'} sub-items")
if found:
    for s_item in found['subitems']:
        print(f"   - {s_item['name']} | {s_item['notes']}")

# 4. Check stock card shows correct qty
items = get('/api/items')
item = next((i for i in items if i['name'] == 'HySBatt v2.1 Pack'), None)
print(f"4. Stock card shows item: {PASS if item else FAIL}")

# 5. Duplicate item_name auto-reuses existing item (no duplicate)
s, d = post('/api/transactions', {
    'item_name': 'hySBATT v2.1 PACK',   # different case
    'type': 'IN', 'quantity': 1,
    'logged_by': 'TestBot', 'reference': 'TEST-CASE'
})
items2 = get('/api/items')
same_items = [i for i in items2 if 'hys' in i['name'].lower()]
print(f"5. Case-insensitive reuse (no duplicates): {PASS if len(same_items)==1 else FAIL} -- {[i['name'] for i in same_items]}")

# 6. Clean up test entries
txs_all = get('/api/transactions')
test_txs = [t for t in txs_all if t.get('logged_by')=='TestBot']
for t in test_txs:
    delete(f'/api/transactions/{t["id"]}')
print(f"6. Cleanup: removed {len(test_txs)} test transaction(s)")

print("\nAll tests done.")
