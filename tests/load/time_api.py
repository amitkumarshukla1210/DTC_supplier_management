"""Time the main API calls against the big database (server on port 8766)."""
import json
import sys
import time

import httpx

BASE = "http://127.0.0.1:8766"
client = httpx.Client(base_url=BASE, timeout=600)
setup = client.post("/api/auth/setup", json={"username": "admin", "full_name": "Load Admin", "password": "Secure-test-password-2026!"})
if setup.status_code != 201:
    client.post("/api/auth/login", json={"username": "admin", "password": "Secure-test-password-2026!"})

models = client.get("/api/planning/models").json()
first_model = models[0]["id"]
parents = client.get("/api/planning/bom-part-parents", params={"tab": "Level 2 Part BOM"}).json()["items"]
supplier = client.get("/api/suppliers").json()[0]["id"]
calls = [
    ("GET  /api/suppliers (all 1,000)", "GET", "/api/suppliers", {}),
    ("GET  /api/planning/products", "GET", "/api/planning/products", {}),
    ("GET  /api/planning/models", "GET", "/api/planning/models", {}),
    ("GET  materials first 500", "GET", "/api/planning/materials", {"limit": 500}),
    ("GET  materials search 'LM-0123'", "GET", "/api/planning/materials", {"search": "LM-0123"}),
    ("GET  purchased materials (picker)", "GET", "/api/planning/materials", {"purchased": "true", "limit": 100}),
    ("GET  material categories", "GET", "/api/planning/material-categories", {}),
    ("GET  model BOM of 1 model", "GET", f"/api/planning/model-bom/{first_model}", {}),
    ("GET  L2 parents", "GET", "/api/planning/bom-part-parents", {"tab": "Level 2 Part BOM"}),
    ("GET  L3 parents", "GET", "/api/planning/bom-part-parents", {"tab": "Level 3 Part BOM"}),
    ("GET  part BOM of 1 parent", "GET", f"/api/planning/part-bom/{parents[0]['id']}", {"tab": "Level 2 Part BOM"}),
    ("GET  supplier links page 500", "GET", "/api/planning/supplier-parts", {"limit": 500}),
    ("GET  supplier BOM products", "GET", "/api/planning/supplier-bom/products", {"supplier_id": supplier}),
    ("GET  BOM tree Excel: Product Master (all)", "GET", "/api/planning/bom-tree/xlsx", {"tab": "Product Master"}),
    ("GET  BOM tree Excel: Level 2 (all)", "GET", "/api/planning/bom-tree/xlsx", {"tab": "Level 2 Part BOM"}),
]
results = []
for label, method, path, params in calls:
    start = time.perf_counter()
    response = client.request(method, path, params=params)
    elapsed = time.perf_counter() - start
    results.append((label, response.status_code, elapsed, len(response.content)))
    print(f"{elapsed:7.2f}s  {response.status_code}  {len(response.content)/1024:8.0f} KB  {label}", flush=True)

ids = [m["id"] for m in models[:50]]
start = time.perf_counter()
client.get("/api/planning/model-boms", params=[("model_id", i) for i in ids])
print(f"{time.perf_counter() - start:7.2f}s  ---  {'':>8}     Level 1 tab now: 1 request for a page of 50 models")
# What the Model BOM Level 1 tab did before: one request per model (all 1,000 selected)
start = time.perf_counter()
for model in models:
    client.get(f"/api/planning/model-bom/{model['id']}", params={"tab": "Model BOM Level 1", "limit": 500})
elapsed = time.perf_counter() - start
print(f"{elapsed:7.2f}s  ---  {'':>8}     Level 1 tab today: {len(models)} requests (one per model)")
json.dump([{"label": r[0], "status": r[1], "seconds": round(r[2], 3), "kb": round(r[3] / 1024)} for r in results] + [{"label": "L1 tab, 1000 requests", "seconds": round(elapsed, 2)}],
          open(sys.argv[1] if len(sys.argv) > 1 else "/dev/null", "w"), indent=1)
