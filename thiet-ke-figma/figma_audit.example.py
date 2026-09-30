# Bản mẫu của figma_audit.py: đã bỏ token và mã file Figma thật.
# Điền mã file (phần sau figma.com/design/) vào chỗ FILE_KEY_xx.
"""
Figma Audit Script v2 - Rate-limit aware
Counts: Modules (pages), Features (sections), Screens (frames), Components (instances)

Strategy:
  - depth=3 for counting (lighter API load)
  - depth=5 only for component counting (separate pass)
  - Reads Retry-After header on 429
  - Caches results in JSON so re-runs skip completed files
"""

import json
import os
import sys
import time
import urllib.request
import urllib.error
from datetime import datetime

# Token Figma cá nhân: đặt vào biến môi trường FIGMA_TOKEN trước khi chạy.
FIGMA_TOKEN = os.environ.get("FIGMA_TOKEN", "")
CACHE_FILE = "figma_audit_cache.json"
DELAY_BETWEEN_FILES = 15  # seconds

FILES = [
    ("THS - Nguoi dung cuoi", "FILE_KEY_01"),
    ("THS Phase 3", "FILE_KEY_02"),
    ("THS Phase 1", "FILE_KEY_03"),
    ("THS - NKG", "FILE_KEY_04"),
    ("THS - CMS", "FILE_KEY_05"),
    ("AI Tutor [Web]", "FILE_KEY_06"),
    ("AI Tutor [App]", "FILE_KEY_07"),
    ("THS - Nguoi dung cuoi [App]", "FILE_KEY_08"),
    ("Quiz", "FILE_KEY_09"),
    ("Kho hoc lieu", "FILE_KEY_10"),
    ("Editor", "FILE_KEY_11"),
    ("Smart class", "FILE_KEY_12"),
    ("[Firefly] App", "FILE_KEY_13"),
    ("THS - Admin truong", "FILE_KEY_14"),
]

# Display names (Vietnamese) for the report
DISPLAY_NAMES = {
    "FILE_KEY_01": "THS - Nguoi dung cuoi",
    "FILE_KEY_02": "THS Phase 3",
    "FILE_KEY_03": "THS Phase 1",
    "FILE_KEY_04": "THS - NKG",
    "FILE_KEY_05": "THS - CMS",
    "FILE_KEY_06": "AI Tutor [Web]",
    "FILE_KEY_07": "AI Tutor [App]",
    "FILE_KEY_08": "THS - Nguoi dung cuoi [App]",
    "FILE_KEY_09": "Quiz",
    "FILE_KEY_10": "Kho hoc lieu",
    "FILE_KEY_11": "Editor",
    "FILE_KEY_12": "Smart class",
    "FILE_KEY_13": "[Firefly] App",
    "FILE_KEY_14": "THS - Admin truong",
}

SKIP_PAGES = {"draft", "thumbnail", "cover", "---", "------------------------", ""}


def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def figma_api(endpoint, max_retries=4):
    """Call Figma REST API with smart retry on rate limit."""
    url = f"https://api.figma.com/v1/{endpoint}"
    req = urllib.request.Request(url, headers={"X-Figma-Token": FIGMA_TOKEN})

    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429:
                # Read Retry-After header
                retry_after = e.headers.get("Retry-After", "60")
                try:
                    wait = int(retry_after)
                except ValueError:
                    wait = 60
                # Add buffer
                wait = max(wait, 30) + 5
                if attempt < max_retries - 1:
                    print(f"  [429] Rate limited. Retry-After={retry_after}s. Waiting {wait}s... (attempt {attempt+2}/{max_retries})")
                    time.sleep(wait)
                else:
                    print(f"  [429] Max retries reached. Skipping.")
                    return None
            elif e.code == 404:
                print(f"  [404] File not found or no access.")
                return None
            elif e.code == 403:
                print(f"  [403] No permission to access this file.")
                return None
            else:
                print(f"  [{e.code}] {e.reason}")
                return None
    return None


def count_frames_in_node(node):
    """Count FRAME children directly inside a node."""
    return sum(1 for c in node.get("children", []) if c.get("type") == "FRAME")


def count_sections_in_node(node):
    """Count SECTION children directly inside a node."""
    return sum(1 for c in node.get("children", []) if c.get("type") == "SECTION")


def count_all_frames_in_sections(page):
    """Count all FRAME nodes that are inside SECTIONs (screens)."""
    total = 0
    for child in page.get("children", []):
        if child.get("type") == "SECTION":
            total += count_frames_in_node(child)
            # Also check nested sections
            for sub in child.get("children", []):
                if sub.get("type") == "SECTION":
                    total += count_frames_in_node(sub)
    return total


def count_instances_recursive(node):
    """Recursively count INSTANCE nodes."""
    count = 1 if node.get("type") == "INSTANCE" else 0
    for child in node.get("children", []):
        count += count_instances_recursive(child)
    return count


def audit_file(name, file_id, cache):
    """Audit a single Figma file. Uses cache if available."""
    if file_id in cache:
        print(f"  [CACHED] {name} - using cached data")
        return cache[file_id]

    print(f"  [SCANNING] {name}...")

    # Step 1: Get structure (depth=3 for pages > sections > frames)
    data = figma_api(f"files/{file_id}?depth=3")
    if not data:
        return None

    file_result = {
        "name": name,
        "file_id": file_id,
        "modules": [],
        "totals": {"modules": 0, "features": 0, "screens": 0, "components": 0},
    }

    pages = data.get("document", {}).get("children", [])
    for page in pages:
        page_name = page.get("name", "").strip()
        if page_name.lower() in SKIP_PAGES:
            continue

        features_count = 0
        screens_count = 0
        feature_details = []

        for child in page.get("children", []):
            if child.get("type") == "SECTION":
                features_count += 1
                sec_name = child.get("name", "Unnamed")
                sec_screens = count_frames_in_node(child)
                # Check nested sections too
                for sub in child.get("children", []):
                    if sub.get("type") == "SECTION":
                        features_count += 1
                        sub_name = sub.get("name", "Unnamed")
                        sub_screens = count_frames_in_node(sub)
                        screens_count += sub_screens
                        feature_details.append({"name": sub_name, "screens": sub_screens})
                screens_count += sec_screens
                feature_details.append({"name": sec_name, "screens": sec_screens})

        module = {
            "name": page_name,
            "features": features_count,
            "screens": screens_count,
            "components": 0,  # Will be filled in component pass
            "feature_details": feature_details,
        }
        file_result["modules"].append(module)
        file_result["totals"]["modules"] += 1
        file_result["totals"]["features"] += features_count
        file_result["totals"]["screens"] += screens_count

    # Step 2: Count components (need deeper scan - but use same depth=3 data)
    # Count INSTANCE nodes from what we already have
    for page in pages:
        page_name = page.get("name", "").strip()
        if page_name.lower() in SKIP_PAGES:
            continue
        inst_count = count_instances_recursive(page)
        # Find matching module and update
        for mod in file_result["modules"]:
            if mod["name"] == page_name:
                mod["components"] = inst_count
                file_result["totals"]["components"] += inst_count
                break

    t = file_result["totals"]
    print(f"  [OK] {t['modules']} modules, {t['features']} features, {t['screens']} screens, {t['components']} components")

    # Cache the result
    cache[file_id] = file_result
    save_cache(cache)

    return file_result


def generate_html(results, output_path):
    """Generate HTML report."""
    grand = {"modules": 0, "features": 0, "screens": 0, "components": 0}
    for r in results:
        for k in grand:
            grand[k] += r["totals"][k]

    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    rows_summary = ""
    for i, r in enumerate(results):
        t = r["totals"]
        rows_summary += f"""
        <tr>
            <td>{i+1}</td>
            <td class="file-name"><a href="https://www.figma.com/design/{r['file_id']}" target="_blank">{r['name']}</a></td>
            <td class="num">{t['modules']}</td>
            <td class="num">{t['features']}</td>
            <td class="num">{t['screens']}</td>
            <td class="num">{t['components']}</td>
        </tr>"""

    rows_summary += f"""
        <tr class="total-row">
            <td colspan="2">TONG CONG</td>
            <td class="num">{grand['modules']}</td>
            <td class="num">{grand['features']}</td>
            <td class="num">{grand['screens']}</td>
            <td class="num">{grand['components']}</td>
        </tr>"""

    detail_sections = ""
    for r in results:
        module_rows = ""
        for m in r["modules"]:
            feature_list = ""
            if m["feature_details"]:
                items = "".join(
                    f"<li>{f['name']} <span class='badge'>{f['screens']} screens</span></li>"
                    for f in m["feature_details"]
                )
                feature_list = f"<ul class='feature-list'>{items}</ul>"

            module_rows += f"""
            <tr>
                <td class="module-name">{m['name']}</td>
                <td class="num">{m['features']}</td>
                <td class="num">{m['screens']}</td>
                <td class="num">{m['components']}</td>
                <td>{feature_list}</td>
            </tr>"""

        t = r["totals"]
        detail_sections += f"""
        <div class="detail-card">
            <div class="detail-header">
                <h3><a href="https://www.figma.com/design/{r['file_id']}" target="_blank">{r['name']}</a></h3>
                <div class="stats-pills">
                    <span class="pill modules">{t['modules']} Modules</span>
                    <span class="pill features">{t['features']} Features</span>
                    <span class="pill screens">{t['screens']} Screens</span>
                    <span class="pill components">{t['components']} Components</span>
                </div>
            </div>
            <table class="detail-table">
                <thead>
                    <tr>
                        <th>Module (Page)</th>
                        <th>Features</th>
                        <th>Screens</th>
                        <th>Components</th>
                        <th>Chi tiet Features</th>
                    </tr>
                </thead>
                <tbody>{module_rows}</tbody>
            </table>
        </div>"""

    html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Figma Audit Report - Truong Hoc So</title>
<style>
  :root {{
    --bg: #0f0f1a; --surface: #1a1a2e; --surface2: #222240;
    --border: #2d2d50; --text: #e8e8f0; --text-dim: #8888aa;
    --accent1: #6c63ff; --accent2: #00d4aa; --accent3: #ff6b9d; --accent4: #ffd93d;
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family:'Segoe UI',system-ui,sans-serif; background:var(--bg); color:var(--text); line-height:1.6; padding:2rem; }}
  .container {{ max-width:1400px; margin:0 auto; }}
  .header {{ text-align:center; padding:3rem 2rem; background:var(--surface); border-radius:20px; border:1px solid var(--border); margin-bottom:2rem; position:relative; overflow:hidden; }}
  .header::before {{ content:''; position:absolute; top:0; left:0; right:0; height:4px; background:linear-gradient(135deg,#6c63ff,#00d4aa); }}
  .header h1 {{ font-size:2.2rem; background:linear-gradient(135deg,#6c63ff,#00d4aa); -webkit-background-clip:text; -webkit-text-fill-color:transparent; margin-bottom:0.5rem; }}
  .header .subtitle {{ color:var(--text-dim); font-size:0.95rem; }}
  .grand-stats {{ display:grid; grid-template-columns:repeat(4,1fr); gap:1.5rem; margin-bottom:2.5rem; }}
  .stat-card {{ background:var(--surface); border:1px solid var(--border); border-radius:16px; padding:1.5rem; text-align:center; transition:transform 0.2s,box-shadow 0.2s; }}
  .stat-card:hover {{ transform:translateY(-4px); box-shadow:0 8px 30px rgba(108,99,255,0.15); }}
  .stat-card .number {{ font-size:2.8rem; font-weight:800; line-height:1; margin-bottom:0.3rem; }}
  .stat-card:nth-child(1) .number {{ color:var(--accent1); }}
  .stat-card:nth-child(2) .number {{ color:var(--accent2); }}
  .stat-card:nth-child(3) .number {{ color:var(--accent3); }}
  .stat-card:nth-child(4) .number {{ color:var(--accent4); }}
  .stat-card .label {{ color:var(--text-dim); font-size:0.9rem; text-transform:uppercase; letter-spacing:1px; }}
  .section-title {{ font-size:1.5rem; margin:2.5rem 0 1.5rem; padding-left:1rem; border-left:4px solid var(--accent1); }}
  table {{ width:100%; border-collapse:collapse; background:var(--surface); border-radius:12px; overflow:hidden; border:1px solid var(--border); }}
  th {{ background:var(--surface2); padding:1rem; text-align:left; font-weight:600; font-size:0.85rem; text-transform:uppercase; letter-spacing:0.5px; color:var(--text-dim); border-bottom:2px solid var(--border); }}
  td {{ padding:0.85rem 1rem; border-bottom:1px solid var(--border); font-size:0.95rem; }}
  tr:hover td {{ background:rgba(108,99,255,0.05); }}
  .num {{ text-align:center; font-weight:600; font-variant-numeric:tabular-nums; }}
  .file-name a {{ color:var(--accent1); text-decoration:none; font-weight:600; }}
  .file-name a:hover {{ text-decoration:underline; }}
  .total-row td {{ background:var(--surface2)!important; font-weight:700; font-size:1.05rem; border-top:2px solid var(--accent1); }}
  .detail-card {{ background:var(--surface); border:1px solid var(--border); border-radius:16px; margin-bottom:1.5rem; overflow:hidden; }}
  .detail-header {{ padding:1.2rem 1.5rem; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; background:var(--surface2); border-bottom:1px solid var(--border); }}
  .detail-header h3 {{ font-size:1.1rem; }}
  .detail-header h3 a {{ color:var(--text); text-decoration:none; }}
  .detail-header h3 a:hover {{ color:var(--accent1); }}
  .stats-pills {{ display:flex; gap:0.5rem; flex-wrap:wrap; }}
  .pill {{ padding:0.3rem 0.8rem; border-radius:20px; font-size:0.8rem; font-weight:600; }}
  .pill.modules {{ background:rgba(108,99,255,0.15); color:var(--accent1); }}
  .pill.features {{ background:rgba(0,212,170,0.15); color:var(--accent2); }}
  .pill.screens {{ background:rgba(255,107,157,0.15); color:var(--accent3); }}
  .pill.components {{ background:rgba(255,217,61,0.15); color:var(--accent4); }}
  .detail-table {{ border:none; border-radius:0; }}
  .module-name {{ font-weight:600; color:var(--accent1); }}
  .feature-list {{ list-style:none; display:flex; flex-wrap:wrap; gap:0.4rem; }}
  .feature-list li {{ background:rgba(0,212,170,0.08); padding:0.2rem 0.6rem; border-radius:6px; font-size:0.8rem; color:var(--text-dim); }}
  .badge {{ background:rgba(255,107,157,0.2); color:var(--accent3); padding:0.1rem 0.4rem; border-radius:4px; font-size:0.7rem; font-weight:600; margin-left:0.3rem; }}
  .footer {{ text-align:center; padding:2rem; color:var(--text-dim); font-size:0.85rem; }}
  @media (max-width:768px) {{ .grand-stats {{ grid-template-columns:repeat(2,1fr); }} body {{ padding:1rem; }} }}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>Figma Audit Report</h1>
    <p class="subtitle">Truong Hoc So -- Tong hop {len(results)} file Figma -- Cap nhat {now}</p>
  </div>
  <div class="grand-stats">
    <div class="stat-card"><div class="number">{grand['modules']}</div><div class="label">Modules</div></div>
    <div class="stat-card"><div class="number">{grand['features']}</div><div class="label">Features</div></div>
    <div class="stat-card"><div class="number">{grand['screens']}</div><div class="label">Screens</div></div>
    <div class="stat-card"><div class="number">{grand['components']}</div><div class="label">Components</div></div>
  </div>
  <h2 class="section-title">Bang Tong Hop</h2>
  <table>
    <thead><tr><th>#</th><th>File Figma</th><th>Modules</th><th>Features</th><th>Screens</th><th>Components</th></tr></thead>
    <tbody>{rows_summary}</tbody>
  </table>
  <h2 class="section-title">Chi Tiet Tung File</h2>
  {detail_sections}
  <div class="footer">Generated by Figma Audit Script v2 -- {now}</div>
</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n  Report saved: {output_path}")


def main():
    print("=" * 50)
    print("  FIGMA AUDIT v2 - Truong Hoc So")
    print("=" * 50)

    cache = load_cache()
    cached_count = sum(1 for _, fid in FILES if fid in cache)
    print(f"  Cache: {cached_count}/{len(FILES)} files already scanned")

    results = []
    for i, (name, file_id) in enumerate(FILES):
        print(f"\n[{i+1}/{len(FILES)}] {name}")
        result = audit_file(name, file_id, cache)
        if result:
            results.append(result)

        # Delay between API calls (skip if cached)
        if file_id not in cache and i < len(FILES) - 1:
            next_cached = FILES[i+1][1] in cache
            if not next_cached:
                print(f"  Waiting {DELAY_BETWEEN_FILES}s before next file...")
                time.sleep(DELAY_BETWEEN_FILES)

    if results:
        generate_html(results, "figma_audit_report.html")
        with open("figma_audit_data.json", "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"  JSON saved: figma_audit_data.json")

    print(f"\n  Done! Scanned {len(results)}/{len(FILES)} files.")
    if len(results) < len(FILES):
        print(f"  TIP: Run again to retry failed files (cached results preserved).")


if __name__ == "__main__":
    main()
