import sys, io
from playwright.sync_api import sync_playwright
BASE="https://uat.truonghocsovietnam.vn"
path, btn, tag = sys.argv[1], sys.argv[2], sys.argv[3]
JSF=io.open("fields.js",encoding="utf-8").read()
JST=io.open("text.js",encoding="utf-8").read()
with sync_playwright() as p:
    b=p.chromium.launch(channel="msedge", headless=True)
    ctx=b.new_context(storage_state="_rs/auth.json", viewport={"width":1680,"height":1600}, locale="vi-VN")
    pg=ctx.new_page()
    pg.goto(BASE+path, wait_until="domcontentloaded", timeout=60000); pg.wait_for_timeout(7000)
    try:
        pg.get_by_role("button", name=btn).first.click(timeout=15000)
    except Exception as e:
        print("!! click thất bại:", str(e)[:150])
    pg.wait_for_timeout(6000)
    out=[f"=== {path}  [click: {btn}] -> {pg.url}"]
    tabs=pg.eval_on_selector_all("[role=tab],.MuiTab-root,.MuiStep-root", JST)
    out.append("--- TABS/STEPS: " + " | ".join(dict.fromkeys(tabs))[:900])
    labs=pg.eval_on_selector_all("label,.MuiFormLabel-root,legend,h6,.MuiTypography-subtitle1,.MuiTypography-subtitle2", JST)
    out.append(f"--- LABELS/NHÓM ({len(set(labs))}) ---")
    out += ["  "+x for x in dict.fromkeys(labs)]
    f=pg.eval_on_selector_all("input,select,textarea", JSF)
    out.append(f"--- FIELDS ({len(f)}) ---"); out += ["  "+x for x in dict.fromkeys(f)]
    pg.screenshot(path=f"_rs/{tag}.png", full_page=True)
    b.close()
r="\n".join(out); io.open(f"_rs/{tag}.txt","w",encoding="utf-8").write(r); print(r[:6000])
