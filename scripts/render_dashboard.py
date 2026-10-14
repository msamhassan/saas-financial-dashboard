"""Builds the 4 report pages from dashboard/page*.html + metrics.json and exports a 2x PNG of each."""
import os
from playwright.sync_api import sync_playwright

B = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
DASH = f"{B}/dashboard"
PAGES = [("page1_overview.html", "Overview", "01_finance_overview.png"),
         ("page2_decomposition.html", "Decomp.<br>Tree", "02_decomposition_tree.png"),
         ("page3_sales.html", "Sales<br>Team", "03_sales_performance.png"),
         ("page4_customers.html", "Customer<br>Insights", "04_customer_insights.png")]
CSS, JS = open(f"{DASH}/shared.css").read(), open(f"{DASH}/shared.js").read()
DATA = open(f"{B}/data/processed/metrics.json").read()
os.makedirs(f"{B}/images", exist_ok=True); os.makedirs(f"{DASH}/build", exist_ok=True)


def nav(active):
    items = "".join(f'<div class="pg{" on" if i == active else ""}">{p[1]}</div>' for i, p in enumerate(PAGES))
    return f'<div class="nav"><div class="logo">BI</div>{items}<div class="pg">Recon</div></div>'


with sync_playwright() as p:
    br = p.chromium.launch()
    pg = br.new_page(viewport={"width": 1600, "height": 900}, device_scale_factor=2)
    for i, (tpl, _, png) in enumerate(PAGES):
        html = open(f"{DASH}/{tpl}").read().replace("/*__CSS__*/", CSS).replace("/*__HELPERS__*/", JS) \
            .replace("/*__DATA__*/", DATA).replace("__NAV__", nav(i))
        out = f"{DASH}/build/{tpl}"
        open(out, "w").write(html)
        errs = []
        pg.on("pageerror", lambda e: errs.append(str(e)))
        pg.goto(f"file://{out}"); pg.wait_for_timeout(400)
        assert not errs, f"{tpl}: {errs}"
        pg.locator(".frame").screenshot(path=f"{B}/images/{png}")
        print("rendered", png)
    br.close()
