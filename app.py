"""Shopify store control panel: sync store data to SQLite, print reports.

Usage:
  python app.py sync                 # pull products, inventory, orders into shopify.db
  python app.py report bestsellers   # top products by units sold
  python app.py report lowstock      # variants at/below LOW_STOCK threshold
  python app.py report outofstock    # products with every variant at 0
  python app.py report revenue       # daily + monthly revenue
  python app.py dashboard            # write dashboard.html (charts) and open it
  python app.py excel                # export shopify.xlsx (data + dashboard sheet)
  python app.py selftest             # run built-in check on fake data
"""
import json
import os
import sqlite3
import sys
import webbrowser

import requests

DB = os.path.join(os.path.dirname(__file__), "shopify.db")
API_VERSION = "2026-07"
LOW_STOCK = 5

SCHEMA = """
CREATE TABLE IF NOT EXISTS products (
  id TEXT PRIMARY KEY, title TEXT, status TEXT
);
CREATE TABLE IF NOT EXISTS variants (
  id TEXT PRIMARY KEY, product_id TEXT REFERENCES products(id),
  sku TEXT, title TEXT, price REAL, inventory INTEGER
);
CREATE TABLE IF NOT EXISTS orders (
  id TEXT PRIMARY KEY, name TEXT, created_at TEXT, total REAL, financial_status TEXT
);
CREATE TABLE IF NOT EXISTS order_items (
  id TEXT PRIMARY KEY, order_id TEXT REFERENCES orders(id), product_id TEXT,
  title TEXT, quantity INTEGER, price REAL
);
"""

TZ = "+3 hours"  # store timezone (Riyadh); UTC timestamps shifted before day/month grouping


def load_env():
    # ponytail: tiny .env parser instead of python-dotenv
    path = os.path.join(os.path.dirname(__file__), ".env")
    if not os.path.exists(path):
        sys.exit("missing .env — copy .env.example and fill it in")
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


_token = None


def get_token(store):
    """Legacy shpat_ token if set, else client-credentials exchange (2026+ apps)."""
    global _token
    if os.environ.get("SHOPIFY_TOKEN"):
        return os.environ["SHOPIFY_TOKEN"]
    if _token is None:
        r = requests.post(
            f"https://{store}/admin/oauth/access_token",
            data={
                "grant_type": "client_credentials",
                "client_id": os.environ["SHOPIFY_CLIENT_ID"],
                "client_secret": os.environ["SHOPIFY_CLIENT_SECRET"],
            },
            timeout=30,
        )
        r.raise_for_status()
        _token = r.json()["access_token"]  # valid 24h; CLI runs are short
    return _token


def gql(query, variables=None):
    store = os.environ["SHOPIFY_STORE"]  # e.g. mystore.myshopify.com
    token = get_token(store)
    r = requests.post(
        f"https://{store}/admin/api/{API_VERSION}/graphql.json",
        json={"query": query, "variables": variables or {}},
        headers={"X-Shopify-Access-Token": token},
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    if data.get("errors"):
        raise RuntimeError(data["errors"])
    return data["data"]


def paginate(query, root, variables=None):
    """Yield nodes from a cursor-paginated connection at data[root]."""
    cursor = None
    while True:
        data = gql(query, {**(variables or {}), "cursor": cursor})[root]
        for edge in data["edges"]:
            yield edge["node"]
        if not data["pageInfo"]["hasNextPage"]:
            return
        cursor = data["edges"][-1]["cursor"]


PRODUCTS_Q = """
query($cursor: String) {
  products(first: 100, after: $cursor) {
    pageInfo { hasNextPage }
    edges { cursor node {
      id title status
      variants(first: 100) { edges { node {
        id sku title price inventoryQuantity
      } } }
    } }
  }
}
"""

ORDERS_Q = """
query($cursor: String) {
  orders(first: 100, after: $cursor) {
    pageInfo { hasNextPage }
    edges { cursor node {
      id name createdAt displayFinancialStatus
      totalPriceSet { shopMoney { amount } }
      lineItems(first: 100) { edges { node {
        id title quantity product { id }
        originalUnitPriceSet { shopMoney { amount } }
      } } }
    } }
  }
}
"""


def sync(db):
    db.executescript(SCHEMA)
    if not any(c[1] == "id" for c in db.execute("PRAGMA table_info(order_items)")):
        db.execute("DROP TABLE order_items")  # migrate pre-line-item-id schema
        db.executescript(SCHEMA)
    # full reload in one transaction: drops records deleted in Shopify,
    # and a failed sync rolls back leaving the previous data intact
    for t in ("order_items", "orders", "variants", "products"):
        db.execute(f"DELETE FROM {t}")
    n = 0
    for p in paginate(PRODUCTS_Q, "products"):
        db.execute(
            "INSERT OR REPLACE INTO products VALUES (?,?,?)",
            (p["id"], p["title"], p["status"]),
        )
        for v in (e["node"] for e in p["variants"]["edges"]):
            db.execute(
                "INSERT OR REPLACE INTO variants VALUES (?,?,?,?,?,?)",
                (v["id"], p["id"], v["sku"], v["title"],
                 float(v["price"]), v["inventoryQuantity"]),
            )
        n += 1
    print(f"synced {n} products")

    n = 0
    for o in paginate(ORDERS_Q, "orders"):
        db.execute(
            "INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?)",
            (o["id"], o["name"], o["createdAt"],
             float(o["totalPriceSet"]["shopMoney"]["amount"]),
             o["displayFinancialStatus"]),
        )
        for li in (e["node"] for e in o["lineItems"]["edges"]):
            db.execute(
                "INSERT OR REPLACE INTO order_items VALUES (?,?,?,?,?,?)",
                (li["id"], o["id"], (li["product"] or {}).get("id"),
                 li["title"], li["quantity"],
                 float(li["originalUnitPriceSet"]["shopMoney"]["amount"])),
            )
        n += 1
    db.commit()
    print(f"synced {n} orders")


REPORTS = {
    "bestsellers": (
        "Top products by units sold",
        """SELECT title, SUM(quantity) units, ROUND(SUM(quantity*price),2) revenue
           FROM order_items GROUP BY title ORDER BY units DESC LIMIT 10""",
    ),
    "lowstock": (
        f"Variants with inventory <= {LOW_STOCK}",
        f"""SELECT p.title, v.title variant, v.sku, v.inventory
            FROM variants v JOIN products p ON p.id = v.product_id
            WHERE v.inventory <= {LOW_STOCK} AND p.status = 'ACTIVE'
            ORDER BY v.inventory""",
    ),
    "outofstock": (
        "Active products with every variant at 0 inventory",
        """SELECT p.title, COUNT(*) variants
           FROM products p JOIN variants v ON v.product_id = p.id
           WHERE p.status = 'ACTIVE'
           GROUP BY p.id HAVING MAX(v.inventory) <= 0 ORDER BY p.title""",
    ),
    "revenue": (
        "Revenue by day (last 30) and by month",
        f"""SELECT substr(datetime(created_at,'{TZ}'),1,10) day,
            COUNT(*) orders, ROUND(SUM(total),2) revenue
            FROM orders WHERE financial_status NOT IN ('REFUNDED','VOIDED')
            GROUP BY day ORDER BY day DESC LIMIT 30""",
    ),
}

REVENUE_MONTHLY = f"""SELECT substr(datetime(created_at,'{TZ}'),1,7) month,
                      COUNT(*) orders, ROUND(SUM(total),2) revenue
                      FROM orders WHERE financial_status NOT IN ('REFUNDED','VOIDED')
                      GROUP BY month ORDER BY month DESC"""


DASHBOARD_TPL = """<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__SHOP__ — dashboard</title>
<style>
:root{
  --page:#f9f9f7; --surface:#fcfcfb; --ink:#0b0b0b; --ink2:#52514e;
  --muted:#898781; --grid:#e1e0d9; --axis:#c3c2b7; --series:#2a78d6;
  --border:rgba(11,11,11,.10);
}
@media (prefers-color-scheme: dark){:root{
  --page:#0d0d0d; --surface:#1a1a19; --ink:#ffffff; --ink2:#c3c2b7;
  --muted:#898781; --grid:#2c2c2a; --axis:#383835; --series:#3987e5;
  --border:rgba(255,255,255,.10);
}}
*{box-sizing:border-box;margin:0}
body{background:var(--page);color:var(--ink);
  font:14px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif;padding:24px}
h1{font-size:18px;font-weight:600}
.sub{color:var(--muted);font-size:12px;margin:2px 0 20px}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
  gap:12px;margin-bottom:16px}
.tile,.card{background:var(--surface);border:1px solid var(--border);
  border-radius:10px;padding:14px 16px}
.tile .l{color:var(--ink2);font-size:12px}
.tile .v{font-size:26px;font-weight:600;margin-top:2px}
.cards{display:grid;grid-template-columns:1fr;gap:16px;max-width:900px}
.card h2{font-size:13px;font-weight:600;margin-bottom:2px}
.card .u{color:var(--muted);font-size:11px;margin-bottom:10px}
svg{display:block;width:100%;height:auto}
svg text{font:11px system-ui,-apple-system,"Segoe UI",sans-serif;fill:var(--muted)}
svg .val{fill:var(--ink2)}
details{margin-top:8px}summary{color:var(--muted);font-size:11px;cursor:pointer}
table{border-collapse:collapse;font-size:12px;margin-top:6px}
td,th{padding:3px 10px 3px 0;text-align:right;font-variant-numeric:tabular-nums}
td:first-child,th:first-child{text-align:left}
th{color:var(--muted);font-weight:500}
#tip{position:fixed;pointer-events:none;background:var(--surface);
  border:1px solid var(--border);border-radius:8px;padding:6px 10px;
  font-size:12px;box-shadow:0 2px 8px rgba(0,0,0,.12);display:none;z-index:9}
#tip .tv{font-weight:600;font-size:13px}#tip .tl{color:var(--ink2)}
</style></head><body>
<h1>__SHOP__ — store dashboard</h1>
<div class="sub">generated __DATE__ · revenue excludes refunded and voided orders · times in store timezone</div>
<div class="tiles" id="tiles"></div>
<div class="cards">
<div class="card"><h2>Revenue by day</h2><div class="u">last 30 days with orders, SAR</div>
  <div id="daily"></div><details><summary>Data table</summary><div id="dailyT"></div></details></div>
<div class="card"><h2>Revenue by month</h2><div class="u">SAR</div>
  <div id="monthly"></div><details><summary>Data table</summary><div id="monthlyT"></div></details></div>
<div class="card"><h2>Top products by units sold</h2><div class="u">units; revenue in tooltip</div>
  <div id="tops"></div><details><summary>Data table</summary><div id="topsT"></div></details></div>
</div>
<div id="tip"></div>
<script>
const D=__DATA__;
const css=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const fmt=n=>n>=1e6?(n/1e6).toFixed(1)+"M":n>=1e3?(n/1e3).toFixed(1)+"K":Math.round(n)+"";
const full=n=>Math.round(n).toLocaleString("en")+" SAR";
const S=(t,a)=>{const e=document.createElementNS("http://www.w3.org/2000/svg",t);
  for(const k in a)e.setAttribute(k,a[k]);return e};
const tip=document.getElementById("tip");
function showTip(ev,lines){tip.style.display="block";tip.replaceChildren();
  lines.forEach(([c,t])=>{const d=document.createElement("div");d.className=c;
    d.textContent=t;tip.appendChild(d)});
  const x=Math.min(ev.clientX+14,innerWidth-tip.offsetWidth-8);
  tip.style.left=x+"px";tip.style.top=(ev.clientY+14)+"px"}
const hideTip=()=>tip.style.display="none";
function nice(m){const p=Math.pow(10,Math.floor(Math.log10(m||1)));
  for(const s of[1,2,5,10])if(m<=s*p)return s*p;return 10*p}
function table(id,heads,rows){const t=document.createElement("table");
  const tr=document.createElement("tr");
  heads.forEach(h=>{const th=document.createElement("th");th.textContent=h;tr.appendChild(th)});
  t.appendChild(tr);
  rows.forEach(r=>{const tr=document.createElement("tr");
    r.forEach(c=>{const td=document.createElement("td");td.textContent=c;tr.appendChild(td)});
    t.appendChild(tr)});
  document.getElementById(id).appendChild(t)}
// stat tiles
D.tiles.forEach(([l,v])=>{const d=document.createElement("div");d.className="tile";
  const a=document.createElement("div");a.className="l";a.textContent=l;
  const b=document.createElement("div");b.className="v";b.textContent=v;
  d.append(a,b);document.getElementById("tiles").appendChild(d)});
// daily line chart
(function(){
  const days=D.daily;if(!days.length)return;
  const W=840,H=240,L=46,R=14,T=14,B=26,iw=W-L-R,ih=H-T-B;
  const max=nice(Math.max(...days.map(d=>d[2])));
  const X=i=>L+(days.length<2?iw/2:i*iw/(days.length-1)),Y=v=>T+ih-v/max*ih;
  const svg=S("svg",{viewBox:`0 0 ${W} ${H}`});
  for(let g=0;g<=4;g++){const y=T+ih*g/4;
    svg.appendChild(S("line",{x1:L,x2:W-R,y1:y,y2:y,stroke:g==4?css("--axis"):css("--grid"),"stroke-width":1}));
    const t=S("text",{x:L-6,y:y+4,"text-anchor":"end"});t.textContent=fmt(max*(4-g)/4);svg.appendChild(t)}
  [0,days.length-1].forEach(i=>{if(i<0)return;
    const t=S("text",{x:X(i),y:H-8,"text-anchor":i?"end":"start"});t.textContent=days[i][0].slice(5);svg.appendChild(t)});
  const pts=days.map((d,i)=>`${X(i)},${Y(d[2])}`).join(" ");
  svg.appendChild(S("polygon",{points:`${L},${T+ih} ${pts} ${X(days.length-1)},${T+ih}`,
    fill:css("--series"),opacity:.1}));
  svg.appendChild(S("polyline",{points:pts,fill:"none",stroke:css("--series"),
    "stroke-width":2,"stroke-linejoin":"round","stroke-linecap":"round"}));
  const last=days.length-1;
  svg.appendChild(S("circle",{cx:X(last),cy:Y(days[last][2]),r:4,fill:css("--series"),
    stroke:css("--surface"),"stroke-width":2}));
  const el=S("text",{x:X(last)-8,y:Y(days[last][2])-10,"text-anchor":"end","class":"val"});
  el.textContent=fmt(days[last][2]);svg.appendChild(el);
  const cross=S("line",{y1:T,y2:T+ih,stroke:css("--axis"),"stroke-width":1,visibility:"hidden"});
  const dot=S("circle",{r:4,fill:css("--series"),stroke:css("--surface"),
    "stroke-width":2,visibility:"hidden"});
  svg.append(cross,dot);
  svg.addEventListener("pointermove",ev=>{
    const r=svg.getBoundingClientRect(),mx=(ev.clientX-r.left)*W/r.width;
    let i=Math.round((mx-L)/(iw/Math.max(days.length-1,1)));
    i=Math.max(0,Math.min(days.length-1,i));
    cross.setAttribute("x1",X(i));cross.setAttribute("x2",X(i));
    cross.setAttribute("visibility","visible");
    dot.setAttribute("cx",X(i));dot.setAttribute("cy",Y(days[i][2]));
    dot.setAttribute("visibility","visible");
    showTip(ev,[["tv",full(days[i][2])],["tl",days[i][0]+" · "+days[i][1]+" order(s)"]])});
  svg.addEventListener("pointerleave",()=>{hideTip();
    cross.setAttribute("visibility","hidden");dot.setAttribute("visibility","hidden")});
  document.getElementById("daily").appendChild(svg);
  table("dailyT",["Day","Orders","Revenue (SAR)"],
    days.map(d=>[d[0],d[1],Math.round(d[2]).toLocaleString("en")]))})();
// monthly columns
(function(){
  const ms=D.monthly;if(!ms.length)return;
  const W=840,H=200,L=46,R=14,T=18,B=26,iw=W-L-R,ih=H-T-B;
  const max=nice(Math.max(...ms.map(m=>m[2])));
  const band=iw/ms.length,bw=Math.min(24,band*.6);
  const svg=S("svg",{viewBox:`0 0 ${W} ${H}`});
  for(let g=0;g<=4;g++){const y=T+ih*g/4;
    svg.appendChild(S("line",{x1:L,x2:W-R,y1:y,y2:y,stroke:g==4?css("--axis"):css("--grid"),"stroke-width":1}));
    const t=S("text",{x:L-6,y:y+4,"text-anchor":"end"});t.textContent=fmt(max*(4-g)/4);svg.appendChild(t)}
  ms.forEach((m,i)=>{
    const x=L+band*i+(band-bw)/2,h=Math.max(m[2]/max*ih,2),y=T+ih-h;
    const bar=S("path",{d:`M${x},${T+ih} V${y+4} Q${x},${y} ${x+4},${y} H${x+bw-4}
      Q${x+bw},${y} ${x+bw},${y+4} V${T+ih} Z`,fill:css("--series")});
    const hit=S("rect",{x:L+band*i,y:T,width:band,height:ih,fill:"transparent"});
    hit.addEventListener("pointermove",ev=>{bar.setAttribute("opacity",".8");
      showTip(ev,[["tv",full(m[2])],["tl",m[0]+" · "+m[1]+" order(s)"]])});
    hit.addEventListener("pointerleave",()=>{bar.removeAttribute("opacity");hideTip()});
    const cap=S("text",{x:x+bw/2,y:y-6,"text-anchor":"middle","class":"val"});
    cap.textContent=fmt(m[2]);
    const mo=S("text",{x:x+bw/2,y:H-8,"text-anchor":"middle"});mo.textContent=m[0];
    svg.append(bar,cap,mo,hit)});
  document.getElementById("monthly").appendChild(svg);
  table("monthlyT",["Month","Orders","Revenue (SAR)"],
    ms.map(m=>[m[0],m[1],Math.round(m[2]).toLocaleString("en")]))})();
// top products bars
(function(){
  const ts=D.tops;if(!ts.length)return;
  const W=840,L=170,R=60,row=30,bw=18,H=ts.length*row+8;
  const iw=W-L-R,max=nice(Math.max(...ts.map(t=>t[1])));
  const svg=S("svg",{viewBox:`0 0 ${W} ${H}`});
  ts.forEach((p,i)=>{
    const y=8+i*row,w=Math.max(p[1]/max*iw,2);
    const lbl=S("text",{x:L-10,y:y+bw/2+4,"text-anchor":"end"});
    lbl.textContent=p[0].length>24?p[0].slice(0,23)+"…":p[0];
    const bar=S("path",{d:`M${L},${y} H${L+w-4} Q${L+w},${y} ${L+w},${y+4} V${y+bw-4}
      Q${L+w},${y+bw} ${L+w-4},${y+bw} H${L} Z`,fill:css("--series")});
    const val=S("text",{x:L+w+8,y:y+bw/2+4,"class":"val"});val.textContent=p[1];
    const hit=S("rect",{x:0,y:y-4,width:W,height:row,fill:"transparent"});
    hit.addEventListener("pointermove",ev=>{bar.setAttribute("opacity",".8");
      showTip(ev,[["tv",p[1]+" units"],["tl",p[0]+" · "+full(p[2])]])});
    hit.addEventListener("pointerleave",()=>{bar.removeAttribute("opacity");hideTip()});
    svg.append(lbl,bar,val,hit)});
  svg.appendChild(S("line",{x1:L,x2:L,y1:4,y2:H-4,stroke:css("--axis"),"stroke-width":1}));
  document.getElementById("tops").appendChild(svg);
  table("topsT",["Product","Units","Revenue (SAR)"],
    ts.map(p=>[p[0],p[1],Math.round(p[2]).toLocaleString("en")]))})();
</script></body></html>
"""


def render_dashboard(db, shop, date):
    daily = db.execute(REPORTS["revenue"][1]).fetchall()[::-1]  # oldest first
    monthly = db.execute(REVENUE_MONTHLY).fetchall()[::-1]
    tops = db.execute(REPORTS["bestsellers"][1]).fetchall()
    oos = db.execute("""SELECT COUNT(*) FROM variants v JOIN products p
                        ON p.id=v.product_id
                        WHERE v.inventory <= 0 AND p.status='ACTIVE'""").fetchone()[0]
    cur = monthly[-1] if monthly else ("-", 0, 0)
    tiles = [
        ["Revenue this month (SAR)", f"{cur[2]:,.0f}"],
        ["Orders this month", str(cur[1])],
        ["Best seller", tops[0][0] if tops else "-"],
        ["Out-of-stock variants", f"{oos:,}"],
    ]
    data = {"tiles": tiles, "daily": [list(r) for r in daily],
            "monthly": [list(r) for r in monthly], "tops": [list(r) for r in tops]}
    payload = json.dumps(data).replace("</", "<\\/")
    return (DASHBOARD_TPL.replace("__SHOP__", shop)
            .replace("__DATE__", date).replace("__DATA__", payload))


def dashboard(db):
    import datetime
    html = render_dashboard(db, os.environ.get("SHOPIFY_STORE", "store"),
                            datetime.date.today().isoformat())
    path = os.path.join(os.path.dirname(__file__), "dashboard.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print("wrote", path)
    webbrowser.open("file:///" + path.replace("\\", "/"))


def print_rows(db, title, query):
    cur = db.execute(query)
    cols = [d[0] for d in cur.description]
    rows = cur.fetchall()
    print(f"\n== {title} ==")
    if not rows:
        print("(no rows)")
        return
    widths = [max(len(str(x)) for x in [c] + [r[i] for r in rows])
              for i, c in enumerate(cols)]
    print("  ".join(c.ljust(w) for c, w in zip(cols, widths)))
    for r in rows:
        print("  ".join(str(x).ljust(w) for x, w in zip(r, widths)))


def report(db, name):
    title, query = REPORTS[name]
    print_rows(db, title, query)
    if name == "revenue":
        print_rows(db, "Revenue by month", REVENUE_MONTHLY)


def selftest():
    db = sqlite3.connect(":memory:")
    db.executescript(SCHEMA)
    db.execute("INSERT INTO products VALUES ('p1','Mug','ACTIVE')")
    db.execute("INSERT INTO variants VALUES ('v1','p1','MUG-1','Default',10.0,2)")
    # 22:30 UTC = 01:30 next day in +3 — must group under 2026-07-02
    db.execute("INSERT INTO orders VALUES ('o1','#1001','2026-07-01T22:30:00Z',30.0,'PAID')")
    # refunded order — must not count toward revenue
    db.execute("INSERT INTO orders VALUES ('o2','#1002','2026-07-01T10:00:00Z',99.0,'REFUNDED')")
    # two line items with identical product+title — must both count
    db.execute("INSERT INTO order_items VALUES ('li1','o1','p1','Mug',2,10.0)")
    db.execute("INSERT INTO order_items VALUES ('li2','o1','p1','Mug',1,10.0)")
    best = db.execute(REPORTS["bestsellers"][1]).fetchone()
    assert best == ("Mug", 3, 30.0), best
    low = db.execute(REPORTS["lowstock"][1]).fetchone()
    assert low == ("Mug", "Default", "MUG-1", 2), low
    # Mug still has 2 units, so it must NOT appear as out of stock
    assert db.execute(REPORTS["outofstock"][1]).fetchone() is None
    day = db.execute(REPORTS["revenue"][1]).fetchone()
    assert day == ("2026-07-02", 1, 30.0), day
    month = db.execute(REVENUE_MONTHLY).fetchone()
    assert month == ("2026-07", 1, 30.0), month
    html = render_dashboard(db, "test", "2026-07-12")
    assert "Mug" in html and "<svg" not in html  # svg is built client-side
    assert '"daily": [["2026-07-02"' in html, "daily series missing"
    print("selftest OK")


def main():
    args = sys.argv[1:]
    if args == ["selftest"]:
        return selftest()
    load_env()
    db = sqlite3.connect(DB)
    if args == ["sync"]:
        sync(db)
    elif args == ["dashboard"]:
        dashboard(db)
    elif args == ["excel"]:
        import excel  # lazy: needs openpyxl, the rest of the tool doesn't
        excel.export(db)
    elif len(args) == 2 and args[0] == "report" and args[1] in REPORTS:
        report(db, args[1])
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
