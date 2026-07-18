# Shopify Store Control Panel

A small Python CLI that syncs a Shopify store (products, inventory, orders)
into a local SQLite database and produces reports and dashboards:
best sellers, low stock, daily/monthly revenue — as terminal tables,
a self-contained HTML dashboard, or an Excel workbook with live formulas.

## Requirements

- Python 3.10+
- `pip install -r requirements.txt` (`requests`; `openpyxl` only for Excel export)

## Setup

1. In the [Shopify Dev Dashboard](https://dev.shopify.com), create an app.
2. **Versions → Create version** → set API access scopes to
   `read_products,read_orders,read_inventory` (add `read_all_orders` for order
   history beyond 60 days) → **Release**.
3. Install the app on your store and approve the permissions.
4. Copy `.env.example` to `.env` and fill in your store's `.myshopify.com`
   domain plus the app's Client ID and Secret (from the app's Settings page).

The app exchanges the Client ID/Secret for a 24-hour Admin API token
automatically on every run (client credentials grant).

## Usage

```
python app.py sync                 # pull store data into shopify.db
python app.py report bestsellers   # top products by units sold
python app.py report lowstock      # variants at/below the LOW_STOCK threshold
python app.py report outofstock    # products with every variant at 0
python app.py report revenue       # daily + monthly revenue
python app.py dashboard            # write dashboard.html and open it
python app.py excel                # export shopify.xlsx (data + dashboard)
python app.py selftest             # run the built-in checks (no network)
```

Revenue reports exclude refunded/voided orders and group days/months in the
store's timezone (`TZ` in `app.py`, default UTC+3).
