# GG Shipping Merger

Merges a Shopify **orders** export with a **charges** export, adding each order's
company shipping cost and the shipping difference, and produces a polished,
formatted Excel (`.xlsx`) file.

There are two ways to use it: a **web app** (easiest, for the client) and a
**command line** tool (for developers).

---

## Web app (for the client)

A zero-install browser front end — upload two CSVs, click Merge, download the
Excel file. No Python or terminal required.

**Live app:** <https://ggshippingmerger.streamlit.app/>

### How the client uses it
1. Open the app URL.
2. (If a password is set) enter the access password.
3. Upload the **Orders CSV** on the left and the **Charges CSV** on the right.
4. Click **Merge**, then **Download merged Excel file**.

### Deploying it (one-time, for the developer)
1. Push this repo to GitHub (already done).
2. Go to <https://share.streamlit.io>, sign in with GitHub, and click **New app**.
3. Pick this repo, branch `main`, main file `app.py`. Deploy.
4. **Set a password:** in the app's **Settings → Secrets**, paste:
   ```toml
   app_password = "something-only-the-client-knows"
   ```
   (See `.streamlit/secrets.toml.example`.) Leave it out for open access.
5. Copy the public URL and send it to the client. Every `git push` to `main`
   redeploys automatically.

> Note: Streamlit Community Cloud sleeps after inactivity, so the first load
> after a quiet period takes ~30 seconds to wake up.

---

## Command line (for developers)

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Argument order is always ORDERS first, then CHARGES:
python merge.py path/to/orders.csv path/to/charges.csv
python merge.py orders.csv charges.csv -o custom_output.xlsx
```

Output defaults to `Merged-CSVs/orders_with_shipping_cost_<date>.xlsx`.

### Auditing a result
`audit.py` verifies that a merged `.xlsx` faithfully preserves the source CSV:

```bash
python audit.py path/to/orders.csv path/to/output.xlsx
```

---

## Run the web app locally

```bash
source .venv/bin/activate
streamlit run app.py
```

Then open the URL it prints (usually <http://localhost:8501>).
