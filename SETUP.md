# Google Sheet setup (one time, ~5 minutes)

1. Go to https://sheets.new and create a Google Sheet (name it e.g. "Spink Records").
2. In the sheet: **Extensions > Apps Script**. Delete the default code and paste in everything from `Code.gs`.
3. **Deploy > New deployment > type: Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click Deploy, approve the permissions, and copy the **Web app URL** (ends in `/exec`).
4. Open the calculator, tap **⚙ Sheet**, paste the URL, and Save.
   (This is stored in that browser only; repeat on each phone/computer.)

The sheet gets two tabs automatically:
- **Records** – one row per customer (total, deposit paid, paid, balance, status).
- **Payments** – one row per payment (date, amount, type, collected by). A customer's first payment is marked "Deposit", later ones "Payment".

If you set up an earlier version of the script, delete the old "Records" and "Payments" tabs once so they are recreated with the new columns.

If you change `Code.gs` later: Deploy > Manage deployments > edit > New version.

## Using it
- Fill in the calculator, then tick **Save to sheet** (top-right of Summary).
- When the mommy pays: tap **Records**, pick her, enter **Payment now**, date and **Collected by**, then tick **Save to sheet** again.
- Editing anything after a save un-ticks the box, so nothing is updated until you tick it again.
