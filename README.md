# Creekside Lions PTO · Spirit Store Tracker

## Using it (everyone)

1. Open the link. Anyone can look. No login.
2. To change a number, tap **Unlock to edit**, type your name and the PIN.
3. Type in the yellow boxes.
4. Tap **Submit changes** at the bottom. Done. Everyone sees it within 30 seconds.

The **Overview** tab lists what still needs filling in. Tap an item and it takes you there.

Volunteer PIN: counts, deposits, cash box. Admin PIN: everything, plus Undo on the Activity tab.

## Where the numbers live

A Google Sheet in the PTO's Google account. The website reads it and saves to it. Nobody has to open the Sheet day to day. It also keeps every change in its Change Log tab, and File > Version history can roll anything back.

## One-time setup

**Google Sheet (signed in to the PTO's Google account)**

1. drive.google.com > New > File upload > choose `Creekside Spirit Store.xlsx`. Double-click it, then File > Save as Google Sheets. Use that new copy from here on.
2. Extensions > Apps Script. Select everything there, delete it, paste in all of `Code.gs`, click the Save icon.
3. Go back to the Sheet tab and reload it. A **Spirit Store** menu appears. Spirit Store > **Set PINs**. Google asks to authorize: Continue, pick the account, then **Advanced > Go to project (unsafe)** > Allow. That warning is normal for your own script. Then enter the two PINs.
4. Back in Apps Script: **Deploy > New deployment**. Gear icon > **Web app**. Execute as: **Me**. Who has access: **Anyone**. Deploy. Copy the **Web app URL** (ends in /exec).

**Website**

5. In `index.html`, paste that URL between the quotes on the `const API_URL = '';` line.
6. Put `index.html` (and `logo.png` if you have the official logo) in a public GitHub repository and turn on Settings > Pages > Deploy from branch > main > / (root). The site link shows on that page.

## Good to know

- Change a PIN anytime: Spirit Store > Set PINs. Works right away.
- Want to hide the numbers from the public: Spirit Store > Require an access code to view.
- After editing Code.gs later: Deploy > Manage deployments > pencil > Version: New version > Deploy. Same link.
- PINs keep honest people from fat-fingering things. They are not bank security.
