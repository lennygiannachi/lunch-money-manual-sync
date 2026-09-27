# lunch-money-manual-sync

Keep [Lunch Money](https://lunchmoney.app) manual accounts in sync with a Google Sheet.

Some balances never connect to Lunch Money automatically: a 401(k) portal, a crypto wallet, a car's value, a private loan. This Google Apps Script lets you track those balances in one sheet. From the sheet's menu you can create or update the matching manual accounts in Lunch Money through the v2 API.

## How it works

1. **Read.** The script reads the `Tracking` tab of the spreadsheet it's bound to.
2. **Group.** Rows with the same Account name (case- and whitespace-insensitive) are summed into one account, which takes the newest Last Updated in its group.
3. **Match.** The script lists your Lunch Money manual accounts and matches rows to them by name. If no name matches, it tries each account's display name.
4. **Plan.** Each account is put in one of three groups:
   - **create:** not in Lunch Money yet
   - **update:** balance, type, institution, subtype, display name or as-of date differs
   - **unchanged:** already matches
5. **Apply.** The script sends creates and updates to Lunch Money one at a time, with a short pause between calls.

The script never deletes or closes anything. Lunch Money accounts that aren't in the sheet are ignored.

An `onEdit` trigger stamps the current time into **Last Updated** whenever you edit a **Balance** cell. The sync sends that time as the account's `balance_as_of`.

## Sheet layout

The `Tracking` tab needs a header row. Columns are matched by header name (case doesn't matter) and can be in any order. The only exception is the Last Updated timestamp trigger, which uses fixed column positions (see [Configuration](#configuration)).

| Column | Required | Sent to Lunch Money as | Notes |
|---|---|---|---|
| Account | yes | `name` | Match key. Max 45 characters. |
| Platform | | `institution_name` | Max 50 characters. |
| Category | yes | `type` | See the category list below. |
| Subtype | | `subtype` | Free text, e.g. `Retirement`, `Student Loan`. |
| Balance | yes | `balance` | Numbers or text like `$1,234.56` / `(50.00)`. |
| Last Updated | | `balance_as_of` | Filled in automatically by the edit trigger. |
| Notes | | not sent | For your own reference. |

**Categories.** Each Category maps to a Lunch Money account type:

| Category values | Lunch Money type |
|---|---|
| cash, checking, savings, bank | `cash` |
| credit, credit card | `credit` |
| crypto, cryptocurrency | `cryptocurrency` |
| investment(s), brokerage, retirement | `investment` |
| equity, rsu, employee compensation | `employee compensation` |
| loan(s), mortgage | `loan` |
| liability, other liability | `other liability` |
| asset, other asset | `other asset` |
| property, real estate | `real estate` |
| car, vehicle | `vehicle` |

Add your own aliases in `LM_TYPE_MAP`.

**Behavior details**
- **Liabilities:** credit, loan and other-liability balances are sent exactly as entered. A positive number means the amount owed.
- **Crypto names:** for crypto rows that have a Platform, the Lunch Money display name is set to the plain Account name. That way an account named `ADA` shows as "ADA" and not "Coinbase ADA", while still grouping under Coinbase.
- **Skipped rows:** rows with an unknown Category, an unreadable Balance, a missing Account name, or conflicting Categories for the same name are skipped. They're listed under **PROBLEMS** in the results.

## Setup

1. Open your spreadsheet and go to **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with [`manual_tracking_sync.gs`](manual_tracking_sync.gs) and save.
   - If the project already has an `onOpen` or `onEdit` function, keep only one of each and merge them.
3. Reload the spreadsheet. A **Lunch Money** menu appears.
4. Choose **Lunch Money → Set API token…** and paste a token from <https://my.lunchmoney.app/developers>.
5. The first time you run anything, Google asks you to authorize the script.
   - You'll see a "Google hasn't verified this app" warning because the script is yours and unpublished. Click **Advanced → Go to (project name) (unsafe) → Allow**.

The token is saved in your Apps Script [user properties](https://developers.google.com/apps-script/guides/properties). It isn't stored in the sheet or the code, and other editors can't read it.

## Usage

| Menu item | What it does |
|---|---|
| **Preview sync…** | Dry run. Shows what would be created or updated, with an **Apply to Lunch Money** button. Apply re-reads the sheet first, so edits made after previewing are included. |
| **Sync now** | Applies immediately, then shows the results. |
| **Set API token…** | Saves or replaces your Lunch Money token. |

A typical flow: update balances in the sheet (Last Updated fills in by itself), run **Preview sync…**, check the plan, then click **Apply**.

### Scheduled sync

To sync without opening the sheet, go to **Triggers → Add Trigger** in the Apps Script editor. Choose the `lmScheduledSync` function, a **Time-driven** event source, and an interval such as daily.

The scheduled run doesn't open any dialogs. If any Lunch Money call fails, the run fails too, so Google emails you the error.

## Configuration

All settings are in the `CONFIG` section at the top of the script.

| Constant | Default | Purpose |
|---|---|---|
| `LM_TAB_NAME` | `Tracking` | Tab to read |
| `LM_BALANCE_COL` | `5` (E) | Column the edit trigger watches |
| `LM_LAST_UPDATED_COL` | `6` (F) | Column the edit trigger stamps |
| `LM_DEFAULT_CURRENCY` | `usd` | Currency for newly created accounts |
| `LM_PAUSE_MS` | `350` | Delay between API calls |
| `LM_TYPE_MAP` | see above | Category → account type aliases |

## Notes

- **Timestamps on multi-row pastes:** the edit trigger only stamps the first row of an edit. When you paste several balances at once, only the top row gets a new Last Updated.
- **Changing an account:** renaming an Account in the sheet creates a *new* Lunch Money account, because the name is the match key. Rename it in Lunch Money too, or set its display name to the new name.
- **Blank fields:** a blank Platform or Subtype never clears an existing value in Lunch Money.
