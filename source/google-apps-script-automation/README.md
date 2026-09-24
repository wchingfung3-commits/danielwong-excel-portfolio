# Google Sheets Order & Reporting Automation

A self-initiated Google Apps Script portfolio project by Daniel Wong. It demonstrates how a small business order sheet can be standardized, validated, summarized and prepared for review without sending any messages automatically.

## What the script does

- Adds an **Operations Automation** menu to Google Sheets.
- Builds a synthetic demo workbook.
- Standardizes customer, email, region and transaction-status fields.
- Flags duplicate IDs, invalid emails, invalid quantities, invalid prices, pending orders and excluded transactions.
- Preserves every source row with a clear **Ready** or **Review** status.
- Refreshes a KPI dashboard and regional revenue chart.
- Creates a Gmail draft only after the user enters a recipient and runs the command.
- Records each action in an audit log.

## Installation

1. Create a blank Google Sheet.
2. Open **Extensions → Apps Script**.
3. Replace the default code with `Code.gs`.
4. Open **Project Settings**, enable the manifest file, and use the supplied `appsscript.json` settings if required.
5. Save the project and reload the spreadsheet.
6. Use **Operations Automation → Set up demo workbook**.
7. Review requested Google permissions before authorizing. The Gmail permission is used only to create a draft; the script does not send email.

## Safe handover notes

- Use synthetic or anonymized data for testing.
- Review the clean output and dashboard before relying on the figures.
- Do not add time-based triggers until the workflow has been tested in the owner's account.
- Apps Script quotas and Google Workspace administrator policies can affect execution.

## Portfolio disclosure

This is a self-initiated demonstration using synthetic data. It does not represent a paid client engagement.
