# MySQL to SharePoint List Migrator

This Python script reads data from a MySQL database table and uploads it to a SharePoint list using Microsoft Graph API with client credentials authentication.

## Prerequisites

1. Python 3.6 or higher
2. MySQL database with the target table
3. Azure app registration with:
   - Client ID and Client Secret
   - Appropriate permissions for SharePoint (Sites.ReadWrite.All)
4. SharePoint site and list created

## Installation

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the script with the required command-line arguments:

```bash
python mysql_to_sharepoint.py \
  --db-user your_mysql_user \
  --db-password your_mysql_password \
  --db-name your_database \
  --table your_table_name \
  --client-id your_azure_client_id \
  --client-secret your_azure_client_secret \
  --tenant-id your_azure_tenant_id \
  --site-url https://yourtenant.sharepoint.com/sites/yoursite \
  --list-name "Your List Name"
```

### Arguments

- `--db-host`: MySQL host (default: localhost)
- `--db-port`: MySQL port (default: 3306)
- `--db-user`: MySQL username (required)
- `--db-password`: MySQL password (required)
- `--db-name`: MySQL database name (required)
- `--table`: MySQL table name (default: rve25_ruderkurs)
- `--client-id`: Azure app client ID (required)
- `--client-secret`: Azure app client secret (required)
- `--tenant-id`: Azure tenant ID (required)
- `--site-url`: SharePoint site URL (required, e.g., https://tenant.sharepoint.com/sites/sitename)
- `--list-name`: SharePoint list display name (required)

## How it works

1. Connects to the specified MySQL database and reads all rows from the table
2. Authenticates with Microsoft Graph API using client credentials
3. Retrieves the SharePoint site and list IDs
4. For each row from MySQL:
   - Combines the `id`, `vorname` (first name), and `nachname` (last name) columns into a Title field with format: `{id} - {nachname}, {vorname}`
   - Checks if an item with this Title already exists in the SharePoint list
   - Creates a new list item only if it doesn't exist (prevents duplicates)
   - Maps all other columns (except id, vorname, nachname) directly to SharePoint fields

## Notes

- The SharePoint list **must have a Title field** (standard for most SharePoint lists)
- The Title field is created by combining: `id - nachname, vorname`
- If a Title already exists in the list, that row is skipped (no duplicates)
- All other MySQL column names must match the SharePoint field names (case-sensitive)
- All values are converted to strings when uploading to SharePoint
- The columns `id`, `vorname`, and `nachname` are excluded from direct field mapping (combined into Title only)
- The script assumes the Azure app has the necessary permissions to write to the SharePoint list

## Troubleshooting

- Ensure your Azure app has Sites.ReadWrite.All permission
- Verify the SharePoint site URL format
- Check that the list name matches exactly
- Confirm MySQL credentials and table existence