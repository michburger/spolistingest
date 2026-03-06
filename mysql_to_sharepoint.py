import argparse
import pymysql
from msal import ConfidentialClientApplication
import requests

def map_mysql_to_sharepoint(mysql_row, columns):
    # Map MySQL row data to SharePoint list item fields
    # Combines id, vorname, and nachname into Title field
    # Excludes id, vorname, nachname from direct mapping since they're used for Title
    # CREATE TABLE `rve25_ruderkurs` (
    # `id` int(11) NOT NULL,
    # `esk` int(11) NOT NULL, // 0 = Erwachsene, 1 = Studierende, 2 = Kinder
    # `nachname` varchar(100) NOT NULL,
    # `vorname` varchar(100) NOT NULL,
    # `adresse` varchar(255) NOT NULL,
    # `email` varchar(100) NOT NULL,
    # `telefon` varchar(20) NOT NULL,
    # `geburtstag` date DEFAULT NULL,
    # `nachricht` varchar(255) NOT NULL,
    # `schwimmen` int(11) NOT NULL,
    # `datenschutz` int(11) NOT NULL,
    # `erzeugt` timestamp NOT NULL DEFAULT current_timestamp()
    # )
    row_dict = {col: val for col, val in zip(columns, mysql_row)}
    
    # Extract components for Title field
    row_id = str(row_dict.get('id', ''))
    vorname = str(row_dict.get('vorname', ''))
    nachname = str(row_dict.get('nachname', ''))
    # Combine into Title
    title = f"{row_id} - {nachname}, {vorname}"
    esk = int(row_dict.get('esk', -1))
    kurs_mapping = {0: 2, 1: 4, 2: 3}
    kurs_value = int(kurs_mapping.get(esk, -1))  # Default to -1 if esk value is unexpected
    print(f"Mapping MySQL row to SharePoint item: Title='{title}', Kurs='{kurs_value}', ESK='{esk}'")
    result = {
        'Title': title,
        'KursLookupId': kurs_value,
        'Status': 'Anfrage',
        'Vorname': vorname,
        'Nachname': nachname,
        'e_x002d_mail': str(row_dict.get('email', '')),
        'Adresse': str(row_dict.get('adresse', '')),
        'Zusatzangaben': str(row_dict.get('nachricht', '')),
        'Telefon': str(row_dict.get('telefon', '')),
        'Geburtsdatum': str(row_dict.get('geburtstag', '')),
        }    
    return result

def item_exists_in_list(site_id, list_id, title, headers):
    """Check if an item with the given Title already exists in the SharePoint list"""
    try:
        # Escape single quotes in title for OData filter
        escaped_title = title.replace("'", "''")
        response = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?$filter=fields/Title eq '{escaped_title}'",
            headers=headers
        )
        response.raise_for_status()
        items = response.json().get('value', [])
        return len(items) > 0
    except Exception as e:
        print(f"Error checking if item exists: {e}")
        print(f"Error details: {e.response.text if hasattr(e, 'response') else 'No response'}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Read from MySQL table and upload to SharePoint list via Graph API")
    parser.add_argument('--db-host', default='localhost', help='MySQL host')
    parser.add_argument('--db-port', type=int, default=3306, help='MySQL port')
    parser.add_argument('--db-socket', help='Path to MySQL unix socket (overrides host/port when provided)')
    parser.add_argument('--db-user', required=True, help='MySQL username')
    parser.add_argument('--db-password', required=True, help='MySQL password')
    parser.add_argument('--db-name', required=True, help='MySQL database name')
    parser.add_argument('--table', default='rve25_ruderkurs', help='MySQL table name')
    parser.add_argument('--client-id', required=True, help='Azure app client ID')
    parser.add_argument('--client-secret', required=True, help='Azure app client secret')
    parser.add_argument('--tenant-id', required=True, help='Azure tenant ID')
    parser.add_argument('--site-url', required=True, help='SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)')
    parser.add_argument('--list-name', required=True, help='SharePoint list display name')

    args = parser.parse_args()

    # Connect to MySQL and read data
    try:
        connect_kwargs = {
            'user': args.db_user,
            'password': args.db_password,
            'database': args.db_name
        }
        # prefer unix socket when provided
        if args.db_socket:
            connect_kwargs['unix_socket'] = args.db_socket
        else:
            connect_kwargs['host'] = args.db_host
            connect_kwargs['port'] = args.db_port

        conn = pymysql.connect(**connect_kwargs)
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM {args.table}")
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        print(f"Read {len(rows)} rows from table {args.table}")
    except Exception as e:
        print(f"Error reading from MySQL: {e}")
        return
    finally:
        if 'conn' in locals():
            conn.close()

    # Authenticate with Microsoft Graph API
    try:
        app = ConfidentialClientApplication(
            args.client_id,
            authority=f"https://login.microsoftonline.com/{args.tenant_id}",
            client_credential=args.client_secret
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            raise Exception(f"Could not acquire token: {result.get('error_description', 'Unknown error')}")
        token = result["access_token"]
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
    except Exception as e:
        print(f"Error authenticating with Graph API: {e}")
        return

    # Get site ID
    try:
        # Parse site URL to get hostname and path
        site_url_clean = args.site_url.replace('https://', '').replace('http://', '')
        if '/' in site_url_clean:
            hostname, path = site_url_clean.split('/', 1)
            site_path = f"{hostname}:/{path}"
        else:
            site_path = site_url_clean
        site_response = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_path}", headers=headers)
        site_response.raise_for_status()
        site_id = site_response.json()['id']
    except Exception as e:
        print(f"Error getting site ID: {e}")
        return

    # Get list ID
    try:
        list_response = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$filter=displayName eq '{args.list_name}'",
            headers=headers
        )
        list_response.raise_for_status()
        lists = list_response.json()['value']
        if not lists:
            raise Exception(f"List '{args.list_name}' not found")
        list_id = lists[0]['id']
    except Exception as e:
        print(f"Error getting list ID: {e}")
        return

    # Upload data to SharePoint list
    success_count = 0
    skipped_count = 0
    for row in rows:
        try:
            # Map MySQL row to SharePoint fields using custom mapping
            item_data = map_mysql_to_sharepoint(row, columns)
            title = item_data['Title']
            
            # Check if item with this Title already exists
            if item_exists_in_list(site_id, list_id, title, headers):
                print(f"Skipped: Item with Title '{title}' already exists")
                skipped_count += 1
                continue
            
            # Create new item in SharePoint list
            response = requests.post(
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items",
                headers=headers,
                json={"fields": item_data}
            )
            response.raise_for_status()
            success_count += 1
        except Exception as e:
            print(f"Error creating list item for row {row}: {e}")
            print(f"Error details: {e.response.text if hasattr(e, 'response') else 'No response'}")

    print(f"Successfully uploaded {success_count} items to SharePoint list '{args.list_name}'")
    if skipped_count > 0:
        print(f"Skipped {skipped_count} items (already existing)")

if __name__ == "__main__":
    main()