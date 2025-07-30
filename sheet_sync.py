import time
from googleapiclient.discovery import build
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path
import traceback

# Google Sheets API settings
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1MupgKr69jm1M4g79cuZXkd9wovjuyQcdI-G6vpc7Ql4'  # Change this to your sheet ID
HEADER_RANGE = 'Payment!C1:ZZ1'  # Header row range (building names)
DATA_RANGE = 'Payment!C2:ZZ300'  # Data range

def get_google_sheets_service():
    """Create a connection to Google Sheets API"""
    try:
        print("Reading credentials.json...")
        if not os.path.exists('credentials.json'):
            raise FileNotFoundError("credentials.json file not found in the folder")
        creds = service_account.Credentials.from_service_account_file(
            'credentials.json', scopes=SCOPES)
        print("credentials.json loaded successfully")
        print("Creating Google Sheets service...")
        service = build('sheets', 'v4', credentials=creds)
        print("Google Sheets service created successfully")
        return service
    except Exception as e:
        print(f"\nError: {str(e)}")
        print("\nError details:")
        traceback.print_exc()
        raise

def find_buildings_structure(header_values):
    """Detect buildings structure in the sheet"""
    if not header_values or len(header_values) < 1:
        print("No data in the sheet")
        return []
    header_row = header_values[0]
    buildings = []
    for col_index, cell_value in enumerate(header_row):
        if cell_value and cell_value.strip():
            current_building = {
                'name': cell_value.strip(),
                'client_col': col_index,
                'status_col': col_index + 1,
                'notes_col': col_index + 2
            }
            buildings.append(current_building)
    # Only print summary, not every building
    print(f"Found {len(buildings)} buildings")
    return buildings

def get_column_letter(index):
    """Convert column number to letter (e.g., 0->A, 1->B, 26->AA)"""
    result = ""
    while index >= 0:
        remainder = index % 26
        result = chr(65 + remainder) + result  # 65 is ASCII code for 'A'
        index = index // 26 - 1
    return result

def update_sheet_status(service, row_index: int, new_status: str, status_col: int):
    """Update client status in the sheet"""
    try:
        col_letter = get_column_letter(status_col + 2)
        range_name = f'Payment!{col_letter}{row_index + 2}'
        print(f"Updating cell {range_name} to {new_status}")
        body = {
            'values': [[new_status]]
        }
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        print(f"Client status updated in cell {range_name}")
    except Exception as e:
        print(f"Error updating client status in sheet: {str(e)}")

def process_client(service, client_name: str, status: str, building_info: dict, row_index: int, mikrotik_manager, values):
    """Process a single client and update status"""
    if not client_name or not status:
        return
    print(f"\nProcessing client: {client_name} | Status: {status}")
    try:
        if status.upper() == 'CUT':
            if mikrotik_manager.block_client(client_name):
                update_sheet_status(service, row_index, 'BLOCK', building_info['status_col'])
                print(f"Client {client_name} blocked and status updated")
            else:
                print(f"Failed to block client {client_name}")
        elif status.upper().startswith('LIMIT='):
            block_date = values[row_index][building_info['notes_col']]
            if not block_date:
                print(f"No block date specified for client {client_name}")
                return
            try:
                day, month, year = block_date.split('-')
                # Support flexible date formats: 5-6-2025 or 05-06-2025
                if not (1 <= len(day) <= 2 and 1 <= len(month) <= 2 and len(year) == 4):
                    raise ValueError
                # Validate day and month ranges
                if not (1 <= int(day) <= 31 and 1 <= int(month) <= 12):
                    raise ValueError
            except:
                print(f"Invalid date format for client {client_name}. Should be D-M-YYYY or DD-MM-YYYY")
                return
            mac_address = mikrotik_manager.get_client_mac(client_name)
            if not mac_address:
                print(f"MAC address not found for client {client_name}")
                return
            if mikrotik_manager.schedule_block_client(client_name, mac_address, block_date):
                if mikrotik_manager.activate_client(client_name):
                    new_status = status.split('=')[1]
                    update_sheet_status(service, row_index, new_status, building_info['status_col'])
                    print(f"Client {client_name} activated and scheduled to be blocked at {block_date}")
                else:
                    print(f"Failed to activate client {client_name} after scheduling block")
            else:
                # فشل الجدولة - قد يكون بسبب تاريخ قديم
                update_sheet_status(service, row_index, 'Old-Date', building_info['status_col'])
                print(f"Failed to schedule block for client {client_name} - Date might be in the past")
        elif status.upper() == 'NOT PAY':
            # Check if client is already active first
            if mikrotik_manager.is_client_active(client_name):
                # Client is already active, just change status to 'Waiting'
                update_sheet_status(service, row_index, 'Waiting', building_info['status_col'])
                print(f"Client {client_name} is already active, status updated to Waiting")
            elif mikrotik_manager.is_client_blocked(client_name):
                # Client is blocked, activate and set to 'Waiting'
                if mikrotik_manager.activate_client(client_name):
                    update_sheet_status(service, row_index, 'Waiting', building_info['status_col'])
                    print(f"Client {client_name} activated in Mikrotik and status set to Waiting")
                else:
                    print(f"Failed to activate client {client_name}")
            else:
                print(f"Client {client_name} not found in Mikrotik")
        elif status.upper().startswith('NEW='):
            mac_or_ip = values[row_index][building_info['notes_col']]
            if not mac_or_ip:
                print(f"No MAC address or IP for client {client_name}")
                update_sheet_status(service, row_index, 'NO-MAC', building_info['status_col'])
                return
            
            mac_or_ip = mac_or_ip.strip()  # Remove any whitespace
            print(f"Processing NEW client {client_name} with MAC/IP: {mac_or_ip}")
            
            success, error = mikrotik_manager.add_new_client(client_name, mac_or_ip)
            if success:
                new_status = status.split('=')[1]
                if new_status.upper() == 'NOT-PAY':
                    # تحديث الحالة إلى Not Pay
                    update_sheet_status(service, row_index, 'Not Pay', building_info['status_col'])
                    print(f"Client {client_name} added successfully and status set to Not Pay")
                else:
                    # تحديث الحالة العادية
                    update_sheet_status(service, row_index, new_status, building_info['status_col'])
                    print(f"Client {client_name} added successfully and status updated to {new_status}")
            else:
                update_sheet_status(service, row_index, error, building_info['status_col'])
                print(f"Error adding client {client_name}: {error}")
        elif status.upper().startswith('UPD-MAC='):
            mac_or_ip = values[row_index][building_info['notes_col']]
            if not mac_or_ip:
                print(f"No MAC address or IP for client {client_name}")
                update_sheet_status(service, row_index, 'NO-MAC', building_info['status_col'])
                return
            
            mac_or_ip = mac_or_ip.strip()  # Remove any whitespace
            print(f"Updating MAC for client {client_name} with new MAC/IP: {mac_or_ip}")
            
            success, error = mikrotik_manager.update_client_mac(client_name, mac_or_ip)
            if success:
                # التحقق من نوع الحالة المطلوبة
                new_status = status.split('=')[1]
                if new_status.upper() == 'NOT-PAY':
                    # تحديث الحالة إلى Not Pay
                    update_sheet_status(service, row_index, 'Not Pay', building_info['status_col'])
                    print(f"Client {client_name} MAC updated successfully and status set to Not Pay")
                else:
                    # تحديث الحالة العادية
                    update_sheet_status(service, row_index, new_status, building_info['status_col'])
                    print(f"Client {client_name} MAC updated successfully and status updated to {new_status}")
            else:
                update_sheet_status(service, row_index, error, building_info['status_col'])
                print(f"Error updating client {client_name} MAC: {error}")
        elif status.upper().startswith('ACTIVATE='):
            # Check if client is already active
            if mikrotik_manager.is_client_active(client_name):
                # Client is already active, just update status
                new_status = status.split('=')[1]
                if new_status.upper() == 'NOT-PAY':
                    # تحديث الحالة إلى Not Pay
                    update_sheet_status(service, row_index, 'Not Pay', building_info['status_col'])
                    print(f"Client {client_name} is already active, status updated to Not Pay")
                else:
                    # تحديث الحالة العادية
                    update_sheet_status(service, row_index, new_status, building_info['status_col'])
                    print(f"Client {client_name} is already active, status updated to {new_status}")
            else:
                # Try to activate the client
                if mikrotik_manager.activate_client(client_name):
                    new_status = status.split('=')[1]
                    if new_status.upper() == 'NOT-PAY':
                        # تحديث الحالة إلى Not Pay
                        update_sheet_status(service, row_index, 'Not Pay', building_info['status_col'])
                        print(f"Client {client_name} activated and status set to Not Pay")
                    else:
                        # تحديث الحالة العادية
                        update_sheet_status(service, row_index, new_status, building_info['status_col'])
                        print(f"Client {client_name} activated and status updated to {new_status}")
                else:
                    print(f"Failed to activate client {client_name}")
        elif status.upper().startswith('RE=NAME'):
            # Update client name (part before @ in comment)
            print(f"[DEBUG] Processing RE=NAME for client: {client_name}")
            # Check if notes column exists in the row
            if len(values[row_index]) <= building_info['notes_col']:
                print(f"[ERROR] Notes column {building_info['notes_col']} not found in row {row_index}")
                update_sheet_status(service, row_index, 'NO-NOTES-COL', building_info['status_col'])
                return
            new_name = values[row_index][building_info['notes_col']]
            print(f"[DEBUG] Raw new_name from notes column: '{new_name}'")
            if not new_name:
                print(f"No new name specified for client {client_name}")
                update_sheet_status(service, row_index, 'NO-NAME', building_info['status_col'])
                return
            new_name = new_name.strip()  # Remove any whitespace
            print(f"Updating name for client '{client_name}' to: '{new_name}'")
            success, error = mikrotik_manager.update_client_name(client_name, new_name)
            if success:
                update_sheet_status(service, row_index, 'Not Pay', building_info['status_col'])
                # تحديث عمود اسم العميل في Google Sheets ليطابق الاسم الجديد
                try:
                    client_col_letter = get_column_letter(building_info['client_col'] + 2)
                    client_name_range = f'Payment!{client_col_letter}{row_index + 2}'
                    body = {'values': [[new_name]]}
                    service.spreadsheets().values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=client_name_range,
                        valueInputOption='USER_ENTERED',
                        body=body
                    ).execute()
                    print(f"[DEBUG] Updated client name in Google Sheets to '{new_name}' at {client_name_range}")
                except Exception as e:
                    print(f"[ERROR] Failed to update client name in Google Sheets: {str(e)}")
                # تحديث خانة الملاحظات برقم الهاتف من ميكروتك بدون تغيير أي تنسيق
                try:
                    phone = mikrotik_manager.get_client_phone(new_name)
                    notes_col_letter = get_column_letter(building_info['notes_col'] + 2)
                    notes_range = f'Payment!{notes_col_letter}{row_index + 2}'
                    # استخدم valueInputOption USER_ENTERED فقط بدون أي تنسيق
                    body = {'values': [[phone]]}
                    service.spreadsheets().values().update(
                        spreadsheetId=SPREADSHEET_ID,
                        range=notes_range,
                        valueInputOption='USER_ENTERED',
                        body=body
                    ).execute()
                    print(f"[DEBUG] Updated notes column in Google Sheets to phone '{phone}' at {notes_range}")
                except Exception as e:
                    print(f"[ERROR] Failed to update notes column in Google Sheets: {str(e)}")
                print(f"Client {client_name} name updated successfully to {new_name}")
            else:
                update_sheet_status(service, row_index, error, building_info['status_col'])
                print(f"Error updating client {client_name} name: {error}")
        elif status.upper().startswith('RE=NUMBER'):
            # Update client phone number (part after @ in comment)
            new_phone = values[row_index][building_info['notes_col']]
            if not new_phone:
                print(f"No new phone number specified for client {client_name}")
                update_sheet_status(service, row_index, 'NO-PHONE', building_info['status_col'])
                return
            new_phone = new_phone.strip()  # Remove any whitespace
            print(f"Updating phone for client {client_name} to: {new_phone}")
            success, error = mikrotik_manager.update_client_phone(client_name, new_phone)
            if success:
                update_sheet_status(service, row_index, 'Not Pay', building_info['status_col'])
                print(f"Client {client_name} phone updated successfully to {new_phone}")
            else:
                update_sheet_status(service, row_index, error, building_info['status_col'])
                print(f"Error updating client {client_name} phone: {error}")
    except Exception as e:
        print(f"Error processing client {client_name}: {str(e)}")
        traceback.print_exc()

def process_sheet_data(service, mikrotik_manager):
    """Process Payment sheet data"""
    try:
        print("\nReading Google Sheets data...")
        header_result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=HEADER_RANGE
        ).execute()
        header_values = header_result.get('values', [])
        data_result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=DATA_RANGE
        ).execute()
        values = data_result.get('values', [])
        if not values:
            print('No data!')
            return
        buildings = find_buildings_structure(header_values)
        if not buildings:
            print("No buildings found in the sheet!")
            return
        for building in buildings:
            # Only print actionable buildings (not every building)
            actionable = False
            for row_index, row in enumerate(values):
                try:
                    if len(row) <= building['client_col']:
                        continue
                    client_name = row[building['client_col']]
                    if not client_name:
                        continue
                    if len(row) <= building['status_col']:
                        continue
                    status = row[building['status_col']]
                    if not status:
                        continue
                    if (status.upper() == 'CUT' or
                        status.upper() == 'NOT PAY' or
                        status.upper().startswith('NEW=') or
                        status.upper().startswith('UPD-MAC=') or
                        status.upper().startswith('ACTIVATE=') or
                        status.upper().startswith('LIMIT=') or
                        status.upper().startswith('RE=NAME') or
                        status.upper().startswith('RE=NUMBER')):
                        if not actionable:
                            print(f"\nProcessing building: {building['name']}")
                            actionable = True
                        process_client(service, client_name, status, building, row_index, mikrotik_manager, values)
                except Exception as e:
                    print(f"Error processing row {row_index + 2}: {str(e)}")
                    traceback.print_exc()
                    continue
    except Exception as e:
        print(f"Error processing data: {str(e)}")
        traceback.print_exc()
