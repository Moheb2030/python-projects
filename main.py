import time
from sheet_sync import get_google_sheets_service, process_sheet_data
from mikrotik_manager import MikrotikManager

def main():
    # Mikrotik connection settings
    MIKROTIK_HOST = '10.0.0.1'
    MIKROTIK_USER = 'api_user'
    MIKROTIK_PASS = 'Api2024#Secure'
    MIKROTIK_PORT = 8728

    while True:
        try:
            print("\n[INFO] Starting new sync cycle...")
            # Create Google Sheets connection
            service = get_google_sheets_service()
            # Create Mikrotik connection
            mikrotik = MikrotikManager(
                MIKROTIK_HOST, 
                MIKROTIK_USER, 
                MIKROTIK_PASS,
                MIKROTIK_PORT
            )
            # Process data
            process_sheet_data(service, mikrotik)
            print("[INFO] Sync cycle finished. Waiting 2 minutes before next run...")
        except Exception as e:
            print(f"[ERROR] {str(e)}")
        # Wait 2 minutes before next run
        try:
            time.sleep(120)
        except KeyboardInterrupt:
            print("\n[INFO] Program stopped by user.")
            break

if __name__ == "__main__":
    main()
