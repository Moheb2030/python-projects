import routeros_api
import time
import re
from datetime import datetime

class MikrotikManager:
    def __init__(self, host, username, password, port=8728):
        """Initialize connection to Mikrotik"""
        self.connection = routeros_api.RouterOsApiPool(
            host,
            username=username,
            password=password,
            port=port,
            plaintext_login=True
        )
        self.api = self.connection.get_api()
        # Load client list once during initialization
        self.clients_cache = self._load_clients()
        
    def _load_clients(self) -> dict:
        """Load client list from Mikrotik and store in memory"""
        clients = {}
        try:
            list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
            bindings = list_ip_bindings.get()
            for binding in bindings:
                comment = binding.get('comment', '')
                if comment:
                    # Extract client name from comment (before ' @')
                    client_name = self._extract_client_name(comment)
                    clients[client_name] = binding
            print(f"Loaded {len(clients)} clients from MikroTik")
            return clients
        except Exception as e:
            print(f"Error loading clients: {str(e)}")
            return {}
            
    def _extract_client_name(self, full_comment: str) -> str:
        """Extract client name from full comment"""
        # Split text at ' @' (space then @) and take the first part
        return full_comment.split(' @')[0].strip()
        
    def find_client_in_ip_bindings(self, client_name: str) -> dict:
        """Search for client in IP bindings"""
        # Try exact match first
        exact_match = self.clients_cache.get(client_name.strip())
        if exact_match:
            return exact_match
        
        # If exact match fails, try flexible search
        return self.find_client_flexible(client_name)
        
    def refresh_clients_cache(self):
        """Update client list in memory"""
        self.clients_cache = self._load_clients()
        
    def block_client(self, client_name: str) -> bool:
        """Block client in MikroTik"""
        try:
            binding = self.find_client_in_ip_bindings(client_name)
            if binding:
                list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
                list_ip_bindings.set(id=binding['id'], type='blocked')
                print(f"Client {client_name} blocked successfully")
                return True
            else:
                print(f"Client {client_name} not found")
                return False
        except Exception as e:
            print(f"Error blocking client {client_name}: {str(e)}")
            return False
            
    def activate_client(self, client_name: str) -> bool:
        """Activate client in MikroTik"""
        try:
            binding = self.find_client_in_ip_bindings(client_name)
            if binding:
                list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
                list_ip_bindings.set(id=binding['id'], type='bypassed')
                print(f"Client {client_name} activated successfully")
                return True
            else:
                print(f"Client {client_name} not found")
                return False
        except Exception as e:
            print(f"Error activating client {client_name}: {str(e)}")
            return False
            
    def add_new_client(self, client_name: str, mac_or_ip: str) -> tuple:
        """Add new client to MikroTik - supports MAC or IP address"""
        try:
            # Check if client already exists
            if self.find_client_in_ip_bindings(client_name):
                print(f"Client {client_name} already exists")
                return False, "Already-Exists"
            
            # Determine if input is IP or MAC
            mac_address = None
            
            if self.is_ip_address(mac_or_ip):
                # If IP address, search for associated MAC
                print(f"Input is IP address: {mac_or_ip}")
                mac_address = self.find_mac_by_ip(mac_or_ip)
                if not mac_address:
                    print(f"MAC not found for IP: {mac_or_ip}")
                    return False, "IP-Not-Found"
                print(f"Will use MAC: {mac_address} (obtained from IP: {mac_or_ip})")
            else:
                # If MAC address
                mac_address = mac_or_ip
                print(f"Input is MAC address: {mac_address}")
            
            # الفحص الجديد: التحقق من وجود MAC address في IP bindings
            existing_mac_binding = self.find_mac_in_ip_bindings(mac_address)
            
            if existing_mac_binding:
                existing_comment = existing_mac_binding.get('comment', '')
                
                # إذا كان مستخدم غير مخول (ZZZZ=Blocked unauthorized)
                if self.is_unauthorized_user(existing_comment):
                    print(f"Found unauthorized user with MAC {mac_address}, deleting it first...")
                    if self.delete_ip_binding_by_id(existing_mac_binding['id']):
                        print(f"Deleted unauthorized user binding, proceeding with adding new client")
                        # تحديث الكاش بعد الحذف
                        self.refresh_clients_cache()
                    else:
                        print(f"Failed to delete unauthorized user binding")
                        return False, "Delete-Failed"
                else:
                    # MAC address ينتمي لعميل حالي
                    print(f"MAC {mac_address} already belongs to an existing client: {existing_comment}")
                    return False, "Already-Exists"
                
            # Add new client with MAC only (no IP binding)
            comment = client_name
            
            list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
            
            # Create binding with MAC only
            list_ip_bindings.add(
                **{
                    'mac-address': mac_address,
                    'type': 'bypassed',
                    'comment': comment
                }
            )
            
            # Update cache
            self.refresh_clients_cache()
            print(f"Client {client_name} added successfully with MAC: {mac_address}")
            return True, None
            
        except Exception as e:
            error_str = str(e).lower()
            if "invalid value of mac-address" in error_str:
                print(f"Invalid MAC address for client {client_name}")
                return False, "Wrong-MAC"
            print(f"Error adding new client: {str(e)}")
            return False, "Wrong-MAC"
            
    def update_client_mac(self, client_name: str, mac_or_ip: str) -> tuple:
        """Update MAC address for existing client in MikroTik - supports MAC or IP address"""
        try:
            # Check if client exists
            existing_binding = self.find_client_in_ip_bindings(client_name)
            if not existing_binding:
                print(f"Client {client_name} not found in MikroTik")
                return False, "Client-Not-Found"
            
            # Determine if input is IP or MAC
            mac_address = None
            
            if self.is_ip_address(mac_or_ip):
                # If IP address, search for associated MAC
                print(f"Input is IP address: {mac_or_ip}")
                mac_address = self.find_mac_by_ip(mac_or_ip)
                if not mac_address:
                    print(f"MAC not found for IP: {mac_or_ip}")
                    return False, "IP-Not-Found"
                print(f"Will update to MAC: {mac_address} (obtained from IP: {mac_or_ip})")
            else:
                # If MAC address
                mac_address = mac_or_ip
                print(f"Input is MAC address: {mac_address}")
            
            # الفحص الجديد: التحقق من وجود MAC address في IP bindings
            existing_mac_binding = self.find_mac_in_ip_bindings(mac_address)
            
            if existing_mac_binding:
                # التأكد من أن هذا MAC ليس للعميل نفسه
                if existing_mac_binding['id'] != existing_binding['id']:
                    existing_comment = existing_mac_binding.get('comment', '')
                    
                    # إذا كان مستخدم غير مخول (ZZZZ=Blocked unauthorized)
                    if self.is_unauthorized_user(existing_comment):
                        print(f"Found unauthorized user with MAC {mac_address}, deleting it first...")
                        if self.delete_ip_binding_by_id(existing_mac_binding['id']):
                            print(f"Deleted unauthorized user binding, proceeding with MAC update")
                            # تحديث الكاش بعد الحذف
                            self.refresh_clients_cache()
                        else:
                            print(f"Failed to delete unauthorized user binding")
                            return False, "Delete-Failed"
                    else:
                        # MAC address ينتمي لعميل حالي آخر
                        print(f"MAC {mac_address} already belongs to another client: {existing_comment}")
                        return False, "Already-Exists"
                
            # Check current client status
            current_type = existing_binding.get('type', 'bypassed')
            is_blocked = (current_type == 'blocked')
            
            # Update existing client's MAC address
            list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
            
            # If client is blocked, update MAC and activate; if active, just update MAC
            if is_blocked:
                # Update MAC address and activate client
                list_ip_bindings.set(
                    id=existing_binding['id'],
                    **{
                        'mac-address': mac_address,
                        'type': 'bypassed'
                    }
                )
                print(f"Client {client_name} MAC updated to: {mac_address} and activated from blocked status")
            else:
                # Update MAC address only (client was already active)
                list_ip_bindings.set(
                    id=existing_binding['id'],
                    **{
                        'mac-address': mac_address
                    }
                )
                print(f"Client {client_name} MAC address updated successfully to: {mac_address}")
            
            # Update cache
            self.refresh_clients_cache()
            return True, None
            
        except Exception as e:
            error_str = str(e).lower()
            if "invalid value of mac-address" in error_str:
                print(f"Invalid MAC address for client {client_name}")
                return False, "Wrong-MAC"
            print(f"Error updating client MAC: {str(e)}")
            return False, "Wrong-MAC"
            
    def is_client_blocked(self, client_name: str) -> bool:
        """Check if client is blocked in MikroTik"""
        try:
            binding = self.find_client_in_ip_bindings(client_name)
            if binding:
                return binding.get('type') == 'blocked'
            return False
        except Exception as e:
            print(f"Error checking client status {client_name}: {str(e)}")
            return False
            
    def get_client_mac(self, client_name: str) -> str:
        """Get MAC address for client by name"""
        try:
            binding = self.find_client_in_ip_bindings(client_name)
            if binding:
                return binding.get('mac-address', '')
            return ''
        except Exception as e:
            print(f"Error getting MAC for client {client_name}: {str(e)}")
            return ''

    def schedule_block_client(self, client_name: str, mac_address: str, block_date: str) -> bool:
        """Add schedule to block client on specific date"""
        try:
            # فحص ما إذا كان التاريخ في المستقبل
            if not self.is_date_in_future(block_date):
                print(f"Block date {block_date} is not in the future. Rejecting schedule.")
                return False
            
            # التحقق من وجود جدولة سابقة للعميل وحذفها
            if self.find_existing_schedule(client_name):
                print(f"Found existing schedule for client {client_name}, deleting it first...")
                self.delete_existing_schedule(client_name)
            
            # Convert date to required format (e.g., Dec/05/2024)
            # Support formats like: 5-6-2025, 05-06-2025, 10-12-2025
            day, month, year = block_date.split('-')
            # Ensure day and month have leading zeros if needed
            day = day.zfill(2)
            month = month.zfill(2)
            formatted_date = f"{['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][int(month)-1]}/{day}/20{year[-2:]}"
            
            # Create script content
            script_content = f"""# Assign the user's MAC address to a variable
:local targetMac "{mac_address}"

# Find the entry in IP Binding list based on the MAC address
:foreach i in=[/ip hotspot ip-binding find where mac-address=$targetMac and type="bypassed"] do={{
    # Change the type to "blocked"
    /ip hotspot ip-binding set $i type=blocked
}}

# Log the operation in the system log
:log info "MAC Address $targetMac has been changed to blocked" """

            # Add scheduler
            scheduler = self.api.get_resource('/system/scheduler')
            scheduler.add(
                name=f"{client_name}",
                start_date=formatted_date,
                start_time="23:59:00",
                interval="00:00:00",
                on_event=script_content,
                policy="ftp,reboot,read,write,policy,test,password,sniff,sensitive,romon"
            )
            print(f"Scheduled block for client {client_name} on {block_date} at 23:59:00")
            return True
            
        except Exception as e:
            print(f"Error scheduling block for client {client_name}: {str(e)}")
            return False

    def get_dhcp_leases(self):
        """Get DHCP leases list"""
        try:
            dhcp_resource = self.api.get_resource('/ip/dhcp-server/lease')
            leases = dhcp_resource.get()
            return leases
        except Exception as e:
            print(f"Error getting DHCP leases: {str(e)}")
            return []
    
    def find_mac_by_ip(self, ip_address: str) -> str:
        """Find MAC address by IP address in DHCP leases"""
        try:
            print(f"Searching for MAC address for IP: {ip_address}")
            leases = self.get_dhcp_leases()
            
            # Search DHCP leases quietly (no detailed printing)
            for lease in leases:
                lease_ip = lease.get('address', '')
                lease_mac = lease.get('mac-address', '')
                
                if lease_ip == ip_address and lease_mac:
                    print(f"Found MAC: {lease_mac} for IP: {ip_address}")
                    return lease_mac
                    
            # Also check ARP table as backup
            print(f"Not found in DHCP leases, checking ARP table...")
            try:
                arp_resource = self.api.get_resource('/ip/arp')
                arp_entries = arp_resource.get()
                
                for arp in arp_entries:
                    arp_ip = arp.get('address', '')
                    arp_mac = arp.get('mac-address', '')
                    
                    if arp_ip == ip_address and arp_mac:
                        print(f"Found MAC in ARP table: {arp_mac} for IP: {ip_address}")
                        return arp_mac
                        
            except Exception as arp_error:
                print(f"Error checking ARP table: {str(arp_error)}")
            
            print(f"MAC not found for IP: {ip_address}")
            return None
        except Exception as e:
            print(f"Error searching MAC for IP {ip_address}: {str(e)}")
            return None
    
    def is_ip_address(self, identifier: str) -> bool:
        """Check if string is valid IP address"""
        ip_pattern = r'^(\d{1,3}\.){3}\d{1,3}$'
        if re.match(ip_pattern, identifier):
            parts = identifier.split('.')
            return all(0 <= int(part) <= 255 for part in parts)
        return False
    
    def is_client_active(self, client_name: str) -> bool:
        """Check if client is active in MikroTik"""
        try:
            binding = self.find_client_in_ip_bindings(client_name)
            if binding:
                return binding.get('type') == 'bypassed'
            return False
        except Exception as e:
            print(f"Error checking client activation status {client_name}: {str(e)}")
            return False

    def __del__(self):
        """Close connection when finished"""
        try:
            self.connection.disconnect()
        except:
            pass
    
    def find_mac_in_ip_bindings(self, mac_address: str) -> dict:
        """البحث عن MAC address في IP bindings"""
        try:
            list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
            bindings = list_ip_bindings.get()
            
            for binding in bindings:
                binding_mac = binding.get('mac-address', '')
                if binding_mac.lower() == mac_address.lower():
                    return binding
            return None
        except Exception as e:
            print(f"Error searching for MAC {mac_address}: {str(e)}")
            return None
    
    def is_unauthorized_user(self, comment: str) -> bool:
        """فحص ما إذا كان الكومنت يشير إلى مستخدم غير مخول (ZZZZ=Blocked unauthorized)"""
        if not comment:
            return False
        return comment.startswith('ZZZZ=Blocked unauthorized')
    
    def delete_ip_binding_by_id(self, binding_id: str) -> bool:
        """حذف IP binding باستخدام الـ ID"""
        try:
            list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
            list_ip_bindings.remove(id=binding_id)
            print(f"Deleted IP binding with ID: {binding_id}")
            return True
        except Exception as e:
            print(f"Error deleting IP binding {binding_id}: {str(e)}")
            return False
    
    def find_existing_schedule(self, client_name: str) -> dict:
        """البحث عن جدولة موجودة للعميل"""
        try:
            scheduler = self.api.get_resource('/system/scheduler')
            schedules = scheduler.get()
            
            for schedule in schedules:
                schedule_name = schedule.get('name', '')
                if schedule_name == client_name:
                    return schedule
            return None
        except Exception as e:
            print(f"Error finding existing schedule for {client_name}: {str(e)}")
            return None
    
    def delete_existing_schedule(self, client_name: str) -> bool:
        """حذف الجدولة الموجودة للعميل"""
        try:
            existing_schedule = self.find_existing_schedule(client_name)
            if existing_schedule:
                scheduler = self.api.get_resource('/system/scheduler')
                scheduler.remove(id=existing_schedule['id'])
                print(f"Deleted existing schedule for client {client_name}")
                return True
            return False
        except Exception as e:
            print(f"Error deleting existing schedule for {client_name}: {str(e)}")
            return False
    
    def is_date_in_future(self, block_date: str) -> bool:
        """فحص ما إذا كان التاريخ في المستقبل"""
        try:
            day, month, year = block_date.split('-')
            # تحويل إلى تاريخ
            block_datetime = datetime(int(year), int(month), int(day))
            current_datetime = datetime.now()
            
            # مقارنة التواريخ (فقط اليوم والشهر والسنة، بدون الوقت)
            block_date_only = block_datetime.date()
            current_date_only = current_datetime.date()
            
            return block_date_only > current_date_only
        except Exception as e:
            print(f"Error checking date validity: {str(e)}")
            return False
    
    def update_client_name(self, client_name: str, new_name: str) -> tuple:
        """Update client name in MikroTik comment (part before @)"""
        try:
            print(f"[DEBUG] Searching for client '{client_name}' in cache...")
            print(f"[DEBUG] Cache has {len(self.clients_cache)} clients")
            
            # Check if client exists
            existing_binding = self.find_client_in_ip_bindings(client_name)
            if not existing_binding:
                print(f"[ERROR] Client '{client_name}' not found in MikroTik")
                # Let's try to search manually in cache keys
                print(f"[DEBUG] Available client names in cache: {list(self.clients_cache.keys())[:10]}...")
                return False, "Client-Not-Found"
            
            # Get current comment
            current_comment = existing_binding.get('comment', '')
            print(f"[DEBUG] Current comment: '{current_comment}'")
            
            if not current_comment:
                print(f"[ERROR] No comment found for client {client_name}")
                return False, "No-Comment"
            
            # Split comment at ' @' (space then @) and update the name part
            if ' @' in current_comment:
                # Keep the phone number part after ' @'
                phone_part = current_comment.split(' @', 1)[1]
                new_comment = f"{new_name.strip()} @{phone_part}"
                print(f"[DEBUG] Found ' @' separator, phone part: '{phone_part}'")
            elif '@' in current_comment:
                # Fallback: if only @ without space
                phone_part = current_comment.split('@', 1)[1]
                new_comment = f"{new_name.strip()} @{phone_part}"
                print(f"[DEBUG] Found '@' separator, phone part: '{phone_part}'")
            else:
                # If no @ found, just replace the whole comment with new name
                new_comment = new_name.strip()
                print(f"[DEBUG] No @ found, replacing entire comment")
            
            print(f"[DEBUG] New comment will be: '{new_comment}'")
            
            # Update comment in MikroTik
            list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
            list_ip_bindings.set(
                id=existing_binding['id'],
                comment=new_comment
            )
            
            # Update cache
            self.refresh_clients_cache()
            print(f"Client name updated from '{client_name}' to '{new_name}'")
            print(f"Comment updated to: {new_comment}")
            print(f"[DEBUG] Cache refreshed. New client should be findable as '{new_name}'")
            
            # Verify the update worked by trying to find the client with new name
            verification = self.find_client_in_ip_bindings(new_name)
            if verification:
                print(f"[DEBUG] Verification successful: Client found with new name '{new_name}'")
            else:
                print(f"[WARNING] Verification failed: Client not found with new name '{new_name}'")
            
            return True, None
            
        except Exception as e:
            print(f"Error updating client name: {str(e)}")
            import traceback
            traceback.print_exc()
            return False, "Update-Failed"
    
    def update_client_phone(self, client_name: str, new_phone: str) -> tuple:
        """Update client phone number in MikroTik comment (part after @)"""
        try:
            # Check if client exists
            existing_binding = self.find_client_in_ip_bindings(client_name)
            if not existing_binding:
                print(f"Client {client_name} not found in MikroTik")
                return False, "Client-Not-Found"
            
            # Get current comment
            current_comment = existing_binding.get('comment', '')
            if not current_comment:
                print(f"No comment found for client {client_name}")
                return False, "No-Comment"
            
            # Split comment at ' @' (space then @) and update the phone part
            if ' @' in current_comment:
                # Keep the name part before ' @'
                name_part = current_comment.split(' @', 1)[0]
                new_comment = f"{name_part} @{new_phone.strip()}"
            elif '@' in current_comment:
                # Fallback: if only @ without space
                name_part = current_comment.split('@', 1)[0]
                new_comment = f"{name_part} @{new_phone.strip()}"
            else:
                # If no @ found, add @ and phone number
                new_comment = f"{current_comment} @{new_phone.strip()}"
            
            # Update comment in MikroTik
            list_ip_bindings = self.api.get_resource('/ip/hotspot/ip-binding')
            list_ip_bindings.set(
                id=existing_binding['id'],
                comment=new_comment
            )
            
            # Update cache
            self.refresh_clients_cache()
            print(f"Client phone updated for '{client_name}'")
            print(f"Comment updated to: {new_comment}")
            return True, None
            
        except Exception as e:
            print(f"Error updating client phone: {str(e)}")
            return False, "Update-Failed"
    
    def find_client_flexible(self, client_name: str) -> dict:
        """Search for client with flexible matching"""
        client_name = client_name.strip()
        
        # Try exact match first
        exact_match = self.clients_cache.get(client_name)
        if exact_match:
            return exact_match
        
        # Try case-insensitive match
        for cached_name, binding in self.clients_cache.items():
            if cached_name.lower() == client_name.lower():
                print(f"[DEBUG] Found case-insensitive match: '{cached_name}' for '{client_name}'")
                return binding
        
        # Try partial match (client_name is contained in cached name)
        for cached_name, binding in self.clients_cache.items():
            if client_name.lower() in cached_name.lower():
                print(f"[DEBUG] Found partial match: '{cached_name}' for '{client_name}'")
                return binding
        
        print(f"[DEBUG] No match found for '{client_name}'")
        return None

    def get_client_phone(self, client_name: str) -> str:
        """Get phone number for client by name (after @ in comment)"""
        try:
            binding = self.find_client_in_ip_bindings(client_name)
            if binding:
                comment = binding.get('comment', '')
                if ' @' in comment:
                    return comment.split(' @', 1)[1].strip()
                elif '@' in comment:
                    return comment.split('@', 1)[1].strip()
            return ''
        except Exception as e:
            print(f"Error getting phone for client {client_name}: {str(e)}")
            return ''
