import subprocess
import wmi
import win32com.client

#Function to retrieve system's unique UUID
def get_system_details():
    uuid_process = subprocess.Popen('C:\Program Files (x86)\GnuWin32\sbin\dmidecode.exe -s system-uuid'.split(), stdout=subprocess.PIPE, stderr = subprocess.PIPE)
    manufacturer_process = subprocess.Popen('C:\Program Files (x86)\GnuWin32\sbin\dmidecode.exe -s system-manufacturer'.split(), stdout=subprocess.PIPE, stderr = subprocess.PIPE)
    version_process = subprocess.Popen('C:\Program Files (x86)\GnuWin32\sbin\dmidecode.exe -s system-version'.split(), stdout=subprocess.PIPE, stderr = subprocess.PIPE)
    serialNumber_process = subprocess.Popen('C:\Program Files (x86)\GnuWin32\sbin\dmidecode.exe -s system-serial-number'.split(), stdout=subprocess.PIPE, stderr = subprocess.PIPE)
    system_details = {
	    "uuid": uuid_process.stdout.read().strip(),
	    "manufacturer": manufacturer_process.stdout.read().strip(),
	    "version": version_process.stdout.read().strip(),
	    "serialNumber": serialNumber_process.stdout.read().strip()
    }
    return system_details

#Function to get processor details
def get_processor_details():
	processor_details = []
	strComputer = "."
	objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
	objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
	colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_Processor")
	for objItem in colItems:
		processor = {
			'name': objItem.Name,
			'description': objItem.description,
			'DeviceID': objItem.DeviceID,
			'manufacturer': objItem.Manufacturer,
			'processorID': objItem.processorID,
			'systemName': objItem.SystemName
		}
		processor_details.append(processor)
	return processor_details

#Function to get motherboard details
def get_motherboard_details():
	motherboard_details = []
	strComputer = "."
	objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
	objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
	colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_BaseBoard")
	for objItem in colItems:
		main_board = {
			'name': objItem.Name,
			'description': objItem.Description,
			'manufacturer': objItem.Manufacturer,
			'model': objItem.Model,
			'product': objItem.Product,
			'serialNumber': objItem.SerialNumber,
			'version': objItem.Version
		}
		motherboard_details.append(main_board)
	return motherboard_details

#Function to get VGA devices
def get_gpu_details():
	gpu_details = []
	computer = wmi.WMI()
	gpu_info = computer.Win32_VideoController()
	for eachitem in gpu_info:
		each_gpu = {
			'name': eachitem.Name,
			'adapterRAM': eachitem.AdapterRAM,
			'description': eachitem.description,
			'pnpDeviceID': eachitem.PNPDeviceID,
			'systemName': eachitem.SystemName
		}
		gpu_details.append(each_gpu)
	return gpu_details

#Function to get monitor details
def get_monitor_details():
	monitor_details = []
	obj = wmi.WMI().Win32_PnPEntity(ConfigManagerErrorCode=0)
	displays = [x for x in obj if 'Monitor' in str(x)]
	for item in displays:
		if item.PNPClass == "Monitor":
			eachMonitor = {
				'name': item.Caption,
				'manufacturer': item.Manufacturer,
				'hardwareID': item.HardwareID[0],
				'description': item.Description,
				'systemName':item.SystemName
			}
			monitor_details.append(eachMonitor)
	return monitor_details

#Function to get cd_drive details
def get_cd_drive_details():
	cd_drive_details = []
	speaker_details = []
	obj = wmi.WMI().Win32_PnPEntity(ConfigManagerErrorCode=0)
	cd_drive = [x for x in obj if 'CDROM' in str(x)]
	for item in cd_drive:
		each_cd_drive = {
			'name': item.Name,
			'description': item.Description,
			'hardwareID': item.HardwareID[0],
			'pnpDeviceID': item.pnpDeviceID,
			'manufacturer': item.Manufacturer,
			'systemName': item.systemName
		}
		cd_drive_details.append(each_cd_drive)
	return cd_drive_details

#Function to get mouse details
def get_mouse_details():
	mouse_details = []
	obj = wmi.WMI().Win32_PnPEntity(ConfigManagerErrorCode=0)
	mouse = [x for x in obj if 'Mouse' in str(x)]
	for item in mouse:
		eachMouse = {
			'name': item.Caption,
			'manufacturer': item.Manufacturer,
			'hardwareID': item.HardwareID[0],
			'description': item.Description,
			'systemName':item.SystemName
		}
		mouse_details.append(eachMouse)
	return mouse_details

#Fuction to get speaker details
def get_speaker_details():
	speaker_details = []
	obj = wmi.WMI().Win32_PnPEntity(ConfigManagerErrorCode=0)
	speakers = [x for x in obj if 'Speakers' in str(x)]
	for item in speakers:
		eachSpeaker = {
			'name': item.Caption,
			'manufacturer': item.Manufacturer,
			'hardwareID': item.HardwareID[0],
			'description': item.Description,
			'systemName':item.SystemName
		}
		speaker_details.append(eachSpeaker)
	return speaker_details

#Function to get keyboard details
def get_keyboard_details():
	keyboard_details = []
	obj = wmi.WMI().Win32_PnPEntity(ConfigManagerErrorCode=0)
	keyboard = [x for x in obj if 'Keyboard' in str(x)]
	for item in keyboard:
		eachKeyboard = {
			'name': item.Caption,
			'manufacturer': item.Manufacturer,
			'hardwareID': item.HardwareID[0],
			'description': item.Description,
			'systemName':item.SystemName
		}
		keyboard_details.append(eachKeyboard)
	return keyboard_details

#Function to get Hard Disk details
def get_hard_disk_details():
	hard_disk_details = []
	strComputer = "."
	objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
	objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
	colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_DiskDrive")
	for objItem in colItems:
		eachHardDisk = {
			'name': objItem.Name,
			'manufacturer': objItem.Manufacturer,
			'deviceID': objItem.DeviceID,
			'description': objItem.Description,
			'pnpDeviceID': objItem.PNPDeviceID,
			'size': objItem.Size
		}
		hard_disk_details.append(eachHardDisk)
	return hard_disk_details

#Get RAM details
def get_ram_details():
	RAM_details = []
	strComputer = "."
	objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
	objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
	colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
	for objItem in colItems:
		eachRAM = {
			'name': objItem.Name,
			'manufacturer': objItem.manufacturer,
			'description': objItem.description,
			'serialNumber': objItem.SerialNumber,
			'capaity': objItem.capacity
		}
		RAM_details.append(eachRAM)
	return RAM_details

# print get_system_details()

# print get_processor_details()

# print get_motherboard_details()

# print get_gpu_details()

# print get_monitor_details()

# print get_cd_drive_details()

# print get_mouse_details()

# print get_speaker_details()

# print get_keyboard_details()

# print get_hard_disk_details()

# print get_ram_details()