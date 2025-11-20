# Coleta de inventario - retorna JSON no mesmo formato esperado pelo Python 
# PowerShell - coleta dos dados, de forma bruta (sem nenhum tratamento)

$ErrorActionPreference = 'SilentlyContinue'  
        

# Comandos para coletar os dados
$info = @{
    CS = Get-CimInstance Win32_ComputerSystem
            
    OS = Get-CimInstance Win32_OperatingSystem
            
    BIOS = Get-CimInstance Win32_BIOS
        
    CSProduct = Get-CimInstance Win32_ComputerSystemProduct | Select-Object IdentifyingNumber

    Boot = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Diagnostics-Performance/Operational';ID=100} -MaxEvents 1
            
    CPU = Get-CimInstance Win32_Processor -OperationTimeoutSec 20
        
    GPU = Get-CimInstance Win32_VideoController | Select-Object Name, AdapterRAM
            
    RAM = Get-CimInstance Win32_PhysicalMemory | Select-Object Capacity, Speed, SMBIOSMemoryType, MemoryType
            
    Volumes = Get-Volume | Select-Object DriveLetter, FileSystemLabel, Size, SizeRemaining, DriveType
            
    Discs = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size
            
    Net = Get-NetAdapter | Where-Object {$_.Status -eq 'Up' -and $_.HardwareInterface -eq $true -and $_.Name -like '*Ethernet*'} | Select-Object Name, InterfaceDescription, MacAddress, LinkSpeed

    IP = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike '169.*' -and $_.IPAddress -ne $null } | Select-Object InterfaceAlias, IPAddress -First 1
            
}

$info | ConvertTo-Json -Depth 4  