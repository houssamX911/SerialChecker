"""
Serial Checker — Windows hardware / system snapshot GUI (HTML + JSON export, diff check).
Copyright (c) houssamX. All rights reserved.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
import os
import hashlib
import ctypes
import webbrowser
import html as html_lib
import tkinter as tk
import tkinter.messagebox as tkmsg
from datetime import datetime
from pathlib import Path

APP_AUTHOR = "houssamX"
APP_CREDIT = f"© {APP_AUTHOR}"

# Theme (match reference UI)
BG_MAIN = "#05050B"
BG_CARD = "#0C0C14"
BORDER_CARD = "#1E293B"
HEADER_BLUE = "#3B82F6"
HEADER_ORANGE = "#F97316"
TEXT_LABEL = "#FFFFFF"
TEXT_VALUE = "#9CA3AF"
GREEN_OK = "#22C55E"
RED_BAD = "#EF4444"
BTN_CLOSE = "#EF4444"
BTN_EXPORT = "#374151"
BTN_CHECK = "#22C55E"
BTN_DISCORD = "#5865F2"
TITLE_BOX_BG = "#0F0F18"
TITLE_BOX_FG = "#C4B5FD"

# One PowerShell script = one process (single WMI gather).
_GATHER_SCRIPT = r"""
$ErrorActionPreference = 'SilentlyContinue'
$cs = Get-CimInstance Win32_ComputerSystem
$csp = Get-CimInstance Win32_ComputerSystemProduct
$bios = Get-CimInstance Win32_BIOS
$bb = Get-CimInstance Win32_BaseBoard
$cpuList = @(Get-CimInstance Win32_Processor)
$cpu = $cpuList | Where-Object { $_.ProcessorType -eq 3 } | Select-Object -First 1
if (-not $cpu) { $cpu = $cpuList | Select-Object -First 1 }
$ch = Get-CimInstance Win32_SystemEnclosure
$core = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' -ErrorAction SilentlyContinue).Enabled
# BIOS VT-x/AMD-V (not Hyper-V running); HypervisorPresent is false when no hypervisor is active
$virtFw = $null
try { $virtFw = $cpu.VirtualizationFirmwareEnabled } catch {}
$virtual = if ($null -ne $virtFw) { $virtFw } else { $cs.HypervisorPresent }
$secure = $null
try { $secure = Confirm-SecureBootUEFI } catch {}
$tpm = Get-Tpm
$tpmPresent = $tpm.TpmPresent
$tpmKey = $null
try { $tpmKey = (Get-TpmEndorsementKeyInfo -Hash 'Sha256' -ErrorAction SilentlyContinue).PublicKeyHash } catch {}
if (-not $tpmKey) { $tpmKey = "$($tpm.ManufacturerIdTxt)|$($tpm.ManufacturerVersion)|$($tpm.SpecVersion)" }
$net = Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -First 1
$gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1
$diskList = @(Get-CimInstance Win32_DiskDrive | Select-Object Caption,Model,InterfaceType,Status,PNPDeviceID,SCSIPort,SerialNumber)
$physList = @(Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName,Model,SerialNumber,UniqueId,UniqueIdFormat,BusType,HealthStatus)
$monitors = @()
Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction SilentlyContinue | ForEach-Object {
  $man = (($_.ManufacturerName | Where-Object {$_ -ne 0}) | ForEach-Object {[char]$_}) -join ''
  $model = (($_.UserFriendlyName | Where-Object {$_ -ne 0}) | ForEach-Object {[char]$_}) -join ''
  $serial = (($_.SerialNumberID | Where-Object {$_ -ne 0}) | ForEach-Object {[char]$_}) -join ''
  $monitors += [ordered]@{
    Active = if ($_.Active) { 'Enabled' } else { 'Disabled' }
    Manufacturer = if ($man) { $man } else { 'N/A' }
    ModelName = if ($model) { $model } else { 'N/A' }
    MonitorSerial = if ($serial) { $serial } else { 'N/A' }
    IDSerialNumber = if ($_.InstanceName) { $_.InstanceName } else { 'N/A' }
  }
}
# Baseboard asset: WMI Part/SKU only (never duplicate Serial); SMBIOS Type 2 fills Asset Tag in Python
$bbAssetRaw = [string]$bb.PartNumber
if ([string]::IsNullOrWhiteSpace($bbAssetRaw)) { $bbAssetRaw = [string]$bb.SKU }
$csLoc = 'N/A'
if ($bb.ConfigOptions -and @($bb.ConfigOptions).Count -gt 0) {
  $csLoc = [string](@($bb.ConfigOptions)[0])
}
if ([string]::IsNullOrWhiteSpace($csLoc)) { $csLoc = 'N/A' }
$smbHex = ''
try {
  $rw = Get-CimInstance -Namespace root\wmi -ClassName MSSmBios_RawSMBiosTables -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($rw -and $rw.SMBiosData) {
    $arr = $rw.SMBiosData
    $bytes = $null
    if ($arr -is [byte[]]) { $bytes = $arr }
    elseif ($arr -is [System.Array]) {
      $bytes = [byte[]]::new($arr.Length); [Array]::Copy($arr, $bytes, $arr.Length)
    }
    if ($bytes -and $bytes.Length -gt 0) {
      $smbHex = [System.BitConverter]::ToString($bytes) -replace '-',''
    }
  }
} catch {}
function Test-CpuOemPlaceholder([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $true }
  $t = $s.Trim()
  if ($t -match '(?i)part\s+of\s+o[e]?[.]?m[.]?|part\s+of\s+oen|^\s*oem\s*$|o\.e\.m\.|to\s+be\s+filled|default\s+string|not\s+applicable') { return $true }
  return $false
}
$cpuMan = [string]$cpu.Manufacturer
if ($cpuMan -match '^GenuineIntel') { $cpuMan = 'Intel(R) Corporation' }
elseif ($cpuMan -match '^AuthenticAMD') { $cpuMan = 'Advanced Micro Devices, Inc.' }
# Match HWiNFO-style WMI: SerialNumber (OEM/board id), PartNumber as reported, AssetTag (not socket)
$cpuSerialId = ([string]$cpu.SerialNumber).Trim()
if ([string]::IsNullOrWhiteSpace($cpuSerialId) -or (Test-CpuOemPlaceholder $cpuSerialId)) {
  $cpuSerialId = (([string]$cpu.ProcessorId) -replace '\s','').ToUpper()
}
$cpuPartRaw = [string]$cpu.PartNumber
if ([string]::IsNullOrWhiteSpace($cpuPartRaw)) { $cpuPartRaw = 'N/A' }
$cpuAssetVal = ([string]$cpu.AssetTag).Trim()
if ([string]::IsNullOrWhiteSpace($cpuAssetVal) -or (Test-CpuOemPlaceholder $cpuAssetVal)) {
  $cpuAssetVal = ([string]$cpu.SocketDesignation).Trim()
}
if ([string]::IsNullOrWhiteSpace($cpuAssetVal) -or (Test-CpuOemPlaceholder $cpuAssetVal)) {
  $cpuAssetVal = [string]$cpu.DeviceID
}
if ([string]::IsNullOrWhiteSpace($cpuAssetVal)) { $cpuAssetVal = 'N/A' }
$chSkuVal = [string]$ch.SKU
if ([string]::IsNullOrWhiteSpace($chSkuVal)) {
  try { $chSkuVal = [string](Get-CimInstance Win32_SystemEnclosure | Select-Object -First 1).SKU } catch {}
}
$netIPv4 = $null
if ($net) {
  $netIPv4 = (Get-NetIPAddress -InterfaceIndex $net.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike '169.254.*' } | Select-Object -First 1).IPAddress
}
$netPhysical = @()
$seenMac = @{}
Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue |
  Where-Object {
    $_.LinkLayerAddress -and $_.IPAddress -and $_.IPAddress -ne '255.255.255.255' -and
    $_.IPAddress -notmatch '^(224\.|239\.)'
  } | ForEach-Object {
    $mac = [string]$_.LinkLayerAddress
    if ($mac -match '(?i)^0{1,2}(-?0{2}){5}') { return }
    $kind = if ($_.State -eq 'Permanent') { 'static' } else { 'dynamic' }
    $k = $mac.ToLowerInvariant()
    if (-not $seenMac.ContainsKey($k)) {
      $seenMac[$k] = $true
      $netPhysical += [ordered]@{ Mac = $mac; Kind = $kind }
    }
  }
[ordered]@{
  SystemManufacturer = $cs.Manufacturer
  ProductName = $csp.Name
  SystemVersionIndex = $csp.Version
  SystemSerial = $bios.SerialNumber
  SystemUUID = $csp.UUID
  FamilySerial = $cs.SystemFamily
  SKUNumber = $cs.SystemSKUNumber
  BIOSVendor = $bios.Manufacturer
  BIOSVersion = $bios.SMBIOSBIOSVersion
  ReleaseDate = (Get-Date $bios.ReleaseDate -Format 'yyyy-MM-dd')
  CoreIsolationRaw = $core
  VirtualizationRaw = $virtual
  SecureBootRaw = $secure
  TpmStatusRaw = $tpmPresent
  BBManufacturer = $bb.Manufacturer
  BBVersion = $bb.Version
  BBProduct = $bb.Product
  BBSerial = $bb.SerialNumber
  BBAsset = $bbAssetRaw
  CSLocation = $csLoc
  SMBiosHex = $smbHex
  CPUManu = $cpuMan
  CPUType = $cpu.Name
  CPUSerial = $cpuSerialId
  CPUPart = $cpuPartRaw
  CPUAsset = $cpuAssetVal
  CPUSocket = $cpu.SocketDesignation
  ChassisManu = $ch.Manufacturer
  ChassisType = $ch.ChassisTypes
  ChassisVersion = $ch.Version
  ChassisSerial = $ch.SerialNumber
  ChassisAsset = $ch.SMBIOSAssetTag
  ChassisSKU = $chSkuVal
  MAC = if ($net) { $net.MacAddress } else { $null }
  NetIPv4 = $netIPv4
  NetPhysical = $netPhysical
  GpuName = $gpu.Name
  PciDevice = $gpu.PNPDeviceID
  GuidSerial = $gpu.VideoProcessor
  TpmBase = $tpmKey
  Disks = $diskList
  PhysicalDisks = $physList
  Monitors = $monitors
} | ConvertTo-Json -Depth 8
"""


def run_ps_file(script: str, timeout: int = 45) -> str:
    path = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", suffix=".ps1", delete=False, encoding="utf-8") as tmp:
            tmp.write(script)
            path = tmp.name
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        out = (result.stdout or "").strip()
        return out if out else "N/A"
    except Exception:
        return "N/A"
    finally:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass


def first_line(value: str) -> str:
    if not value:
        return "N/A"
    line = value.splitlines()[0].strip()
    return line if line else "N/A"


def format_cpu_serial_display(s: str) -> str:
    v = first_line(s or "")
    if v == "N/A":
        return v
    compact = re.sub(r"\s+", "", v).upper()
    if len(compact) == 16 and re.match(r"^[0-9A-F]+$", compact):
        return f"{compact[:8]} {compact[8:]}"
    return v


def _smbios_read_string_table(data: bytes, start: int) -> tuple[list[str], int]:
    """SMBIOS string table: str1\\0str2\\0...\\0\\0 — final \\0 ends last string, next \\0 ends the table."""
    strings: list[str] = [""]
    pos = start
    while pos < len(data):
        end = data.find(b"\x00", pos)
        if end == -1:
            break
        strings.append(data[pos:end].decode("latin-1", errors="replace"))
        pos = end + 1
        if pos < len(data) and data[pos] == 0:
            return strings, pos + 1
    return strings, pos


def parse_smbios_baseboard_type2(data: bytes) -> dict[str, str]:
    """First SMBIOS type 2 (baseboard): serial, asset tag, location in chassis (matches HWiNFO)."""
    out: dict[str, str] = {"serial": "", "asset": "", "location": ""}
    i = 0
    while i + 4 <= len(data):
        typ = data[i]
        ln = data[i + 1]
        if typ == 127:
            break
        if ln < 4:
            i += 1
            continue
        if i + ln > len(data):
            break
        if typ == 2 and ln >= 9:
            idx_serial = data[i + 7]
            idx_asset = data[i + 8]
            # 15-byte baseboard records (common): Location in Chassis at 0x0Ah; longer SMBIOS 2.7+ uses 0x0Bh
            if ln == 15 and ln > 10:
                idx_loc = data[i + 10]
            elif ln > 11:
                idx_loc = data[i + 11]
            elif ln > 10:
                idx_loc = data[i + 10]
            else:
                idx_loc = 0
            str_start = i + ln
            strings, next_pos = _smbios_read_string_table(data, str_start)

            def _get(idx: int) -> str:
                if idx <= 0 or idx >= len(strings):
                    return ""
                return strings[idx].strip()

            out["serial"] = _get(idx_serial)
            out["asset"] = _get(idx_asset)
            out["location"] = _get(idx_loc) if idx_loc else ""
            return out
        str_start = i + ln
        _, next_pos = _smbios_read_string_table(data, str_start)
        i = next_pos
    return out


def bool_to_status(value: str) -> str:
    clean = first_line(value).lower()
    if clean in ("true", "1", "yes"):
        return "Enabled"
    if clean in ("false", "0", "no"):
        return "Disabled"
    return first_line(value)


def build_disk_rows(disk: dict) -> list[tuple[str, str]]:
    return [
        ("Caption", first_line(str(disk.get("Caption", "N/A")))),
        ("DISK_STORAGE_MODEL", first_line(str(disk.get("Model", "N/A")))),
        ("STORAGE_QUERY_PROPERTY", first_line(str(disk.get("InterfaceType", "N/A")))),
        ("SMART_RCV_DRIVE_DATA", first_line(str(disk.get("Status", "N/A")))),
        ("STORAGE_QUERY_WWN", first_line(str(disk.get("PNPDeviceID", "N/A")))),
        ("SCSI_PASS_THROUGH", first_line(str(disk.get("SCSIPort", "N/A")))),
        ("ATA_PASS_THROUGH", first_line(str(disk.get("SerialNumber", "N/A")))),
    ]


def hash_triplet(base_value: str) -> tuple[str, str, str]:
    clean = first_line(base_value)
    if clean == "N/A":
        return "N/A", "N/A", "N/A"
    raw = clean.encode("utf-8", errors="ignore")
    return (
        hashlib.md5(raw).hexdigest(),
        hashlib.sha1(raw).hexdigest(),
        hashlib.sha256(raw).hexdigest(),
    )


def _format_gpu_prefix_uuid(guid_hex: str) -> str:
    """Normalize to GPU-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (lowercase)."""
    g = guid_hex.strip().strip("{}").lower()
    g = re.sub(r"[^0-9a-f-]", "", g)
    if re.fullmatch(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", g):
        return f"GPU-{g}"
    return "N/A"


def gpu_guid_fallback_from_pnp(pnp_device_id: str) -> str:
    """Stable pseudo-GPU-UUID from PNP path when dxdiag is unavailable."""
    pnp = first_line(pnp_device_id)
    if pnp == "N/A":
        return "N/A"
    h = hashlib.md5(pnp.encode("utf-8", errors="ignore")).hexdigest()
    u = f"{h[0:8]}-{h[8:12]}-{h[12:16]}-{h[16:20]}-{h[20:32]}"
    return f"GPU-{u}"


def fetch_gpu_guid_from_dxdiag() -> str:
    """
    DirectX 'Device Identifier' from dxdiag Display Devices block — matches tools that use DirectX runtime.
    """
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as tmp:
            temp_path = tmp.name
        subprocess.run(
            ["dxdiag", "/whql:off", "/t", temp_path],
            capture_output=True,
            text=True,
            timeout=22,
            check=False,
        )
        if not temp_path or not os.path.exists(temp_path):
            return "N/A"
        with open(temp_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        in_display = False
        current_card = ""
        pairs: list[tuple[str, str]] = []
        for raw in lines:
            line = raw.strip()
            if line.startswith("---------------") and "Display Devices" in line:
                in_display = True
                continue
            if in_display and line.startswith("---------------") and "Sound Devices" in line:
                break
            if not in_display:
                continue
            low = line.lower()
            if low.startswith("card name:"):
                current_card = line.split(":", 1)[1].strip()
            if low.startswith("device identifier:"):
                rest = line.split(":", 1)[1].strip()
                out = _format_gpu_prefix_uuid(rest)
                if out != "N/A":
                    pairs.append((current_card, out))
        for card, guid in pairs:
            cl = card.lower()
            if "microsoft" in cl or "basic render" in cl or "remote" in cl:
                continue
            return guid
        if pairs:
            return pairs[0][1]
    except Exception:
        return "N/A"
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
    return "N/A"


def get_display_devices_fallback() -> list[dict[str, str]]:
    class DISPLAY_DEVICEW(ctypes.Structure):
        _fields_ = [
            ("cb", ctypes.c_ulong),
            ("DeviceName", ctypes.c_wchar * 32),
            ("DeviceString", ctypes.c_wchar * 128),
            ("StateFlags", ctypes.c_ulong),
            ("DeviceID", ctypes.c_wchar * 128),
            ("DeviceKey", ctypes.c_wchar * 128),
        ]

    class DEVMODEW(ctypes.Structure):
        _fields_ = [
            ("dmDeviceName", ctypes.c_wchar * 32),
            ("dmSpecVersion", ctypes.c_ushort),
            ("dmDriverVersion", ctypes.c_ushort),
            ("dmSize", ctypes.c_ushort),
            ("dmDriverExtra", ctypes.c_ushort),
            ("dmFields", ctypes.c_uint),
            ("dmOrientation", ctypes.c_short),
            ("dmPaperSize", ctypes.c_short),
            ("dmPaperLength", ctypes.c_short),
            ("dmPaperWidth", ctypes.c_short),
            ("dmScale", ctypes.c_short),
            ("dmCopies", ctypes.c_short),
            ("dmDefaultSource", ctypes.c_short),
            ("dmPrintQuality", ctypes.c_short),
            ("dmColor", ctypes.c_short),
            ("dmDuplex", ctypes.c_short),
            ("dmYResolution", ctypes.c_short),
            ("dmTTOption", ctypes.c_short),
            ("dmCollate", ctypes.c_short),
            ("dmFormName", ctypes.c_wchar * 32),
            ("dmLogPixels", ctypes.c_ushort),
            ("dmBitsPerPel", ctypes.c_uint),
            ("dmPelsWidth", ctypes.c_uint),
            ("dmPelsHeight", ctypes.c_uint),
            ("dmDisplayFlags", ctypes.c_uint),
            ("dmDisplayFrequency", ctypes.c_uint),
            ("dmICMMethod", ctypes.c_uint),
            ("dmICMIntent", ctypes.c_uint),
            ("dmMediaType", ctypes.c_uint),
            ("dmDitherType", ctypes.c_uint),
            ("dmReserved1", ctypes.c_uint),
            ("dmReserved2", ctypes.c_uint),
            ("dmPanningWidth", ctypes.c_uint),
            ("dmPanningHeight", ctypes.c_uint),
        ]

    devices = []
    user32 = ctypes.windll.user32
    enum_display_devices = user32.EnumDisplayDevicesW
    enum_display_devices.argtypes = [
        ctypes.c_wchar_p,
        ctypes.c_ulong,
        ctypes.POINTER(DISPLAY_DEVICEW),
        ctypes.c_ulong,
    ]
    enum_display_devices.restype = ctypes.c_int
    enum_display_settings = user32.EnumDisplaySettingsW
    enum_display_settings.argtypes = [ctypes.c_wchar_p, ctypes.c_uint, ctypes.POINTER(DEVMODEW)]
    enum_display_settings.restype = ctypes.c_int

    DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = 0x00000001
    DISPLAY_DEVICE_MIRRORING_DRIVER = 0x00000008
    ENUM_CURRENT_SETTINGS = 0xFFFFFFFF

    def get_hz(device_name: str) -> str:
        dm = DEVMODEW()
        dm.dmSize = ctypes.sizeof(DEVMODEW)
        if enum_display_settings(device_name, ENUM_CURRENT_SETTINGS, ctypes.byref(dm)):
            hz = int(dm.dmDisplayFrequency)
            return f"{hz}Hz" if hz > 1 else "N/A"
        return "N/A"

    adapter_index = 0
    while True:
        adapter = DISPLAY_DEVICEW()
        adapter.cb = ctypes.sizeof(DISPLAY_DEVICEW)
        if not enum_display_devices(None, adapter_index, ctypes.byref(adapter), 0):
            break

        if (adapter.StateFlags & DISPLAY_DEVICE_MIRRORING_DRIVER) != 0:
            adapter_index += 1
            continue

        dname = first_line(adapter.DeviceName)
        display_label = "N/A"
        if dname.startswith("\\\\.\\"):
            display_label = "\\" + dname[4:]
        elif dname.startswith("\\"):
            display_label = dname
        elif dname:
            display_label = "\\" + dname

        monitor_index = 0
        while True:
            monitor = DISPLAY_DEVICEW()
            monitor.cb = ctypes.sizeof(DISPLAY_DEVICEW)
            if not enum_display_devices(adapter.DeviceName, monitor_index, ctypes.byref(monitor), 0):
                break

            attached = (monitor.StateFlags & DISPLAY_DEVICE_ATTACHED_TO_DESKTOP) != 0
            if attached:
                refresh = get_hz(adapter.DeviceName)
                if display_label != "N/A" and refresh != "N/A":
                    hz_short = refresh[:-2] if refresh.endswith("Hz") else refresh
                    active_line = f"{display_label}\\{hz_short}Hz"
                else:
                    active_line = display_label
                devices.append(
                    {
                        "Active": active_line,
                        "Manufacturer": "N/A",
                        "ModelName": first_line(monitor.DeviceString),
                        "MonitorSerial": "N/A",
                        "IDSerialNumber": first_line(monitor.DeviceID),
                        "RefreshRate": refresh,
                        "DisplayLabel": display_label,
                    }
                )
            monitor_index += 1
        adapter_index += 1
    return devices


def gather_sections() -> dict:
    raw = run_ps_file(_GATHER_SCRIPT, timeout=45)
    if raw == "N/A":
        data = {}
    else:
        try:
            data = json.loads(raw)
        except Exception:
            data = {}

    def g(key: str, default: str = "N/A") -> str:
        v = data.get(key)
        if v is None:
            return default
        return first_line(str(v))

    system_manufacturer = g("SystemManufacturer")
    product_name = g("ProductName")
    system_version_index = g("SystemVersionIndex")
    system_serial = g("SystemSerial")
    system_uuid = g("SystemUUID")
    family_serial = g("FamilySerial")
    sku_number = g("SKUNumber")

    bios_vendor = g("BIOSVendor")
    bios_version = g("BIOSVersion")
    release_date = g("ReleaseDate")
    core_isolation = bool_to_status(g("CoreIsolationRaw"))
    virtualization = bool_to_status(g("VirtualizationRaw"))
    secure_boot = bool_to_status(g("SecureBootRaw"))
    tpm_status = bool_to_status(g("TpmStatusRaw"))

    bb_manufacturer = g("BBManufacturer")
    bb_version = g("BBVersion")
    bb_product = g("BBProduct")
    bb_serial = g("BBSerial")
    bb_asset_wmi = g("BBAsset")

    smb_hex_raw = data.get("SMBiosHex")
    smb_hex = ""
    if isinstance(smb_hex_raw, str):
        smb_hex = smb_hex_raw.replace(" ", "").strip()
    bb_smb: dict[str, str] = {}
    if smb_hex and len(smb_hex) % 2 == 0:
        try:
            bb_smb = parse_smbios_baseboard_type2(bytes.fromhex(smb_hex))
        except (ValueError, TypeError):
            bb_smb = {}

    bb_asset = first_line(bb_smb.get("asset", "") or "")
    if not bb_asset or bb_asset.lower() in ("default string", "to be filled by o.e.m.", "none"):
        bb_asset = bb_asset_wmi
    if bb_asset.lower() in ("base board", "baseboard", "to be filled by o.e.m.", "default string", ""):
        bb_asset = "N/A"
    if bb_asset not in ("N/A", "") and bb_asset == bb_serial:
        bb_asset = "N/A"

    cs_location = first_line(bb_smb.get("location", "") or "")
    if not cs_location or cs_location == "N/A":
        cs_location = g("CSLocation")

    cpu_manu = g("CPUManu")
    cpu_type = g("CPUType")
    cpu_serial = format_cpu_serial_display(g("CPUSerial"))
    cpu_part = g("CPUPart")
    cpu_asset = g("CPUAsset")
    cpu_socket = g("CPUSocket")

    chassis_manu = g("ChassisManu")
    chassis_type = g("ChassisType")
    chassis_version = g("ChassisVersion")
    chassis_serial = g("ChassisSerial")
    chassis_asset = g("ChassisAsset")
    chassis_sku = g("ChassisSKU")
    if chassis_sku in ("N/A", "", "None"):
        chassis_sku = "N/A"

    mac_cache = g("MAC")
    net_ipv4 = g("NetIPv4")

    network_physical: list[tuple[str, str]] = []
    raw_np = data.get("NetPhysical")
    if isinstance(raw_np, dict):
        raw_np = [raw_np]
    if isinstance(raw_np, list):
        for item in raw_np[:24]:
            if not isinstance(item, dict):
                continue
            mac = first_line(str(item.get("Mac", "N/A")))
            kind = first_line(str(item.get("Kind", "dynamic"))).lower()
            if mac in ("N/A", "", "None"):
                continue
            if kind not in ("static", "dynamic"):
                kind = "dynamic"
            network_physical.append((mac, kind))

    pci_device = g("PciDevice")
    gpu_name = g("GpuName")
    guid_serial = g("GuidSerial")
    # GUID Serial like GPU-57885081-4160-... comes from DirectX Device Identifier (dxdiag), not WMI VideoProcessor.
    dx_guid = fetch_gpu_guid_from_dxdiag()
    if dx_guid != "N/A":
        guid_serial = dx_guid
    else:
        fb = gpu_guid_fallback_from_pnp(pci_device)
        if fb != "N/A":
            guid_serial = fb

    tpm_base = g("TpmBase")
    tpm_md5, tpm_sha1, tpm_sha256 = hash_triplet(tpm_base)

    disk_items = data.get("Disks") or []
    if isinstance(disk_items, dict):
        disk_items = [disk_items]
    physical_disk_items = data.get("PhysicalDisks") or []
    if isinstance(physical_disk_items, dict):
        physical_disk_items = [physical_disk_items]
    monitor_items = data.get("Monitors") or []
    if isinstance(monitor_items, dict):
        monitor_items = [monitor_items]

    disk_rows = [build_disk_rows(item) for item in disk_items] or [[
        ("Caption", "N/A"),
        ("DISK_STORAGE_MODEL", "N/A"),
        ("STORAGE_QUERY_PROPERTY", "N/A"),
        ("SMART_RCV_DRIVE_DATA", "N/A"),
        ("STORAGE_QUERY_WWN", "N/A"),
        ("SCSI_PASS_THROUGH", "N/A"),
        ("ATA_PASS_THROUGH", "N/A"),
    ]]

    for i, rows in enumerate(disk_rows):
        if i >= len(physical_disk_items):
            continue
        p = physical_disk_items[i]
        enriched = []
        for key, value in rows:
            if key == "DISK_STORAGE_MODEL":
                v = first_line(str(p.get("Model", "N/A")))
                enriched.append((key, v if v != "N/A" else value))
            elif key == "STORAGE_QUERY_PROPERTY":
                v = first_line(str(p.get("BusType", "N/A")))
                enriched.append((key, v if v != "N/A" else value))
            elif key == "SMART_RCV_DRIVE_DATA":
                v = first_line(str(p.get("HealthStatus", "N/A")))
                enriched.append((key, v if v != "N/A" else value))
            elif key == "STORAGE_QUERY_WWN":
                unique_id = first_line(str(p.get("UniqueId", "N/A")))
                unique_fmt = first_line(str(p.get("UniqueIdFormat", "N/A")))
                if unique_id != "N/A":
                    enriched.append((key, f"{unique_id} ({unique_fmt})"))
                else:
                    enriched.append((key, value))
            elif key == "ATA_PASS_THROUGH":
                v = first_line(str(p.get("SerialNumber", "N/A")))
                enriched.append((key, v if v != "N/A" else value))
            elif key == "Caption":
                v = first_line(str(p.get("FriendlyName", "N/A")))
                enriched.append((key, v if v != "N/A" else value))
            else:
                enriched.append((key, value))
        disk_rows[i] = enriched

    fallback_displays = get_display_devices_fallback()

    monitor_rows = []
    for i, item in enumerate(monitor_items):
        mon_id = first_line(str(item.get("IDSerialNumber", "N/A")))
        mon_key = mon_id.upper()
        active_display = first_line(str(item.get("Active", "N/A")))
        if i < len(fallback_displays):
            al = first_line(str(fallback_displays[i].get("Active", "N/A")))
            if al != "N/A":
                active_display = al
        else:
            for disp in fallback_displays:
                dk = first_line(disp.get("IDSerialNumber", "N/A")).upper()
                if mon_key != "N/A" and mon_key in dk:
                    al = first_line(str(disp.get("Active", "N/A")))
                    if al != "N/A":
                        active_display = al
                    break
        monitor_rows.append(
            [
                ("Active Monitor", active_display),
                ("Manufacturer", first_line(str(item.get("Manufacturer", "N/A")))),
                ("Model Name", first_line(str(item.get("ModelName", "N/A")))),
                ("Monitor Serial", first_line(str(item.get("MonitorSerial", "N/A")))),
                ("ID Serial Number", mon_id),
            ]
        )

    existing_ids = {
        row[4][1] for row in monitor_rows if len(row) > 4 and row[4][1] != "N/A"
    }
    # If WMI already found displays, only use fallback to fill missing entries.
    needed_fallback = len(monitor_rows) == 0
    for disp in fallback_displays:
        if not needed_fallback:
            break
        disp_id = first_line(disp.get("IDSerialNumber", "N/A"))
        if disp_id in existing_ids:
            continue
        monitor_rows.append(
            [
                ("Active Monitor", first_line(disp.get("Active", "N/A"))),
                ("Manufacturer", first_line(disp.get("Manufacturer", "N/A"))),
                ("Model Name", first_line(disp.get("ModelName", "N/A"))),
                ("Monitor Serial", first_line(disp.get("MonitorSerial", "N/A"))),
                ("ID Serial Number", disp_id),
            ]
        )
    if not monitor_rows:
        monitor_rows = [[
            ("Active Monitor", "N/A"),
            ("Manufacturer", "N/A"),
            ("Model Name", "N/A"),
            ("Monitor Serial", "N/A"),
            ("ID Serial Number", "N/A"),
        ]]

    return {
        "base_sections": {
            "System Information": [
                ("Manufacturer", system_manufacturer),
                ("Product Name", product_name),
                ("Version Index", system_version_index),
                ("System Serial", system_serial),
                ("System UUID", system_uuid),
                ("Family Serial", family_serial),
                ("SKU Number", sku_number),
            ],
            "BIOS Information": [
                ("BIOS Vendor", bios_vendor),
                ("BIOS Version", bios_version),
                ("Release Date", release_date),
                ("Core Isolation", core_isolation),
                ("Virtualization", virtualization),
                ("Secure Boot", secure_boot),
                ("TPM Status", tpm_status),
            ],
            "Baseboard Information": [
                ("Manufacturer", bb_manufacturer),
                ("Version Index", bb_version),
                ("Product Name", bb_product),
                ("Serial Number", bb_serial),
                ("Asset Number", bb_asset),
                ("(CS) Location", cs_location),
            ],
            "Processor Information": [
                ("CPU Manufacturer", cpu_manu),
                ("Processor Type", cpu_type),
                ("Serial Number", cpu_serial),
                ("Part Number", cpu_part),
                ("Asset Number", cpu_asset),
                ("Processor Socket", cpu_socket),
            ],
            "Chassis/Enclosure Information": [
                ("Manufacturer", chassis_manu),
                ("Chassis Type", chassis_type),
                ("Version Index", chassis_version),
                ("Serial Number", chassis_serial),
                ("Asset Number", chassis_asset),
                ("SKU Number", chassis_sku),
            ],
            "Network Information": [
                ("MAC [Cache1]", mac_cache),
                ("Local IP", net_ipv4 if net_ipv4 not in ("N/A", "", "None") else "N/A"),
            ],
            "GPU Information": [
                ("PCI Device", pci_device),
                ("GPU Name", gpu_name),
                ("GUID Serial", guid_serial),
            ],
            "TPM Information": [
                ("MD5", tpm_md5),
                ("SHA1", tpm_sha1),
                ("SHA256", tpm_sha256),
            ],
        },
        "disk_rows": disk_rows,
        "monitor_rows": monitor_rows,
        "network_physical": network_physical,
    }


def _value_color(val: str) -> str:
    v = first_line(val)
    if v == "Enabled":
        return GREEN_OK
    if v == "Disabled":
        return RED_BAD
    return TEXT_VALUE


def add_info_block(
    parent: tk.Widget,
    title: str,
    rows: list[tuple[str, str]],
    nav: tuple[str, callable, callable] | None = None,
    value_refs: list[tk.Label] | None = None,
    header_style: str = "blue",
) -> None:
    frame = tk.Frame(parent, bg=BG_CARD, bd=0, highlightbackground=BORDER_CARD, highlightthickness=1)
    frame.pack(fill="x", pady=8)

    header_fg = HEADER_BLUE if header_style == "blue" else HEADER_ORANGE
    header_row = tk.Frame(frame, bg=BG_CARD)
    header_row.pack(fill="x")
    tk.Label(
        header_row,
        text=title,
        font=("Segoe UI", 10, "bold"),
        fg=header_fg,
        bg=BG_CARD,
        anchor="w",
        padx=10,
        pady=8,
    ).pack(side="left")

    if nav:
        page_text, on_prev, on_next = nav
        nav_wrap = tk.Frame(header_row, bg=BG_CARD)
        nav_wrap.pack(side="right", padx=8)
        tk.Button(nav_wrap, text="<", command=on_prev, bg="#1F2937", fg="white", relief="flat", width=3).pack(side="left", padx=2)
        tk.Label(nav_wrap, text=page_text, fg=TEXT_VALUE, bg=BG_CARD, font=("Segoe UI", 9)).pack(side="left", padx=4)
        tk.Button(nav_wrap, text=">", command=on_next, bg="#1F2937", fg="white", relief="flat", width=3).pack(side="left", padx=2)

    for idx, (key, value) in enumerate(rows):
        if idx > 0:
            tk.Frame(frame, bg="#111827", height=1).pack(fill="x", padx=8)
        row = tk.Frame(frame, bg=BG_CARD)
        row.pack(fill="x", padx=10, pady=4)
        tk.Label(row, text=f"{key}:", fg=TEXT_LABEL, bg=BG_CARD, font=("Segoe UI", 9), anchor="w").pack(side="left")
        val_lbl = tk.Label(
            row,
            text=value if value else "N/A",
            fg=_value_color(value if value else ""),
            bg=BG_CARD,
            font=("Segoe UI", 9),
            anchor="w",
            wraplength=230,
            justify="left",
        )
        val_lbl.pack(side="left", padx=(8, 0))
        if value_refs is not None:
            value_refs.append(val_lbl)


def clear_widget(widget: tk.Widget) -> None:
    for child in widget.winfo_children():
        child.destroy()


def _phys_kind_color(kind: str) -> str:
    k = kind.lower()
    if k == "dynamic":
        return GREEN_OK
    if k == "static":
        return HEADER_ORANGE
    return TEXT_VALUE


def render_network_information(
    parent: tk.Widget,
    rows: list[tuple[str, str]],
    physical_rows: list[tuple[str, str]],
) -> None:
    for idx, (key, value) in enumerate(rows):
        if idx > 0:
            tk.Frame(parent, bg="#111827", height=1).pack(fill="x", padx=8)
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill="x", padx=10, pady=4)
        tk.Label(row, text=f"{key}:", fg=TEXT_LABEL, bg=BG_CARD, font=("Segoe UI", 9), anchor="w").pack(side="left")
        val_lbl = tk.Label(
            row,
            text=value if value else "N/A",
            fg=_value_color(value if value else ""),
            bg=BG_CARD,
            font=("Segoe UI", 9),
            anchor="w",
            wraplength=230,
            justify="left",
        )
        val_lbl.pack(side="left", padx=(8, 0))
    for mac, kind in physical_rows:
        tk.Frame(parent, bg="#111827", height=1).pack(fill="x", padx=8)
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill="x", padx=10, pady=4)
        tk.Label(row, text="Physical Address:", fg=TEXT_LABEL, bg=BG_CARD, font=("Segoe UI", 9), anchor="w").pack(
            side="left"
        )
        mid = tk.Frame(row, bg=BG_CARD)
        mid.pack(side="left", fill="x", expand=True)
        tk.Label(mid, text=mac, fg=TEXT_VALUE, bg=BG_CARD, font=("Segoe UI", 9), anchor="w").pack(side="left", padx=(8, 0))
        tk.Label(row, text=kind, fg=_phys_kind_color(kind), bg=BG_CARD, font=("Segoe UI", 9, "bold"), anchor="e").pack(
            side="right", padx=8
        )


def render_kv_rows(parent: tk.Widget, rows: list[tuple[str, str]], value_refs: list[tk.Label] | None = None) -> None:
    for idx, (key, value) in enumerate(rows):
        if idx > 0:
            tk.Frame(parent, bg="#111827", height=1).pack(fill="x", padx=8)
        row = tk.Frame(parent, bg=BG_CARD)
        row.pack(fill="x", padx=10, pady=4)
        tk.Label(row, text=f"{key}:", fg=TEXT_LABEL, bg=BG_CARD, font=("Segoe UI", 9), anchor="w").pack(side="left")
        val_lbl = tk.Label(
            row,
            text=value if value else "N/A",
            fg=_value_color(value if value else ""),
            bg=BG_CARD,
            font=("Segoe UI", 9),
            anchor="w",
            wraplength=230,
            justify="left",
        )
        val_lbl.pack(side="left", padx=(8, 0))
        if value_refs is not None:
            value_refs.append(val_lbl)


def minimize_window(root: tk.Tk) -> None:
    """Borderless windows: avoid root.iconify() on Windows — it can crash with overrideredirect."""

    def _minimize() -> None:
        try:
            root.update_idletasks()
            hwnd = int(root.winfo_id())
            GA_ROOT = 2
            root_hwnd = ctypes.windll.user32.GetAncestor(hwnd, GA_ROOT)
            if root_hwnd:
                hwnd = root_hwnd
            WM_SYSCOMMAND = 0x0112
            SC_MINIMIZE = 0xF020
            # PostMessage is safer than iconify() for undecorated HWND
            ctypes.windll.user32.PostMessageW(hwnd, WM_SYSCOMMAND, SC_MINIMIZE, 0)
        except Exception:
            try:
                hwnd = int(root.winfo_id())
                SW_MINIMIZE = 6
                ctypes.windll.user32.ShowWindow(hwnd, SW_MINIMIZE)
            except Exception:
                pass

    root.after(1, _minimize)


def start_window_drag(root: tk.Tk, event: tk.Event) -> None:
    # Screen-consistent offset so dragging works from any child widget, not only the title bar
    root._drag_x = event.x_root - root.winfo_rootx()
    root._drag_y = event.y_root - root.winfo_rooty()


def on_window_drag(root: tk.Tk, event: tk.Event) -> None:
    x = event.x_root - getattr(root, "_drag_x", 0)
    y = event.y_root - getattr(root, "_drag_y", 0)
    root.geometry(f"+{x}+{y}")


def bind_drag_recursive(container: tk.Widget, root: tk.Tk) -> None:
    """Allow moving the borderless window by dragging any non-button surface."""

    def _bd(w: tk.Widget) -> None:
        w.bind("<Button-1>", lambda e: start_window_drag(root, e))
        w.bind("<B1-Motion>", lambda e: on_window_drag(root, e))

    for child in container.winfo_children():
        bind_drag_recursive(child, root)
        if isinstance(child, tk.Button):
            continue
        _bd(child)


def _he(s: object) -> str:
    return html_lib.escape(first_line(str(s)) if s is not None else "N/A", quote=True)


def _app_bundle_dir() -> Path:
    """Folder next to the script (dev) or the .exe (PyInstaller --onefile)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _report_txt_dir() -> Path:
    """Where serial_report.txt is written: same folder as .exe when bundled, else cwd."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


def _app_state_path() -> Path:
    return _app_bundle_dir() / "serial_checker_state.json"


def _desktop_path() -> Path:
    return Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"


def gather_data_jsonable(data: dict) -> dict:
    def conv(x: object) -> object:
        if isinstance(x, dict):
            return {k: conv(v) for k, v in x.items()}
        if isinstance(x, (list, tuple)):
            return [conv(i) for i in x]
        return x

    return conv(data)


def gather_data_from_saved(obj: dict) -> dict:
    if not obj:
        return {}

    def rows_to_tuples(rows: object) -> list:
        if not isinstance(rows, list):
            return []
        out: list = []
        for r in rows:
            if isinstance(r, (list, tuple)) and len(r) == 2:
                out.append((str(r[0]), str(r[1])))
        return out

    bs_raw = obj.get("base_sections") or {}
    base_sections: dict[str, list[tuple[str, str]]] = {}
    for name, rows in bs_raw.items():
        base_sections[name] = rows_to_tuples(rows)

    disk_rows: list[list[tuple[str, str]]] = []
    for disk in obj.get("disk_rows") or []:
        if isinstance(disk, list):
            tuples = rows_to_tuples(disk)
            disk_rows.append([("Caption", v) if k == "Driver" else (k, v) for k, v in tuples])

    monitor_rows: list[list[tuple[str, str]]] = []
    for mon in obj.get("monitor_rows") or []:
        if isinstance(mon, list):
            monitor_rows.append(rows_to_tuples(mon))

    netp: list[tuple[str, str]] = []
    for row in obj.get("network_physical") or []:
        if isinstance(row, (list, tuple)) and len(row) == 2:
            netp.append((str(row[0]), str(row[1])))

    return {
        "base_sections": base_sections,
        "disk_rows": disk_rows,
        "monitor_rows": monitor_rows,
        "network_physical": netp,
    }


def _save_export_state(html_p: str, json_p: str) -> None:
    try:
        with open(_app_state_path(), "w", encoding="utf-8") as f:
            json.dump({"html": html_p, "json": json_p}, f, indent=2)
    except OSError:
        pass


def load_last_export_json_path() -> str | None:
    p = _app_state_path()
    if not p.is_file():
        return None
    try:
        with open(p, encoding="utf-8") as f:
            st = json.load(f)
        jp = st.get("json")
        if jp and Path(jp).is_file():
            return str(jp)
    except (OSError, json.JSONDecodeError):
        pass
    return None


_HASH_VALUE_LABELS = frozenset(
    {
        "PCI Device",
        "GUID Serial",
        "STORAGE_QUERY_WWN",
        "MD5",
        "SHA1",
        "SHA256",
        "DISK_STORAGE_MODEL",
        "STORAGE_QUERY_PROPERTY",
        "SMART_RCV_DRIVE_DATA",
        "SCSI_PASS_THROUGH",
        "ATA_PASS_THROUGH",
        "Caption",
    }
)

_VAL_DEFAULT_MARKERS = frozenset(
    {"default string", "n/a", "to be filled by o.e.m.", "none", ""}
)


def _default_val_class(val: str) -> str:
    if first_line(val).strip().lower() in _VAL_DEFAULT_MARKERS:
        return " default"
    return ""


def _format_dump_value(val: str, key: str) -> str:
    v = first_line(str(val))
    if v == "Enabled":
        return '<span class="status enabled">Enabled</span>'
    if v == "Disabled":
        return '<span class="status disabled">Disabled</span>'
    use_hash = key in _HASH_VALUE_LABELS or len(v) > 96 or "PCI\\" in v or v.startswith("GPU-")
    body = _he(v)
    if use_hash:
        return f'<div class="hash-value">{body}</div>'
    return body


def _dump_section_block(title: str, rows: list[tuple[str, str]]) -> str:
    inner = []
    for key, val in rows:
        vc = _default_val_class(str(val))
        inner.append(
            '<div class="info-row"><div class="info-label">'
            + _he(key)
            + '</div><div class="info-value'
            + vc
            + '">'
            + _format_dump_value(str(val), key)
            + "</div></div>"
        )
    return (
        '<div class="section"><div class="section-header">'
        + _he(title)
        + '</div><div class="section-content">'
        + "".join(inner)
        + "</div></div>"
    )


def _dump_disk_combined(disk_rows: list[list[tuple[str, str]]]) -> str:
    merged: list[tuple[str, str]] = []
    for d in disk_rows:
        merged.extend(d)
    return _dump_section_block("Disk Drive Information", merged)


def _dump_monitors(monitors: list[list[tuple[str, str]]]) -> str:
    parts = []
    for mon in monitors:
        sub: list[str] = []
        hdr = "Monitor Information"
        for key, val in mon:
            if key == "Active Monitor":
                hdr = first_line(str(val))
                continue
            vc = _default_val_class(str(val))
            sub.append(
                '<div class="info-row"><div class="info-label">'
                + _he(key)
                + '</div><div class="info-value'
                + vc
                + '">'
                + _format_dump_value(str(val), key)
                + "</div></div>"
            )
        parts.append(
            '<div class="monitor-item"><div class="monitor-header">'
            + _he(hdr)
            + "</div>"
            + "".join(sub)
            + "</div>"
        )
    return (
        '<div class="section"><div class="section-header">Monitor Information</div><div class="section-content">'
        + "".join(parts)
        + "</div></div>"
    )


def _net_kv(data: dict, key: str) -> str:
    for k, v in data.get("base_sections", {}).get("Network Information", []):
        if k == key:
            return first_line(str(v))
    return "N/A"


def _dump_network(data: dict) -> str:
    mac = _net_kv(data, "MAC [Cache1]")
    lip = _net_kv(data, "Local IP")
    items = [
        f'<div class="mac-item"><div class="mac-label">MAC [Cache1]</div><div class="mac-address">{_he(mac)}</div></div>',
        f'<div class="mac-item"><div class="mac-label">Local IP</div><div class="mac-address">{_he(lip)}</div></div>',
    ]
    for macp, kind in data.get("network_physical", []):
        items.append(
            f'<div class="mac-item"><div class="mac-label">Physical Address ({_he(kind)})</div>'
            f'<div class="mac-address">{_he(macp)}</div></div>'
        )
    return (
        '<div class="section"><div class="section-header">Network Information</div><div class="section-content">'
        f'<div class="network-grid">{"".join(items)}</div></div></div>'
    )


def _dump_arp(data: dict) -> str:
    lip = _net_kv(data, "Local IP")
    hdr = f"INTERFACE: {lip} --- 0x0"
    parts = []
    for mac, kind in data.get("network_physical", []):
        parts.append(
            f'<div class="arp-entry"><div class="hash-value">{_he(mac)}</div><span class="arp-type">{_he(kind)}</span></div>'
        )
    if not parts:
        parts.append('<div class="arp-entry"><div class="hash-value">N/A</div><span class="arp-type">—</span></div>')
    return (
        '<div class="section"><div class="section-header">ARP Information</div><div class="section-content">'
        f'<div class="arp-grid"><div class="arp-interface"><div class="arp-interface-header">{_he(hdr)}</div>'
        f'{"".join(parts)}</div></div></div></div>'
    )


def _dump_gpu(data: dict) -> str:
    rows = data.get("base_sections", {}).get("GPU Information", [])
    return _dump_section_block("GPU Information", rows)


def _dump_tpm(data: dict) -> str:
    rows = data.get("base_sections", {}).get("TPM Information", [])
    return _dump_section_block("TPM Information", rows)


_DUMP_STYLES = """*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Poppins',sans-serif;background:#0a0a0f;color:#e0e0e0;line-height:1.6;min-height:100vh;overflow-x:hidden}
body::before{content:'';position:fixed;top:0;left:0;right:0;bottom:0;background:radial-gradient(circle at 20% 50%,rgba(120,119,198,.3)0%,transparent 50%),radial-gradient(circle at 80% 80%,rgba(98,126,234,.2)0%,transparent 50%),radial-gradient(circle at 40% 20%,rgba(139,92,246,.2)0%,transparent 50%);z-index:-1}
.container{max-width:1400px;margin:0 auto;padding:20px}
header{background:rgba(26,26,46,.6);backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,.1);padding:25px 40px;border-radius:16px;margin-bottom:30px;box-shadow:0 8px 32px
rgba(0,0,0,.3);display:flex;align-items:center;justify-content:space-between}
.header-content{display:flex;align-items:center;gap:20px;flex-wrap:wrap}
.status-badges{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.badge{padding:6px 16px;border-radius:8px;font-size:13px;font-weight:500;border:1px solid rgba(255,255,255,.1)}
h1{font-size:2.2em;font-weight:900;background:linear-gradient(135deg,#e0e7ff 0%,#c7d2fe 50%,#a5b4fc 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
letter-spacing:.5px}
.sections-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(450px,1fr));gap:20px;margin-bottom:20px}
.section{background:rgba(26,26,46,.4);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.1);border-radius:12px;overflow:hidden}
.section-header{background:rgba(30,30,50,.6);padding:16px 24px;font-size:.9em;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:#a5b4fc;border-bottom:1px solid
rgba(255,255,255,.1);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.section-content{padding:20px 24px}
.info-row{display:flex;padding:12px 0;align-items:center;border-bottom:1px solid rgba(255,255,255,.05);flex-wrap:wrap;gap:4px}
.info-row:last-child{border-bottom:none}
.info-label{flex:0 0 40%;min-width:140px;color:#94a3b8;font-size:.9em;font-weight:500}
.info-value{flex:1;color:#e2e8f0;font-size:.9em;word-break:break-word}
.info-value.default{color:#64748b;font-style:italic}
.status{padding:2px 10px;border-radius:4px;font-size:.85em;font-weight:500;display:inline-block}
.status.enabled{background:rgba(34,197,94,.2);color:#4ade80;border:1px solid rgba(34,197,94,.3)}
.status.disabled{background:rgba(239,68,68,.2);color:#f87171;border:1px solid rgba(239,68,68,.3)}
.network-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px}
.mac-item{background:rgba(30,30,50,.4);padding:12px 16px;border-radius:8px;border:1px solid rgba(255,255,255,.08)}
.mac-label{color:#94a3b8;font-size:.8em;margin-bottom:4px}
.mac-address{color:#e2e8f0;font-size:.9em;font-weight:500}
.arp-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:12px}
.arp-interface{background:rgba(30,30,50,.4);padding:16px;border-radius:8px;border:1px solid rgba(255,255,255,.08)}
.arp-interface-header{color:#a5b4fc;font-weight:600;margin-bottom:12px;font-size:.95em}
.arp-entry{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid rgba(255,255,255,.05);font-size:.85em}
.arp-entry:last-child{border-bottom:none}
.arp-type{color:#94a3b8;font-size:.9em}
.hash-value{font-family:monospace;font-size:.85em;font-weight:500;background:rgba(0,0,0,.3);padding:8px 12px;border-radius:6px;word-break:break-all;color:#cbd5e1}
.monitor-item{background:rgba(30,30,50,.4);padding:20px;border-radius:8px;margin-bottom:16px;border:1px solid rgba(255,255,255,.08)}
.monitor-header{color:#a5b4fc;font-weight:600;margin-bottom:16px;font-size:1.1em}
.gpu-item{background:rgba(30,30,50,.4);padding:20px;border-radius:8px;border:1px solid rgba(255,255,255,.08)}
.footer-info{background:rgba(26,26,46,.4);border:1px solid rgba(255,255,255,.1);border-radius:12px;padding:16px 24px;margin-top:20px;color:#94a3b8;font-size:.9em}
"""


def build_html_serial_dump(data: dict, title_suffix: str) -> str:
    base = data["base_sections"]
    grid1 = "".join(
        [
            _dump_section_block("System Information", base.get("System Information", [])),
            _dump_section_block("BIOS Information", base.get("BIOS Information", [])),
            _dump_section_block("Baseboard Information", base.get("Baseboard Information", [])),
            _dump_section_block("Processor Information", base.get("Processor Information", [])),
            _dump_disk_combined(data.get("disk_rows", [])),
            _dump_section_block("Chassis/Enclosure Information", base.get("Chassis/Enclosure Information", [])),
        ]
    )
    grid2 = '<div class="sections-grid">' + _dump_network(data) + _dump_monitors(data.get("monitor_rows", [])) + "</div>"
    grid3 = '<div class="sections-grid">' + _dump_arp(data) + _dump_gpu(data) + "</div>"
    tpm_only = _dump_tpm(data).replace(
        '<div class="section">', '<div class="section" style="grid-column:1/-1;">', 1
    )

    cr = html_lib.escape(APP_CREDIT, quote=False)
    body = (
        f'<div class="sections-grid">{grid1}</div>{grid2}{grid3}{tpm_only}'
        f'<div class="footer-info">SERIAL CHECKER — exported {html_lib.escape(title_suffix, quote=True)} — '
        f"{html_lib.escape(str(datetime.now().strftime('%Y-%m-%d %H:%M:%S')), quote=True)}"
        f"<br>{cr}</div>"
    )
    fonts = '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">'  # noqa: E501
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="author" content="{html_lib.escape(APP_AUTHOR, quote=True)}">
<title>SERIAL CHECKER DUMP</title>
{fonts}
<style>{_DUMP_STYLES}</style>
</head>
<body>
<div class="container">
<header>
<div class="header-content">
<h1>SERIAL CHECKER</h1>
</div>
</header>
{body}
</div>
</body>
</html>
"""


def _cmp_val_inner(val: str, key: str, as_hash: bool = False) -> str:
    v = first_line(str(val))
    if v == "Enabled":
        return '<span class="status enabled">Enabled</span>'
    if v == "Disabled":
        return '<span class="status disabled">Disabled</span>'
    use_hash = as_hash or key in _HASH_VALUE_LABELS or len(v) > 96 or "PCI\\" in v or v.startswith("GPU-")
    body = _he(v)
    if use_hash:
        return f'<div class="hash-value">{body}</div>'
    return body


def _cmp_row(label: str, old_v: str, new_v: str, key: str) -> str:
    ch = "changed" if first_line(str(old_v)) != first_line(str(new_v)) else ""
    dc_old = " default" if _default_val_class(str(old_v)) else ""
    dc_new = " default" if _default_val_class(str(new_v)) else ""
    wwn = " wwn-row" if key == "STORAGE_QUERY_WWN" else ""
    exp = " expandable" if key == "STORAGE_QUERY_WWN" else ""
    o = _cmp_val_inner(str(old_v), key)
    n = _cmp_val_inner(str(new_v), key)
    if exp:
        o = o.replace("hash-value", f"hash-value{exp}", 1)
        n = n.replace("hash-value", f"hash-value{exp}", 1)
    return (
        f'<div class="comparison-row{wwn}"><div class="comparison-label">{_he(label)}</div>'
        f'<div class="comparison-old{dc_old}">{o}</div>'
        f'<div class="comparison-new{dc_new}{(" " + ch) if ch else ""}">{n}</div></div>'
    )


def _compare_section_rows(title: str, old_rows: list[tuple[str, str]], new_rows: list[tuple[str, str]]) -> str:
    n = max(len(old_rows), len(new_rows))
    parts = []
    for i in range(n):
        ok = old_rows[i][0] if i < len(old_rows) else ""
        ov = old_rows[i][1] if i < len(old_rows) else ""
        nk = new_rows[i][0] if i < len(new_rows) else ""
        nv = new_rows[i][1] if i < len(new_rows) else ""
        lbl = ok or nk
        parts.append(_cmp_row(lbl, ov, nv, ok or nk))
    t = title
    return (
        f'<div class="section"><div class="section-header">{_he(t)}</div>'
        f'<div class="section-content">{"".join(parts)}</div></div>'
    )


def _compare_disk_sections(
    old_disks: list[list[tuple[str, str]]], new_disks: list[list[tuple[str, str]]]
) -> str:
    n = max(len(old_disks), len(new_disks))
    return "".join(
        _compare_section_rows(
            f"Disk Drive #{i + 1}",
            old_disks[i] if i < len(old_disks) else [],
            new_disks[i] if i < len(new_disks) else [],
        )
        for i in range(n)
    )


def _compare_monitor_sections(
    old_mons: list[list[tuple[str, str]]], new_mons: list[list[tuple[str, str]]]
) -> str:
    n = max(len(old_mons), len(new_mons))
    return "".join(
        _compare_section_rows(
            f"Monitor #{i + 1}",
            old_mons[i] if i < len(old_mons) else [],
            new_mons[i] if i < len(new_mons) else [],
        )
        for i in range(n)
    )


def _compare_network_cols(old_d: dict, new_d: dict) -> str:
    def _net_lines(data: dict) -> list[tuple[str, str]]:
        rows = [("MAC [CACHE1]", _net_kv(data, "MAC [Cache1]")), ("LOCAL IP", _net_kv(data, "Local IP"))]
        for mac, kind in data.get("network_physical", []):
            rows.append((str(kind).upper(), mac))
        return rows

    def _entry_row(lbl: str, val: str, mark: bool) -> str:
        cls = "network-entry changed" if mark else "network-entry"
        return (
            f'<div class="{cls}"><span class="network-label">{_he(lbl)}</span>'
            f'<span class="hash-value">{_he(val)}</span></div>'
        )

    o_lines = _net_lines(old_d)
    n_lines = _net_lines(new_d)
    m = max(len(o_lines), len(n_lines))
    lo_html = ""
    ln_html = ""
    for i in range(m):
        ol = o_lines[i] if i < len(o_lines) else ("—", "")
        nl = n_lines[i] if i < len(n_lines) else ("—", "")
        ch = ol != nl
        lo_html += _entry_row(ol[0], ol[1], ch)
        ln_html += _entry_row(nl[0], nl[1], ch)

    return (
        f'<div class="section"><div class="section-header">Network Information</div><div class="section-content">'
        f'<div class="network-comparison"><div class="network-column"><div class="network-header">Previous Network Data</div>{lo_html}</div>'
        f'<div class="network-column"><div class="network-header">Current Network Data</div>{ln_html}</div></div></div></div>'
    )


def _compare_arp_block(old_d: dict, new_d: dict) -> str:
    oip = _net_kv(old_d, "Local IP")
    nip = _net_kv(new_d, "Local IP")
    op = old_d.get("network_physical", [])
    np_ = new_d.get("network_physical", [])
    m = max(len(op), len(np_), 1)
    rows = []
    for i in range(m):
        om = f"{op[i][0]} {op[i][1]}" if i < len(op) else ""
        nm = f"{np_[i][0]} {np_[i][1]}" if i < len(np_) else "[REMOVED]"
        if i >= len(np_):
            nm = "[REMOVED]"
        if i >= len(op):
            om = ""
        ch = om != nm
        rows.append(
            f'<div class="arp-entry-comparison"><div class="arp-old"><span class="hash-value">{_he(om)}</span></div>'
            f'<div class="arp-new{" changed" if ch else ""}"><span class="hash-value">{_he(nm)}</span></div></div>'
        )
    hdr = f"INTERFACE: {nip} --- 0x0"
    return (
        f'<div class="section"><div class="section-header">ARP Information</div><div class="section-content">'
        f'<div class="arp-comparison"><div class="arp-interface-comparison"><div class="arp-interface-header">{_he(hdr)}</div>'
        f'{"".join(rows)}</div></div></div></div>'
    )


_COMPARE_STYLES = _DUMP_STYLES + """
.badge.changes{background:rgba(34,197,94,.2);color:#4ade80}
.comparison-row{display:flex;padding:12px 0;align-items:flex-start;border-bottom:1px solid rgba(255,255,255,.05);flex-wrap:wrap;gap:4px}
.comparison-label{flex:0 0 20%;min-width:120px;color:#94a3b8;font-size:.85em;font-weight:500}
.comparison-old{flex:0 0 42%;color:#e2e8f0;font-size:.85em;word-break:break-word;padding-right:14px}
.comparison-new{flex:1;min-width:120px;color:#e2e8f0;font-size:.85em;word-break:break-word;padding-left:14px}
.comparison-new.changed{color:#4ade80;font-weight:500}
.comparison-old.default,.comparison-new.default{color:#64748b;font-style:italic}
.network-comparison{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.network-column{background:rgba(30,30,50,.4);padding:16px;border-radius:8px;border:1px solid rgba(255,255,255,.08)}
.network-header{color:#a5b4fc;font-weight:600;margin-bottom:12px;text-align:center}
.network-entry{padding:6px 0;font-size:.85em;border-bottom:1px solid rgba(255,255,255,.05)}
.network-entry.changed{color:#4ade80;font-weight:500}
.network-label{display:inline-block;min-width:125px;margin-right:8px}
.arp-comparison{margin-bottom:16px}
.arp-interface-comparison{background:rgba(30,30,50,.4);padding:16px;border-radius:8px;border:1px solid rgba(255,255,255,.08)}
.arp-entry-comparison{display:flex;padding:6px 0;border-bottom:1px solid rgba(255,255,255,.05);font-size:.8em}
.arp-old,.arp-new{flex:1;padding:0 8px}
.arp-old{border-right:1px solid rgba(255,255,255,.1)}
.arp-new.changed{color:#4ade80;font-weight:500}
.comparison-row.wwn-row .hash-value.expandable{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:pointer;max-height:28px}
.comparison-row.wwn-row:hover .hash-value.expandable{white-space:normal;word-break:break-all;max-height:200px}
"""


def build_html_serial_comparison(old: dict, new: dict) -> str:
    ob = old.get("base_sections", {})
    nb = new.get("base_sections", {})
    order = [
        "System Information",
        "Baseboard Information",
        "Chassis/Enclosure Information",
        "BIOS Information",
        "Processor Information",
    ]
    chunks = []
    for sec in order:
        chunks.append(
            _compare_section_rows(
                sec.replace("Chassis/Enclosure Information", "Chassis Information"),
                ob.get(sec, []),
                nb.get(sec, []),
            )
        )
    chunks.append(_compare_disk_sections(old.get("disk_rows", []), new.get("disk_rows", [])))
    chunks.append(_compare_network_cols(old, new))
    chunks.append(_compare_arp_block(old, new))
    chunks.append(_compare_monitor_sections(old.get("monitor_rows", []), new.get("monitor_rows", [])))
    chunks.append(_compare_section_rows("GPU Information", ob.get("GPU Information", []), nb.get("GPU Information", [])))
    chunks.append(_compare_section_rows("TPM Information", ob.get("TPM Information", []), nb.get("TPM Information", [])))

    fonts = '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">'  # noqa: E501
    inner = "".join(chunks)
    stamp = html_lib.escape(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), quote=True)
    cr = html_lib.escape(APP_CREDIT, quote=False)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="author" content="{html_lib.escape(APP_AUTHOR, quote=True)}">
<title>SERIAL CHECKER COMPARISON</title>
{fonts}
<style>{_COMPARE_STYLES}</style>
</head>
<body>
<div class="container">
<header>
<div class="header-content">
<h1>SERIAL COMPARISON</h1>
</div>
<div class="status-badges"><span class="badge changes">Change Detection</span></div>
</header>
{inner}
<div class="footer-info">SERIAL CHECKER — comparison generated {stamp}<br>{cr}</div>
</div>
</body>
</html>
"""


def export_data(data: dict) -> None:
    """Write TXT (legacy), full HTML dump, and JSON snapshot on Desktop for later Check."""
    lines = []
    base_sections = data["base_sections"]
    for section_name, rows in base_sections.items():
        lines.append(f"[{section_name}]")
        for key, value in rows:
            lines.append(f"{key}: {value}")
        if section_name == "Network Information":
            for mac, kind in data.get("network_physical", []):
                lines.append(f"Physical Address: {mac} ({kind})")
        lines.append("")

    for i, rows in enumerate(data["disk_rows"], start=1):
        lines.append(f"[Disk Drive Information #{i}]")
        for key, value in rows:
            lines.append(f"{key}: {value}")
        lines.append("")

    for i, rows in enumerate(data["monitor_rows"], start=1):
        lines.append(f"[Monitor Information #{i}]")
        for key, value in rows:
            lines.append(f"{key}: {value}")
        lines.append("")

    lines.append(APP_CREDIT)

    report_dir = _report_txt_dir()
    report_txt = report_dir / "serial_report.txt"
    with open(report_txt, "w", encoding="utf-8") as report_file:
        report_file.write("\n".join(lines))

    stamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    desk = _desktop_path()
    try:
        desk.mkdir(parents=True, exist_ok=True)
    except OSError:
        desk = report_dir
    base_name = f"SERIAL_DUMP_{stamp}"
    html_path = desk / f"{base_name}.html"
    json_path = desk / f"{base_name}.json"

    html_doc = build_html_serial_dump(data, stamp)
    with open(html_path, "w", encoding="utf-8") as hf:
        hf.write(html_doc)
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump(gather_data_jsonable(data), jf, indent=2, ensure_ascii=False)
    _save_export_state(str(html_path), str(json_path))

    tkmsg.showinfo("Export", f"Saved:\n{html_path}\n{json_path}\n{report_txt}")


def run_serial_check_then_refresh(root: tk.Tk) -> None:
    jp = load_last_export_json_path()
    if not jp:
        tkmsg.showwarning(
            "Check",
            "No previous export found. Use Export first — it saves HTML + JSON on your Desktop.",
        )
        return
    try:
        with open(jp, encoding="utf-8") as f:
            raw = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        tkmsg.showerror("Check", f"Could not read export JSON:\n{e}")
        return
    old_data = gather_data_from_saved(raw)
    new_data = gather_sections()
    html_doc = build_html_serial_comparison(old_data, new_data)
    stamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    out = _desktop_path() / f"SERIAL_COMPARISON_{stamp}.html"
    try:
        _desktop_path().mkdir(parents=True, exist_ok=True)
    except OSError:
        out = _report_txt_dir() / f"SERIAL_COMPARISON_{stamp}.html"
    try:
        with open(out, "w", encoding="utf-8") as wf:
            wf.write(html_doc)
    except OSError as e:
        tkmsg.showerror("Check", f"Could not write comparison:\n{e}")
        return
    tkmsg.showinfo("Check", f"Comparison saved:\n{out}")
    refresh_ui(root)


def refresh_ui(root: tk.Tk) -> None:
    root.destroy()
    main()


def build_ui(data: dict) -> None:
    base = data["base_sections"]
    disks = data["disk_rows"]
    monitors = data["monitor_rows"]
    net_physical = data.get("network_physical", [])
    disk_index = [0]
    monitor_index = [0]
    root = tk.Tk()
    root.title(f"SERIAL CHECKER — {APP_AUTHOR}")
    root.geometry("1500x860")
    root.configure(bg=BG_MAIN)
    root.overrideredirect(True)

    title_bar = tk.Frame(root, bg=BG_MAIN, height=52)
    title_bar.pack(fill="x", padx=10, pady=(10, 4))
    title_bar.grid_columnconfigure(0, weight=1)
    title_bar.grid_columnconfigure(1, weight=1)
    title_bar.grid_columnconfigure(2, weight=1)
    def _bind_drag(w: tk.Widget) -> None:
        w.bind("<Button-1>", lambda e: start_window_drag(root, e))
        w.bind("<B1-Motion>", lambda e: on_window_drag(root, e))

    _bind_drag(title_bar)

    left_area = tk.Frame(title_bar, bg=BG_MAIN)
    left_area.grid(row=0, column=0, sticky="w")

    mid_title = tk.Frame(title_bar, bg=BG_MAIN)
    mid_title.grid(row=0, column=1, sticky="nsew")
    center_title = tk.Frame(mid_title, bg=TITLE_BOX_BG, highlightbackground=TITLE_BOX_FG, highlightthickness=1)
    # Optical center over middle column (nudge right within middle title cell)
    center_title.place(relx=0.68, rely=0.5, anchor="center")
    serial_lbl = tk.Label(
        center_title,
        text="SERIAL CHECKER",
        fg=TITLE_BOX_FG,
        bg=TITLE_BOX_BG,
        font=("Segoe UI", 10, "bold"),
        padx=14,
        pady=6,
    )
    serial_lbl.pack()
    for w in (mid_title, center_title, serial_lbl):
        _bind_drag(w)

    right_btns = tk.Frame(title_bar, bg=BG_MAIN)
    right_btns.grid(row=0, column=2, sticky="e")

    def btn(parent, text, bg_c, fg_c="white", cmd=None, w=0):
        b = tk.Button(
            parent,
            text=text,
            command=cmd,
            bg=bg_c,
            fg=fg_c,
            activebackground=bg_c,
            activeforeground=fg_c,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=10 if w == 0 else w,
            pady=4,
            cursor="hand2",
        )
        b.pack(side="left", padx=3)
        return b

    btn(right_btns, "Close", BTN_CLOSE, "white", root.destroy)
    btn(right_btns, "Export", BTN_EXPORT, "white", lambda: export_data(data))
    btn(right_btns, "Check", BTN_CHECK, "white", lambda: run_serial_check_then_refresh(root))
    btn(left_area, "Discord", BTN_DISCORD, "white", lambda: webbrowser.open("https://discord.gg/I_DONT_HAVE_SERVER_YET"))

    credit_bar = tk.Frame(root, bg=BG_MAIN)
    credit_bar.pack(side="bottom", fill="x", pady=(0, 8))
    tk.Label(
        credit_bar,
        text=APP_CREDIT,
        fg="#6B7280",
        bg=BG_MAIN,
        font=("Segoe UI", 8),
    ).pack()

    wrapper = tk.Frame(root, bg=BG_MAIN)
    wrapper.pack(fill="both", expand=True, padx=14, pady=8)
    canvas = tk.Canvas(wrapper, bg=BG_MAIN, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    content = tk.Frame(canvas, bg=BG_MAIN)
    canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")

    def _scroll_region(_: tk.Event | None = None) -> None:
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _sync_canvas_width(event: tk.Event) -> None:
        canvas.itemconfigure(canvas_window, width=event.width)
        _scroll_region()

    canvas.bind("<Configure>", _sync_canvas_width)
    content.bind("<Configure>", lambda _: _scroll_region())

    def _on_mousewheel(event: tk.Event) -> None:
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    content.grid_columnconfigure(0, weight=1, uniform="cols")
    content.grid_columnconfigure(1, weight=1, uniform="cols")
    content.grid_columnconfigure(2, weight=1, uniform="cols")
    col1 = tk.Frame(content, bg=BG_MAIN, width=420)
    col2 = tk.Frame(content, bg=BG_MAIN, width=420)
    col3 = tk.Frame(content, bg=BG_MAIN, width=420)
    col1.grid(row=0, column=0, sticky="nsew", padx=(4, 6))
    col2.grid(row=0, column=1, sticky="nsew", padx=6)
    col3.grid(row=0, column=2, sticky="nsew", padx=(6, 4))

    add_info_block(col1, "System Information", base["System Information"])
    add_info_block(col1, "BIOS Information", base["BIOS Information"])
    add_info_block(col1, "Baseboard Information", base["Baseboard Information"])

    disk_container = tk.Frame(col2, bg=BG_MAIN)
    disk_container.pack(fill="x")
    add_info_block(col2, "Processor Information", base["Processor Information"])
    add_info_block(col2, "Chassis/Enclosure Information", base["Chassis/Enclosure Information"])

    net_card = tk.Frame(col3, bg=BG_CARD, highlightbackground=BORDER_CARD, highlightthickness=1)
    net_card.pack(fill="x", pady=8)
    net_hdr = tk.Frame(net_card, bg=BG_CARD)
    net_hdr.pack(fill="x")
    tk.Label(net_hdr, text="Network Information", font=("Segoe UI", 10, "bold"), fg=HEADER_BLUE, bg=BG_CARD, padx=10, pady=8).pack(
        side="left"
    )
    net_body = tk.Frame(net_card, bg=BG_CARD)
    net_body.pack(fill="x")
    render_network_information(net_body, base["Network Information"], net_physical)

    monitor_container = tk.Frame(col3, bg=BG_MAIN)
    monitor_container.pack(fill="x", pady=8)

    add_info_block(col3, "GPU Information", base["GPU Information"])
    add_info_block(col3, "TPM (INTC)", base["TPM Information"])

    def render_disk() -> None:
        clear_widget(disk_container)
        add_info_block(
            disk_container,
            "Disk Drive Information",
            disks[disk_index[0]],
            nav=(f"{disk_index[0] + 1}/{len(disks)}", prev_disk, next_disk),
        )
        bind_drag_recursive(disk_container, root)

    def render_monitor() -> None:
        clear_widget(monitor_container)
        add_info_block(
            monitor_container,
            "Monitor Information",
            monitors[monitor_index[0]],
            nav=(f"{monitor_index[0] + 1}/{len(monitors)}", prev_monitor, next_monitor),
        )
        bind_drag_recursive(monitor_container, root)

    def prev_disk() -> None:
        disk_index[0] = (disk_index[0] - 1) % len(disks)
        render_disk()

    def next_disk() -> None:
        disk_index[0] = (disk_index[0] + 1) % len(disks)
        render_disk()

    def prev_monitor() -> None:
        monitor_index[0] = (monitor_index[0] - 1) % len(monitors)
        render_monitor()

    def next_monitor() -> None:
        monitor_index[0] = (monitor_index[0] + 1) % len(monitors)
        render_monitor()

    render_disk()
    render_monitor()
    bind_drag_recursive(wrapper, root)
    bind_drag_recursive(content, root)

    root.mainloop()


def main() -> None:
    data = gather_sections()
    build_ui(data)


if __name__ == "__main__":
    main()
