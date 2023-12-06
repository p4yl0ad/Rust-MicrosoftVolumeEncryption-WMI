# Rust-MicrosoftVolumeEncryption-WMI
Does what it says on the tin.

## How
Queries the `root\cimv2` WMI namespace and fetches each of the drive letters by selecting the `Caption` property from the `Win32_LogicalDisk` class, for each of the drive letters it queries the `root\CIMV2\Security\MicrosoftVolumeEncryption` WMI namespace and selects all properties from the `Win32_EncryptableVolume` class where the drive letter is the current iterations drive letter, in order to fetch the `DeviceID`, `PersistentVolumeID`, `DriveLetter` and `ProtectionStatus`.

## Usage 
```
C:\Users\devmachine\Desktop>MicrosoftVolumeEncryption.exe
[i] Drive Letter        : C:
[*] WMI Query           : SELECT * FROM Win32_EncryptableVolume Where DriveLetter='C:'
[*] DeviceID            : \\?\Volume{bfd080a8-b428-481f-bf69-93b4b04d3f07}\
[*] PersistentVolumeID  :
[*] DriveLetter         : C:
[*] ProtectionStatus    : 0


[i] Drive Letter        : D:
[*] WMI Query           : SELECT * FROM Win32_EncryptableVolume Where DriveLetter='D:'
```


## Dependencies 
```
[dependencies.windows]
version = "0.52.0"
features = [
    "Win32_Foundation",
    "Win32_Security",
    "Win32_System_Com",
    "Win32_System_Com_Marshal",
    "Win32_System_Com_StructuredStorage",
    "Win32_System_Com_Urlmon",
    "Win32_System_Wmi",
    "Win32_System_Ole",
    "Win32_System_Variant"
]
```
