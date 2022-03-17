# MKT
Official thread page https://forums.mydigitallife.net/threads/microsoft-toolkit-continued.82782/

# Microsoft Toolkit v2.7.2

- Official KMS Solution for Microsoft Products by @CODYQX4
- This is a set of tools and functions for managing licensing, deploying, and activating Microsoft Office and Windows.
- All output from these functions is displayed in the Information Console.
- All functions are run in the background and the GUI is disabled to prevent running multiple functions, as they could conflict or cause damage if run concurrently.
- Office Setup Customization Functions (Customize Setup Tab), AutoKMS Uninstaller (if AutoKMS is installed), AutoRearm Uninstaller (if AutoRearm is installed), Office Uninstaller and Product Key Checker work even if Microsoft Office or Windows is not installed/supported.
- For information about individual functions, see the program Readme.
- Activation Instructions & other Tutorials by @Razorsharp1

## Requirements:

- Microsoft .NET Framework 4.0-4.8 (Not 3.5)
- Microsoft Office 2010 or Later for Office Toolkit Support
- Windows Vista or Later for Windows Toolkit Support

## Changelog:

2.7.2

- Support for Windows 11/10 ARM64 (except WinDivert Client)
- Support for Microsoft Office 2021
- Fixed issue running WinDivert Loader
- Fixed activation for ServerRdsh edition
- Implemented fix for Office C2R non-genuine banner
- Updated TAP Drivers
- Updated Keys and Key Checker

2.7.1

- Improved Windows Defender exclusions for Windows 10 or 8.1
- Updated DLL Injection Localhost Bypass to use Avrf custom provider
- Updated KMSEmulator PID Generator for better compatibility between KMS Host OS/pKey
- Support to detect and skip Windows 10 KMS 2038
- Fixed Activate button not working if the installed key is not KMS
- Proper handling for Office 2010 Service Pack 2 when added to ISO
- Disabled Office ISO Channel Switcher and removed its payload
- Updated Office Uninstaller scripts
- Improved Office Click To Run detection
- Preliminary support for Microsoft Office 2021 Preview
- Support for Windows Server 2022
- Updated Keys and Key Checker

2.7 / Implemented by @CODYQX4

- Add Windows Defender exclusions while activating if on Windows 10 or later
- DLL Injection Localhost Bypass now uses Image File Execution Options
- Fixed KMSEmulator not working with ports that aren't 4 digits
- Support for Microsoft Office 2019
- Support for Windows Server 2019
- Updates Keys and Key Checker

## Credits:

- Bosh for the original GUI Design and co-development of Office Toolkit
- ZWT for the original KMSEmulator
- letsgoawayhell, Phazor, nosferati87, mikmik38, Hotbird64, and qad for KMSEmulator fixes and improvements
- cynecx, zm0d, murphy78, abbodi1406, and namazso for LocalHost Bypass improvements
- MasterDisaster, FreeStyler, Daz, nononsense, and janek2012 for work on Key Checker

## Downloads:

```
- v2.7.2

MTKV272.7z / Microsoft Toolkit.exe
7z file password: 2021

https://drive.google.com/uc?export=download&id=1OYZfFRUGp1wR9YIQDO65JCZIlY5Zgey8
https://download.ru/files/cU1dYwF6

=
Bonus:
MTKV272_Activation-Resources.7z / AutoKMS.exe, AutoRearm.exe ...
7z file password: 2021

https://drive.google.com/uc?export=download&id=1IZ9URnpeV8uj52Mpu_WRcfz_5oxoU3FG
https://download.ru/files/vgVtV2dI
```

## Checksums:

- v2.7.2
```
  File: MTKV272.7z
  SHA-1: 2f259dd8ba60a771e3e51224e56aef9e5901f687
SHA-256: 66fd5d2b2d949d15866f1354618df987d39d897cd647feaebd0367a1cc436fa4

   File: Microsoft Toolkit.exe
  SHA-1: edbeb2c9178d3903ea180bb709121f154fc5dec1
SHA-256: d8da2ae7d74835f071b490ef89d5e3412e0b810eaf1c1b9104d3d229b1b36251

   File: AutoKMS.exe
  SHA-1: c94317ca1107ae3a92a556de20bbc17f570eab49
SHA-256: f46c91176f9bfbf4c9234000e724d0314bacd98dd4f5f145a89416584187a2ac

   File: AutoRearm.exe
  SHA-1: 9855d787855e42444e3f25e8401c48b334b2cbd7
SHA-256: a370526405bbabad034b3aa39d737a4ae2d178654a21eca70ee8f8a9b378ddf2

   File: KMS Server Service.exe
  SHA-1: 1e978061a0d4cb9a566d4b6b5d6049ea673931dc
SHA-256: 8693c647f29f39a90e8ad6da0e8a792ca6e90b961166d68bdbc0059d3111975c

   File: WinDivert Loader.exe
  SHA-1: ac989ad8e970cf76d09f10f715a8cc38298c274e
SHA-256: e0b9d9f25ae11466ed94ce6c3bed06a82de902e0413c8d72acd7e32b67398814

   File: SECOHook.dll (x86)
  SHA-1: 64a01421587c81f5a6980130bf41416ff8d74d50
SHA-256: d244d08c226c0ff85f1ae56c46476899a5ef7706ebb337f461d0587bfa6971d2

   File: SECOHook.dll (x64)
  SHA-1: b7beb3b3b2e9558bd5bf22e5d131c9e783aa9316
SHA-256: 465607fef8d64391d3f61ce187d90b0d1028bb58ecac26e4e747188c04a11c1d

   File: SECOHook.dll (arm64)
  SHA-1: 886b1d5de77daa66aa37979de333a41ce8b72379
SHA-256: ba1f71085747a780ff1560df7fec75ccdcf3de8fc2d77345fb6566bc0480ba5d
```
- v2.7.1
```
   File: MTKV271.7z
  SHA-1: ea5bddae69b655901a3b2c1c9bcb8584d0915c0d
SHA-256: 9f6410335714cb040448e473825e8fe0c41564e0b537ddabea5129f3ef9407a7

   File: Microsoft Toolkit.exe
  SHA-1: 48a2429e9b17d91ff158dfdf896e62874c66bec7
SHA-256: 0774dc41d6a7956a9e551e52db40944eed0362f60067faf7f08920e062df7108
```

## Recommendation for usage with Antivirus products:

- Create a new folder where you will download or place Microsoft Toolkit.exe into
- Exclude the folder in your AV product
- Extract the downloaded 7z file into the folder
- If you are using Windows Defender on Windows 11/10/8.1, Microsoft Toolkit will add the needed exclusion for activation files during process
- If you are using other security products, either temporary turn off AV program, or exclude these paths:

```
- DLL Injection Localhost Bypass for Win 11/10/8.1 only:
C:\Windows\System32\SECOHook.dll

- WinDivert Localhost Bypass for Win 11/10/8.1 only:
"C:\ProgramData\Microsoft Toolkit\WinDivert-KMS\WinDivert Loader.exe"

- if any of those are installed (default paths, change to correct path if changed in settings):
C:\Windows\AutoKMS\AutoKMS.exe
C:\Windows\AutoRearm\AutoRearm.exe
"C:\Windows\KMSServerService\KMS Server Service.exe"
```

## Source Code:
```
   File: MTKSource272.7z
  SHA-1: feab244b59a4e32f31dfaff330297801a64e72db
SHA-256: 0c662e0305b6c895e890c29e7f1ec9666a8b08df3eb62afdd8224e8b60ae7f8c

pass: 2021
https://drive.google.com/uc?export=download&id=1_vHwfKq6Qps7kmtV1h4EUBOaK40c-zNE

  File: MTKSource272.zip
  SHA-1: 4003796481475e9aea1cc5f18533eef882e4d3e2
SHA-256: 42dac40812143c0a7b529096fcfde5dc76abc08645996a1268cb0b58e9ba68e3

https://download.ru/files/U7SBcTFg
```
