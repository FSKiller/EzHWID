# EzHWID
Easy to use and robust DL/KMS38 activation script with added ability to switch edition
[Overview] Easy to use script that defeats activation on Windows 10 and
Server 2016/2019/2022 on UEFI-GPT systems using SLIC emulation
technique.

[How to] Run the script. Press [D] to activate permanently (only on
Client operating systems). If everything went well you should be
activated!

Press [K] to activate till 2038 (also supported on Server)

Press [X] to check activation status.

Press [C] to create $OEM$ folder which can be placed in source directory
of Windows installation media/ISO to preactivate. A $OEM$ folder which
uses traditional script will be created when traditional script is used
to create $OEM$ folder. A $OEM$ folder which uses AIO script will be
created when AIO script is used to create $OEM$ folder.

Press [R] to view this read-me document.

Press [G] to open the Github repo of this project.

Press [S] to switch your edition (Windows 10 1803+ only).

Press [H] to open the discord server of this project.

Press [F] to toggle forceful mode.

Switches: /dl: Start installation process for DL activation and skip
main menu. /k38: Start installation process for KMS38 activation and
skip main menu. /silent: Used with /dl or /k38. Don't prompt user for
anything. /force: This option bypasses the message "Windows is also
permanently activated." and also force installs a generic key, even if
one is already installed. /pkey: Used for specifying product key to be
insalled

Example of syntax: EzHWID.cmd /dl /silent /force /pkey
P42PH-HYD6B-Y3DHY-B79JH-CT8YK

Q: Which operating systems are supported? A: HWID: Windows 10/11 Home (N
+ SL + China) Windows 10/11 Pro (WS + Edu) (N) Windows 10/11 Education
(N) Windows 10/11 (IoT) Enterprise (Multi-session / N) Windows 11 SE (N)
KMS38: Windows 10/11: Enterprise, Enterprise LTSC/LTSB, Enterprise G,
Enterprise multi-session, Enterprise, Education, Pro, Pro Workstation,
Pro Education, Home, Home Single Language, Home China + N editions
Windows Server 2022/2019/2016: LTSC editions (Standard, Datacenter,
Essentials, Cloud Storage, Azure Core, Server ARM64), SAC editions
(Standard ACor, Datacenter ACor, Azure Datacenter)

Q: I get error "PowerShell is not installed in your system". How to fix?
A: Enable powershell using optional features control panel.

Q: How to upgrade core editions to non-core editions? A: Open an admin
command prompt.

Check which key you want to install, depending on the edition you want
to upgrade to: Core: YTMG3-N6DKC-DKB77-7M9GH-8HVX7 CoreN:
4CPRK-NM3K3-X6XXQ-RXX86-WXCHW CoreSingleLanguage:
BT79Q-G7N6G-PGBYW-4YWX6-6F4BT CoreCountrySpecific:
N2434-X9D7W-8PF6X-8DV9T-8TYMD Professional:
VK7JG-NPHTM-C97JM-9MPGT-3V66T ProfessionalN:
2B87N-8KFHP-DKV6R-Y2C8J-PKCKT ProfessionalWorkstation:
DXG7C-N36C4-C4HTG-X4T3X-2YV77 ProfessionalWorkstationN:
WYPNQ-8C467-V2W6J-TX4WX-WT2RQ ProfessionalEducation:
8PTT6-RNW4C-6V7J2-C2D3X-MHBPB ProfessionalEducationN:
GJTYN-HDMQY-FRR76-HVGC7-QPF8P Education: YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY
EducationN: 84NGF-MHBT6-FXBX8-QWJK7-DRR8H Enterprise:
XGVPP-NMH47-7TTHJ-W3FW7-8HV2C EnterpriseN: 3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT
ServerRdsh: NJCF7-PW8QT-3324D-688JX-2YV66 IoTEnterprise:
XQQYW-NFFMW-XJPBH-K8732-CKFFD CloudEdition:
KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W CloudEditionN:
K9VKN-3BGWV-Y624W-MCRMQ-BHDCD

Disconnect from internet (IMPORTANT) run this command: changepk
/productkey [enter your key here] Wait for it to complete, ignore any
errors. After it finishes, reboot your system.

[Credits] @mspaintmsi for integrated patcher method @Windows_Addict and
@abbodi1406 for scripting ideas and great assistance in scripting
