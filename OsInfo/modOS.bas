Attribute VB_Name = "modOS"

'Detecting OS Windows 95 - Windows 7 and some OS Info, pappsegull@yahoo.se dec 20 2011
'I did notice that GetVersionEx return Win XP when I run Win 7 64, but this functions return it correct on my PC
'Add more info if you need here is the whole class: http://msdn.microsoft.com/en-us/library/windows/desktop/aa394239(v=vs.85).aspx

'Class Win32_OperatingSystem: CIM_OperatingSystem
'{
'  string   BootDevice;
'  string   BuildNumber;
'  string   BuildType;
'  string   Caption;
'  string   CodeSet;
'  string   CountryCode;
'  string   CreationClassName;
'  string   CSCreationClassName;
'  string   CSDVersion;
'  string   CSName;
'  sint16   CurrentTimeZone;
'  boolean  DataExecutionPrevention_Available;
'  boolean  DataExecutionPrevention_32BitApplications;
'  boolean  DataExecutionPrevention_Drivers;
'  uint8    DataExecutionPrevention_SupportPolicy;
'  boolean  Debug;
'  string   Description;
'  boolean  Distributed;
'  uint32   EncryptionLevel;
'  uint8    ForegroundApplicationBoost;
'  uint64   FreePhysicalMemory;
'  uint64   FreeSpaceInPagingFiles;
'  uint64   FreeVirtualMemory;
'  datetime InstallDate;
'  uint32   LargeSystemCache;
'  datetime LastBootUpTime;
'  datetime LocalDateTime;
'  string   Locale;
'  string   Manufacturer;
'  uint32   MaxNumberOfProcesses;
'  uint64   MaxProcessMemorySize;
'  string   MUILanguages[];
'  string   Name;
'  uint32   NumberOfLicensedUsers;
'  uint32   NumberOfProcesses;
'  uint32   NumberOfUsers;
'  uint32   OperatingSystemSKU;
'  string   Organization;
'  string   OSArchitecture;
'  uint32   OSLanguage;
'  uint32   OSProductSuite;
'  uint16   OSType;
'  string   OtherTypeDescription;
'  Boolean  PAEEnabled;
'  string   PlusProductID;
'  string   PlusVersionNumber;
'  boolean  Primary;
'  uint32   ProductType;
'  string   RegisteredUser;
'  string   SerialNumber;
'  uint16   ServicePackMajorVersion;
'  uint16   ServicePackMinorVersion;
'  uint64   SizeStoredInPagingFiles;
'  string   Status;
'  uint32   SuiteMask;
'  string   SystemDevice;
'  string   SystemDirectory;
'  string   SystemDrive;
'  uint64   TotalSwapSpaceSize;
'  uint64   TotalVirtualMemorySize;
'  uint64   TotalVisibleMemorySize;
'  string   Version;
'  string   WindowsDirectory;
'};

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type
Enum OSVerEnum
    [Not Detected]
    [Windows 95]
    [Windows 98]
    [Windows ME]
    [Windows NT 3.51]
    [Windows NT 4.0]
    [Windows 2000]
    [Windows XP]
    [Windows 2003]
    [Windows Vista]
    [Windows 7]
End Enum
Enum VerInfo
    [Return Enum]
    [Return Platform]
    [Return Major]
    [Return Minor]
End Enum
Enum OsInfo
    [OS Enum]
    [Version]
    [Boot Device]
    [Build Number]
    [Build Type]
    [Caption]
    [Code Set]
    [Country Code]
    [IsDebug]
    [Encryption Level]
    [Install Date]
    [Licensed Users]
    [Organization]
    [OS Language]
    [OS Product Suite]
    [OS Platform ID]
    [OS Major Version]
    [OS Minor Version]
    [ServicePack Major Version]
    [ServicePack Minor Version]
    [OS Type]
    [IsPrimary]
    [Number Of Licensed Users]
    [Registered User]
    [Serial Number]
End Enum


Function GetOSVerEnum(Optional VersionInfo As VerInfo) As OSVerEnum
'                Win95 Win98 WinME NT 3.51 NT 4.0 Win2000 WinXP Win2003 Vista  Win7
' ------------------------------------------------------------------------------------
'PlatFormID      1     1     1      2      2      2       2     2       2      2
'MajorVersion    4     4     4      3      4      5       5     5       6      6
'MinorVersion    0    10    90     51      0      0       1     2       0      1

Dim v As Variant, OS As OSVERSIONINFO, MajorVersion%, MinorVersion%
    v = GetOSInfo(Version): v = Split(v, ".")
    MajorVersion% = v(0): MinorVersion% = v(1)
    With OS
        .OSVSize = Len(OS): If Not CBool(GetVersionEx(OS)) Then Exit Function
        Select Case VersionInfo
            Case [Return Platform]: GetOSVerEnum = .PlatformID
            Case [Return Major]: GetOSVerEnum = MajorVersion%
            Case [Return Minor]: GetOSVerEnum = MinorVersion%
        End Select
        If VersionInfo <> [Return Enum] Then Exit Function
        Select Case .PlatformID
            Case 1 '< NT 3.51
                Select Case MajorVersion
                    Case 4
                        Select Case MinorVersion
                            Case 0: v = [Windows 95]
                            Case 10: v = [Windows 98]
                            Case 90: v = [Windows ME]
                            Case Else: v = [Not Detected]
                        End Select
                    Case Else: v = [Not Detected]
                End Select
            Case 2
                Select Case MajorVersion
                    Case 3: v = [Windows NT 3.51]
                    Case 4: v = [Windows NT 4.0]
                    Case 5 '< Vista
                        Select Case MinorVersion
                            Case 0: v = [Windows 2000]
                            Case 1: v = [Windows XP]
                            Case 2: v = [Windows 2003]
                            Case Else: v = [Not Detected]
                        End Select
                    Case 6
                        Select Case MinorVersion
                            Case 0: v = [Windows Vista]
                            Case 1: v = [Windows 7]
                            Case Else: v = [Not Detected]
                        End Select
                    Case Else: v = [Not Detected]
                End Select
        End Select
    End With
    GetOSVerEnum = v
End Function

Function GetOSInfo(InfoID As OsInfo) As Variant
Static v As Variant, O As Object, Obj As Object, ObjCvDate As Object, b As Boolean
    If Not b Then
        Set ObjCvDate = CreateObject("WbemScripting.SWbemDateTime")
        Set Obj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
        Set Obj = Obj.ExecQuery("Select * from Win32_OperatingSystem"): b = True
    End If
    For Each O In Obj
        Select Case InfoID
            Case [Version]:                       v = O.Version
            Case [Boot Device]:                   v = O.BootDevice
            Case [Build Number]:                  v = O.BuildNumber
            Case [Build Type]:                    v = O.BuildType
            Case [Caption]:                       v = O.Caption
            Case [Code Set]:                      v = O.CodeSet
            Case [Country Code]:                  v = O.CountryCode
            Case [IsDebug]:                       v = CBool(O.Debug)
            Case [Encryption Level]:              v = O.EncryptionLevel
            Case [Install Date]
                ObjCvDate.Value = O.InstallDate:  v = CDate(ObjCvDate.GetVarDate)
            Case [Number Of Licensed Users]:      v = O.NumberOfLicensedUsers
            Case [Organization]:                  v = O.Organization
            Case [OS Language]:                   v = O.OSLanguage
            Case [OS Product Suite]:              v = O.OSProductSuite
            Case [OS Type]:                       v = O.OSType
            Case [IsPrimary]:                     v = CBool(O.Primary)
            Case [Registered User]:               v = O.RegisteredUser
            Case [Serial Number]:                 v = O.SerialNumber
            Case [OS Enum]:                       v = GetOSVerEnum
            Case [OS Platform ID]:                v = GetOSVerEnum([Return Platform])
            Case [OS Major Version]:              v = GetOSVerEnum([Return Major])
            Case [OS Minor Version]:              v = GetOSVerEnum([Return Minor])
            Case [ServicePack Major Version]:     v = O.ServicePackMajorVersion
            Case [ServicePack Minor Version]:     v = O.ServicePackMinorVersion
        End Select
    Next
    GetOSInfo = v
End Function

'Just a demo of the module.
Sub Main()
Dim s$: s$ = "I am running Win7 x64 bit and VB6 works fine;)"
    If GetOSVerEnum < [Windows 7] Then _
      s$ = s$ & vbLf & "Time to upgrade " & GetOSInfo([Registered User]) & "?"
    MsgBox s$ & vbLf & vbLf & "Your OS is: " & vbTab & _
      GetOSInfo(Caption) & "  " & GetOSInfo(Version) & vbLf & _
      "GetOSVerEnum(): " & vbTab & GetOSInfo([OS Enum]) & vbLf & _
      "Install Date: " & vbTab & GetOSInfo([Install Date]) & vbLf & _
      "Licensed users: " & vbTab & GetOSInfo([Number Of Licensed Users]) & vbLf & _
      "Registered User: " & vbTab & GetOSInfo([Registered User]) & vbLf & _
      "Serial Number: " & vbTab & GetOSInfo([Serial Number]) & vbLf & _
      "Organization: " & vbTab & GetOSInfo(Organization) & vbLf & _
      "Product Suite: " & vbTab & GetOSInfo([OS Product Suite]) & vbLf & _
      "Platform ID: " & vbTab & GetOSInfo([OS Platform ID]) & vbLf & _
      "Major Version: " & vbTab & GetOSInfo([OS Major Version]) & vbLf & _
      "Minor Version: " & vbTab & GetOSInfo([OS Minor Version]) & vbLf & _
      "ServicePack Major: " & vbTab & GetOSInfo([ServicePack Major Version]) & vbLf & _
      "ServicePack Minor: " & vbTab & GetOSInfo([ServicePack Minor Version]) & vbLf & _
      "OS Version: " & vbTab & GetOSInfo(Version) & vbLf & _
      "OS Build Number: " & vbTab & GetOSInfo([Build Number]) & vbLf & _
      "OS Language: " & vbTab & GetOSInfo([OS Language]) & vbLf & _
      "OS Type: " & vbTab & vbTab & GetOSInfo([OS Type]) & vbLf & _
      "Code Set: " & vbTab & vbTab & GetOSInfo([Code Set]) & vbLf & _
      "Country Code: " & vbTab & GetOSInfo([Country Code]) & vbLf & _
      "Encryption Level: " & vbTab & GetOSInfo([Encryption Level]) & vbLf & _
      "Build Type: " & vbTab & GetOSInfo([Build Type]) & vbLf & _
      "Boot Device: " & vbTab & GetOSInfo([Boot Device]) & _
      vbLf, vbInformation, "Some info from my GetOSInfo() function"
End Sub
