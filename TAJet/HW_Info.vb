Imports System
Imports System.Management
Imports Microsoft.Win32
Imports System.Runtime.InteropServices

Public Class ComputerInfo
    Private Function CaptureDos(ByVal Filename As String, Optional ByVal Parms As String = vbNullString) As String
        Dim Exec As New System.Diagnostics.Process()
        Dim Buffer As String = vbNullString

        Try
            With Exec
                .StartInfo.RedirectStandardOutput = True
                .StartInfo.UseShellExecute = False
                .StartInfo.CreateNoWindow = True
                .StartInfo.FileName = Filename
                .StartInfo.Arguments = Parms
                .Start()
                'Read in output.
                Buffer = .StandardOutput.ReadToEnd()
                'Wait for exit.
                Exec.WaitForExit()
                'Return string.
                Return Buffer
            End With

        Catch ex As Exception
            Return vbNullString
        End Try
    End Function
    Public Function GetHDSerial() As String
        Dim Serial As String = vbNullString
        Dim Line As String = CaptureDos("cmd.exe", "/c vol")
        Dim idx As Integer = 0

        'Loop backwards until we find a space.
        For Count As Integer = Line.Length - 1 To 0 Step -1
            'Exit if space is found.
            If Line(Count).Equals(" "c) Then
                idx = Count
                Exit For
            End If
        Next Count

        If (idx <> 0) Then
            Return Line.Substring(idx + 1)
        End If

        Return vbNullString
    End Function
    Public Function GetProcessorId() As String
        Dim strProcessorId As String = String.Empty
        Dim query As New SelectQuery("Win32_processor")
        Dim search As New ManagementObjectSearcher(query)
        Dim info As ManagementObject

        For Each info In search.Get()
            strProcessorId = info("processorId").ToString()
        Next
        Return strProcessorId

    End Function

    Public Function GetMACAddress() As String

        Dim mc As ManagementClass = New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()
        Dim MACAddress As String = String.Empty
        For Each mo As ManagementObject In moc

            If (MACAddress.Equals(String.Empty)) Then
                If CBool(mo("IPEnabled")) Then MACAddress = mo("MacAddress").ToString()

                mo.Dispose()
            End If
            MACAddress = MACAddress.Replace(":", String.Empty)

        Next
        Return MACAddress
    End Function

    Public Function GetVolumeSerial(Optional ByVal strDriveLetter As String = "C") As String

        Dim disk As ManagementObject = New ManagementObject(String.Format("win32_logicaldisk.deviceid=""{0}:""", strDriveLetter))
        disk.Get()
        Return disk("VolumeSerialNumber").ToString()
    End Function

    Public Function GetMotherBoardID() As String

        Dim strMotherBoardID As String = String.Empty
        Dim query As New SelectQuery("Win32_BaseBoard")
        Dim search As New ManagementObjectSearcher(query)
        Dim info As ManagementObject
        For Each info In search.Get()

            strMotherBoardID = info("SerialNumber").ToString()

        Next
        Return strMotherBoardID

    End Function

End Class
Public Class SystemInfo

#Region " Public Enums "
    Public Enum OSArch
        x64
        x86
    End Enum
#End Region

#Region " Public Properties "
    ''' <summary>
    ''' Gets a value indicating the current BIOS revision of the unit.
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Public ReadOnly Property BiosRevision As String
        Get
            Return Me.GetBiosRevision
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the current serial number of the unit
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Public ReadOnly Property SerialNumber As String
        Get
            Return Me.GetSerialNumber()
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the hardware manufacturer
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Public ReadOnly Property Manufacturer As String
        Get
            Return Me.GetManufacturer()
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the notebook's operating system
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Public ReadOnly Property OperatingSystem As String
        Get
            Return Me.GetOperatingSystem
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the notebook's computer name
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Public ReadOnly Property ComputerName As String
        Get
            Return Environment.MachineName
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the current revision of DirectX installed
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Public ReadOnly Property DirectXRevision As String
        Get
            Return Me.GetDirectXRevision
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the product number of the notebook
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks>String</remarks>
    Public ReadOnly Property SystemSKUNumber As String
        Get
            Return Me.GetSystemSKUNumber
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the product number of the notebook
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PCID As String
        Get
            Return Me.GetPCID
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating the operating system architecture
    ''' </summary>
    ''' <value>OSArch</value>
    ''' <returns>OSArch</returns>
    ''' <remarks>N/A</remarks>
    Public ReadOnly Property OperatingSystemArchitecture As OSArch
        Get
            Return Me.GetOperatingSystemArchitecture
        End Get
    End Property
#End Region
#Region " Private Methods "
    ''' <summary>
    ''' Method used to determine the BIOS revision of the unit.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Private Function GetBiosRevision() As String
        Dim bios_revision = String.Empty

        Dim searcher As New ManagementObjectSearcher( _
                   "root\CIMV2", _
                   "SELECT * FROM Win32_BIOS")
        For Each queryObj As ManagementObject In searcher.Get()

            bios_revision = CStr(queryObj("SMBIOSBIOSVersion"))
        Next

        Return bios_revision
    End Function
    ''' <summary>
    ''' Method used to determine the serial number of the system.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSerialNumber() As String
        Dim system_serialnumber = String.Empty

        Dim searcher As New ManagementObjectSearcher( _
                   "root\CIMV2", _
                   "SELECT * FROM Win32_BIOS")
        For Each queryObj As ManagementObject In searcher.Get()
            system_serialnumber = CStr(queryObj("SerialNumber"))
        Next

        Return system_serialnumber
    End Function
    ''' <summary>
    ''' Gets the hardware manufacturer.
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Private Function GetManufacturer() As String
        Dim system_mfg = String.Empty

        Dim searcher As New ManagementObjectSearcher( _
                   "root\CIMV2", _
                   "SELECT * FROM Win32_BIOS")
        For Each queryObj As ManagementObject In searcher.Get()
            system_mfg = CStr(queryObj("Manufacturer"))
        Next

        Return system_mfg
    End Function
    ''' <summary>
    ''' Method used to determine which operating system the test is running in
    ''' </summary>
    ''' <returns>N/A</returns>
    ''' <remarks>N/A</remarks>
    Private Function GetOperatingSystem() As String
        Dim strVersion = "Unknown"
        Select Case Environment.OSVersion.Platform
            Case PlatformID.Win32S
                strVersion = "Windows 3.1"
            Case PlatformID.Win32Windows
                Select Case Environment.OSVersion.Version.Minor
                    Case 0I
                        strVersion = "Windows 95"
                    Case 10I
                        If Environment.OSVersion.Version.Revision.ToString() = "2222A" Then
                            strVersion = "Windows 98 Second Edition"
                        Else
                            strVersion = "Windows 98"
                        End If
                    Case 90I
                        strVersion = "Windows ME"
                End Select
            Case PlatformID.Win32NT
                Select Case Environment.OSVersion.Version.Major
                    Case 3I
                        strVersion = "Windows NT 3.51"
                    Case 4I
                        strVersion = "Windows NT 4.0"
                    Case 5I
                        Select Case Environment.OSVersion.Version.Minor
                            Case 0I
                                strVersion = "Windows 2000"
                            Case 1I
                                strVersion = "Windows XP"
                            Case 2I
                                strVersion = "Windows 2003"
                        End Select
                    Case 6I
                        Select Case Environment.OSVersion.Version.Minor
                            Case 0I
                                strVersion = "Windows Vista"
                            Case 1I
                                strVersion = "Windows 7"
                            Case 2I
                                strVersion = "Windows 2008"
                        End Select
                End Select
            Case PlatformID.WinCE
                strVersion = "Windows CE"
            Case PlatformID.Unix
                strVersion = "Unix"
        End Select
        Return strVersion
    End Function
    ''' <summary>
    ''' Method used to determine if the latest version of DirectX needs to be installed.
    ''' to work correctly.
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>N/A</remarks>
    Private Function GetDirectXRevision() As String
        Dim version = String.Empty
        Dim r As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\DirectX")
        If Not r Is Nothing Then
            If r.GetValue("Version") Is Nothing Then
                version = "N/A"
            Else
                version = CStr(r.GetValue("Version"))
            End If
        Else
            version = "N/A"
        End If
        r.Close()

        Return version
    End Function
    ''' <summary>
    ''' Method used to obtain the product number from the registry
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Private Function GetSystemSKUNumber() As String
        Dim r As RegistryKey = Registry.LocalMachine.OpenSubKey("HARDWARE\DESCRIPTION\System\BIOS")
        If r Is Nothing Then
            Return "N/A"
        Else
            If r.GetValue("SystemSKU") Is Nothing Then
                Return "N/A"
            ElseIf CStr(r.GetValue("SystemSKU")) = String.Empty Then
                Return "N/A"
            Else
                Dim pnum = CStr(r.GetValue("SystemSKU")).Trim
                Dim strings_to_remove() = {"#aba", "#abc", "#aa", "#abv", "#A>A", "#abl", "#"}

                For i As Integer = 0 To strings_to_remove.Length - 1
                    If pnum.ToLower.Contains(strings_to_remove(i)) Then
                        pnum = pnum.ToLower.Replace(strings_to_remove(i), String.Empty)
                    End If
                Next

                Return pnum
            End If
        End If
    End Function
    ''' <summary>
    ''' Method used to obtain the PCID from the registry
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Private Function GetPCID() As String
        Dim pcid = String.Empty
        Dim r As RegistryKey = Registry.LocalMachine.OpenSubKey("HARDWARE\DESCRIPTION\System\BIOS")
        If r Is Nothing Then
            pcid = "N/A"
        Else
            Select Case True
                Case r.GetValue("SystemVersion") Is Nothing, CStr(r.GetValue("SystemVersion")) = String.Empty
                    pcid = "N/A"
                Case Else
                    pcid = CStr(r.GetValue("SystemVersion"))
            End Select
        End If
        r.Close()

        Return pcid
    End Function
    ''' <summary>
    ''' Gets the OS service pack
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetServicePack()
        Dim os As New OSVERSIONINFO()
        os.dwOSVersionInfoSize = Marshal.SizeOf(GetType(OSVERSIONINFO))
        GetVersionEx(os)
        If os.szCSDVersion = "" Then
            Return String.Empty
        Else
            Return os.szCSDVersion
        End If
    End Function
    ''' <summary>
    ''' Gets OS architecture
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>N/A</remarks>
    Private Function GetOperatingSystemArchitecture() As OSArch
        Dim arch As String = String.Empty

        Dim searcher As New ManagementObjectSearcher( _
                "root\CIMV2", _
                "SELECT * FROM Win32_OperatingSystem")

        For Each queryObj As ManagementObject In searcher.Get()
            arch = CStr(queryObj("OSArchitecture"))
        Next

        If arch.ToLower.Contains("64") Then
            Return OSArch.x64
        Else
            Return OSArch.x86
        End If
    End Function
#End Region

#Region " Private Methods "
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure OSVERSIONINFO
        Public dwOSVersionInfoSize As Integer
        Public dwMajorVersion As Integer
        Public dwMinorVersion As Integer
        Public dwBuildNumber As Integer
        Public dwPlatformId As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> _
        Public szCSDVersion
    End Structure
    <DllImport("kernel32.Dll")> _
    Private Shared Function GetVersionEx(ByRef o As OSVERSIONINFO) As Short
    End Function
#End Region

End Class