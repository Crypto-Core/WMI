Imports System.Management
Imports Microsoft.VisualBasic.Devices
Public Class WMI
    Public Class Win32_Account
        ''' <summary>
        ''' Short description of the object. This property is inherited from the CIM_ManagedSystemElement class.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Caption() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Caption = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Description of the object. This property is inherited from the CIM_ManagedSystemElement class.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Description() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Description = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the Windows domain to which a group or user belongs.
        '''Example: "NA-SALES"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Domain() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Domain = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                Domain = objItem.Domain
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Date and time that the object was installed. This property does not require a value to indicate that the object is installed. This property is inherited from the CIM_ManagedSystemElement class.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function InstallDate() As DateTime
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InstallDate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                Dim str As String = objItem.InstallDate
                InstallDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If TRUE, the account is defined on the local machine. To retrieve only accounts defined on the local machine, design a query that includes the condition "LocalAccount=TRUE".
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LocalAccount() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LocalAccount = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                LocalAccount = objItem.LocalAccount
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the Windows system account on the domain specified by the Domain property of this class. This property overrides the Name property inherited from CIM_ManagedSystemElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Name() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Name = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Security identifier (SID) for this account. A SID is a string value of variable length used to identify a trustee. Each account has a unique SID issued by an authority (such as a Windows domain), stored in a security database. When a user logs on, the system retrieves the user's SID from the database and places it in the user's access token. The system uses the SID in the user's access token to identify the user in all subsequent interactions with Windows security. When a SID has been used as the unique identifier for a user or group, it cannot be used again to identify another user or group.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                SID = objItem.SID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Enumerated values that specify the type of security identifier (SID).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SIDType() As SByte
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SIDType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                SIDType = objItem.SIDType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Current status of the object. Various operational and nonoperational statuses can be defined. Operational statuses include: "OK", "Degraded", and "Pred Fail" (an element, such as a SMART-enabled hard disk drive, may be functioning properly but predicts a failure in the near future). Nonoperational statuses include: "Error", "Starting", "Stopping", and "Service". The latter, "Service", can apply during mirror-resilvering of a disk, reload of a user permissions list, or other administrative work. Not all such work is online, yet the managed element is neither "OK" nor in one of the other states.
        '''This property is inherited from the CIM_ManagedSystemElement class.
        '''The values are:
        '''"OK"
        '''"Error"
        '''"Degraded"
        '''"Unknown"
        '''"Pred Fail"
        '''"Starting"
        '''"Stopping"
        '''"Service"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Status() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Status = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account")
            For Each objItem In objItems
                Status = objItem.Status
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_OperatingSystem
        ''' <summary>
        ''' Short description of the object—a one-line string. The string includes the operating system version. For example, "Microsoft Windows 7 Enterprise ". This property can be localized.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Caption() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Caption = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the disk drive from which the Windows operating system starts.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BootDevice() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BootDevice = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                BootDevice = objItem.BootDevice
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Build number of an operating system. It can be used for more precise version information than product release version numbers.
        '''Example: "1381"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BuildNumber() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BuildNumber = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                BuildNumber = objItem.BuildNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Type of build used for an operating system.
        '''Examples: ""retail build"", ""checked build""
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BuildType() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BuildType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                BuildType = objItem.BuildType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Code page value an operating system uses. A code page contains a character table that an operating system uses to translate strings for different languages. The American National Standards Institute (ANSI) lists values that represent defined code pages. If an operating system does not use an ANSI code page, this member is set to 0 (zero). The CodeSet string can use a maximum of six characters to define the code page value.
        '''Example: "1255"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CodeSet() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CodeSet = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                CodeSet = objItem.CodeSet
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Code for the country/region that an operating system uses. Values are based on international phone dialing prefixes—also referred to as IBM country/region codes. This property can use a maximum of six characters to define the country/region code value.
        '''Example: "1" (United States)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CountryCode() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CountryCode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                CountryCode = objItem.CountryCode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the first concrete class that appears in the inheritance chain used in the creation of an instance. When used with other key properties of the class, this property allows all instances of this class and its subclasses to be identified uniquely.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreationClassName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CreationClassName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                CreationClassName = objItem.CreationClassName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Creation class name of the scoping computer system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CSCreationClassName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CSCreationClassName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                CSCreationClassName = objItem.CSCreationClassName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' NULL-terminated string that indicates the latest service pack installed on a computer. If no service pack is installed, the string is NULL.
        '''Example: "Service Pack 3"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CSDVersion() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CSDVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                CSDVersion = objItem.CSDVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the scoping computer system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CSName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CSName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                CSName = objItem.CSName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number, in minutes, an operating system is offset from Greenwich mean time (GMT). The number is positive, negative, or zero.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentTimeZone() As Int16

            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentTimeZone = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                CurrentTimeZone = objItem.CurrentTimeZone
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Data execution prevention is a hardware feature to prevent buffer overrun attacks by stopping the execution of code on data-type memory pages. If True, then this feature is available. On 64-bit computers, the data execution prevention feature is configured in the BCD store and the properties in Win32_OperatingSystem are set accordingly.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DataExecutionPrevention_Available() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DataExecutionPrevention_Available = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                DataExecutionPrevention_Available = objItem.DataExecutionPrevention_Available
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' When the data execution prevention hardware feature is available, this property indicates that the feature is set to work for 32-bit applications if True. On 64-bit computers, the data execution prevention feature is configured in the Boot Configuration Data (BCD) store and the properties in Win32_OperatingSystem are set accordingly.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DataExecutionPrevention_32BitApplications() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DataExecutionPrevention_32BitApplications = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                DataExecutionPrevention_32BitApplications = objItem.DataExecutionPrevention_32BitApplications
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' When the data execution prevention hardware feature is available, this property indicates that the feature is set to work for drivers if True. On 64-bit computers, the data execution prevention feature is configured in the BCD store and the properties in Win32_OperatingSystem are set accordingly.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DataExecutionPrevention_Drivers() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DataExecutionPrevention_Drivers = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                DataExecutionPrevention_Drivers = objItem.DataExecutionPrevention_Drivers
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Indicates which Data Execution Prevention (DEP) setting is applied. The DEP setting specifies the extent to which DEP applies to 32-bit applications on the system. DEP is always applied to the Windows kernel.
        ''' 0 = Always Off(DEP is turned off for all 32-bit applications on the computer with no exceptions. This setting is not available for the user interface.)
        ''' 1 = Always On(DEP is enabled for all 32-bit applications on the computer. This setting is not available for the user interface.)
        ''' 2 = Opt In(DEP is enabled for a limited number of binaries, the kernel, and all Windows-based services. However, it is off by default for all 32-bit applications. A user or administrator must explicitly choose either the AlwaysOn or the OptOut setting before DEP can be applied to 32-bit applications.)
        ''' 3 = Opt Out(DEP is enabled by default for all 32-bit applications. A user or administrator can explicitly remove support for a 32-bit application by adding the application to an exceptions list.)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DataExecutionPrevention_SupportPolicy() As SByte
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DataExecutionPrevention_SupportPolicy = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                DataExecutionPrevention_SupportPolicy = objItem.DataExecutionPrevention_SupportPolicy
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Operating system is a checked (debug) build. If True, the debugging version is installed. Checked builds provide error checking, argument verification, and system debugging code. Additional code in a checked binary generates a kernel debugger error message and breaks into the debugger. This helps immediately determine the cause and location of the error. Performance may be affected in a checked build due to the additional code that is executed.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Debug() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Debug = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Debug = objItem.Debug
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Description of the Windows operating system. Some user interfaces for example, those that allow editing of this description, limit its length to 48 characters.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Description() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Description = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the operating system is distributed across several computer system nodes. If so, these nodes should be grouped as a cluster.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Distributed() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Distributed = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Distributed = objItem.Distributed
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Encryption level for secure transactions: 40-bit, 128-bit, or n-bit.
        '''40-bit (0)
        '''128-bit (1)
        '''n-bit (2)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EncryptionLevel() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            EncryptionLevel = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                EncryptionLevel = objItem.EncryptionLevel
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Increase in priority is given to the foreground application. Application boost is implemented by giving an application more execution time slices (quantum lengths).
        ''' 0 = None (The system boosts the quantum length by 6.)
        ''' 1 = Minimum (The system boosts the quantum length by 12.)
        ''' 2 = Maximum (The system boosts the quantum length by 18.)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ForegroundApplicationBoost() As SByte
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ForegroundApplicationBoost = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                ForegroundApplicationBoost = objItem.ForegroundApplicationBoost
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number, in kilobytes, of physical memory currently unused and available.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FreePhysicalMemory() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            FreePhysicalMemory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                FreePhysicalMemory = objItem.FreePhysicalMemory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number, in kilobytes, that can be mapped into the operating system paging files without causing any other pages to be swapped out.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FreeSpaceInPagingFiles() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            FreeSpaceInPagingFiles = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                FreeSpaceInPagingFiles = objItem.FreeSpaceInPagingFiles
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number, in kilobytes, of virtual memory currently unused and available.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FreeVirtualMemory() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            FreeVirtualMemory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                FreeVirtualMemory = objItem.FreeVirtualMemory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Date object was installed. This property does not require a value to indicate that the object is installed.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function InstallDate() As DateTime

            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InstallDate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Dim str As String = objItem.InstallDate
                InstallDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' his property is obsolete and not supported.
        ''' 0 = Optimize for Applications
        ''' 1 = Optimize for System Performance
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LargeSystemCache() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LargeSystemCache = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                LargeSystemCache = objItem.LargeSystemCache
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Date and time the operating system was last restarted.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LastBootUpTime() As DateTime
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LastBootUpTime = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Dim str As String = objItem.LastBootUpTime
                LastBootUpTime = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Operating system version of the local date and time-of-day.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LocalDateTime() As DateTime
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LocalDateTime = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Dim str As String = objItem.LocalDateTime
                LocalDateTime = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Language identifier used by the operating system. A language identifier is a standard international numeric abbreviation for a country/region. Each language has a unique language identifier (LANGID), a 16-bit value that consists of a primary language identifier and a secondary language identifier.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Locale() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Locale = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Locale = objItem.Locale
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the operating system manufacturer. For Windows-based systems, this value is "Microsoft Corporation".
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Manufacturer() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Manufacturer = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Manufacturer = objItem.Manufacturer
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Maximum number of process contexts the operating system can support. The default value set by the provider is 4294967295 (0xFFFFFFFF). If there is no fixed maximum, the value should be 0 (zero). On systems that have a fixed maximum, this object can help diagnose failures that occur when the maximum is reached—if unknown, enter 4294967295 (0xFFFFFFFF).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxNumberOfProcesses() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MaxNumberOfProcesses = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                MaxNumberOfProcesses = objItem.MaxNumberOfProcesses
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Maximum number, in kilobytes, of memory that can be allocated to a process. For operating systems with no virtual memory, typically this value is equal to the total amount of physical memory minus the memory used by the BIOS and the operating system. For some operating systems, this value may be infinity, in which case 0 (zero) should be entered. In other cases, this value could be a constant, for example, 2G or 4G.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxProcessMemorySize() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MaxProcessMemorySize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                MaxProcessMemorySize = objItem.MaxProcessMemorySize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Multilingual User Interface Pack (MUI Pack ) languages installed on the computer. For example, "en-us". MUI Pack languages are resource files that can be installed on the English version of the operating system. When an MUI Pack is installed, you can can change the user interface language to one of 33 supported languages.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MUILanguages() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MUILanguages = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                MUILanguages = objItem.MUILanguages
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Operating system instance within a computer system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Name() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Name = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of user licenses for the operating system. If unlimited, enter 0 (zero). If unknown, enter -1.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumberOfLicensedUsers() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfLicensedUsers = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                NumberOfLicensedUsers = objItem.NumberOfLicensedUsers
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of process contexts currently loaded or running on the operating system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumberOfProcesses() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfProcesses = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                NumberOfProcesses = objItem.NumberOfProcesses
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of user sessions for which the operating system is storing state information currently.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumberOfUsers() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfUsers = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                NumberOfUsers = objItem.NumberOfUsers
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function OperatingSystemSKU() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OperatingSystemSKU = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                OperatingSystemSKU = objItem.OperatingSystemSKU
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Company name for the registered user of the operating system.
        ''' Example: "Microsoft Corporation"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Organization() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Organization = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Organization = objItem.Organization
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Architecture of the operating system, as opposed to the processor. This property can be localized.
        ''' Example: 32-bit
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OSArchitecture() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OSArchitecture = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                OSArchitecture = objItem.OSArchitecture
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Language version of the operating system installed. The following table lists the possible values. Example: 0x0807 (German, Switzerland).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OSLanguage() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OSLanguage = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                OSLanguage = objItem.OSLanguage
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Installed and licensed system product additions to the operating system. For example, the value of 146 (0x92) for OSProductSuite indicates Enterprise, Terminal Services, and Data Center (bits one, four, and seven set). The following table lists possible values.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OSProductSuite() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OSProductSuite = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                OSProductSuite = objItem.OSProductSuite
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Type of operating system. The following list identifies the possible values.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OSType() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OSType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                OSType = objItem.OSType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Additional description for the current operating system version.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OtherTypeDescription() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OtherTypeDescription = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                OtherTypeDescription = objItem.OtherTypeDescription
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the physical address extensions (PAE) are enabled by the operating system running on Intel processors. PAE allows applications to address more than 4 GB of physical memory. When PAE is enabled, the operating system uses three-level linear address translation rather than two-level. Providing more physical memory to an application reduces the need to swap memory to the page file and increases performance. To enable, PAE, use the "/PAE" switch in the Boot.ini file. For more information about the Physical Address Extension feature, see Http://Go.Microsoft.Com/FWLink/p/?LinkID=45912.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PAEEnabled() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PAEEnabled = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                PAEEnabled = objItem.PAEEnabled
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PlusProductID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PlusProductID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                PlusProductID = objItem.PlusProductID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PlusVersionNumber() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PlusVersionNumber = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                PlusVersionNumber = objItem.PlusVersionNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Specifies whether the operating system booted from an external USB device. If true, the operating system has detected it is booting on a supported locally connected storage device.
        ''' Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows 8 and Windows Server 2012.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PortableOperatingSystem() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PortableOperatingSystem = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                PortableOperatingSystem = objItem.PortableOperatingSystem
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Specifies whether this is the primary operating system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Primary() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Primary = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Primary = objItem.Primary
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Additional system information.
        ''' Work Station (1)
        ''' Domain Controller (2)
        ''' Server (3)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ProductType() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ProductType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                ProductType = objItem.ProductType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the registered user of the operating system.
        ''' Example: "Ben Smith"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function RegisteredUser() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            RegisteredUser = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                RegisteredUser = objItem.RegisteredUser
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Operating system product serial identification number.
        ''' Example: "10497-OEM-0031416-71674"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SerialNumber() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SerialNumber = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                SerialNumber = objItem.SerialNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Major version number of the service pack installed on the computer system. If no service pack has been installed, the value is 0 (zero).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ServicePackMajorVersion() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ServicePackMajorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                ServicePackMajorVersion = objItem.ServicePackMajorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Minor version number of the service pack installed on the computer system. If no service pack has been installed, the value is 0 (zero).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ServicePackMinorVersion() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ServicePackMinorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                ServicePackMinorVersion = objItem.ServicePackMinorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Total number of kilobytes that can be stored in the operating system paging files—0 (zero) indicates that there are no paging files. Be aware that this number does not represent the actual physical size of the paging file on disk.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SizeStoredInPagingFiles() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SizeStoredInPagingFiles = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                SizeStoredInPagingFiles = objItem.SizeStoredInPagingFiles
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Current status of the object. Various operational and nonoperational statuses can be defined. Operational statuses include: "OK", "Degraded", and "Pred Fail" (an element, such as a SMART-enabled hard disk drive may function properly, but predicts a failure in the near future). Nonoperational statuses include: "Error", "Starting", "Stopping", and "Service". The Service status applies to administrative work, such as mirror-resilvering of a disk, reload of a user permissions list, or other administrative work. Not all such work is online, but the managed element is neither "OK" nor in one of the other states.
        '''"OK"
        '''"Error"
        '''"Degraded"
        '''"Unknown"
        '''"Pred Fail"
        '''"Starting"
        '''"Stopping"
        '''"Service"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Status() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Status = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Status = objItem.Status
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Qualifiers: BitMap ("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10") , BitValues ("Windows Server 2003, Small Business Edition", "Windows Server 2003, Enterprise Edition", "Windows Server 2003, Backoffice Edition", "Windows Server 2003, Communications Edition", "Microsoft Terminal Services", "Windows Server 2003, Small Business Edition Restricted", "Windows XP Embedded", "Windows Server 2003, Datacenter Edition", "Single User", "Windows XP Home Edition", "Windows Server 2003, Web Edition")
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuiteMask() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SuiteMask = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                SuiteMask = objItem.SuiteMask
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Physical disk partition on which the operating system is installed.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemDevice() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemDevice = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                SystemDevice = objItem.SystemDevice
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' System directory of the operating system.
        '''Example: "C:\WINDOWS\SYSTEM32"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemDirectory() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemDirectory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                SystemDirectory = objItem.SystemDirectory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Letter of the disk drive on which the operating system resides. Example: "C:"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemDrive() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemDrive = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                SystemDrive = objItem.SystemDrive
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Total swap space in kilobytes. This value may be NULL (unspecified) if the swap space is not distinguished from page files. However, some operating systems distinguish these concepts. For example, in UNIX, whole processes can be swapped out when the free page list falls and remains below a specified amount.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TotalSwapSpaceSize() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            TotalSwapSpaceSize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                TotalSwapSpaceSize = objItem.TotalSwapSpaceSize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number, in kilobytes, of virtual memory. For example, this may be calculated by adding the amount of total RAM to the amount of paging space, that is, adding the amount of memory in or aggregated by the computer system to the property, SizeStoredInPagingFiles.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TotalVirtualMemorySize() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            TotalVirtualMemorySize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                TotalVirtualMemorySize = objItem.TotalVirtualMemorySize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Total amount, in kilobytes, of physical memory available to the operating system. This value does not necessarily indicate the true amount of physical memory, but what is reported to the operating system as available to it.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TotalVisibleMemorySize() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            TotalVisibleMemorySize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                TotalVisibleMemorySize = objItem.TotalVisibleMemorySize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Version number of the operating system.
        '''Example: "4.0"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Version() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Version = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                Version = objItem.Version
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Windows directory of the operating system.
        '''Example: "C:\WINDOWS"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function WindowsDirectory() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            WindowsDirectory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                WindowsDirectory = objItem.WindowsDirectory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The QuantumLength property defines the number of clock ticks per quantum. A quantum is a unit of execution time that the scheduler is allowed to give to an application before switching to other applications. When a thread runs one quantum, the kernel preempts it and moves it to the end of a queue for applications with equal priorities. The actual length of a thread's quantum varies across different Windows platforms. For Windows NT/Windows 2000 only.
        '''Unknown (0)
        '''One tick (1)
        '''Two ticks (2)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function QuantumLength() As SByte
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            QuantumLength = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                QuantumLength = objItem.QuantumLength
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The QuantumType property specifies either fixed or variable length quantums. Windows defaults to variable length quantums where the foreground application has a longer quantum than the background applications. Windows Server defaults to fixed-length quantums. A quantum is a unit of execution time that the scheduler is allowed to give to an application before switching to another application. When a thread runs one quantum, the kernel preempts it and moves it to the end of a queue for applications with equal priorities. The actual length of a thread's quantum varies across different Windows platforms.
        '''The property can take the following values:
        '''0 = Unkown - Quantum Type not known.
        '''1 = Fixed - Quantum length is fixed.
        '''2 = Variable - Quantum length is variable.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function QuantumType() As SByte
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            QuantumType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
            For Each objItem In objItems
                QuantumType = objItem.QuantumType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_BootConfiguration
        ''' <summary>
        ''' Path to the system files required for booting the system.
        '''Example: "C:\Windows"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BootDirectory() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BootDirectory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                BootDirectory = objItem.BootDirectory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Short description of the CIM_Setting object. This property is inherited from CIM_Setting.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Caption() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Caption = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Path to the configuration files. This value may be similar to the value in the BootDirectory property.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConfigurationPath() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ConfigurationPath = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                ConfigurationPath = objItem.ConfigurationPath
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Description of the CIM_Setting object. This property is inherited from CIM_Setting.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Description() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Description = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Last drive letter to which a physical drive is assigned.
        '''Example: "E:"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LastDrive() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LastDrive = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                LastDrive = objItem.LastDrive
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the boot configuration. It is an identifier for the boot configuration.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Name() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Name = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Directory where temporary files can reside during boot time.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ScratchDirectory() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ScratchDirectory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                ScratchDirectory = objItem.ScratchDirectory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Identifier by which the CIM_Setting object is known. This property is inherited from CIM_Setting.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SettingID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SettingID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                SettingID = objItem.SettingID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Directory where temporary files are stored.
        '''Example: "C:\TEMP"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TempDirectory() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            TempDirectory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BootConfiguration")
            For Each objItem In objItems
                TempDirectory = objItem.TempDirectory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_ProcessStartup
        ''' <summary>
        ''' Qualifiers: MappingStrings ("Win32API|Process and Thread Functions|CreateProcess|dwCreationFlags") , BitMap ("0", "1", "2", "3", "4", "9", "10", "26") , BitValues ("Debug_Process", "Debug_Only_This_Process", "Create_Suspended", "Detached_Process", "Create_New_Console", "Create_New_Process_Group", "Create_Unicode_Environment", "Create_Default_Error_Mode")
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreateFlags() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CreateFlags = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                CreateFlags = objItem.CreateFlags
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' List of settings for the configuration of a computer. Environment variables specify search paths for files, directories for temporary files, application-specific options, and other similar information. The system maintains a block of environment settings for each user and one for the computer. The system environment block represents environment variables for all of the users of a specific computer. A user's environment block represents the environment variables that the system maintains for a specific user, and includes the set of system environment variables. By default, each process receives a copy of the environment block for its parent process. Typically, this is the environment block for the user who is logged on. A process can specify different environment blocks for its child processes.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EnvironmentVariables() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            EnvironmentVariables = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                EnvironmentVariables = objItem.EnvironmentVariables
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ErrorMode() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ErrorMode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                ErrorMode = objItem.ErrorMode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function FillAttribute() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            FillAttribute = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                FillAttribute = objItem.FillAttribute
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PriorityClass() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PriorityClass = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                PriorityClass = objItem.PriorityClass
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ShowWindow() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ShowWindow = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                ShowWindow = objItem.ShowWindow
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function Title() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Title = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                Title = objItem.Title
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function WinstationDesktop() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            WinstationDesktop = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                WinstationDesktop = objItem.WinstationDesktop
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function X() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            X = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                X = objItem.X
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function XCountChars() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            XCountChars = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                XCountChars = objItem.XCountChars
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function XSize() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            XSize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                XSize = objItem.XSize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function Y() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Y = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                Y = objItem.Y
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function YCountChars() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            YCountChars = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                YCountChars = objItem.YCountChars
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function YSize() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            YSize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ProcessStartup")
            For Each objItem In objItems
                YSize = objItem.YSize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_VideoConfiguration
        ''' <summary>
        ''' The ActualColorResolution property indicates the current color depth of the video display.
        '''This property has been deprecated in favor of a corresponding property(s) contained in the Win32_VideoController, Win32_DesktopMonitor and//or CIM_VideoControllerResolution
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SettingID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SettingID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                SettingID = objItem.SettingID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        
        Public Shared Function Caption() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Caption = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function Description() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Description = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ActualColorResolution() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ActualColorResolution = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                ActualColorResolution = objItem.ActualColorResolution
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The AdapterChipType property contains the name of the adapter chip.
        '''Example: s3
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AdapterChipType() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdapterChipType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                AdapterChipType = objItem.AdapterChipType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function AdapterCompatibility() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdapterCompatibility = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                AdapterCompatibility = objItem.AdapterCompatibility
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function AdapterDACType() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdapterDACType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                AdapterDACType = objItem.AdapterDACType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function AdapterDescription() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdapterDescription = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                AdapterDescription = objItem.AdapterDescription
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function AdapterRAM() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdapterRAM = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                AdapterRAM = objItem.AdapterRAM
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function AdapterType() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdapterType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                AdapterType = objItem.AdapterType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function BitsPerPixel() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BitsPerPixel = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                BitsPerPixel = objItem.BitsPerPixel
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ColorPlanes() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ColorPlanes = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                ColorPlanes = objItem.ColorPlanes
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ColorTableEntries() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ColorTableEntries = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                ColorTableEntries = objItem.ColorTableEntries
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function DeviceSpecificPens() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DeviceSpecificPens = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                DeviceSpecificPens = objItem.DeviceSpecificPens
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function DriverDate() As DateTime
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DriverDate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                Dim str As String = objItem.DriverDate
                DriverDate = New DateTime(str.Substring(0, 4), str.Substring(4, 6), str.Substring(6, 8), str.Substring(8, 10), str.Substring(10, 12), str.Substring(12, 14))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function HorizontalResolution() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            HorizontalResolution = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                HorizontalResolution = objItem.HorizontalResolution
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function InfFilename() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InfFilename = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                InfFilename = objItem.InfFilename
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function InstalledDisplayDrivers() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InstalledDisplayDrivers = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                InstalledDisplayDrivers = objItem.InstalledDisplayDrivers
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function MonitorManufacturer() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MonitorManufacturer = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                MonitorManufacturer = objItem.MonitorManufacturer
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function MonitorType() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MonitorType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                MonitorType = objItem.MonitorType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function Name() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Name = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PixelsPerXLogicalInch() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PixelsPerXLogicalInch = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                PixelsPerXLogicalInch = objItem.PixelsPerXLogicalInch
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PixelsPerYLogicalInch() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PixelsPerYLogicalInch = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                PixelsPerYLogicalInch = objItem.PixelsPerYLogicalInch
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function RefreshRate() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            RefreshRate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                RefreshRate = objItem.RefreshRate
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ScanMode() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ScanMode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                ScanMode = objItem.ScanMode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ScreenHeight() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ScreenHeight = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                ScreenHeight = objItem.ScreenHeight
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ScreenWidth() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ScreenWidth = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                ScreenWidth = objItem.ScreenWidth
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function SystemPaletteEntries() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemPaletteEntries = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                SystemPaletteEntries = objItem.SystemPaletteEntries
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function VerticalResolution() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VerticalResolution = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoConfiguration")
            For Each objItem In objItems
                VerticalResolution = objItem.VerticalResolution
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_ComputerSystem
        ''' <summary>
        ''' System hardware security settings for administrator password status.
        '''Disabled (0)
        '''Enabled (1)
        '''Not Implemented (2)
        '''Unknown (3)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AdminPasswordStatus() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdminPasswordStatus = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                AdminPasswordStatus = objItem.AdminPasswordStatus
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the system manages the page file.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AutomaticManagedPagefile() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AutomaticManagedPagefile = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                AutomaticManagedPagefile = objItem.AutomaticManagedPagefile
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the automatic reset boot option is enabled.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AutomaticResetBootOption() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AutomaticResetBootOption = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                AutomaticResetBootOption = objItem.AutomaticResetBootOption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the automatic reset is enabled.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AutomaticResetCapability() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AutomaticResetCapability = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                AutomaticResetCapability = objItem.AutomaticResetCapability
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Boot option limit is ON. Identifies the system action when the ResetLimit value is reached.
        '''Reserved (0)
        '''Operating system (1)
        '''System utilities (2)
        '''Do not reboot (3)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BootOptionOnLimit() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BootOptionOnLimit = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                BootOptionOnLimit = objItem.BootOptionOnLimit
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Type of reboot action after the time on the watchdog timer is elapsed.
        '''Reserved (0)
        '''Operating system (1)
        '''System utilities (2)
        '''Do not reboot (3)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BootOptionOnWatchDog() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BootOptionOnWatchDog = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                BootOptionOnWatchDog = objItem.BootOptionOnWatchDog
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, indicates whether a boot ROM is supported.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BootROMSupported() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BootROMSupported = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                BootROMSupported = objItem.BootROMSupported
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Status and Additional Data fields that identify the boot status.
        '''This value comes from the Boot Status member of the System Boot Information structure in the SMBIOS information.
        '''Windows Server 2012 R2, Windows 8.1, Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows 10 and Windows Server 2016 Technical Preview.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BootupState() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BootupState = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                BootupState = objItem.BootupState
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' System is started. Fail-safe boot bypasses the user startup files—also called SafeBoot.
        '''The following list contains the required values:
        '''"Normal boot"
        '''"Fail-safe boot"
        '''"Fail-safe with network boot"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BootStatus() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BootStatus = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                BootStatus = objItem.BootStatus
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Short description of the object—a one-line string. This property is inherited from CIM_ManagedSystemElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Caption() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Caption = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Boot up state of the chassis.
        '''This value comes from the Boot-up State member of the System Enclosure or Chassis structure in the SMBIOS information.
        '''Other (1)
        '''Unknown (2)
        '''Safe (3)
        '''Warning (4)
        '''Critical (5)
        '''Non-recoverable (6)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ChassisBootupState() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ChassisBootupState = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                ChassisBootupState = objItem.ChassisBootupState
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The chassis or enclosure SKU number as a string.
        '''This value comes from the SKU Number member of the System Enclosure or Chassis structure in the SMBIOS information.
        '''Windows Server 2012 R2, Windows 8.1, Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows 10 and Windows Server 2016 Technical Preview.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ChassisSKUNumber() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ChassisSKUNumber = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                ChassisSKUNumber = objItem.ChassisSKUNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the first concrete class in the inheritance chain of an instance. You can use this property with other properties of the class to identify all instances of the class and its subclasses. This property is inherited from CIM_System.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreationClassName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CreationClassName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                CreationClassName = objItem.CreationClassName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Amount of time the unitary computer system is offset from Coordinated Universal Time (UTC).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentTimeZone() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentTimeZone = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                CurrentTimeZone = objItem.CurrentTimeZone
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the daylight savings mode is ON.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DaylightInEffect() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DaylightInEffect = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                DaylightInEffect = objItem.DaylightInEffect
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Description of the object.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Description() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Description = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of local computer according to the domain name server (DNS).
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DNSHostName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DNSHostName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                DNSHostName = objItem.DNSHostName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the domain to which a computer belongs.
        '''Note  If the computer is not part of a domain, then the name of the workgroup is returned.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Domain() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Domain = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Domain = objItem.Domain
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        '''Role of a computer in an assigned domain workgroup. A domain workgroup is a collection of computers on the same network. For example, a DomainRole property may show that a computer is a member workstation. This property is inherited from CIM_ManagedSystemElement.
        '''Standalone Workstation (0)
        '''Member Workstation (1)
        '''Standalone Server (2)
        '''Member Server (3)
        '''Backup Domain Controller (4)
        '''Primary Domain Controller (5)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DomainRole() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DomainRole = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                DomainRole = objItem.DomainRole
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Enables daylight savings time (DST) on a computer. A value of True indicates that the system time changes to an hour ahead or behind when DST starts or ends. A value of False indicates that the system time does not change to an hour ahead or behind when DST starts or ends. A value of NULL indicates that the DST status is unknown on a system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EnableDaylightSavingsTime() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            EnableDaylightSavingsTime = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                EnableDaylightSavingsTime = objItem.EnableDaylightSavingsTime
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The following table lists the hardware security settings for the reset button on a computer.
        '''Disabled (0)
        '''Enabled (1)
        '''Not Implemented (2)
        '''Unknown (3)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FrontPanelResetStatus() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            FrontPanelResetStatus = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                FrontPanelResetStatus = objItem.FrontPanelResetStatus
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, a hypervisor is present.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function HypervisorPresent() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            HypervisorPresent = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                HypervisorPresent = objItem.HypervisorPresent
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, an infrared (IR) port exists on a computer system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function InfraredSupported() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InfraredSupported = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                InfraredSupported = objItem.InfraredSupported
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Data required to find the initial load device or boot service to request that the operating system start up. This property is inherited from CIM_UnitaryComputerSystem.
        '''Windows Server 2008 R2:  This property is available, but empty.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function InitialLoadInfo() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InitialLoadInfo = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                InitialLoadInfo = objItem.InitialLoadInfo
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Object is installed. An object does not need a value to indicate that it is installed. This property is inherited from CIM_ManagedSystemElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function InstallDate() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InstallDate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Dim str As String = objItem.InstallDate
                InstallDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' System hardware security settings for Keyboard Password Status.
        '''Disabled (0)
        '''Enabled (1)
        '''Not Implemented (2)
        '''Unknown (3)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function KeyboardPasswordStatus() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            KeyboardPasswordStatus = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                KeyboardPasswordStatus = objItem.KeyboardPasswordStatus
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Array entry of the InitialLoadInfo property that contains the data to start the loaded operating system. This property is inherited from CIM_UnitaryComputerSystem.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LastLoadInfo() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LastLoadInfo = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                LastLoadInfo = objItem.LastLoadInfo
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of a computer manufacturer.
        '''Example: Adventure Works
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Manufacturer() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Manufacturer = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Manufacturer = objItem.Manufacturer
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Product name that a manufacturer gives to a computer. This property must have a value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Model() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Model = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Model = objItem.Model
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Key of a CIM_System instance in an enterprise environment. This property is inherited from CIM_ManagedSystemElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Name() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Name = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Computer system Name value that is generated automatically. The CIM_ComputerSystem object and its derivatives are top-level objects of the Common Information Model (CIM). They provide the scope for several components. Unique CIM_System keys are required, but you can define a heuristic to create the CIM_ComputerSystem name that generates the same name, and is independent from the discovery protocol. This prevents inventory and management problems when the same asset or entity is discovered multiple times, but cannot be resolved to one object. Using a heuristic is recommended, but not required.
        '''The heuristic is outlined in the CIM V2 Common Model specification, and assumes that the documented rules are used to determine and assign a name. The NameFormat values list defines the order to assign a computer system name. Several rules map to the same value.
        '''The CIM_ComputerSystem Name value that is calculated using the heuristic is the key value of the system. However, use aliases to assign a different name for CIM_ComputerSystem, which can be more unique to your company. This property is inherited from CIM_System.
        '''The following list identifies the values for this property.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NameFormat() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NameFormat = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                NameFormat = objItem.NameFormat
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the network Server Mode is enabled.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NetworkServerModeEnabled() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NetworkServerModeEnabled = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                NetworkServerModeEnabled = objItem.NetworkServerModeEnabled
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of logical processors available on the computer.
        '''You can use NumberOfLogicalProcessors and NumberOfProcessors to determine if the computer is hyperthreading. For more information, see Remarks.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumberOfLogicalProcessors() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfLogicalProcessors = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                NumberOfLogicalProcessors = objItem.NumberOfLogicalProcessors
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of physical processors currently available on a system. This is the number of enabled processors for a system, which does not include the disabled processors. If a computer system has two physical processors each containing two logical processors, then the value of NumberOfProcessors is 2 and NumberOfLogicalProcessors is 4. The processors may be multicore or they may be hyperthreading processors. For more information, see Remarks.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumberOfProcessors() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfProcessors = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                NumberOfProcessors = objItem.NumberOfProcessors
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' List of data for a bitmap that the original equipment manufacturer (OEM) creates.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OEMLogoBitmap() As SByte()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OEMLogoBitmap = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                OEMLogoBitmap = objItem.OEMLogoBitmap
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' List of free-form strings that an OEM defines. For example, an OEM defines the part numbers for system reference documents, manufacturer contact information, and so on.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OEMStringArray() As String()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OEMStringArray = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                OEMStringArray = objItem.OEMStringArray
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the computer is part of a domain. If the value is NULL, the computer is not in a domain or the status is unknown. If you remove the computer from a domain, the value becomes false.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PartOfDomain() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PartOfDomain = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PartOfDomain = objItem.PartOfDomain
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Time delay before a reboot is initiated—in milliseconds. It is used after a system power cycle, local or remote system reset, and automatic system reset. A value of –1 (minus one) indicates that the pause value is unknown.
        '''Windows Vista:  This property may return an unknown number.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PauseAfterReset() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PauseAfterReset = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PauseAfterReset = objItem.PauseAfterReset
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Type of the computer in use, such as laptop, desktop, or Tablet.
        ''' 0 = Unspecified
        ''' 1 = Desktop
        ''' 2 = Mobile
        ''' 3 = Workstation
        ''' 4 = Enterprise Server
        ''' 5 = SOHO Server
        ''' 6 = Appliance PC
        ''' 7 = Performance Server
        ''' 8 = Maximum
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PCSystemType() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PCSystemType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PCSystemType = objItem.PCSystemType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Type of the computer in use, such as laptop, desktop, or Tablet.
        '''Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows 8.1 and Windows Server 2012 R2.
        '''Unspecified (0)
        '''Desktop (1)
        '''Mobile (2)
        '''Workstation (3)
        '''Enterprise Server (4)
        '''SOHO Server (5)
        '''Appliance PC (6)
        '''Performance Server (7)
        '''Slate (8)
        '''Maximum (9)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PCSystemTypeEx() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PCSystemTypeEx = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PCSystemTypeEx = objItem.PCSystemTypeEx
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PowerManagementCapabilities() As UInt16()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerManagementCapabilities = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PowerManagementCapabilities = objItem.PowerManagementCapabilities
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, device can be power-managed, for example, a device can be put into suspend mode, and so on. This property does not indicate that power management features are enabled currently, but it does indicate that the logical device is capable of power management. This property is inherited from CIM_UnitaryComputerSystem.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PowerManagementSupported() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerManagementSupported = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PowerManagementSupported = objItem.PowerManagementSupported
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' System hardware security settings for Power-On Password Status.
        '''Disabled (0)
        '''Enabled (1)
        '''Not Implemented (2)
        '''Unknown (3)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PowerOnPasswordStatus() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerOnPasswordStatus = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PowerOnPasswordStatus = objItem.PowerOnPasswordStatus
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Current power state of a computer and its associated operating system. The power saving states have the following values: Value 4 (Unknown) indicates that the system is known to be in a power save mode, but its exact status in this mode is unknown; 2 (Low Power Mode) indicates that the system is in a power save state, but still functioning and may exhibit degraded performance; 3 (Standby) indicates that the system is not functioning, but could be brought to full power quickly; and 7 (Warning) indicates that the computer system is in a warning state and a power save mode. This property is inherited from CIM_UnitaryComputerSystem.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PowerState() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerState = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PowerState = objItem.PowerState
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PowerSupplyState() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerSupplyState = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PowerSupplyState = objItem.PowerSupplyState
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Contact information for the primary system owner, for example, phone number, email address, and so on. This property is inherited from CIM_System.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PrimaryOwnerContact() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PrimaryOwnerContact = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PrimaryOwnerContact = objItem.PrimaryOwnerContact
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the primary system owner. This property is inherited from CIM_System.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PrimaryOwnerName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PrimaryOwnerName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                PrimaryOwnerName = objItem.PrimaryOwnerName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If enabled, the value is 4 and the unitary computer system can be reset using the power and reset buttons. If disabled, the value is 3, and a reset is not allowed. This property is inherited from CIM_UnitaryComputerSystem.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ResetCapability() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ResetCapability = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                ResetCapability = objItem.ResetCapability
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of automatic resets since the last reset. A value of –1 (minus one) indicates that the count is unknown.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ResetCount() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ResetCount = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                ResetCount = objItem.ResetCount
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of consecutive times a system reset is attempted. A value of –1 (minus one) indicates that the limit is unknown.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ResetLimit() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ResetLimit = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                ResetLimit = objItem.ResetLimit
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' List that specifies the roles of a system in the information technology environment. This property is inherited from CIM_System.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Roles() As String()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Roles = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Roles = objItem.Roles
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Current status of an object. Various operational and nonoperational statuses can be defined. Operational statuses include: OK, Degraded, and Pred Fail, which is an element such as a SMART-enabled hard disk drive that may be functioning properly, but predicts a failure in the near future. Nonoperational statuses include: Error, Starting, Stopping, and Service, which can apply during mirror-resilvering of a disk, reloading a user permissions list, or other administrative work. Not all status work is online, but the managed element is not OK or in one of the other states. This property is inherited from CIM_ManagedSystemElement.
        '''Values include the following:
        '''OK ("OK")
        '''Error ("Error")
        '''Degraded ("Degraded")
        '''Unknown ("Unknown")
        '''Pred Fail ("Pred Fail")
        '''Starting ("Starting")
        '''Stopping ("Stopping")
        '''Service ("Service")
        '''Stressed ("Stressed")
        '''NonRecover ("NonRecover")
        '''No Contact ("No Contact")
        '''Lost Comm ("Lost Comm")
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Status() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Status = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Status = objItem.Status
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' List of the support contact information for the Windows operating system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SupportContactDescription() As String()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SupportContactDescription = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                SupportContactDescription = objItem.SupportContactDescription
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The family to which a particular computer belongs. A family refers to a set of computers that are similar but not identical from a hardware or software point of view.
        '''This value comes from the Family member of the System Information structure in the SMBIOS information.
        '''Windows Server 2012 R2, Windows 8.1, Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows 10 and Windows Server 2016 Technical Preview.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemFamily() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemFamily = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                SystemFamily = objItem.SystemFamily
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Identifies a particular computer configuration for sale. It is sometimes also called a product ID or purchase order number.
        '''This value comes from the SKU Number member of the System Information structure in the SMBIOS information.
        '''Windows Server 2012 R2, Windows 8.1, Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows 10 and Windows Server 2016 Technical Preview.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemSKUNumber() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemSKUNumber = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                SystemSKUNumber = objItem.SystemSKUNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' SystemStartupDelay is no longer available for use because Boot.ini is not used to configure system startup. Instead, use the BCD classes supplied by the Boot Configuration Data (BCD) WMI provider or the Bcdedit command.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemStartupDelay() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemStartupDelay = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                SystemStartupDelay = objItem.SystemStartupDelay
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' SystemStartupOptions is no longer available for use because Boot.ini is not used to configure system startup. Instead, use the BCD classes supplied by the Boot Configuration Data (BCD) WMI provider or the Bcdedit command.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemStartupOptions() As String()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemStartupOptions = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                SystemStartupOptions = objItem.SystemStartupOptions
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' SystemStartupSetting is no longer available for use because Boot.ini is not used to configure system startup. Instead, use the BCD classes supplied by the Boot Configuration Data (BCD) WMI provider or the Bcdedit command.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemStartupSetting() As SByte
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemStartupSetting = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                SystemStartupSetting = objItem.SystemStartupSetting
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' System running on the Windows-based computer. This property must have a value.
        '''The following list identifies some of the possible values for this property.
        '''"x64-based PC"
        '''"X86-based PC"
        '''"MIPS-based PC"
        '''"Alpha-based PC"
        '''"Power PC"
        '''"SH-x PC"
        '''"StrongARM PC"
        '''"64-bit Intel PC"
        '''"64-bit Alpha PC"
        '''"Unknown"
        '''"X86-Nec98 PC"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemType() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                SystemType = objItem.SystemType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Thermal state of the system when last booted.
        '''This value comes from the Thermal State member of the System Enclosure or Chassis structure in the SMBIOS information.
        '''Other (1)
        '''Unknown (2)
        '''Safe (3)
        '''Warning (4)
        '''Critical (5)
        '''Non-recoverable (6)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ThermalState() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ThermalState = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                ThermalState = objItem.ThermalState
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Total size of physical memory. Be aware that, under some circumstances, this property may not return an accurate value for the physical memory. For example, it is not accurate if the BIOS is using some of the physical memory. For an accurate value, use the Capacity property in Win32_PhysicalMemory instead.
        '''Example: 67108864
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TotalPhysicalMemory() As UInt64
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            TotalPhysicalMemory = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                TotalPhysicalMemory = objItem.TotalPhysicalMemory
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of a user that is logged on currently. This property must have a value. In a terminal services session, UserName returns the name of the user that is logged on to the console—not the user logged on during the terminal service session.
        '''Example: jeffsmith
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function UserName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            UserName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                UserName = objItem.UserName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Event that causes the system to power up.
        '''This value comes from the Wake-up Type member of the System Information structure in the SMBIOS information.
        '''Reserved (0)
        '''Other (1)
        '''Unknown (2)
        '''APM Timer (3)
        '''Modem Ring (4)
        '''LAN Remote (5)
        '''Power Switch (6)
        '''PCI PME# (7)
        '''AC Power Restored (8)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function WakeUpType() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            WakeUpType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                WakeUpType = objItem.WakeUpType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the workgroup for this computer. If the value of the PartOfDomain property is False, then the name of the workgroup is returned.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Workgroup() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Workgroup = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
            For Each objItem In objItems
                Workgroup = objItem.Workgroup
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
End Class