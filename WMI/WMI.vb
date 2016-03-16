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
        Public Shared Function PowerManagementCapabilities() As Object()
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
                Roles = objItem.Roles()
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
    Public Class Win32_Process
        Public Shared Function Caption() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Caption = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process")
            For Each objItem In objItems
                Caption = objItem.Caption
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_Processor
        ''' <summary>
        ''' On a 32-bit operating system, the value is 32 and on a 64-bit operating system it is 64. This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddressWidth() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AddressWidth = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                AddressWidth = objItem.AddressWidth
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Processor architecture used by the platform.
        ''' 0 = x86
        ''' 1 = MIPS
        ''' 2 = Alpha
        ''' 3 = PowerPC
        ''' 5 = ARM
        ''' 6 = ia64
        ''' 9 = x64
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Architecture() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Architecture = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Architecture = objItem.Architecture
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Represents the asset tag of this processor.
        '''This value comes from the Asset Tag member of the Processor Information structure in the SMBIOS information.
        '''Windows Server 2012 R2, Windows 8.1, Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows Server 2016 Technical Preview and Windows 10.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AssetTag() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AssetTag = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                AssetTag = objItem.AssetTag
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Availability and status of the device. Inherited from CIM_LogicalDevice.
        ''' 1 = Other
        ''' 2 = Unknown
        ''' 3 = Running or Full Power
        ''' 4 = Warning
        ''' 5 = In Test
        ''' 6 = Not Applicable
        ''' 7 = Power Off
        ''' 8 = Off Line
        ''' 9 = Off Duty
        ''' 10 = Degraded
        ''' 11 = Not Installed
        ''' 12 = Install Error
        ''' 13 = Power Save - Unknown
        ''' 14 = Power Save - Low Power Mode
        ''' 15 = Power Save - Standby
        ''' 16 = Power Cycle
        ''' 17 = Power Save - Warning
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Availability() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Availability = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Availability = objItem.Availability
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Short description of an object (a one-line string). This property is inherited from CIM_ManagedSystemElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' This value comes from the Processor Characteristics member of the Processor Information structure in the SMBIOS information.
        ''' Windows Server 2012 R2, Windows 8.1, Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows Server 2016 Technical Preview and Windows 10.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Characteristics() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Characteristics = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Characteristics = objItem.Characteristics
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Windows API Configuration Manager error code.
        '''Value	Meaning
        '''0 (0x0) Device is working properly.
        '''1 (0x1) Device is not configured correctly.
        '''2 (0x2) Windows cannot load the driver for this device.
        '''3 (0x3) Driver for this device might be corrupted or the system may be low on memory or other resources.
        '''4 (0x4) Device is not working properly. One of its drivers or the registry might be corrupted.
        '''5 (0x5) Driver for the device requires a resource that Windows cannot manage.
        '''6 (0x6) Boot configuration for the device conflicts with other devices.
        '''7 (0x7) Cannot filter.
        '''8 (0x8) Driver loader for the device is missing.
        '''9 (0x9) Device is not working properly. The controlling firmware is incorrectly reporting the resources for the device.
        '''10 (0xA) Device cannot start.
        '''11 (0xB) Device failed.
        '''12 (0xC) Device cannot find enough free resources to use.
        '''13 (0xD) Windows cannot verify the device's resources.
        '''14 (0xE) Device cannot work properly until the computer is restarted.
        '''15 (0xF) Device is not working properly due to a possible re-enumeration problem.
        '''16 (0x10) Windows cannot identify all of the resources that the device uses.
        '''17 (0x11) Device is requesting an unknown resource type.
        '''18 (0x12) Device drivers must be reinstalled.
        '''19 (0x13) Failure using the VxD loader.
        '''20 (0x14) Registry might be corrupted.
        '''21 (0x15) System failure. If changing the device driver is ineffective, see the hardware documentation. Windows is removing the device.
        '''22 (0x16) Device is disabled.
        '''23 (0x17) System failure. If changing the device driver is ineffective, see the hardware documentation.
        '''24 (0x18) Device is not present, not working properly, or does not have all of its drivers installed.
        '''25 (0x19) Windows is still setting up the device.
        '''26 (0x1A) Windows is still setting up the device.
        '''27 (0x1B) Device does not have valid log configuration.
        '''28 (0x1C) Device drivers are not installed.
        '''29 (0x1D) Device is disabled. The device firmware did not provide the required resources.
        '''30 (0x1E) Device is using an IRQ resource that another device is using.
        '''31 (0x1F) Device is not working properly. Windows cannot load the required device drivers.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConfigManagerErrorCode() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ConfigManagerErrorCode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ConfigManagerErrorCode = objItem.ConfigManagerErrorCode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If TRUE, the device is using a configuration that the user defines. This property is inherited from CIM_LogicalDevice.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConfigManagerUserConfig() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ConfigManagerUserConfig = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ConfigManagerUserConfig = objItem.ConfigManagerUserConfig
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' This value comes from the Status member of the Processor Information structure in the SMBIOS information.
        '''Unknown (0)
        '''CPU Enabled (1)
        '''CPU Disabled by User via BIOS Setup (2)
        '''CPU Disabled By BIOS (POST Error) (3)
        '''CPU is Idle (4)
        '''Reserved (5)
        '''Reserved (6)
        '''Other (7)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CpuStatus() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CpuStatus = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                CpuStatus = objItem.CpuStatus
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the first concrete class that appears in the inheritance chain used to create an instance. When used with the other key properties of the class, the property allows all instances of this class and its subclasses to be identified uniquely. This property is inherited from CIM_LogicalDevice.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                CreationClassName = objItem.CreationClassName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Current speed of the processor, in MHz.
        '''This value comes from the Current Speed member of the Processor Information structure in the SMBIOS information.
        '''This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentClockSpeed() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentClockSpeed = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                CurrentClockSpeed = objItem.CurrentClockSpeed
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Voltage of the processor. If the eighth bit is set, bits 0-6 contain the voltage multiplied by 10. If the eighth bit is not set, then the bit setting in VoltageCaps represents the voltage value. CurrentVoltage is only set when SMBIOS designates a voltage value.
        '''Example: Value for a processor voltage of 1.8 volts is 0x12 (1.8 x 10).
        '''This value comes from the Voltage member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentVoltage() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentVoltage = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                CurrentVoltage = objItem.CurrentVoltage
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' On a 32-bit processor, the value is 32 and on a 64-bit processor it is 64. This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DataWidth() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DataWidth = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                DataWidth = objItem.DataWidth
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Description of the object. This property is inherited from CIM_ManagedSystemElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' 
        '''Unique identifier of a processor on the system. This property is inherited from CIM_LogicalDevice.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DeviceID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DeviceID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                DeviceID = objItem.DeviceID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If TRUE, the error reported in LastErrorCode is clear. This property is inherited from CIM_LogicalDevice.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ErrorCleared() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ErrorCleared = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ErrorCleared = objItem.ErrorCleared
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' More information about the error recorded in LastErrorCode, and information about corrective actions that can be taken. This property is inherited from CIM_LogicalDevice.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ErrorDescription() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ErrorDescription = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ErrorDescription = objItem.ErrorDescription
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' External clock frequency, in MHz. If the frequency is unknown, this property is set to NULL.
        '''This value comes from the External Clock member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ExtClock() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ExtClock = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ExtClock = objItem.ExtClock
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' This value comes from the Processor Information structure in the SMBIOS version information. For SMBIOS versions 2.0 thru 2.5 the value comes from the Processor Family member. For SMBIOS version 2.6+ the value comes from the Processor Family 2 member.
        '''This property is inherited from CIM_Processor.
        '''Value	Meaning
        '''1 (0x1) Other
        '''2 (0x2) Unknown
        '''3 (0x3) 8086
        '''4 (0x4) 80286
        '''5 (0x5) Intel386™ Processor
        '''6 (0x6) Intel486™ Processor
        '''7 (0x7) 8087
        '''8 (0x8) 80287
        '''9 (0x9) 80387
        '''10 (0xA) 80487
        '''11 (0xB) Pentium Brand
        '''12 (0xC) Pentium Pro
        '''13 (0xD) Pentium II
        '''14 (0xE) Pentium Processor with MMX™ Technology
        '''15 (0xF) Celeron™
        '''16 (0x10) Pentium II Xeon™
        '''17 (0x11) Pentium III
        '''18 (0x12) M1 Family
        '''19 (0x13) M2 Family
        '''24 (0x18) AMD Duron™ Processor Family
        '''25 (0x19) K5 Family
        '''26 (0x1A) K6 Family
        '''27 (0x1B) K6-2
        '''28 (0x1C) K6-3
        '''29 (0x1D) AMD Athlon™ Processor Family
        '''30 (0x1E) AMD2900 Family
        '''31 (0x1F) K6-2+
        '''32 (0x20) Power PC Family
        '''33 (0x21) Power PC 601
        '''34 (0x22) Power PC 603
        '''35 (0x23) Power PC 603+
        '''36 (0x24) Power PC 604
        '''37 (0x25) Power PC 620
        '''38 (0x26) Power PC X704
        '''39 (0x27) Power PC 750
        '''48 (0x30) Alpha Family
        '''49 (0x31) Alpha 21064
        '''50 (0x32) Alpha 21066
        '''51 (0x33) Alpha 21164
        '''52 (0x34) Alpha 21164PC
        '''53 (0x35) Alpha 21164a
        '''54 (0x36) Alpha 21264
        '''55 (0x37) Alpha 21364
        '''64 (0x40) MIPS Family
        '''65 (0x41) MIPS R4000
        '''66 (0x42) MIPS R4200
        '''67 (0x43) MIPS R4400
        '''68 (0x44) MIPS R4600
        '''69 (0x45) MIPS R10000
        '''80 (0x50) SPARC Family
        '''81 (0x51) SuperSPARC
        '''82 (0x52) microSPARC II
        '''83 (0x53) microSPARC IIep
        '''84 (0x54) UltraSPARC
        '''85 (0x55) UltraSPARC II
        '''86 (0x56) UltraSPARC IIi
        '''87 (0x57) UltraSPARC III
        '''88 (0x58) UltraSPARC IIIi
        '''96 (0x60) 68040
        '''97 (0x61) 68xxx Family
        '''98 (0x62) 68000
        '''99 (0x63) 68010
        '''100 (0x64) 68020
        '''101 (0x65) 68030
        '''112 (0x70) Hobbit Family
        '''120 (0x78) Crusoe™ TM5000 Family
        '''121 (0x79) Crusoe™ TM3000 Family
        '''122 (0x7A) Efficeon™ TM8000 Family
        '''128 (0x80) Weitek
        '''130 (0x82) Itanium™ Processor
        '''131 (0x83) AMD Athlon™ 64 Processor Family
        '''132 (0x84) AMD Opteron™ Processor Family
        '''144 (0x90) PA-RISC Family
        '''145 (0x91) PA-RISC 8500
        '''146 (0x92) PA-RISC 8000
        '''147 (0x93) PA-RISC 7300LC
        '''148 (0x94) PA-RISC 7200
        '''149 (0x95) PA-RISC 7100LC
        '''150 (0x96) PA-RISC 7100
        '''160 (0xA0) V30 Family
        '''176 (0xB0) Pentium III Xeon™ Processor
        '''177 (0xB1) Pentium III Processor with Intel SpeedStep™ Technology
        '''178 (0xB2) Pentium 4
        '''179 (0xB3) Intel Xeon™
        '''180 (0xB4) AS400 Family
        '''181 (0xB5) Intel Xeon™ Processor MP
        '''182 (0xB6) AMD Athlon™ XP Family
        '''183 (0xB7) AMD Athlon™ MP Family
        '''184 (0xB8) Intel Itanium 2
        '''185 (0xB9) Intel Pentium M Processor
        '''190 (0xBE) K7
        '''198 (0xC6) Intel Core™ i7-2760QM
        '''200 (0xC8) IBM390 Family
        '''201 (0xC9) G4
        '''202 (0xCA) G5
        '''203 (0xCB) G6
        '''204 (0xCC) z/Architecture Base
        '''250 (0xFA) i860
        '''251 (0xFB) i960
        '''260 (0x104) SH-3
        '''261 (0x105) SH-4
        '''280 (0x118) ARM
        '''281 (0x119) StrongARM
        '''300 (0x12C) 6x86
        '''301 (0x12D) MediaGX
        '''302 (0x12E) MII
        '''320 (0x140) WinChip
        '''350 (0x15E) DSP
        '''500 (0x1F4) Video Processor
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Family() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Family = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Family = objItem.ExtClock
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Date and time the object is installed. This property does not require a value to indicate that the object is installed. This property is inherited from CIM_ManagedSystemElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Dim str As String = objItem.InstallDate
                InstallDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' 
        '''Size of the Level 2 processor cache. A Level 2 cache is an external memory area that has a faster access time than the main RAM memory.
        '''This value comes from the L2 Cache Handle member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function L2CacheSize() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            L2CacheSize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                L2CacheSize = objItem.L2CacheSize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Clock speed of the Level 2 processor cache. A Level 2 cache is an external memory area that has a faster access time than the main RAM memory.
        '''This value comes from the L2 Cache Handle member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function L2CacheSpeed() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            L2CacheSpeed = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                L2CacheSpeed = objItem.L2CacheSpeed
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Size of the Level 3 processor cache. A Level 3 cache is an external memory area that has a faster access time than the main RAM memory.
        '''This value comes from the L3 Cache Handle member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function L3CacheSize() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            L3CacheSize = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                L3CacheSize = objItem.L3CacheSize
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Clockspeed of the Level 3 property cache. A Level 3 cache is an external memory area that has a faster access time than the main RAM memory.
        '''This value comes from the L3 Cache Handle member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function L3CacheSpeed() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            L3CacheSpeed = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                L3CacheSpeed = objItem.L3CacheSpeed
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Last error code reported by the logical device. This property is inherited from CIM_LogicalDevice.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LastErrorCode() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LastErrorCode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                LastErrorCode = objItem.LastErrorCode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Definition of the processor type. The value depends on the architecture of the processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Level() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Level = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Level = objItem.Level
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' 
        '''Load capacity of each processor, averaged to the last second. Processor loading refers to the total computing burden for each processor at one time. This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LoadPercentage() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LoadPercentage = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                LoadPercentage = objItem.LoadPercentage
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the processor manufacturer.
        '''Example: A. Datum Corporation
        '''This value comes from the Processor Manufacturer member of the Processor Information structure in the SMBIOS information.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Manufacturer = objItem.Manufacturer
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Maximum speed of the processor, in MHz.
        '''This value comes from the Max Speed member of the Processor Information structure in the SMBIOS information.
        '''This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MaxClockSpeed() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MaxClockSpeed = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                MaxClockSpeed = objItem.MaxClockSpeed
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Label by which the object is known. When this property is a subclass, it can be overridden to be a key property.
        '''This value comes from the Processor Version member of the Processor Information structure in the SMBIOS information.
        '''This property is inherited from CIM_ManagedSystemElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of cores for the current instance of the processor. A core is a physical processor on the integrated circuit. For example, in a dual-core processor this property has a value of 2. For more information, see Remarks.
        '''This value comes from the Processor Information structure in the SMBIOS version information. For SMBIOS versions 2.5 thru 2.9 the value comes from the Core Count member. For SMBIOS version 3.0+ the value comes from the Core Count 2 member.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumberOfCores() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfCores = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                NumberOfCores = objItem.NumberOfCores
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The number of enabled cores per processor socket.
        '''This value comes from the Processor Information structure in the SMBIOS version information. For SMBIOS versions 2.5 thru 2.9 the value comes from the Core Enabled member. For SMBIOS version 3.0+ the value comes from the Core Enabled 2 member.
        '''Windows Server 2012 R2, Windows 8.1, Windows Server 2012, Windows 8, Windows Server 2008 R2, Windows 7, Windows Server 2008, and Windows Vista:  This property is not supported before Windows Server 2016 Technical Preview and Windows 10.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function NumberOfEnabledCore() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfEnabledCore = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                NumberOfEnabledCore = objItem.NumberOfEnabledCore
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of logical processors for the current instance of the processor. For processors capable of hyperthreading, this value includes only the processors which have hyperthreading enabled. For more information, see Remarks.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                NumberOfLogicalProcessors = objItem.NumberOfLogicalProcessors
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Processor family type. Used when the Family property is set to 1, which means Other. This string should be set to NULL when the Family property is a value that is not 1. This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OtherFamilyDescription() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OtherFamilyDescription = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                OtherFamilyDescription = objItem.OtherFamilyDescription
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The part number of this processor as set by the manufacturer.
        '''This value comes from the Part Number member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PartNumber() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PartNumber = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                PartNumber = objItem.PartNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Windows Plug and Play device identifier of the logical device. This property is inherited from CIM_LogicalDevice.
        '''Example: *PNP030b
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PNPDeviceID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PNPDeviceID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                PNPDeviceID = objItem.PNPDeviceID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Array of the specific power-related capabilities of a logical device. This property is inherited from CIM_LogicalDevice.
        '''Value	Meaning
        '''0 (0x0) Unknown
        '''1 (0x1) Not Supported
        '''2 (0x2) Disabled
        '''3 (0x3) Enabled - The power management features are currently enabled but the exact feature set is unknown or the information is unavailable.
        '''4 (0x4) Power Saving Modes Entered Automatically - The device can change its power state based on usage or other criteria.
        '''5 (0x5) Power State Settable - The SetPowerState method is supported. This method is found on the parent CIM_LogicalDevice class and can be implemented. For more information, see Designing Managed Object Format (MOF) Classes.
        '''6 (0x6) Power Cycling Supported - The SetPowerState method can be invoked with the PowerState parameter set to 5 (Power Cycle).
        '''7 (0x7) Timed Power-On Supported - The SetPowerState method can be invoked with the PowerState parameter set to 5 (Power Cycle) and Time set to a specific date and time, or interval, for power-on.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PowerManagementCapabilities() As Object()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerManagementCapabilities = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                PowerManagementCapabilities = objItem.PowerManagementCapabilities
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If TRUE, the power of the device can be managed, which means that it can be put into suspend mode, and so on. The property does not indicate that power management features are enabled, but it does indicate that the logical device power can be managed. This property is inherited from CIM_LogicalDevice.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                PowerManagementSupported = objItem.PowerManagementSupported
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Processor information that describes the processor features. For an x86 class CPU, the field format depends on the processor support of the CPUID instruction. If the instruction is supported, the property contains 2 (two) DWORD formatted values. The first is an offset of 08h-0Bh, which is the EAX value that a CPUID instruction returns with input EAX set to 1. The second is an offset of 0Ch-0Fh, which is the EDX value that the instruction returns. Only the first two bytes of the property are significant and contain the contents of the DX register at CPU reset—all others are set to 0 (zero), and the contents are in DWORD format.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ProcessorId() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ProcessorId = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ProcessorId = objItem.ProcessorId
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Primary function of the processor.
        '''This value comes from the Processor Type member of the Processor Information structure in the SMBIOS information.
        '''Other (1)
        '''Unknown (2)
        '''Central Processor (3)
        '''Math Processor (4)
        '''DSP Processor (5)
        '''Video Processor (6)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ProcessorType() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ProcessorType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ProcessorType = objItem.ProcessorType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' System revision level that depends on the architecture. The system revision level contains the same values as the Version property, but in a numerical format.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Revision() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Revision = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Revision = objItem.Revision
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Role of the processor. This property is inherited from CIM_Processor.
        '''Examples: Central Processor or Math Processor
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Role() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Role = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Role = objItem.Role
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the processor supports address translation extensions used for virtualization.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SecondLevelAddressTranslationExtensions() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SecondLevelAddressTranslationExtensions = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                SecondLevelAddressTranslationExtensions = objItem.SecondLevelAddressTranslationExtensions
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The serial number of this processor This value is set by the manufacturer and normally not changeable.
        '''This value comes from the Serial Number member of the Processor Information structure in the SMBIOS information.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                SerialNumber = objItem.SerialNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Type of chip socket used on the circuit.
        '''Example: J202
        '''This value comes from the Socket Designation member of the Processor Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SocketDesignation() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SocketDesignation = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                SocketDesignation = objItem.SocketDesignation
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Values include the following:
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Status = objItem.Status
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' State of the logical device. If this property does not apply to the logical device, use the value 5, which means Not Applicable. This property is inherited from CIM_LogicalDevice.
        '''Value	Meaning
        '''1 (0x1) Other
        '''2 (0x2) Unknown
        '''3 (0x3) Enabled
        '''4 (0x4) Disabled
        '''5 (0x5) Not Applicable
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function StatusInfo() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            StatusInfo = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                StatusInfo = objItem.StatusInfo
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Revision level of the processor in the processor family. This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Stepping() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Stepping = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Stepping = objItem.Stepping
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Value of the CreationClassName property for the scoping computer. This property is inherited from CIM_LogicalDevice.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemCreationClassName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemCreationClassName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                SystemCreationClassName = objItem.SystemCreationClassName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the scoping system. This property is inherited from CIM_LogicalDevice.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                SystemName = objItem.SystemName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The number of threads per processor socket.
        '''This value comes from the Processor Information structure in the SMBIOS version information. For SMBIOS versions 2.5 thru 2.9 the value comes from the Thread Count member. For SMBIOS version 3.0+ the value comes from the Thread Count 2 member.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadCount() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ThreadCount = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                ThreadCount = objItem.ThreadCount
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Globally unique identifier for the processor. This identifier may only be unique within a processor family. This property is inherited from CIM_Processor.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function UniqueId() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            UniqueId = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                UniqueId = objItem.UniqueId
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' CPU socket information, including the method by which this processor can be upgraded, if upgrades are supported. This property is an integer enumeration.
        '''This value comes from the Processor Upgrade member of the Processor Information structure in the SMBIOS information.
        '''This property is inherited from CIM_Processor.
        '''Value	Meaning
        '''1 (0x1) Other
        '''2 (0x2) Unknown
        '''3 (0x3) Daughter Board
        '''4 (0x4) ZIF Socket
        '''5 (0x5) Replacement or Piggy Back
        '''6 (0x6) None
        '''7 (0x7) LIF Socket
        '''8 (0x8) Slot 1
        '''9 (0x9) Slot 2
        '''10 (0xA) 370 Pin Socket
        '''11 (0xB) Slot A
        '''12 (0xC) Slot M
        '''13 (0xD) Socket 423
        '''14 (0xE) Socket A (Socket 462)
        '''15 (0xF) Socket 478
        '''16 (0x10) Socket 754
        '''17 (0x11) Socket 940
        '''18 (0x12) Socket 939
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function UpgradeMethod() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            UpgradeMethod = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                UpgradeMethod = objItem.UpgradeMethod
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Processor revision number that depends on the architecture.
        '''Example: Model 2, Stepping 12
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                Version = objItem.Version
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the Firmware has enabled virtualization extensions.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function VirtualizationFirmwareEnabled() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VirtualizationFirmwareEnabled = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                VirtualizationFirmwareEnabled = objItem.VirtualizationFirmwareEnabled
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If True, the processor supports Intel or AMD Virtual Machine Monitor extensions.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function VMMonitorModeExtensions() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VMMonitorModeExtensions = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                VMMonitorModeExtensions = objItem.VMMonitorModeExtensions
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Voltage capabilities of the processor. Bits 0-3 of the field represent specific voltages that the processor socket can accept. All other bits should be set to 0 (zero). The socket is configurable if multiple bits are set. For more information about the actual voltage at which the processor is running, see CurrentVoltage. If the property is NULL, then the voltage capabilities are unknown.
        '''Value	Meaning
        '''1 (0x1) 5 volts
        '''2 (0x2) 3.3 volts
        '''4 (0x4) 2.9 volts
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function VoltageCaps() As UInt32
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VoltageCaps = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
            For Each objItem In objItems
                VoltageCaps = objItem.VoltageCaps
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_BIOS
        ''' <summary>
        ''' Array of BIOS characteristics supported by the system as defined by the System Management BIOS Reference Specification.
        '''This value comes from the BIOS Characteristics member of the BIOS Information structure in the SMBIOS information.
        '''Value	Meaning
        '''Reserved
        '''0 Reserved
        '''1 Unknown
        '''2 BIOS Characteristics Not Supported
        '''3 ISA is supported
        '''4 MCA is supported
        '''5 EISA is supported
        '''6 PCI is supported
        '''7 PC Card (PCMCIA) is supported
        '''8 Plug and Play is supported
        '''9 APM is supported
        '''10 BIOS is Upgradeable (Flash)
        '''11 BIOS is Upgradable (Flash) - BIOS shadowing is allowed
        '''12 VL-VESA is supported
        '''13 ESCD support is available
        '''14 Boot from CD is supported
        '''15 Selectable Boot is supported
        '''16 BIOS ROM is socketed
        '''17 Boot From PC Card (PCMCIA) is supported
        '''18 EDD (Enhanced Disk Drive) Specification is supported
        '''19 Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5\", 1k Bytes/Sector, 360 RPM) is supported
        '''20 Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported - Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5\", 360 RPM) is supported
        '''21 Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported - Int 13h - 5.25\" / 360 KB Floppy Services are supported
        '''22 Int 13h - 5.25 / 360 KB Floppy Services are supported - Int 13h - 5.25\" /1.2MB Floppy Services are supported
        '''23Int 13h - 5.25 /1.2MB Floppy Services are supported - Int 13h - 3.5\" / 720 KB Floppy Services are supported
        '''24 Int 13h - 3.5 / 720 KB Floppy Services are supported - Int 13h - 3.5\" / 2.88 MB Floppy Services are supported
        '''25 Int 13h - 3.5 / 2.88 MB Floppy Services are supported - Int 5h, Print Screen Service is supported
        '''26 Int 9h, 8042 Keyboard services are supported
        '''27 Int 14h, Serial Services are supported
        '''28 Int 17h, printer services are supported
        '''29 Int 10h, CGA/Mono Video Services are supported
        '''30 NEC PC-98
        '''31 ACPI supported
        '''32 ACPI is supported USB Legacy is supported
        '''33 AGP is supported
        '''34 I2O boot is supported
        '''35 LS-120 boot is supported
        '''36 ATAPI ZIP Drive boot is supported
        '''37 1394 boot is supported
        '''38 Smart Battery supported
        '''39 Smart Battery is supported
        '''40:47 Reserved for BIOS vendor
        '''48:63 Reserved for system vendor
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BiosCharacteristics() As Object()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BiosCharacteristics = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                BiosCharacteristics = objItem.BiosCharacteristics

            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Array of the complete system BIOS information. In many computers there can be several version strings that are stored in the registry and represent the system BIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function BIOSVersion() As Object()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            BIOSVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                BIOSVersion = objItem.BIOSVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Internal identifier for this compilation of this software element. This property is inherited from CIM_SoftwareElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                BuildNumber = objItem.BuildNumber
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Code set used by this software element. This property is inherited from CIM_SoftwareElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                CodeSet = objItem.CodeSet
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name of the current BIOS language.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentLanguage() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentLanguage = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                CurrentLanguage = objItem.CurrentLanguage
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Description of the object. This property is inherited from CIM_ManagedSystemElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The major release of the embedded controller firmware.
        '''This value comes from the Embedded Controller Firmware Major Release member of the BIOS Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EmbeddedControllerMajorVersion() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            EmbeddedControllerMajorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                EmbeddedControllerMajorVersion = objItem.EmbeddedControllerMajorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The minor release of the embedded controller firmware.
        '''This value comes from the Embedded Controller Firmware Minor Release member of the BIOS Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EmbeddedControllerMinorVersion() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            EmbeddedControllerMinorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                EmbeddedControllerMinorVersion = objItem.EmbeddedControllerMinorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Manufacturer's identifier for this software element. Often this will be a stock keeping unit (SKU) or a part number. This property is inherited from CIM_SoftwareElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IdentificationCode() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            IdentificationCode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                IdentificationCode = objItem.IdentificationCode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Number of languages available for installation on this system. Language may determine properties such as the need for Unicode and bidirectional text.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function InstallableLanguages() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InstallableLanguages = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                InstallableLanguages = objItem.InstallableLanguages
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Date and time the object was installed. This property does not need a value to indicate that the object is installed. This property is inherited from CIM_ManagedSystemElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Dim str As String = objItem.InstallDate
                InstallDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Language edition of this software element. The language codes defined in ISO 639 should be used. Where the software element represents a multilingual or international version of a product, the string "multilingual" should be used. This property is inherited from CIM_SoftwareElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LanguageEdition() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LanguageEdition = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                LanguageEdition = objItem.LanguageEdition
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Array of names of available BIOS-installable languages.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ListOfLanguages() As Object()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ListOfLanguages = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                ListOfLanguages = objItem.ListOfLanguages
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Manufacturer of this software element.
        '''This value comes from the Vendor member of the BIOS Information structure in the SMBIOS information.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Manufacturer = objItem.Manufacturer
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Name used to identify this software element.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Records the manufacturer and operating system type for a software element when the TargetOperatingSystem property has a value of 1 (Other). When TargetOperatingSystem has a value of 1, OtherTargetOS must have a nonnull value. For all other values of TargetOperatingSystem, OtherTargetOS is NULL. This property is inherited from CIM_SoftwareElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function OtherTargetOS() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            OtherTargetOS = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                OtherTargetOS = objItem.OtherTargetOS
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If TRUE, this is the primary BIOS of the computer system. This property is inherited from CIM_BIOSElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function PrimaryBIOS() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PrimaryBIOS = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                PrimaryBIOS = objItem.PrimaryBIOS
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Release date of the Windows BIOS in the Coordinated Universal Time (UTC) format of YYYYMMDDHHMMSS.MMMMMM(+-)OOO.
        '''This value comes from the BIOS Release Date member of the BIOS Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ReleaseDate() As DateTime
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ReleaseDate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Dim str As String = objItem.ReleaseDate
                ReleaseDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Assigned serial number of the software element. This property is inherited from CIM_SoftwareElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SerialNumber = objItem.SerialNumber
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' BIOS version as reported by SMBIOS.
        ''' This value comes from the BIOS Version member of the BIOS Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SMBIOSBIOSVersion() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SMBIOSBIOSVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SMBIOSBIOSVersion = objItem.SMBIOSBIOSVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Major SMBIOS version number. This property is NULL if SMBIOS is not found.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SMBIOSMajorVersion() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SMBIOSMajorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SMBIOSMajorVersion = objItem.SMBIOSMajorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Minor SMBIOS version number. This property is NULL if SMBIOS is not found.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SMBIOSMinorVersion() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SMBIOSMinorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SMBIOSMinorVersion = objItem.SMBIOSMinorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' If true, the SMBIOS is available on this computer system.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SMBIOSPresent() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SMBIOSPresent = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SMBIOSPresent = objItem.SMBIOSPresent
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Identifier for this software element; designed to be used in conjunction with other keys to create a unique representation of this CIM_SoftwareElement instance. This property is inherited from CIM_SoftwareElement.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SoftwareElementID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SoftwareElementID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SoftwareElementID = objItem.SoftwareElementID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' State of a software element. This property is inherited from CIM_SoftwareElement.
        '''Value	Meaning
        '''0 Deployable
        '''1 Installable
        '''2 Executable
        '''3 Running
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SoftwareElementState() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SoftwareElementState = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SoftwareElementState = objItem.SoftwareElementState
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Current status of the object. Various operational and nonoperational statuses can be defined. Operational statuses include: "OK", "Degraded", and "Pred Fail" (an element, such as a SMART-enabled hard disk drive, may be functioning properly but predicting a failure in the near future). Nonoperational statuses include: "Error", "Starting", "Stopping", and "Service". The latter, "Service", could apply during mirror-resilvering of a disk, reload of a user permissions list, or other administrative work. Not all such work is online, yet the managed element is neither "OK" nor in one of the other states. This property is inherited from CIM_ManagedSystemElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Status = objItem.Status
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The major release of the System BIOS.
        '''This value comes from the System BIOS Major Release member of the BIOS Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemBiosMajorVersion() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemBiosMajorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SystemBiosMajorVersion = objItem.SystemBiosMajorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' The minor release of the System BIOS.
        '''This value comes from the System BIOS Minor Release member of the BIOS Information structure in the SMBIOS information.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SystemBiosMinorVersion() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemBiosMinorVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                SystemBiosMinorVersion = objItem.SystemBiosMinorVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Target operating system of the owning software element. This property is inherited from CIM_SoftwareElement. The possible values for this property are as follows.
        '''Value	Meaning
        '''0 Unknown
        '''1 Other
        '''2 MACOS
        '''3 ATTUNIX
        '''4 DGUX
        '''5 DECNT
        '''6 Digital Unix
        '''7 OpenVMS
        '''8 HPUX
        '''9 AIX
        '''10 MVS
        '''11 OS400
        '''12 OS/2
        '''13 JavaVM
        '''14 MSDOS
        '''15 WIN3x
        '''16 WIN95
        '''17 WIN98
        '''18 WINNT
        '''19 WINCE
        '''20 NCR3000
        '''21 NetWare
        '''22 OSF
        '''23 DC/OS
        '''24 Reliant UNIX
        '''25 SCO UnixWare
        '''26 SCO OpenServer
        '''27 Sequent
        '''28 IRIX
        '''29 Solaris
        '''30 SunOS
        '''31 U6000
        '''32 ASERIES
        '''33 TandemNSK
        '''34 TandemNT
        '''35 BS2000
        '''36 LINUX
        '''37 Lynx
        '''38 XENIX
        '''39 VM/ESA
        '''40 Interactive UNIX
        '''41 BSDUNIX
        '''42 FreeBSD
        '''43 NetBSD
        '''44 GNU Hurd
        '''45 OS9
        '''46 MACH Kernel
        '''47 Inferno
        '''48 QNX
        '''49 EPOC
        '''50 IxWorks
        '''51 VxWorks
        '''52 MiNT
        '''53 BeOS
        '''54 HP MPE
        '''55 NextStep
        '''56 PalmPilot
        '''57 Rhapsody
        '''58 Windows 2000
        '''59 Dedicated
        '''60 VSE
        '''61 TPF
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function TargetOperatingSystem() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            TargetOperatingSystem = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                TargetOperatingSystem = objItem.TargetOperatingSystem
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        ''' <summary>
        ''' Version of the BIOS. This string is created by the BIOS manufacturer. This property is inherited from CIM_SoftwareElement.
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_BIOS")
            For Each objItem In objItems
                Version = objItem.Version
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
    Public Class Win32_VideoController
        Public Shared Function AcceleratorCapabilities() As Object()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AcceleratorCapabilities = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                AcceleratorCapabilities = objItem.AcceleratorCapabilities
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                AdapterDACType = objItem.AdapterDACType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function AdapterRAM() As Long
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            AdapterRAM = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                AdapterRAM = objItem.AdapterRAM
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function Availability() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Availability = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Availability = objItem.Availability
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CapabilityDescriptions() As Object()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CapabilityDescriptions = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CapabilityDescriptions = objItem.CapabilityDescriptions
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Caption = objItem.Caption
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ColorTableEntries() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ColorTableEntries = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ColorTableEntries = objItem.ColorTableEntries
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ConfigManagerErrorCode() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ConfigManagerErrorCode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ConfigManagerErrorCode = objItem.ConfigManagerErrorCode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ConfigManagerUserConfig() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ConfigManagerUserConfig = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ConfigManagerUserConfig = objItem.ConfigManagerUserConfig
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CreationClassName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CreationClassName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CreationClassName = objItem.CreationClassName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentBitsPerPixel() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentBitsPerPixel = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentBitsPerPixel = objItem.CurrentBitsPerPixel
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentHorizontalResolution() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentHorizontalResolution = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentHorizontalResolution = objItem.CurrentHorizontalResolution
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentNumberOfColors() As Long
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentNumberOfColors = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentNumberOfColors = objItem.CurrentNumberOfColors
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentNumberOfColumns() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentNumberOfColumns = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentNumberOfColumns = objItem.CurrentNumberOfColumns
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentNumberOfRows() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentNumberOfRows = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentNumberOfRows = objItem.CurrentNumberOfRows
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentRefreshRate() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentRefreshRate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentRefreshRate = objItem.CurrentRefreshRate
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentScanMode() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentScanMode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentScanMode = objItem.CurrentScanMode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function CurrentVerticalResolution() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            CurrentVerticalResolution = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                CurrentVerticalResolution = objItem.CurrentVerticalResolution
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Description = objItem.Description
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function DeviceID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DeviceID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                DeviceID = objItem.DeviceID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function DeviceSpecificPens() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DeviceSpecificPens = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                DeviceSpecificPens = objItem.DeviceSpecificPens
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function DitherType() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DitherType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                DitherType = objItem.DitherType
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Dim str As String = objItem.DriverDate
                DriverDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function DriverVersion() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            DriverVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                DriverVersion = objItem.DriverVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ErrorCleared() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ErrorCleared = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ErrorCleared = objItem.ErrorCleared
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ErrorDescription() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ErrorDescription = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ErrorDescription = objItem.ErrorDescription
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ICMIntent() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ICMIntent = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ICMIntent = objItem.ICMIntent
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ICMMethod() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ICMMethod = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ICMMethod = objItem.ICMMethod
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                InfFilename = objItem.InfFilename
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function InfSection() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InfSection = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                InfSection = objItem.InfSection
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function InstallDate() As DateTime
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            InstallDate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Dim str As String = objItem.InstallDate
                InstallDate = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                InstalledDisplayDrivers = objItem.InstalledDisplayDrivers
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function LastErrorCode() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            LastErrorCode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                LastErrorCode = objItem.LastErrorCode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function MaxMemorySupported() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MaxMemorySupported = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                MaxMemorySupported = objItem.MaxMemorySupported
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function MaxNumberControlled() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MaxNumberControlled = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                MaxNumberControlled = objItem.MaxNumberControlled
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function MaxRefreshRate() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MaxRefreshRate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                MaxRefreshRate = objItem.MaxRefreshRate
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function MinRefreshRate() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            MinRefreshRate = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                MinRefreshRate = objItem.MinRefreshRate
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function Monochrome() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Monochrome = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Monochrome = objItem.Monochrome
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
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Name = objItem.Name
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function NumberOfColorPlanes() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfColorPlanes = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                NumberOfColorPlanes = objItem.NumberOfColorPlanes
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function NumberOfVideoPages() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            NumberOfVideoPages = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                NumberOfVideoPages = objItem.NumberOfVideoPages
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PNPDeviceID() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PNPDeviceID = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                PNPDeviceID = objItem.PNPDeviceID
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PowerManagementCapabilities() As Object()
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerManagementCapabilities = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                PowerManagementCapabilities = objItem.PowerManagementCapabilities
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function PowerManagementSupported() As Boolean
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            PowerManagementSupported = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                PowerManagementSupported = objItem.PowerManagementSupported
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ProtocolSupported() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ProtocolSupported = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ProtocolSupported = objItem.ProtocolSupported
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function ReservedSystemPaletteEntries() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            ReservedSystemPaletteEntries = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                ReservedSystemPaletteEntries = objItem.ReservedSystemPaletteEntries
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function SpecificationVersion() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SpecificationVersion = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                SpecificationVersion = objItem.SpecificationVersion
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function Status() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            Status = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Status = objItem.Status
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function StatusInfo() As UInt16
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            StatusInfo = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                StatusInfo = objItem.StatusInfo
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function SystemCreationClassName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemCreationClassName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                SystemCreationClassName = objItem.SystemCreationClassName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function SystemName() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemName = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                SystemName = objItem.SystemName
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function SystemPaletteEntries() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            SystemPaletteEntries = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                SystemPaletteEntries = objItem.SystemPaletteEntries
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function TimeOfLastReset() As DateTime
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            TimeOfLastReset = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                Dim str As String = objItem.TimeOfLastReset
                TimeOfLastReset = New DateTime(str.Substring(0, 4), str.Substring(4, 2), str.Substring(6, 2), str.Substring(8, 2), str.Substring(10, 2), str.Substring(12, 2))
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function VideoArchitecture() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VideoArchitecture = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                VideoArchitecture = objItem.VideoArchitecture
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function VideoMemoryType() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VideoMemoryType = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                VideoMemoryType = objItem.VideoMemoryType
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function VideoMode() As Integer
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VideoMode = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                VideoMode = objItem.VideoMode
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function VideoModeDescription() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VideoModeDescription = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                VideoModeDescription = objItem.VideoModeDescription
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
        Public Shared Function VideoProcessor() As String
            Dim objWMIService As Object
            Dim objItems As Object
            Dim objItem As Object
            Dim server As New Devices.ServerComputer
            Dim ComputerName As String = server.Name
            VideoProcessor = Nothing
            objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2")
            objItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController")
            For Each objItem In objItems
                VideoProcessor = objItem.VideoProcessor
            Next
            objWMIService = Nothing
            objItems = Nothing
            objItem = Nothing
        End Function
    End Class
End Class