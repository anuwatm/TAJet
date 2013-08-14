Imports System.DirectoryServices

Public Class AD
    Private pUsername As String
    Private pPassword As String
    Private pDomain As String
    Private pPrefix As String
    ''' <summary>
    ''' User Profile in AD
    ''' </summary>
    ''' <remarks></remarks>
    Structure Profile
        'FirstName=givenName
        'MiddleInitial=initials
        'LastName=sn
        'UserPrincipalName=UserPrincipalName
        'PostalAddress=PostalAddress
        'MailingAddress=MailingAddress
        'ResidentialAddress=HomePostalAddress
        'Title=Title
        'HomePhone=HomePhone
        'OfficePhone=TelephoneNumber
        'Mobile=Mobile
        'HomePhone=HomePhone
        'Fax=FacsimileTelephoneNumber
        'Email=Email
        'Url=Url
        'UserName=sAMAccountName
        'DistinguishedName=DistinguishedName
        'IsAccountActive = to check the user status in the active directory.
        Dim LoginName As String     'Login Name=cn
        Dim FirstName As String     'First Name=givenName
        Dim MiddleName As String    'Middle Initials=initials
        Dim LastName As String      'Last Name=sn
        'Dim Username As String     'username=sAMAccountName
        Dim Fullname As String      'Fullname=displayName
        Dim Mail As String          'Mail=mail
        Dim Description As String   'Description=description
        Dim CreateAt As String      'Time of Create User=whenCreated
    End Structure
    ''' <summary>
    ''' Domain name + Extension Ex. Domain.com
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Domain() As Integer
        Get
            Return pDomain
        End Get
        Set(ByVal value As Integer)
            pDomain = value
        End Set
    End Property
    ''' <summary>
    ''' login name of User
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Username() As Integer
        Get
            Return pUsername
        End Get
        Set(ByVal value As Integer)
            pUsername = value
        End Set
    End Property
    ''' <summary>
    ''' Password of User
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Password() As Integer
        Get
            Return pPassword
        End Get
        Set(ByVal value As Integer)
            pPassword = value
        End Set
    End Property
    ''' <summary>
    ''' Domain name no Extension 
    ''' </summary>
    ''' <value>domain name no extension</value>
    ''' <returns>string</returns>
    ''' <remarks></remarks>
    Public Property Prefix() As Integer
        Get
            Return pPrefix
        End Get
        Set(ByVal value As Integer)
            pPrefix = value
        End Set
    End Property
    ''' <summary>
    ''' Perfix + username (Readonly)
    ''' </summary>
    ''' <value>-</value>
    ''' <returns>string</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FullUsername() As String
        Get
            Return pPrefix.Trim & "\" & pUsername.Trim
        End Get

    End Property
    ''' <summary>
    ''' AuthenticateUser With Active Directory
    ''' </summary>
    ''' <param name="Domain">Domain Name Ex.Domainname.com</param>
    ''' <param name="prefix">Name of Domain no Extension  Ex. Domainname</param>
    ''' <param name="user">Username to Authen </param>
    ''' <param name="pass">Password of Username</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function AuthenticateUser(ByVal Domain As String, ByVal prefix As String, ByVal user As String, ByVal pass As String) As Boolean
        Dim dirEntry As New DirectoryEntry(Domain, prefix & "\" & user, pass)
        Try
            Dim nat As Object
            nat = dirEntry.NativeObject
            Return True
        Catch
            Return False
        End Try
    End Function
    ''' <summary>
    ''' AuthenticateUser With Active Directory
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function AuthenticateUser() As Boolean
        Dim dirEntry As New DirectoryEntry(pDomain, FullUsername, pPassword)
        Try
            Dim nat As Object
            nat = dirEntry.NativeObject
            Return True
        Catch
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Read User AD Profile
    ''' </summary>
    ''' <param name="oSearch">Object SearchResult</param>
    ''' <param name="PropertyName">Property Profile name</param>
    ''' <returns>profile Value</returns>
    ''' <remarks></remarks>
    Private Function ReadProfileProperty(ByVal oSearch As SearchResult, ByVal PropertyName As String) As String
        If oSearch.Properties(PropertyName).Count > 0 Then
            Return oSearch.Properties(PropertyName)(0)
        Else
            Return ""
        End If
    End Function
    ''' <summary>
    ''' Read User Profile in AD
    ''' </summary>
    ''' <param name="Domain">Domain name Ex. Domain.com</param>
    ''' <param name="Username">Username (Login name)</param>
    ''' <returns>User Profile</returns>
    ''' <remarks></remarks>
    Function getUserData(ByVal Domain As String, ByVal Username As String) As Profile
        Dim oDirectory As DirectoryEntry = New DirectoryEntry(Domain)
        Dim oDirSeach As DirectorySearcher = New DirectorySearcher(oDirectory)
        oDirSeach.Filter = "samaccountname=" + Username '+ ""
        oDirSeach.PropertiesToLoad.Add("CN")
        oDirSeach.PropertiesToLoad.Add("givenName")
        oDirSeach.PropertiesToLoad.Add("initials")
        oDirSeach.PropertiesToLoad.Add("LastName")
        oDirSeach.PropertiesToLoad.Add("sn")
        oDirSeach.PropertiesToLoad.Add("sAMAccountName")
        oDirSeach.PropertiesToLoad.Add("displayName")
        oDirSeach.PropertiesToLoad.Add("mail")
        oDirSeach.PropertiesToLoad.Add("Description")
        oDirSeach.PropertiesToLoad.Add("whenCreated")
        Dim sResult As SearchResult = oDirSeach.FindOne
        If Not IsNothing(sResult) Then
            Dim oProfile As Profile = New Profile
            With oProfile
                .LoginName = ReadProfileProperty(sResult, "cn")
                .FirstName = ReadProfileProperty(sResult, "givenName")
                .MiddleName = ReadProfileProperty(sResult, "initials")
                .LastName = ReadProfileProperty(sResult, "sn")
                .Fullname = ReadProfileProperty(sResult, "displayName")
                .Mail = ReadProfileProperty(sResult, "mail")
                .Description = ReadProfileProperty(sResult, "Description")
                .CreateAt = ReadProfileProperty(sResult, "whenCreated")
            End With
            Return oProfile
        Else
            Return Nothing
        End If
    End Function
End Class
