# read-Windows-SID-in-VBA
Read the windows SID in VBA code

code to be added to a procedure:

  Dim objWMIService As Object

  Dim objAccount  As Object

  Dim strComputer As String

  
  strComputer = "."

  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

  Set objAccount = objWMIService.Get("Win32_UserAccount.Name='" & Environ("USERNAME") & "',Domain='" & Environ("USERDOMAIN") & "'")

  
  Debug.Print "AccountType; " & objAccount.AccountType

  Debug.Print "Caption; " & objAccount.Caption

  Debug.Print "Description; " & objAccount.Description

  Debug.Print "Disabled; " & objAccount.Disabled

  Debug.Print "Domain; " & objAccount.Domain

  Debug.Print "FullName; " & objAccount.FullName

  Debug.Print "InstallDate; " & objAccount.InstallDate

  Debug.Print "LocalAccount; " & objAccount.LocalAccount

  Debug.Print "Lockout; " & objAccount.Lockout

  Debug.Print "Name; " & objAccount.Name

  Debug.Print "PasswordChangeable; " & objAccount.PasswordChangeable

  Debug.Print "PasswordExpires; " & objAccount.PasswordExpires

  Debug.Print "PasswordRequired; " & objAccount.PasswordRequired

  Debug.Print "SID; " & objAccount.SID

  Debug.Print "SIDType; " & objAccount.SIDType

  Debug.Print "Status; " & objAccount.Status
  

  strComputer = "."

  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

  Set objAccount = objWMIService.Get("Win32_SID.SID='S-1-5-21-746137067-1035525444-725345543-4119'")
  Debug.Print "AccountName; " & objAccount.AccountName
  Debug.Print "BinaryRepresentation[]; " & objAccount.BinaryRepresentation(0)
  Debug.Print "ReferencedDomainName; " & objAccount.ReferencedDomainName
  Debug.Print "SID; " & objAccount.SID
  Debug.Print "SidLength; " & objAccount.SidLength
