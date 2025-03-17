'Setup ADO objects
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection
 
'Search Entire Active Direcotry domain
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("DefaultNamingContext")
strBase = "<LDAP://" & strDNSDomain & ">"
 
'Filter on user objects
strFilter = "(&(objectClass=user))"
 
'Comma delimited list of attribute values to retrieve
strAttributes = "sAMAccountName, distinguishedName"
 
'Construct the LDAP syntax query
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
 
'Properties of the query
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 10000
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False
 
'Run the query
Set adoRecordset = adoCommand.Execute
 
'Move to the start of the recordset
adoRecordset.MoveFirst
 
strResults = "User Login Names" 
'Enumerate the resulting recordset
Do Until adoRecordset.EOF 
	'Retrieve values and display
	strName = adoRecordset.Fields("sAMAccountName").Value 
	strDN = adoRecordset.Fields("distinguishedName").Value
	strResults = strResults & VbCrLf & strName & " | " & strDN
	adoRecordset.MoveNext
Loop
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("eredmeny.txt", True)
objFile.Write strResults
objFile.Close
Set objFile = Nothing
MsgBox "Elkészültem nézz bele a 'eredmeny.txt' -be."

