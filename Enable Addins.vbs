' As Excel from time to time disable addins due to issues
' it is better to be sure they are enabled before run refresh

Call ClearDisabledXLAddins

Call ActivatePowerPivotAddin

Call ActivatePowerQueryAddin

' TODO
' create in registry under Excel - Security subkey with DWORD DataConnectionWarnings = 0
' to avoid Power Query messages on Privacy Levels

Sub ActivatePowerPivotAddin()
	' Dim myRegKey As String
    On Error Resume Next
    
	myRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\Microsoft.AnalysisServices.Modeler.FieldList\LoadBehavior"
    If RegKeyExists(myRegKey) = True then RegKeySave myRegKey, 3
	
	myRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\PowerPivotExcelClientAddIn.NativeEntry.1\LoadBehavior"
    If RegKeyExists(myRegKey) = True then RegKeySave myRegKey, 3
	
    myRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\Excel\Addins\Microsoft.AnalysisServices.Modeler.FieldList\LoadBehavior"
    If RegKeyExists(myRegKey) = True then RegKeySave myRegKey, 3
	
	myRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\Excel\Addins\PowerPivotExcelClientAddIn.NativeEntry.1\LoadBehavior"
    If RegKeyExists(myRegKey) = True then RegKeySave myRegKey, 3
	
end sub

Sub ActivatePowerQueryAddin()
    ' Dim myRegKey As String
    On Error Resume Next       
    
    myRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\Excel\AddIns\Microsoft.Mashup.Client.Excel\LoadBehavior"
    If RegKeyExists(myRegKey) = True then RegKeySave myRegKey, 3
    
    myRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\Microsoft.Mashup.Client.Excel\LoadBehavior"
    If RegKeyExists(myRegKey) = True then RegKeySave myRegKey, 3
    
End Sub

Sub ClearDisabledXLAddins()
    ' Dim myRegKey As String
	On Error Resume Next
	
    myRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Excel\Resiliency\DisabledItems\"
    If RegKeyExists(myRegKey) = True Then RegKeyDelete myRegKey
    
    myRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems\"
    If RegKeyExists(myRegKey) = True Then RegKeyDelete myRegKey
    
    myRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems\"
    If RegKeyExists(myRegKey) = True Then RegKeyDelete myRegKey
	
	myRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\Excel\Resiliency\DisabledItems\"
	If RegKeyExists(myRegKey) = True Then RegKeyDelete myRegKey
		
End Sub

'http://www.slipstick.com/developer/read-and-change-a-registry-key-using-vba/

'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Function RegKeyRead(i_RegKey)
    ' Dim myWS As Object
 
    On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'read key from registry
    RegKeyRead = myWS.RegRead(i_RegKey)
End Function

'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
 
' change REG_DWORD to the correct key type
Sub RegKeySave(i_RegKey, i_Value, i_Type)    
	' Dim myWS As Object
	
	if i_Type = vbNullString then i_Type = "REG_DWORD"
	
	On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'write registry key
    myWS.RegWrite i_RegKey, i_Value, i_Type
 
End Sub

'returns True if the registry key i_RegKey was found
'and False if not
Function RegKeyExists(i_RegKey)
	' Dim myWS As Object
 
    On Error Resume Next
    Set myWS = CreateObject("WScript.Shell")    
    myWS.RegRead i_RegKey    
    RegKeyExists = ( Err.Number = 0 )    
    
End Function

'deletes i_RegKey from the registry
'returns True if the deletion was successful,
'and False if not (the key couldn't be found)
Function RegKeyDelete(i_RegKey)
    ' Dim myWS As Object

    On Error Resume Next    
    Set myWS = CreateObject("WScript.Shell")    
    myWS.RegDelete i_RegKey    
    RegKeyDelete = ( Err.Number = 0 )    
	
End Function
