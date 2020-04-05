Option Explicit

Dim objFSO

Set objFSO = createObject("Scripting.FileSystemObject")

Dim arg0 : arg0 = WScript.Arguments.Item(0)

Dim strCurPath : strCurPath = arg0 & "/"
' Dim strCurPath : strCurPath = ObjFso.getParentFolderName(WScript.ScriptFullName) & "/"


' 作成したいフォルダ一覧
Dim folderNameArray : folderNameArray = Array( _
    "0_", _
    "1_", _
    "2_", _
    "3_", _
    "4_", _
    "5_", _
    "6_", _
    "7_", _
    "8_", _
    "9_", _
    "A_", _
    "B_", _
    "C_", _
    "D_", _
    "E_", _
    "F_", _
    "G_", _
    "H_", _
    "I_", _
    "J_", _
    "K_", _
    "L_", _
    "M_", _
    "N_", _
    "O_", _
    "P_", _
    "Q_", _
    "R_", _
    "S_", _
    "T_", _
    "U_", _
    "V_", _
    "W_", _
    "X_", _
    "Y_", _
    "Z_"  _
)

Dim folderNameNum : folderNameNum = UBound(folderNameArray)

' フォルダ作成（単一）
Function MakeFolder(strFolder)
    Dim strTargetPath : strTargetPath = strCurPath & strFolder

    If Not ObjFso.FolderExists(strTargetPath) Then
        ObjFso.CreateFolder(strTargetPath)
    End If
End Function

' 作成したいフォルダ一覧
Sub Main
    Dim i

    For i = 0 To folderNameNum
        MakeFolder(folderNameArray(i))
    Next
End Sub

Main()
Set objFSO = Nothing
