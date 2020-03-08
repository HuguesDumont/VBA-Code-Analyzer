Attribute VB_Name = "CommonProcedures"
Option Explicit

Global checkBoolean         As Boolean
Global checkByref           As Boolean
Global checKCommentMulti    As Boolean
Global checkDebug           As Boolean
Global checkEmptyMethod     As Boolean
Global checkEnd             As Boolean
Global checkGoTo            As Boolean
Global checkBoolParen       As Boolean
Global checkImplicitWbk     As Boolean
Global checkImplicitWs      As Boolean
Global checkInLine          As Boolean
Global checkMissingNext     As Boolean
Global checkMissingOption   As Boolean
Global checkMultiDim        As Boolean
Global checkNameConst       As Boolean
Global checkNameLength      As Boolean
Global checkNameMethod      As Boolean
Global checkNameVar         As Boolean
Global checkParamUse        As Boolean
Global checkMethodUse       As Boolean
Global checkResume          As Boolean
Global checkScope           As Boolean
Global checkSquareBrackets  As Boolean
Global checkStringConcat    As Boolean
Global checkTodo            As Boolean
Global checkType            As Boolean
Global checkVarUse          As Boolean

Global gdic_UsfMethods      As Scripting.Dictionary
Global gdic_WbkMethods      As Scripting.Dictionary
Global gdic_WsMethods       As Scripting.Dictionary
Global calledMacros         As Scripting.Dictionary
'Global wbNamedRanges        As Scripting.Dictionary

Global reg                  As VBScript_RegExp_55.RegExp

Global currentWorkbook      As Workbook

' Method to call Analysis frame from distant workbooks
Public Sub OpenAnalysis()
    CodeSmell.Show
End Sub

Public Sub AddIssue(ByVal ps_CompName As String, ByVal pl_Line As Long, ByVal ps_Desc As String, ByVal ps_Sol As String, ByVal ps_crit As String, _
        Optional ByVal ps_ObjName As String = vbNullString, Optional ByVal ps_MethodName As String = "-1")
    Dim ls_Line As String
    
    ls_Line = "12345"
    RSet ls_Line = CStr(pl_Line)
    
    With CodeSmell.CodeProblems
        .ListItems.Add , , currentWorkbook.Name
        
        With .ListItems(.ListItems.Count).ListSubItems
            .Add , , ps_CompName
            
            If (pl_Line = 0) Then
                .Add , , vbNullString
            ElseIf (ps_MethodName <> "-1") Then
                .Add , , ps_MethodName
            Else
                .Add , , currentWorkbook.VBProject.VBComponents(ps_CompName).CodeModule.ProcOfLine(pl_Line, vbext_pk_Proc)
            End If
            
            .Add , , ps_ObjName
            .Add , , IIf(pl_Line = 0, "1", ls_Line)
            .Add , , ps_Desc
            .Add , , ps_Sol
            .Add , , ps_crit
        End With
    End With
End Sub

Public Function GetCompType(ByRef po_VBC As VBComponent) As String
    On Error Resume Next
    
    If (po_VBC.Type = vbext_ct_StdModule) Then
        GetCompType = "Standard Module"
    ElseIf (po_VBC.Type = vbext_ct_ClassModule) Then
        GetCompType = "Class Module"
    ElseIf (po_VBC.Type = vbext_ct_Document) Then
        GetCompType = IIf(po_VBC.Properties("Name") = currentWorkbook.Name, "Workbook Document", "Worksheet Document")
    ElseIf (po_VBC.Type = vbext_ct_MSForm) Then
        GetCompType = "UserForm"
    ElseIf (po_VBC.Type = vbext_ct_ActiveXDesigner) Then
        GetCompType = "ActiveX Designer"
    Else
        GetCompType = "Unknown"
    End If
End Function

Public Sub FormatLineAs(ByRef ps_MethodLine As String)
    ps_MethodLine = " " & ps_MethodLine
    Call ReplaceAllRegPattern(ps_MethodLine, " Optional ", " ")
    ps_MethodLine = Trim(ps_MethodLine)
    Call ReplaceAllRegPattern(ps_MethodLine, " \+ ", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "((\.)(\w+))+", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "\(\)", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, ("((" & Chr(34) & Chr(34) & ")|(" & Chr(34) & "[^" & Chr(34) & "]+" & Chr(34) & "))"), vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "( = (\+|-)?(\d+)((\.)?(\d+))?(E(\+|-)\d+)?#?)", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "(\+|-)\d+", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "( = \w+)", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, " =( )*", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "\(((\w+)*(\d*( To \d+)?)((, ?((\w+)*\d*( To \d+)?))*))\)", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "(& .* &)", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "& \w+", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, " &", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "\(|\)", vbNullString)
    Call ReplaceAllRegPattern(ps_MethodLine, "(  )+", " ")
    Call ReplaceAllRegPattern(ps_MethodLine, "'.*", " ")
    Call ReplaceAllRegPattern(ps_MethodLine, "\:.*", vbNullString) ' Instructions après :
    
    ps_MethodLine = Trim(ps_MethodLine)
End Sub

' Method to check if string is not between quotation marks and not in comment
Public Function StringIsCode(ByVal ps_CodeLine As String, ByVal ps_String As String) As Integer
    Dim lb_Inquote      As Boolean
    Dim li_CodeEnd      As Long
    Dim li_Pos          As Long
    
    lb_Inquote = False
    StringIsCode = 0
    li_CodeEnd = CommentStart(ps_CodeLine)
    
    If (li_CodeEnd = 0) Then
        li_CodeEnd = Len(ps_CodeLine)
    End If
    
    For li_Pos = 1 To li_CodeEnd
        If (Mid(ps_CodeLine, li_Pos, 1) = Chr(34)) Then
            lb_Inquote = (Not lb_Inquote)
        ElseIf (Mid(ps_CodeLine, li_Pos, Len(ps_String)) = ps_String) Then
            If (Not lb_Inquote) Then
                StringIsCode = li_Pos
                Exit Function
            End If
        End If
    Next li_Pos
End Function

' Method to get the position of the beginning of the comment
Public Function CommentStart(ByVal ps_CodeLine As String) As Integer
    Dim lb_Inquote                      As Boolean
    Dim li_Pos                          As Long
    Dim ls_Char                         As String
    
    lb_Inquote = False
    CommentStart = 0
    
    For li_Pos = 1 To Len(ps_CodeLine)
        ls_Char = Mid(ps_CodeLine, li_Pos, 1)
        
        If (ls_Char = Chr(34)) Then
            lb_Inquote = (Not lb_Inquote)
        ElseIf (ls_Char = "'") Then
            If (Not lb_Inquote) Then
                CommentStart = li_Pos
                Exit Function
            End If
        End If
    Next li_Pos
End Function

Public Sub ReplaceAllRegPattern(ByRef ps_MethodLine As String, ByVal ps_Pattern As String, ByVal ps_Replace As String)
    reg.Pattern = ps_Pattern
    
    While (reg.test(ps_MethodLine))
        ps_MethodLine = Trim(reg.Replace(ps_MethodLine, ps_Replace))
    Wend
End Sub

' Method to get all macros called by objects in CurrentWorkbook
Public Sub ListCalledMacros()
    Dim lo_Shape    As Shape
    Dim lw_Ws       As Worksheet
    
    Set calledMacros = New Scripting.Dictionary
    
    For Each lw_Ws In currentWorkbook.Sheets
        DoEvents
        
        For Each lo_Shape In lw_Ws.Shapes
            On Error Resume Next
            calledMacros.Add UCase(Mid(lo_Shape.OnAction, InStr(lo_Shape.OnAction, "!") + 1)), vbNullString
        Next lo_Shape
    Next lw_Ws
    
    On Error Resume Next
    calledMacros.Remove vbNullString
End Sub

Public Function IsStandardMethod(ByRef po_VBC As VBComponent, ByVal ps_MethodName As String) As Boolean
    Dim lo_Usf              As Object
    Dim lv_Each             As Variant
    
    IsStandardMethod = False
    ps_MethodName = UCase(ps_MethodName)
    
    If (po_VBC.Type = 3) Then
        Set lo_Usf = po_VBC
        
        If (Not lo_Usf.Designer Is Nothing) Then
            For Each lv_Each In lo_Usf.Designer.Controls
                If (Left(ps_MethodName, Len(lv_Each.Name)) = UCase(lv_Each.Name)) Then
                    IsStandardMethod = True
                    Exit Function
                End If
            Next lv_Each
        End If
        
        ' Check if method is a standard userform method
        IsStandardMethod = gdic_UsfMethods.Exists(ps_MethodName)
    ElseIf (po_VBC.Type = 100) Then
        If (currentWorkbook.CodeName = po_VBC.Name) Then ' In the "Workbook" Document
            ' Check if method is a standard workbook method
            IsStandardMethod = gdic_WbkMethods.Exists(ps_MethodName)
        Else ' Check if method is a standard worksheet method
            IsStandardMethod = gdic_WsMethods.Exists(ps_MethodName)
        End If
    End If
End Function
