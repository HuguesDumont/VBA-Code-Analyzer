Attribute VB_Name = "CheckReadibility"
Option Explicit

' Method to correct multiple instructions in a single line (separated with colon (:)) into multiple lines (one instruction by line)
Public Sub SearchInLine(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    Dim li_ColonPos                     As Integer
    
    If (InStr(ps_CodeLine, ": ") > 0) Then
        li_ColonPos = StringIsCode(ps_CodeLine, ": ")
        
        If (li_ColonPos > 0) Then ' Multiple instructions in a single line or etiquette
            Call AddIssue(ps_VbcName, pl_CodeLine, INLINE, MULTILINE, HAUTE)
        End If
    End If
End Sub

Public Sub SearchTODO(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "('( )*TODO)($|( ))"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, TODO_ISSUE, TODO_SOL, BASSE)
    End If
End Sub

Public Sub SearchBoolParen(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "((^| |#)(((While|Until|Loop Until|Loop While) .* )|((If|ElseIf) .*(\=|\>|\<|\<\>|\>\=|\<\=|And|Or|Xor|Not) .* (Then( |$)))))"
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = "((^| |#)(((While|Until|Loop Until|Loop While) \(.* .*\))|((If|ElseIf) \(.*(\=|\>|\<|\<\>|\>\=|\<\=|And|Or|Xor|Not) .*(\) Then( |$)))))"
        
        If (Not reg.test(ps_CodeLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, PARENTHESIS_MISSING, PARENTHESIS_SOL, MOYENNE)
        End If
    End If
End Sub

Public Sub ValidConstantName(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_ConstName As String)
    reg.Pattern = "^([A-Z]{2}[A-Z_0-9]*[A-Z0-9]+)$"
    
    If (Not reg.test(ps_ConstName)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, CONST_NAMING_ISSUE, CONST_NAMING, BASSE, ps_ConstName)
    End If
End Sub

Public Sub ValidVarName(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_ls_VarName As String)
    reg.Pattern = "^([a-z][a-zA-Z_0-9]*[a-zA-Z0-9]+)$"
    
    If (Not reg.test(ps_ls_VarName)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, VAR_NAMING_ISSUE, VAR_NAMING, BASSE, ps_ls_VarName)
    End If
End Sub

Public Sub ValidMethodName(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_MethodName As String)
    reg.Pattern = "^([A-Z][a-zA-Z_0-9]*[a-zA-Z0-9]+)$"
    
    If (Not reg.test(ps_MethodName)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, METHOD_NAMING_ISSUE, METHOD_NAMING, BASSE, ps_MethodName)
    End If
End Sub

Public Sub ValidNameLength(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    If (Len(ps_CodeLine) > 20) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, NAME_LENGTH_ISSUE, NAME_LENGTH_SOL, BASSE, ps_CodeLine)
    End If
End Sub
