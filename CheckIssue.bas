Attribute VB_Name = "CheckIssue"
Option Explicit
' Developped by Hugues DUMONT

Public Sub SearchSquareBrackets(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    If (InStr(ps_CodeLine, "[") > 0) Then
        If (StringIsCode(ps_CodeLine, "[") > 0) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, SQUARE_BRACKETS, NO_SQUARE_BRACKETS, MOYENNE, vbNullString, vbNullString)
        End If
    End If
End Sub

Public Sub SearchImplicitWbk(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "(^|[^\.|\w])(Sheets|Worksheets)($|\.|\(| )"
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = " As (Worksheet|Sheet)(s)?"
        
        If (Not reg.test(ps_CodeLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, IMPLICIT_WORKBOOK, EXPLICIT_WORKBOOK, HAUTE)
        End If
    End If
End Sub

Public Sub SearchImplicitWs(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "(^|[^\.|\w])(Cells|Range)($|\.|\(| )"
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = " As Range"
        
        If (Not reg.test(ps_CodeLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, IMPLICIT_WORKSHEET, EXPLICIT_WORKSHEET, HAUTE)
        End If
    End If
End Sub

Public Sub SearchDebug(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "(^|( )+)(Debug\.(Print|Assert))"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, DEBUG_IN, DEBUG_REMOVE, MOYENNE)
    End If
End Sub

Public Sub SearchResume(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "^((Resume)(( )+.*|)$)"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, RESUME_ISSUE, RESUME_SOL, HAUTE)
    End If
End Sub

Public Sub SearchEnd(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "^((End)($|( )+'))"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, END_ISSUE, END_SOL, HAUTE)
    End If
End Sub

Public Sub SearchGoTo(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "(^| )GoTo "
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = "On Error GoTo "
        
        If (Not reg.test(ps_CodeLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, GOTO_ISSUE, GOTO_SOL, HAUTE)
        End If
    End If
End Sub

Public Sub AnalyseBoolean(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    Dim li_ThenPos  As Integer
    
    reg.Pattern = "((\(| )Not )+((\(| )*(True|False)($|(\)| )+))"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, NEGATED_BOOL, NEGATION_BOOL, HAUTE)
    End If
    
    ps_CodeLine = " " & ps_CodeLine
    
    reg.Pattern = "^( If )"
    
    If (reg.test(ps_CodeLine)) Then
        If (Right(ps_CodeLine, 5) <> " Then") Then
            li_ThenPos = StringIsCode(ps_CodeLine, " Then")
            
            reg.Pattern = "(\w+ (Or|And|Xor) )(True|False)|(True|False)( (Or|And|Xor) \w+)|Not (\()*(True|False)(\))*|(= (.* = (True|False)))|(= ((True|False) = .*))|(\((\()+)(True|False)(\)(\)+))"
            
            If (reg.test(Mid(ps_CodeLine, li_ThenPos + 5))) Then
                Call AddIssue(ps_VbcName, pl_CodeLine, CONDITIONAL_BOOL, SIMPLE_BOOL, MOYENNE)
            Else
                reg.Pattern = "(\w+ (Or|And|Xor) )(True|False)|(True|False)( (Or|And|Xor) \w+)|Not (\()*(True|False)(\))*|(.* = (True|False))|(= ((True|False) = .*))|(\((\()+)(True|False)(\)(\)+))"
                
                If (reg.test(Left(ps_CodeLine, li_ThenPos - 1))) Then
                    Call AddIssue(ps_VbcName, pl_CodeLine, CONDITIONAL_BOOL, SIMPLE_BOOL, MOYENNE)
                End If
            End If
        Else
            reg.Pattern = "(\w+ (Or|And|Xor) )(True|False)|(True|False)( (Or|And|Xor) \w+)|Not (\()*(True|False)(\))*|(.* = (True|False))|(= ((True|False) = .*))|(\((\()+)(True|False)(\)(\)+))"
            
            If (reg.test(ps_CodeLine)) Then
                Call AddIssue(ps_VbcName, pl_CodeLine, CONDITIONAL_BOOL, SIMPLE_BOOL, MOYENNE)
            End If
        End If
    Else
        reg.Pattern = "^( (ElseIf|Do While|Do Until|Loop While|Loop Until|While|Select Case|Case) )"
        
        If (reg.test(ps_CodeLine)) Then
            reg.Pattern = "(\w+ (Or|And|Xor) )(True|False)|(True|False)( (Or|And|Xor) \w+)|Not (\()*(True|False)(\))*|(.* = (True|False))|(= ((True|False) = .*))|(\((\()+)(True|False)(\)(\)+))"
            
            If (reg.test(ps_CodeLine)) Then
                Call AddIssue(ps_VbcName, pl_CodeLine, CONDITIONAL_BOOL, SIMPLE_BOOL, MOYENNE)
            End If
        Else
            reg.Pattern = "(\w+ (Or|And|Xor) )(True|False)|(True|False)( (Or|And|Xor) \w+)|Not (\()*(True|False)(\))*|(= (.* = (True|False)))|(= ((True|False) = .*))|(\((\()+)(True|False)(\)(\)+))"
            
            If (reg.test(ps_CodeLine)) Then ' test ((False))
                Call AddIssue(ps_VbcName, pl_CodeLine, CONDITIONAL_BOOL, SIMPLE_BOOL, MOYENNE)
            End If
        End If
    End If
End Sub

Public Sub SearchStringConcat(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "(" & Chr(34) & ".*" & Chr(34) & " \+)|( \+ " & Chr(34) & ".*" & Chr(34) & ")"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, CONCAT_PLUS, CONCAT_AND, HAUTE)
    End If
End Sub
