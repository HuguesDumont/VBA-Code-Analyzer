Attribute VB_Name = "CheckMissing"
Option Explicit

' Method to check for missing scope at module level
Public Sub MissingModuleScope(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "^((Dim|Const) )"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, IMPLICIT_SCOPE, EXPLICIT_PRIVATE, MOYENNE)
    Else
        reg.Pattern = "^((Property|Enum|Declare|Type|Static|Event|WithEvents) )"
        
        If (reg.test(ps_CodeLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, IMPLICIT_SCOPE, EXPLICIT_PUBLIC, MOYENNE)
        End If
    End If
End Sub

' Method to check for missing scope at method level
Public Sub MissingMethodScope(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "^((Sub|Function|Property|Enum|Type|Declare|(Static (Function|Sub|Property))|Event|WithEvents) )"
    
    If (reg.test(ps_CodeLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, IMPLICIT_SCOPE, EXPLICIT_PUBLIC, MOYENNE)
    End If
End Sub

' Method to check if "Option Explicit" is present in module
Public Sub MissingOption(ByVal ps_VbcName As String, ByVal pl_FoundOption As Long)
    If (pl_FoundOption = 0) Then
        Call AddIssue(ps_VbcName, 1, MISSING_OPTION, ADD_OPTION, HAUTE, vbNullString)
    ElseIf (pl_FoundOption > 1) Then
        Call AddIssue(ps_VbcName, pl_FoundOption, OFFSET_OPTION, MOVE_OPTION, BASSE)
    End If
End Sub

' Method to check for missing type for variables and constants
Public Sub MissingModuleType(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    Dim lb_IsConst      As Boolean
    Dim ls_VarList()    As String
    Dim lv_Each         As Variant
    
    reg.Pattern = "^(( )*(Public|Private|Const|Global|Dim)( )(Const |(\w+)($|,| ( )*As \w+)))"
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = "^(( )*((Public|Private|Global|Dim)( ))?(Const ))"
        lb_IsConst = reg.test(ps_CodeLine)
        
        reg.Pattern = "( )*(Public|Private|Const|Global|Dim)( )(Const )?"
        ps_CodeLine = reg.Replace(ps_CodeLine, vbNullString)
        
        Call FormatLineAs(ps_CodeLine)
        
        If (InStr(ps_CodeLine, ",") > 0) Then
            ls_VarList = Split(ps_CodeLine, ", ")
            
            For Each lv_Each In ls_VarList
                If (checkNameVar And (Not lb_IsConst)) Then
                    Call ValidVarName(ps_VbcName, pl_CodeLine, GetVarName(lv_Each))
                ElseIf (checkNameConst And lb_IsConst) Then
                    Call ValidConstantName(ps_VbcName, pl_CodeLine, GetVarName(lv_Each))
                End If
                
                If (checkNameLength) Then
                    Call ValidNameLength(ps_VbcName, pl_CodeLine, GetVarName(lv_Each))
                End If
                       
                If (checkType) Then
                    Call CheckMissingVar(lv_Each, ps_VbcName, pl_CodeLine, lb_IsConst)
                End If
            Next lv_Each
        Else
            If (checkNameVar And (Not lb_IsConst)) Then
                Call ValidVarName(ps_VbcName, pl_CodeLine, GetVarName(ps_CodeLine))
            ElseIf (checkNameConst And lb_IsConst) Then
                Call ValidConstantName(ps_VbcName, pl_CodeLine, GetVarName(ps_CodeLine))
            End If
            
            If (checkNameLength) Then
                Call ValidNameLength(ps_VbcName, pl_CodeLine, GetVarName(ps_CodeLine))
            End If
            
            If (checkType) Then
                Call CheckMissingVar(ps_CodeLine, ps_VbcName, pl_CodeLine, lb_IsConst)
            End If
        End If
    End If
End Sub

Public Sub MissingMethodType(ByRef po_VBC As VBComponent, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String, ByVal ps_MethodName As String)
    reg.Pattern = "^(( )*((Public|Friend|Private)( ))?(Static )?(Property|Function)( [GS]et)?( )(\w+)(\(.*\))( As \w+)?)"
    
    If (reg.test(ps_CodeLine) And (Not IsStandardMethod(po_VBC, ps_MethodName))) Then
        reg.Pattern = "^(( )*((Public|Friend|Private)( ))?(Static )?(Property|Function)( [GS]et)?( )(\w+)(\(.*\))( As \w+))"
        
        If (Not reg.test(ps_CodeLine)) Then
            Call AddIssue(po_VBC.Name, pl_CodeLine, IMPLICIT_RETURN, EXPLICIT_RETURN, HAUTE, ps_MethodName)
        End If
    End If
End Sub

Public Sub MissingParamType(ByRef po_VBC As VBComponent, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String, ByVal ps_MethodName As String)
    Dim ll_CharAt       As Long
    Dim ls_VarList()    As String
    Dim lv_Param        As Variant
    
    If (Not IsStandardMethod(po_VBC, ps_MethodName)) Then
        reg.Pattern = "( )*((Public|Friend|Private)( ))?(Static )?(Property|Sub|Function)( [GS]et)?( )"
        ps_CodeLine = reg.Replace(ps_CodeLine, vbNullString)
        ps_CodeLine = Mid(ps_CodeLine, InStr(ps_CodeLine, "(") + 1)
        
        If (Len(Trim(ps_CodeLine)) > 0) Then
            ll_CharAt = Len(ps_CodeLine)
            
            While ((Mid(ps_CodeLine, ll_CharAt, 1) <> ")"))
                ll_CharAt = ll_CharAt - 1
            Wend
            
            ps_CodeLine = Left(ps_CodeLine, ll_CharAt - 1)
            
            Call FormatLineAs(ps_CodeLine)
            
            If (Len(Trim(ps_CodeLine)) > 0) Then
                If (InStr(ps_CodeLine, ",") > 0) Then
                    ls_VarList = Split(ps_CodeLine, ", ")
                    
                    For Each lv_Param In ls_VarList
                        If (checkNameVar) Then
                            Call ValidVarName(po_VBC.Name, pl_CodeLine, GetVarName(lv_Param))
                        End If
                        
                        If (checkNameLength) Then
                            Call ValidNameLength(po_VBC.Name, pl_CodeLine, GetVarName(lv_Param))
                        End If
                        
                        If (checkByref Or checkType) Then
                            Call CheckMissingParam(lv_Param, po_VBC.Name, pl_CodeLine)
                        End If
                    Next lv_Param
                Else
                    If (checkNameVar) Then
                        Call ValidVarName(po_VBC.Name, pl_CodeLine, GetVarName(ps_CodeLine))
                    End If
                    
                    If (checkNameLength) Then
                        Call ValidNameLength(po_VBC.Name, pl_CodeLine, GetVarName(ps_CodeLine))
                    End If
                        
                    If (checkByref Or checkType) Then
                        Call CheckMissingParam(ps_CodeLine, po_VBC.Name, pl_CodeLine)
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub MissingVarType(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    Dim lb_IsConst      As Boolean
    Dim ls_VarList()    As String
    Dim lv_Each         As Variant
    
    reg.Pattern = "^(( )*(Dim|Const)( )(\w+)( ( )*As \w+)?)"
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = "^(( )*(Const ))"
        lb_IsConst = reg.test(ps_CodeLine)

        reg.Pattern = "(Dim|Const)( )"
        ps_CodeLine = Trim(reg.Replace(ps_CodeLine, vbNullString))
        
        Call FormatLineAs(ps_CodeLine)
        
        If (InStr(ps_CodeLine, ",") > 0) Then
            ls_VarList = Split(ps_CodeLine, ", ")
            
            For Each lv_Each In ls_VarList
                If (checkNameVar And (Not lb_IsConst)) Then
                    Call ValidVarName(ps_VbcName, pl_CodeLine, GetVarName(lv_Each))
                ElseIf (checkNameConst And lb_IsConst) Then
                    Call ValidConstantName(ps_VbcName, pl_CodeLine, GetVarName(lv_Each))
                End If
                
                If (checkNameLength) Then
                    Call ValidNameLength(ps_VbcName, pl_CodeLine, GetVarName(lv_Each))
                End If
                
                If (checkType) Then
                    Call CheckMissingVar(lv_Each, ps_VbcName, pl_CodeLine, lb_IsConst)
                End If
            Next lv_Each
        Else
            If (checkNameVar And (Not lb_IsConst)) Then
                Call ValidVarName(ps_VbcName, pl_CodeLine, GetVarName(ps_CodeLine))
            ElseIf (checkNameConst And lb_IsConst) Then
                Call ValidConstantName(ps_VbcName, pl_CodeLine, GetVarName(ps_CodeLine))
            End If
            
            If (checkNameLength) Then
                Call ValidNameLength(ps_VbcName, pl_CodeLine, GetVarName(ps_CodeLine))
            End If
            
            If (checkType) Then
                Call CheckMissingVar(ps_CodeLine, ps_VbcName, pl_CodeLine, lb_IsConst)
            End If
        End If
    End If
End Sub

' Method to check for Next with var missing
Public Sub SearchMissingNext(ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    reg.Pattern = "(^| )Next($| :)"
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = "Resume Next"
        
        If (Not reg.test(ps_CodeLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, MISSING_NEXT, ADD_NEXT, BASSE)
        End If
    End If
End Sub

Private Sub CheckMissingParam(ByVal ps_MethodLine As String, ByVal ps_VbcName As String, ByVal pl_CodeLine As Long)
    Dim ls_VarName As String
    
    ls_VarName = GetVarName(ps_MethodLine)
    
    If (checkByref) Then
        reg.Pattern = "^(ByRef|ByVal|ParamArray)( )"
        
        If (Not reg.test(ps_MethodLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, IMPLICIT_PASSING, EXPLICIT_PASSING, HAUTE, ls_VarName)
        End If
    End If
    
    If (checkType) Then
        reg.Pattern = " As "
        
        If (Not reg.test(ps_MethodLine)) Then
            Call AddIssue(ps_VbcName, pl_CodeLine, (IMPLICIT_TYPE & PARAMETER), (EXPLICIT_TYPE & PARAMETER), HAUTE, ls_VarName)
        End If
    End If
End Sub

Private Sub CheckMissingVar(ByVal ps_MethodLine As String, ByVal ps_VbcName As String, ByVal pl_CodeLine As Long, ByVal pb_IsConst As Boolean)
    reg.Pattern = " As "
    
    If (Not reg.test(ps_MethodLine)) Then
        Call AddIssue(ps_VbcName, pl_CodeLine, (IMPLICIT_TYPE & IIf(pb_IsConst, CONSTANTE, VARIABLE)), (EXPLICIT_TYPE & IIf(pb_IsConst, CONSTANTE, VARIABLE)), HAUTE, GetVarName(ps_MethodLine))
    End If
End Sub

Public Function GetVarName(ByVal ps_MethodLine As String) As String
    reg.Pattern = "^(ByRef|ByVal|ParamArray)( )"
    ps_MethodLine = Trim(ps_MethodLine)
    ps_MethodLine = reg.Replace(ps_MethodLine, vbNullString)
    GetVarName = Split(ps_MethodLine, " ")(0)
End Function
