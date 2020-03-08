Attribute VB_Name = "CheckUse"
Option Explicit

Public Sub FindUnusedModuleVar(ByRef po_VBC As VBComponent, ByRef pdic_Public As Scripting.Dictionary, ByRef pdic_Private As Scripting.Dictionary, _
        ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String, ByVal ps_MethodName As String)
    Dim lb_IsConst     As Boolean
    Dim ls_VarList()   As String
    Dim ls_VarName     As Variant

    reg.Pattern = "^(( )*(Private|Dim)( \w+)( ( )*As \w+)?)"
    
    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = "( )*(Private|Dim)( (Declare|Enum|Type|Property|Sub|Function|Implements) )"

        If (Not reg.test(ps_CodeLine)) Then
            reg.Pattern = "(Private|Dim)( Const )"
            
            If (reg.test(ps_CodeLine)) Then
                ps_CodeLine = reg.Replace(ps_CodeLine, vbNullString)
                lb_IsConst = True
            Else
                reg.Pattern = "(Private|Dim)( )"
                ps_CodeLine = reg.Replace(ps_CodeLine, vbNullString)
                lb_IsConst = False
            End If
            
            Call FormatDeclareLine(ps_CodeLine)
            
            If (InStr(ps_CodeLine, ",") > 0) Then
                If (checkMultiDim) Then
                    Call AddIssue(po_VBC.Name, pl_CodeLine, INLINE, MULTILINE, BASSE, , ps_MethodName)
                End If
                
                If (checkVarUse) Then
                    ls_VarList = Split(ps_CodeLine, ", ")
                    
                    For Each ls_VarName In ls_VarList
                        pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ls_VarName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ls_VarName, po_VBC.Name, pl_CodeLine, False, lb_IsConst, ps_MethodName)
                    Next ls_VarName
                End If
            ElseIf (checkVarUse) Then
                pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_CodeLine, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_CodeLine, po_VBC.Name, pl_CodeLine, False, lb_IsConst, ps_MethodName)
            End If
        End If
    Else
        reg.Pattern = "^(( )*(Global|Public)( \w+)( ( )*As \w+)?)"
        
        If (reg.test(ps_CodeLine)) Then
            reg.Pattern = "(( )*(Public )(Declare|Enum|Type|Property|Sub|Function|Implements) )"
            
            If (Not reg.test(ps_CodeLine)) Then
                reg.Pattern = "(Global|Public)( Const )"
                
                If (reg.test(ps_CodeLine)) Then
                    ps_CodeLine = reg.Replace(ps_CodeLine, vbNullString)
                    lb_IsConst = True
                Else
                    reg.Pattern = "(Global|Public)( )"
                    ps_CodeLine = reg.Replace(ps_CodeLine, vbNullString)
                    lb_IsConst = False
                End If
                
                Call FormatDeclareLine(ps_CodeLine)
                
                If (InStr(ps_CodeLine, ",") > 0) Then
                    If (checkMultiDim) Then
                        Call AddIssue(po_VBC.Name, pl_CodeLine, INLINE, MULTILINE, BASSE, , ps_MethodName)
                    End If
                    
                    If (checkVarUse Or checkNameVar Or checkNameConst Or checkNameLength) Then
                        ls_VarList = Split(ps_CodeLine, ", ")
                        
                        For Each ls_VarName In ls_VarList
                            pdic_Public.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ls_VarName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ls_VarName, po_VBC.Name, pl_CodeLine, False, lb_IsConst, ps_MethodName)
                        Next ls_VarName
                    End If
                ElseIf (checkVarUse) Then
                    pdic_Public.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_CodeLine, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_CodeLine, po_VBC.Name, pl_CodeLine, False, lb_IsConst, ps_MethodName)
                End If
            End If
        End If
    End If
End Sub

Public Sub FindUnusedMethodVar(ByRef po_VBC As VBComponent, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String)
    Dim lb_IsConst As Boolean
    Dim ls_VarList() As String
    Dim ls_VarName As Variant
    
    reg.Pattern = "^(( )*(Dim|Const)( )(\w+)( ( )*As \w+)?)"

    If (reg.test(ps_CodeLine)) Then
        reg.Pattern = "^(( )*(Const ))"
        lb_IsConst = reg.test(ps_CodeLine)

        reg.Pattern = "(Dim|Const)( )"
        ps_CodeLine = Trim(reg.Replace(ps_CodeLine, vbNullString))

        Call FormatDeclareLine(ps_CodeLine)
        
        If (InStr(ps_CodeLine, ",") > 0) Then
            If (checkMultiDim) Then
                Call AddIssue(po_VBC.Name, pl_CodeLine, INLINE, MULTILINE, BASSE)
            End If
            
            If (checkVarUse) Then
                ls_VarList = Split(ps_CodeLine, ", ")
                
                For Each ls_VarName In ls_VarList
                    If (Not SearchMethodUse(po_VBC.CodeModule, ls_VarName, pl_CodeLine, lb_IsConst)) Then
                        Call AddIssue(po_VBC.Name, pl_CodeLine, (IIf(lb_IsConst, CONSTANTE, VARIABLE) & NOTUSED), (DELETE & IIf(lb_IsConst, CONSTANTE, VARIABLE)), HAUTE, ls_VarName)
                    End If
                Next ls_VarName
            End If
        ElseIf (checkVarUse) Then
            If (Not SearchMethodUse(po_VBC.CodeModule, ps_CodeLine, pl_CodeLine, lb_IsConst)) Then
                Call AddIssue(po_VBC.Name, pl_CodeLine, (IIf(lb_IsConst, CONSTANTE, VARIABLE) & NOTUSED), (DELETE & IIf(lb_IsConst, CONSTANTE, VARIABLE)), HAUTE, ps_CodeLine)
            End If
        End If
    End If
End Sub

Public Sub FindUnusedParams(ByRef po_VBC As VBComponent, ByVal pl_CodeLine As Long, ByVal ps_CodeLine As String, ByVal ps_MethodName As String)
    Dim lb_Trouve       As Boolean
    Dim lo_CMod         As CodeModule
    Dim ll_CharAt       As Long
    Dim ls_VarList()    As String
    Dim ls_VarName      As Variant
    
    Set lo_CMod = po_VBC.CodeModule
    
    reg.Pattern = "^(( )*((Public|Friend|Private)( ))?(Static )?(Property|Sub|Function)( [GS]et)?( )(\w+)(\(.*\))( As \w+)?)"
    
    lb_Trouve = False
    
    If (reg.test(ps_CodeLine) And (Not IsStandardMethod(po_VBC, ps_MethodName))) Then
        reg.Pattern = "( )*((Public|Friend|Private)( ))?(Static )?(Property|Sub|Function)( [GS]et)?( )"
        ps_CodeLine = reg.Replace(ps_CodeLine, vbNullString)
        
        ps_CodeLine = Mid(ps_CodeLine, InStr(ps_CodeLine, "(") + 1)
        
        If (Len(ps_CodeLine) > 0) Then
            ll_CharAt = Len(ps_CodeLine)
            
            While ((Mid(ps_CodeLine, ll_CharAt, 1) <> ")"))
                ll_CharAt = ll_CharAt - 1
            Wend
            
            ps_CodeLine = Left(ps_CodeLine, ll_CharAt - 1)
            
            Call FormatDeclareLine(ps_CodeLine)
            ps_CodeLine = Trim(ps_CodeLine)
            
            If (Len(ps_CodeLine) > 0) Then
                If (InStr(ps_CodeLine, ",") > 0) Then
                    ls_VarList = Split(ps_CodeLine, ", ")
                    
                    For Each ls_VarName In ls_VarList
                        If (Not SearchMethodUse(lo_CMod, ls_VarName, pl_CodeLine, True)) Then
                            Call AddIssue(po_VBC.Name, pl_CodeLine, (PARAMETER & NOTUSED), (DELETE & PARAMETER), HAUTE, ls_VarName, ps_MethodName)
                        End If
                    Next ls_VarName
                ElseIf (Not SearchMethodUse(lo_CMod, ps_CodeLine, pl_CodeLine, True)) Then
                    Call AddIssue(po_VBC.Name, pl_CodeLine, (PARAMETER & NOTUSED), (DELETE & PARAMETER), HAUTE, ps_CodeLine, ps_MethodName)
                End If
            End If
        End If
    End If
End Sub

Public Sub FindUnusedMethods(ByRef po_VBC As VBComponent, ByRef pdic_Public As Scripting.Dictionary, ByRef pdic_Private As Scripting.Dictionary, ByVal ps_MethodName As String)
    Dim lb_Trouve               As Boolean
    Dim lo_CMod                 As CodeModule
    Dim li_Start                As Integer
    Dim ll_DeclareLine          As Long
    Dim ls_MethodDeclaration    As String
    Dim ls_UMethodName          As String
    Dim lv_Each                 As Variant
    Dim lw_Ws                   As Worksheet
    
    Set lo_CMod = po_VBC.CodeModule
    
    On Error Resume Next
    
    li_Start = 0
    lb_Trouve = False
    
    While (li_Start < 4)
        ll_DeclareLine = lo_CMod.ProcBodyLine(ps_MethodName, li_Start)
        
        If (Err.Number = 0) Then
            ls_MethodDeclaration = lo_CMod.Lines(ll_DeclareLine, 1)
            li_Start = 5
        Else
            On Error GoTo 0
            On Error Resume Next
            li_Start = li_Start + 1
        End If
    Wend
    
    If (li_Start = 4) Then
        Exit Sub
    End If
    
    On Error GoTo 0
    
    reg.Pattern = "^(Private )"
    ls_UMethodName = UCase(ps_MethodName)
    
    If (reg.test(ls_MethodDeclaration)) Then
        If (po_VBC.Type = 2) Then ' Class Module
            If ((ls_UMethodName <> "CLASS_TERMINATE") And (ls_UMethodName <> "CLASS_INITIALIZE")) Then
                pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, False, ps_MethodName)
            End If
        ElseIf (po_VBC.Type = 3) Then ' Userform
            If (Not IsStandardMethod(po_VBC, ls_UMethodName)) Then
                pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, False, ps_MethodName)
            End If
        ElseIf (po_VBC.Type = 100) Then ' Document
            If (currentWorkbook.CodeName = po_VBC.Name) Then ' Workbook Document
                If (Not IsStandardMethod(po_VBC, ls_UMethodName)) Then
                    pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, False, ps_MethodName)
                End If
            Else ' Worksheet Document
                On Error GoTo ContinueProcess
                Set lw_Ws = currentWorkbook.Sheets(po_VBC.Properties("index"))
                
                If (InStr(ps_MethodName, "_") > 0) Then
                    If (Not IsStandardMethod(po_VBC, ls_UMethodName)) Then
                        ' Check for each activex object if name matches
                        For Each lv_Each In lw_Ws.OLEObjects
                            If (Left(ls_UMethodName, Len(lv_Each.Name) + 1) = UCase(lv_Each.Name) & "_") Then
                                lb_Trouve = True
                                Exit For
                            End If
                        Next lv_Each
                        
ContinueProcess:
                        If (Not lb_Trouve) Then
                            On Error GoTo 0
                            pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, False, ps_MethodName)
                        End If
                    End If
                Else
                    pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, False, ps_MethodName)
                End If
            End If
        Else
            pdic_Private.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, False, ps_MethodName)
        End If
    ElseIf ((po_VBC.Type = 2) Or (po_VBC.Type = 3)) Then ' Class Module or Userform
        pdic_Public.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, True, ps_MethodName)
    ElseIf (Not calledMacros.Exists(ls_UMethodName)) Then ' Check if method is a called macro
        pdic_Public.Add currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, Array(currentWorkbook.Name & "/" & po_VBC.Name & "/" & ps_MethodName, po_VBC.Name, ll_DeclareLine, True, False, ps_MethodName)
    End If
End Sub

Public Sub SearchEmptyMethod(ByRef po_VBC As VBComponent, ByVal ps_MethodName As String)
    Dim lb_IsNotEmptyMethod As Boolean
    Dim lb_Trouve           As Boolean
    Dim lo_CMod             As CodeModule
    Dim li_Start            As Integer
    Dim ll_DeclareLine      As Long
    Dim ll_CodeLine         As Long
    Dim ls_CodeLine         As String
    
    Set lo_CMod = po_VBC.CodeModule
    
    On Error Resume Next
    
    li_Start = 0
    lb_Trouve = False
    
    While (li_Start < 4)
        ll_CodeLine = lo_CMod.ProcBodyLine(ps_MethodName, li_Start)
        ll_DeclareLine = ll_CodeLine
        
        If (Err.Number = 0) Then
            ls_CodeLine = lo_CMod.Lines(ll_CodeLine, 1)
            li_Start = 5
        Else
            On Error GoTo 0
            On Error Resume Next
            li_Start = li_Start + 1
        End If
    Wend
    
    If (li_Start = 4) Then
        Exit Sub
    End If
    
    On Error GoTo 0
    
    While (Right(ls_CodeLine, 2) = " _")
        ll_CodeLine = ll_CodeLine + 1
        
        If (Right(ls_CodeLine, 3) = "& _") Then
            ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 3) & " " & Trim(lo_CMod.Lines(ll_CodeLine, 1))
        Else
            ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 1) & Trim(lo_CMod.Lines(ll_CodeLine, 1))
        End If
    Wend
    
    lb_IsNotEmptyMethod = False
    reg.Pattern = "^(End (Sub|Function|Property))"
    
    While ((ll_CodeLine < lo_CMod.CountOfLines) And (Not reg.test(ls_CodeLine)) And (Not lb_IsNotEmptyMethod))
        ll_CodeLine = ll_CodeLine + 1
        ls_CodeLine = Trim(lo_CMod.Lines(ll_CodeLine, 1))
        
        If (Len(ls_CodeLine) > 0) Then
            While (Right(ls_CodeLine, 2) = " _")
                ll_CodeLine = ll_CodeLine + 1
                            
                If (Right(ls_CodeLine, 3) = "& _") Then
                    ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 3) & " " & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                Else
                    ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 1) & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                End If
            Wend
        End If
        
        If ((ls_CodeLine <> vbNullString) And (Left(ls_CodeLine, 1) <> "'") And (Not reg.test(ls_CodeLine))) Then
            lb_IsNotEmptyMethod = True
        End If
    Wend

    If (Not lb_IsNotEmptyMethod) Then
        Call AddIssue(po_VBC.Name, ll_DeclareLine, EMPTY_METHOD, (DELETE & METHOD), MOYENNE, ps_MethodName, ps_MethodName)
    End If
End Sub

' Method to check to format a codeLine to get only the var/const/param part
Private Sub FormatDeclareLine(ByRef ps_MethodLine As String)
    ps_MethodLine = " " & ps_MethodLine
    reg.Pattern = "( )(ByRef|ByVal)( )"
    
    While (reg.test(ps_MethodLine))
        ps_MethodLine = Trim(reg.Replace(ps_MethodLine, " "))
    Wend
    
    ps_MethodLine = Trim(ps_MethodLine)

    Call ReplaceAllRegPattern(ps_MethodLine, "As String \* \d+", vbNullString)  ' Toutes les chaines de longueur fixe (As String * xxx)
    Call ReplaceAllRegPattern(ps_MethodLine, "( As New )(\w+)(\.\w+)*", vbNullString)  ' Tous les " As New sss" où sss est le nom d'un objet/type/classe
    Call ReplaceAllRegPattern(ps_MethodLine, "((\.)(\w+))+", vbNullString)  ' Toute chaine précédée d'un . (.sss)
    ps_MethodLine = " " & ps_MethodLine
    Call ReplaceAllRegPattern(ps_MethodLine, " Optional ", " ")
    ps_MethodLine = " " & Trim(ps_MethodLine)
    Call ReplaceAllRegPattern(ps_MethodLine, " ParamArray ", " ")
    ps_MethodLine = Trim(ps_MethodLine)
    Call ReplaceAllRegPattern(ps_MethodLine, "( ( )*As )(\w+)", vbNullString)  ' Tous les " As sss" restants
    Call ReplaceAllRegPattern(ps_MethodLine, " \+ ", vbNullString)  ' symbole +
    Call ReplaceAllRegPattern(ps_MethodLine, ("((" & Chr(34) & Chr(34) & ")|(" & Chr(34) & "[^" & Chr(34) & "]+" & Chr(34) & "))"), vbNullString) ' Texte entre quote ("sss")
    Call ReplaceAllRegPattern(ps_MethodLine, "( = (\+|-)?(\d+)((\.)?(\d+))?(E(\+|-)\d+)?#?)", vbNullString) ' Nombres (décimaux, entiers, puissances)
    Call ReplaceAllRegPattern(ps_MethodLine, "(\+|-)\d+", vbNullString) ' Nombres restants avec symbôle
    Call ReplaceAllRegPattern(ps_MethodLine, "( = \w+)", vbNullString) ' affectations de constantes
    Call ReplaceAllRegPattern(ps_MethodLine, " =( )*", vbNullString) ' affectations de constantes vides (dues aux suppressions en amont)
    Call ReplaceAllRegPattern(ps_MethodLine, "\(((\w+)*(\d*( To \d+)?)((, ?((\w+)*\d*( To \d+)?))*))\)", vbNullString) ' Déclarations de tableaux
    Call ReplaceAllRegPattern(ps_MethodLine, "(& .* &)", vbNullString) ' Concaténations de chaines de caractère
    Call ReplaceAllRegPattern(ps_MethodLine, "& \w+", vbNullString) ' Concaténations restantes
    Call ReplaceAllRegPattern(ps_MethodLine, " &", vbNullString) ' Symboles de concaténation restants
    Call ReplaceAllRegPattern(ps_MethodLine, "\(|\)", vbNullString) ' Parenthèses restantes
    Call ReplaceAllRegPattern(ps_MethodLine, "(  )+", " ") ' Espaces multiples
    Call ReplaceAllRegPattern(ps_MethodLine, "'.*", " ") ' Commentaires
    Call ReplaceAllRegPattern(ps_MethodLine, "\:.*", vbNullString) ' Instructions après :
End Sub

' Method to look for a var/const/param use in a method
Private Function SearchMethodUse(ByRef po_CMod As CodeModule, ByVal ps_ToSearch As String, ByVal pl_CodeLine As Long, Optional ByVal pb_CanEnd As Boolean = False) As Boolean
    Dim ll_CommentPos As Long
    Dim ls_MethodName As String
    Dim ls_MethodLine As String
    
    SearchMethodUse = False
    ls_MethodName = po_CMod.ProcOfLine(pl_CodeLine, vbext_pk_Proc)
    pl_CodeLine = pl_CodeLine + 1
    
    If (pb_CanEnd) Then
        reg.Pattern = "(\(|\.| |:=|^)" & ps_ToSearch & "((((\)|\.|\,|\())?)$|(((\)|\.|\,|\())?))"
    Else
        reg.Pattern = "(\(|\.| |^)" & ps_ToSearch & "(\)|\.|\,|\(| )"
    End If
    
    While ((Not SearchMethodUse) And (po_CMod.ProcOfLine(pl_CodeLine, vbext_pk_Proc) = ls_MethodName) And (pl_CodeLine < po_CMod.CountOfLines))
        ls_MethodLine = po_CMod.Lines(pl_CodeLine, 1)
        ll_CommentPos = CommentStart(ls_MethodLine) - 1
        
        If (ll_CommentPos > 0) Then
            SearchMethodUse = reg.test(Trim(Left(ls_MethodLine, ll_CommentPos)))
        Else
            SearchMethodUse = reg.test(Trim(ls_MethodLine))
        End If
        
        pl_CodeLine = pl_CodeLine + 1
    Wend
End Function

Public Sub SearchPublicUse(ByRef pdic_ToSearch As Scripting.Dictionary)
    Dim lo_CMod         As CodeModule
    Dim li_CommentPos   As Integer
    Dim ll_CodeLine     As Long
    Dim ls_CodeLine     As String
    Dim lv_Each         As Variant
    Dim lo_VBC          As VBComponent
    
    For Each lo_VBC In currentWorkbook.VBProject.VBComponents
        Set lo_CMod = lo_VBC.CodeModule
        
        For ll_CodeLine = 1 To lo_CMod.CountOfDeclarationLines
            ls_CodeLine = lo_CMod.Lines(ll_CodeLine, 1)
            
            If (Len(Trim(ls_CodeLine) > 0)) Then
                While (Right(ls_CodeLine, 2) = " _")
                    ll_CodeLine = ll_CodeLine + 1
                    
                    If (Right(ls_CodeLine, 3) = "& _") Then
                        ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 3) & " " & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                    Else
                        ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 1) & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                    End If
                Wend
                
                ls_CodeLine = Trim(ls_CodeLine)
                
                If (Left(ls_CodeLine, 1) <> "'") Then
                    li_CommentPos = CommentStart(ls_CodeLine) - 1
                    
                    If (li_CommentPos > 0) Then
                        ls_CodeLine = Trim(Left(ls_CodeLine, li_CommentPos))
                    End If
                    
                    For Each lv_Each In pdic_ToSearch.Items
                        If (Not lv_Each(3)) Then
                            If (lv_Each(4)) Then
                                reg.Pattern = "(=( .*) " & Split(lv_Each(0), "/")(2) & ")|(= " & Split(lv_Each(0), "/")(2) & "( |\)|$))|(\(.*" & Split(lv_Each(0), "/")(2) & ".*\))"
                                
                                If (reg.test(ls_CodeLine)) Then
                                    pdic_ToSearch.Remove (lv_Each(0))
                                End If
                            End If
                        End If
                    Next lv_Each
                End If
            End If
        Next ll_CodeLine
        
        For ll_CodeLine = lo_CMod.CountOfDeclarationLines + 1 To lo_CMod.CountOfLines - 1
            ls_CodeLine = lo_CMod.Lines(ll_CodeLine, 1)
            
            If (Len(Trim(ls_CodeLine)) > 0) Then
                While (Right(ls_CodeLine, 2) = " _")
                    ll_CodeLine = ll_CodeLine + 1
                    
                    If (Right(ls_CodeLine, 3) = "& _") Then
                        ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 3) & " " & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                    Else
                        ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 1) & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                    End If
                Wend
                
                ls_CodeLine = Trim(ls_CodeLine)
                
                If (Left(ls_CodeLine, 1) <> "'") Then
                    li_CommentPos = CommentStart(ls_CodeLine) - 1
                    
                    If (li_CommentPos > 0) Then
                        ls_CodeLine = Trim(Left(ls_CodeLine, li_CommentPos))
                    End If
                    
                    For Each lv_Each In pdic_ToSearch.Items
                        reg.Pattern = "(\(|\.| |:=|^)" & Split(lv_Each(0), "/")(2) & "((((\)|\.|\,|\())?)$|(((\)|\.|\,|\(| ))+))"
                        
                        If (lv_Each(3)) Then
                            If (lv_Each(4)) Then
                                If ((lv_Each(1) <> lo_VBC.Name) And ((lo_VBC.Type = 2) Or (lo_VBC.Type = 3))) Then
                                    reg.Pattern = "\." & Split(lv_Each(0), "/")(2) & "((((\)|\.|\,|\())?)$|(((\)|\.|\,|\(| ))+))"
                                End If
                            End If
                            
                            If (reg.test(ls_CodeLine) And ((lo_CMod.ProcOfLine(ll_CodeLine, vbext_pk_Proc) <> Split(lv_Each(0), "/")(2)) Or (lo_VBC.Name <> Split(lv_Each(0), "/")(1)))) Then
                                reg.Pattern = "^(( )*((Public|Friend|Private)( ))?(Static )?(Property|Sub|Function)( [GS]et)?( )(\w+)(\(.*\))( As \w+)?)"
                                
                                If (Not reg.test(ls_CodeLine)) Then
                                    pdic_ToSearch.Remove (lv_Each(0))
                                Else
                                    ' La méthode est surchargée car son nom est utilisé en tant que paramètre
                                    ' Gérer également le cas des Dim/Const/Redim dans la partie If
                                End If
                            End If
                        ElseIf (reg.test(ls_CodeLine)) Then
                            If (lv_Each(4)) Then
                                pdic_ToSearch.Remove (lv_Each(0))
                            Else
                                reg.Pattern = "^(( )*((Public|Friend|Private)( ))?(Static )?(Property|Sub|Function)( [GS]et)?( )(\w+)(\(.*\))( As \w+)?)"
                                
                                If (Not reg.test(ls_CodeLine)) Then
                                    pdic_ToSearch.Remove (lv_Each(0))
                                Else
                                    ' La variable est surchargée car son nom est utilisé en tant que paramètre
                                    ' Gérer également le cas des Dim/Const/Redim dans la partie If
                                End If
                            End If
                        End If
                        
                        If (pdic_ToSearch.Count = 0) Then
                            Exit Sub
                        End If
                    Next lv_Each
                End If
            End If
        Next ll_CodeLine
    Next lo_VBC
End Sub

Public Sub SearchLineUse(ByRef pdic_DicoSearch As Scripting.Dictionary, ByVal ps_TrimLine As String, ByVal ps_MethodName As String)
    Dim lv_Each     As Variant
    
    For Each lv_Each In pdic_DicoSearch.Items
        reg.Pattern = "^(( )*((Public|Friend|Private)( ))?(Static )?(Property|Sub|Function)( [GS]et)?( )(\w+)(\(.*\))( As \w+)?)"
        
        If (lv_Each(3)) Then
            If (Not reg.test(ps_TrimLine)) Then
                reg.Pattern = "(\(|\.| |:=|^)" & Split(lv_Each(0), "/")(2) & "((((\)|\.|\,|\())?)$|(((\)|\.|\,|\(| ))+))"
                
                If (reg.test(ps_TrimLine) And (ps_MethodName <> Split(lv_Each(0), "/")(2))) Then
                    pdic_DicoSearch.Remove (lv_Each(0))
                    ' Gérer également le cas des Dim/Const/Redim dans la partie If
                End If
            End If
        ElseIf (reg.test(ps_TrimLine)) Then
            If (lv_Each(4)) Then
                reg.Pattern = "(= (\w+(\.))*)" & Split(lv_Each(0), "/")(2) & "((((\)|\,|\())?)$|(((\)|\.|\,|\(| ))+))"
                
                If (reg.test(ps_TrimLine)) Then
                    pdic_DicoSearch.Remove (lv_Each(0))
                    ' Gérer également le cas des surcharges Dim/Const/Redim
                End If
            End If
        Else
            reg.Pattern = "(\(|\.| |:=|^)" & Split(lv_Each(0), "/")(2) & "((((\)|\.|\,|\())?)$|(((\)|\.|\,|\(| ))+))"
            
            If (reg.test(ps_TrimLine)) Then
                pdic_DicoSearch.Remove (lv_Each(0))
                ' Gérer également le cas des surcharges Dim/Const/Redim
            End If
        End If
        
        If (pdic_DicoSearch.Count = 0) Then
            Exit Sub
        End If
    Next lv_Each
End Sub
