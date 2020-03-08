VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CodeSmell 
   Caption         =   "Analyseur de code réalisé par @Hugues DUMONT"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   27915
   OleObjectBlob   =   "CodeSmell.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CodeSmell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As LongPtr) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As LongPtr) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
#Else
    Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
#End If

Private Const LOGPIXELSX As Long = 88 'Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72 'A point is defined as 1/72 inches

Private mb_UsfEvents       As Boolean
Private md_ratioHeight     As Double
Private md_ratioWidth      As Double

Private mi_prevItemIndex   As Integer
Private ms_ProblemArray()  As String

' Englober le multiligne dans les résultats

' Ajouter la prise en compte du "CallByName" pour les fonctions/procédures/property
' Ajouter la prise en compte du "Evaluate"

' Ajouter le contrôle de l'utilisation des Enum

' Ajouter les contrôles de longueur (Classe, Méthode et ligne)

' Ajouter la recherche de tous les boutons sans macro affectée (.OnAction = "")
' Ajouter la recherche de tous les oleobjects dont .OnAction <> "" et dont la procédure n'existe pas dans le code

' Ajouter la recherche des contrôles ActiveX sans évènement

' Ajouter le contrôle des "" à la place des vbnullstring (sauf dans les chaines de caractères)
' Ajouter le contrôle des "Select Case" qui doivent avoir au moins 3 conditions (si 2 conditions alors remplacer par If)

' Ajouter le contrôle de présence des .Value quand nécessaire (complexe)
' Ajouter les contrôles de complexité (nested loops, nested conditions)
' Ajouter les contrôles de nommage

'Return DPI
Private Function PointsPerPixel() As Double
    Dim hDC As Long
    Dim lDotsPerInch As Long
    
    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC
End Function

'****************************************************************************************************
'***************************************** Userform methods** ***************************************
'****************************************************************************************************
Private Sub UserForm_Initialize()
    Dim lo_Ctrl     As Control
    Dim ll_Height   As Long
    Dim ll_Weight   As Long
    
    mb_UsfEvents = True
    
    ll_Height = GetSystemMetrics32(1) ' Screen Resolution height in points
    ll_Weight = GetSystemMetrics32(0) ' Screen Resolution width in points
    
    With Me
        .Height = ll_Height * PointsPerPixel 'Userform height= Height in Resolution * DPI * 50%
        .Width = ll_Weight * PointsPerPixel 'Userform width= Width in Resolution * DPI * 50%
        
        md_ratioHeight = .Height / 560.25
        md_ratioWidth = .Width / 1407.75
    End With
    
    ReDim ms_ProblemArray(1, 1)
    
    Call InitDicoMethods
    Call InitWorkbookList
    Call InitCompList
    Call InitProblemList
    Call InitRuleList
    
    For Each lo_Ctrl In Me.Controls
        With lo_Ctrl
            .Left = .Left * md_ratioWidth
            .Top = .Top * md_ratioHeight
            .Height = .Height * md_ratioHeight
            .Width = .Width * md_ratioWidth
        End With
    Next lo_Ctrl
End Sub

Private Sub UserForm_Terminate()
    Set gdic_UsfMethods = Nothing
    Set gdic_WbkMethods = Nothing
    Set gdic_WsMethods = Nothing
End Sub

'****************************************************************************************************
'***************************************** Controls methods *****************************************
'****************************************************************************************************
Private Sub AnalyzeButton_Click()
    Dim lb_SameComment      As Boolean
    Dim lo_CMod             As CodeModule
    Dim li_CountHigh        As Integer
    Dim li_CountMiddle      As Integer
    Dim li_CountLow         As Integer
    Dim li_SelectItems      As Integer
    Dim li_SubItems         As Integer
    Dim li_CommentPos       As Integer
    Dim ll_CodeLine         As Long
    Dim ll_StartLine        As Long
    Dim ll_FoundOption      As Long
    Dim ls_CodeLine         As String
    Dim ls_TrimLine         As String
    Dim ls_MethodName       As String
    Dim ls_PrevMethod       As String
    Dim lv_WbEach           As Variant
    Dim lv_Each             As Variant
    Dim lo_VBC              As VBComponent
    Dim ldic_Private        As New Scripting.Dictionary
    Dim ldic_Public         As New Scripting.Dictionary
    Dim ldic_Tmp            As New Scripting.Dictionary
    
    Me.Hide
    Progression.Show
    Progression.BarProgress.Max = 1
    
    Set reg = New VBScript_RegExp_55.RegExp
    
    With Me.CodeComponents
        For li_SelectItems = 1 To .ListItems.Count
            If (.ListItems(li_SelectItems).Checked) Then
                Progression.BarProgress.Max = Progression.BarProgress.Max + 1
            End If
        Next li_SelectItems
        
        Call InitProblemList
        Call GetRules
        
        Progression.LabelVbcName = "Getting all called macros."
        Progression.Repaint
        Call ListCalledMacros
        
        For li_SelectItems = 1 To .ListItems.Count
            If (.ListItems(li_SelectItems).Checked) Then
                Progression.LabelVbcName = .ListItems(li_SelectItems).Text & ", " & .ListItems(li_SelectItems).ListSubItems(1).Text
                
                If (currentWorkbook.Name <> .ListItems(li_SelectItems).Text) Then
                    Set currentWorkbook = Application.Workbooks(.ListItems(li_SelectItems).Text)
                    
                    If (checkMethodUse) Then
                        Progression.LabelVbcName = "Getting all called macros."
                        Progression.Repaint
                        Call ListCalledMacros
                    End If
                End If
                
                If (Not ldic_Public.Exists(currentWorkbook.Name)) Then
                    ldic_Public.Add currentWorkbook.Name, Array(currentWorkbook.Name, New Scripting.Dictionary)
                End If
                
                Set lo_VBC = currentWorkbook.VBProject.VBComponents(.ListItems(li_SelectItems).ListSubItems(1).Text)
                Set lo_CMod = lo_VBC.CodeModule
                
                ls_MethodName = vbNullString
                ll_FoundOption = 0
                
                For ll_CodeLine = 1 To lo_CMod.CountOfDeclarationLines
                    DoEvents
                    
                    ls_CodeLine = lo_CMod.Lines(ll_CodeLine, 1)
                    ll_StartLine = ll_CodeLine
                    lb_SameComment = False
                    
                    If (Trim(ls_CodeLine) <> vbNullString) Then
                        While (Right(ls_CodeLine, 2) = " _")
                            ll_CodeLine = ll_CodeLine + 1
                            
                            If (checKCommentMulti And (Not lb_SameComment)) Then
                                If (CommentStart(ls_CodeLine) > 0) Then
                                    Call AddIssue(lo_VBC.Name, ll_CodeLine, COMMENT_MULTI, COMMENT_ONELINE, BASSE, , ls_MethodName)
                                    lb_SameComment = True
                                End If
                            End If
                            
                            If (Right(ls_CodeLine, 3) = "& _") Then
                                ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 3) & " " & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                            Else
                                ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 1) & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                            End If
                        Wend
                        
                        ls_TrimLine = Trim(ls_CodeLine)
                        
                        If (Left(ls_TrimLine, 15) = "Option Explicit") Then ' Check for presence of "Option Explicit"
                            ll_FoundOption = ll_StartLine
                        ElseIf (Left(ls_TrimLine, 1) = "'") Then
                            If (checkTodo) Then
                                Call SearchTODO(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                        Else
                            li_CommentPos = CommentStart(ls_TrimLine) - 1 ' Get start position of comment in line
                            
                            If (li_CommentPos > 0) Then ' If there is a comment
                                If (checkTodo) Then ' Check if comment contains TODO
                                    Call SearchTODO(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                ls_TrimLine = Trim(Left(ls_TrimLine, li_CommentPos)) ' Delete comment part from line
                            End If
                            
                            ' Delete every string between quotes using temporary string
                            Call ReplaceAllRegPattern(ls_TrimLine, ("((" & Chr(34) & Chr(34) & ")|(" & Chr(34) & "[^" & Chr(34) & "]+" & Chr(34) & "))"), "/\")
                            Call ReplaceAllRegPattern(ls_TrimLine, "\/\\", Chr(34) & Chr(34))
                            
                            If (checkVarUse Or checkMultiDim) Then
                                Set ldic_Tmp = ldic_Public.item(currentWorkbook.Name)(1)
                                Call FindUnusedModuleVar(lo_VBC, ldic_Tmp, ldic_Private, ll_StartLine, ls_TrimLine, ls_MethodName)
                                ldic_Public.Remove currentWorkbook.Name
                                ldic_Public.Add currentWorkbook.Name, Array(currentWorkbook.Name, ldic_Tmp)
                            End If
                            
                            If (checkInLine) Then
                                Call SearchInLine(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkBoolParen) Then
                                Call SearchBoolParen(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkScope) Then
                                Call MissingModuleScope(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkType Or checkNameVar Or checkNameConst Or checkNameLength) Then
                                Call MissingModuleType(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkResume) Then
                                Call SearchResume(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkEnd) Then
                                Call SearchEnd(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkBoolean) Then
                                Call AnalyseBoolean(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkStringConcat) Then
                                Call SearchStringConcat(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkMissingNext) Then
                                Call SearchMissingNext(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                                
                            If (checkGoTo) Then
                                Call SearchGoTo(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkSquareBrackets) Then
                                Call SearchSquareBrackets(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkDebug) Then
                                Call SearchDebug(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkImplicitWs) Then
                                Call SearchImplicitWs(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkImplicitWbk) Then
                                Call SearchImplicitWbk(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                        End If
                    End If
                Next ll_CodeLine
                
                ' Reporting problems about "Option Explicit"
                If (checkMissingOption) Then
                    Call MissingOption(lo_VBC.Name, ll_FoundOption)
                End If
                
                ls_PrevMethod = vbNullString
                
                For ll_CodeLine = lo_CMod.CountOfDeclarationLines + 1 To lo_CMod.CountOfLines
                    DoEvents
                    
                    ls_MethodName = lo_CMod.ProcOfLine(ll_CodeLine, vbext_pk_Proc)
                    ls_CodeLine = lo_CMod.Lines(ll_CodeLine, 1)
                    ll_StartLine = ll_CodeLine
                    lb_SameComment = False
                    
                    If (Trim(ls_CodeLine) <> vbNullString) Then
                        While (Right(ls_CodeLine, 2) = " _")
                            ll_CodeLine = ll_CodeLine + 1
                            
                            If (checKCommentMulti And (Not lb_SameComment)) Then
                                If (CommentStart(ls_CodeLine) > 0) Then
                                    Call AddIssue(lo_VBC.Name, ll_CodeLine, COMMENT_MULTI, COMMENT_ONELINE, BASSE, , ls_MethodName)
                                    lb_SameComment = True
                                End If
                            End If
                            
                            If (Right(ls_CodeLine, 3) = "& _") Then
                                ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 3) & " " & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                            Else
                                ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 1) & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                            End If
                        Wend
                        
                        ls_TrimLine = Trim(ls_CodeLine)
                        
                        If (Left(ls_TrimLine, 1) = "'") Then
                            If (checkTodo) Then
                                Call SearchTODO(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                        Else
                            li_CommentPos = CommentStart(ls_TrimLine) - 1
                            
                            If (li_CommentPos > 0) Then
                                If (checkTodo) Then
                                    Call SearchTODO(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                ls_TrimLine = Trim(Left(ls_TrimLine, li_CommentPos))
                            End If
                            
                            Call ReplaceAllRegPattern(ls_TrimLine, ("((" & Chr(34) & Chr(34) & ")|(" & Chr(34) & "[^" & Chr(34) & "]+" & Chr(34) & "))"), "/\")
                            Call ReplaceAllRegPattern(ls_TrimLine, "\/\\", Chr(34) & Chr(34))
                            
                            reg.Pattern = "^(( )*((Public|Friend|Private)( ))?(Static )?(Property|Sub|Function)( [GS]et)?( )(\w+)(\(.*\))( As \w+)?)"
                            
                            If (reg.test(ls_TrimLine)) Then
                                If (checkMethodUse And (ls_PrevMethod <> ls_MethodName)) Then
                                    Set ldic_Tmp = ldic_Public.item(currentWorkbook.Name)(1)
                                    Call FindUnusedMethods(lo_VBC, ldic_Tmp, ldic_Private, ls_MethodName)
                                    ldic_Public.Remove currentWorkbook.Name
                                    ldic_Public.Add currentWorkbook.Name, Array(currentWorkbook.Name, ldic_Tmp)
                                End If
                                
                                If (checkNameLength) Then
                                    Call ValidNameLength(lo_VBC.Name, ll_StartLine, ls_MethodName)
                                End If
                                
                                If (checkNameMethod) Then
                                    Call ValidMethodName(lo_VBC.Name, ll_StartLine, ls_MethodName)
                                End If
                                
                                If (checkScope) Then
                                    Call MissingMethodScope(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                            
                                If (checkType) Then
                                    Call MissingMethodType(lo_VBC, ll_StartLine, ls_TrimLine, ls_MethodName)
                                End If
                                
                                If (checkType Or checkByref Or checkNameVar Or checkNameLength) Then
                                    Call MissingParamType(lo_VBC, ll_StartLine, ls_TrimLine, ls_MethodName)
                                End If
                                
                                If (checkParamUse) Then
                                    Call FindUnusedParams(lo_VBC, ll_StartLine, ls_TrimLine, ls_MethodName)
                                End If
                                
                                If (checkEmptyMethod And (ls_PrevMethod <> ls_MethodName) And (ls_MethodName <> "Class_Initialize") And (ls_MethodName <> "Class_Terminate")) Then
                                    Call SearchEmptyMethod(lo_VBC, ls_MethodName)
                                End If
                                
                                ls_PrevMethod = ls_MethodName
                            Else
                                If (checkVarUse Or checkMultiDim) Then
                                    Call FindUnusedMethodVar(lo_VBC, ll_StartLine, ls_TrimLine)
                                End If

                                If (checkBoolParen) Then
                                    Call SearchBoolParen(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkType Or checkNameLength Or checkNameConst Or checkNameVar) Then
                                    Call MissingVarType(lo_VBC.Name, ll_StartLine, ls_CodeLine)
                                End If
                                
                                If (checkResume) Then
                                    Call SearchResume(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkEnd) Then
                                    Call SearchEnd(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkBoolean) Then
                                    Call AnalyseBoolean(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkInLine) Then
                                    Call SearchInLine(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkMissingNext) Then
                                    Call SearchMissingNext(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkGoTo) Then
                                    Call SearchGoTo(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkDebug) Then
                                    Call SearchDebug(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                            
                                If (checkImplicitWs) Then
                                    Call SearchImplicitWs(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                                
                                If (checkImplicitWbk) Then
                                    Call SearchImplicitWbk(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                                End If
                            End If
                            
                            If (checkSquareBrackets) Then
                                Call SearchSquareBrackets(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                            
                            If (checkStringConcat) Then
                                Call SearchStringConcat(lo_VBC.Name, ll_StartLine, ls_TrimLine)
                            End If
                        End If
                    End If
                Next ll_CodeLine
                
                If (checkMethodUse Or checkVarUse) Then
                    For ll_CodeLine = lo_CMod.CountOfDeclarationLines + 1 To lo_CMod.CountOfLines - 1
                        DoEvents
                        
                        ls_MethodName = lo_CMod.ProcOfLine(ll_CodeLine, vbext_pk_Proc)
                        ls_CodeLine = lo_CMod.Lines(ll_CodeLine, 1)
                        ll_StartLine = ll_CodeLine
                        
                        If (Trim(ls_CodeLine) <> vbNullString) Then
                            While (Right(ls_CodeLine, 2) = " _")
                                ll_CodeLine = ll_CodeLine + 1
                                
                                If (Right(ls_CodeLine, 3) = "& _") Then
                                    ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 3) & " " & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                                Else
                                    ls_CodeLine = Left(ls_CodeLine, Len(ls_CodeLine) - 1) & Trim(lo_CMod.Lines(ll_CodeLine, 1))
                                End If
                            Wend
                            
                            ls_TrimLine = Trim(ls_CodeLine)
                            
                            If (Left(ls_TrimLine, 1) <> "'") Then
                                li_CommentPos = CommentStart(ls_TrimLine) - 1
                                
                                If (li_CommentPos > 0) Then
                                    ls_TrimLine = Trim(Left(ls_TrimLine, li_CommentPos))
                                End If
                            
                                If (ldic_Public.item(currentWorkbook.Name)(1).Count > 0) Then
                                    Set ldic_Tmp = ldic_Public.item(currentWorkbook.Name)(1)
                                    Call SearchLineUse(ldic_Tmp, ls_TrimLine, ls_MethodName)
                                    ldic_Public.Remove currentWorkbook.Name
                                    ldic_Public.Add currentWorkbook.Name, Array(currentWorkbook.Name, ldic_Tmp)
                                End If
                                
                                If (ldic_Private.Count > 0) Then
                                    Call SearchLineUse(ldic_Private, ls_TrimLine, ls_MethodName)
                                End If
                            End If
                        End If
                    Next ll_CodeLine
                    
                    If (ldic_Private.Count > 0) Then
                        For Each lv_Each In ldic_Private.Items
                            DoEvents
                            
                            'var/Method Name ; module name ; declare line ; isMethod ; isObj/lb_IsConst
                            If (lv_Each(3) And checkMethodUse) Then
                                Call AddIssue(lv_Each(1), lv_Each(2), PRIVATE_SCOPE & (METHOD & NOTUSED), (DELETE & METHOD), HAUTE, Split(lv_Each(0), "/")(2))
                            ElseIf ((Not lv_Each(3)) And checkVarUse) Then
                                Call AddIssue(lv_Each(1), lv_Each(2), (PRIVATE_SCOPE & IIf(lv_Each(4), CONSTANTE, VARIABLE) & NOTUSED), (DELETE & IIf(lv_Each(4), CONSTANTE, VARIABLE)), HAUTE, Split(lv_Each(0), "/")(2))
                            End If
                        Next lv_Each
                    End If
                End If
                
                ldic_Private.RemoveAll
                
                Progression.BarProgress.Value = Progression.BarProgress.Value + 1
                Progression.Repaint
            End If
        Next li_SelectItems
        
        If (checkMethodUse Or checkVarUse) Then
            If (checkMethodUse And checkVarUse) Then
                Progression.LabelVbcName = "Looking for public variables/constants and methods not used."
            ElseIf (checkMethodUse) Then
                Progression.LabelVbcName = "Looking for public methods not used."
            ElseIf (checkVarUse) Then
                Progression.LabelVbcName = "Looking for public variables/constants not used."
            End If
            
            Progression.Repaint
            
            For Each lv_WbEach In ldic_Public.Items
                Set currentWorkbook = Application.Workbooks(lv_WbEach(0))
                Set ldic_Tmp = lv_WbEach(1)
                
                Call SearchPublicUse(ldic_Tmp)
                
                If (ldic_Tmp.Count > 0) Then
                    For Each lv_Each In ldic_Tmp.Items
                        DoEvents
                        
                        If (lv_Each(3) And checkMethodUse) Then
                            Call AddIssue(lv_Each(1), lv_Each(2), (PUBLIC_SCOPE & METHOD & NOTUSED), (DELETE & METHOD), HAUTE, Split(lv_Each(0), "/")(2), lv_Each(5))
                        ElseIf ((Not lv_Each(3)) And checkVarUse) Then
                            Call AddIssue(lv_Each(1), lv_Each(2), (PUBLIC_SCOPE & IIf(lv_Each(4), CONSTANTE, VARIABLE) & NOTUSED), (DELETE & IIf(lv_Each(4), CONSTANTE, VARIABLE)), HAUTE, Split(lv_Each(0), "/")(2), lv_Each(5))
                        End If
                    Next lv_Each
                End If
            Next lv_WbEach
        End If
    End With
        
    Progression.LabelVbcName = "Updating problem list"
    Progression.BarProgress.Value = Progression.BarProgress.Value + 1
    
    With Me.CodeProblems
        .SortKey = 4
        .sortOrder = lvwDescending
        .Sorted = True
    End With
    
    With Me.CodeProblems
        li_CountHigh = 0
        li_CountMiddle = 0
        li_CountLow = 0
        
        If (.ListItems.Count > 0) Then
            ReDim ms_ProblemArray(1 To .ListItems.Count, 1 To 8)
            
            For li_SelectItems = 1 To .ListItems.Count
                DoEvents
                ms_ProblemArray(li_SelectItems, 1) = .ListItems(li_SelectItems).Text
                
                For li_SubItems = 1 To 7
                    ms_ProblemArray(li_SelectItems, li_SubItems + 1) = .ListItems(li_SelectItems).ListSubItems(li_SubItems)
                Next li_SubItems
                
                Select Case .ListItems(li_SelectItems).ListSubItems(7)
                    Case HAUTE
                        li_CountHigh = li_CountHigh + 1
                    Case MOYENNE
                        li_CountMiddle = li_CountMiddle + 1
                    Case BASSE
                        li_CountLow = li_CountLow + 1
                End Select
            Next li_SelectItems
        Else
            ReDim ms_ProblemArray(1 To 1, 1 To 8)
        End If
        
        Me.LabelTotal.Caption = "TOTAL : " & CStr(.ListItems.Count)
        Me.LabelHigh.Caption = "HAUTES : " & CStr(li_CountHigh)
        Me.LabelMiddle.Caption = "MOYENNES : " & CStr(li_CountMiddle)
        Me.LabelLow.Caption = "BASSES : " & CStr(li_CountLow)
    End With
    
    mi_prevItemIndex = -1
    
    Set reg = Nothing
    Set calledMacros = Nothing
    
    Unload Progression
    Me.Show
End Sub

' ************************************* Workbook listview *********************************************
Private Sub WorkbookList_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Dim ll_Index As Long
    Dim lb_AllSelected As Boolean
    Dim lb_EnableComp As Boolean
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        lb_AllSelected = True
        lb_EnableComp = False
        
        With Me.WorkbookList
            For ll_Index = 1 To .ListItems.Count
                If (Not .ListItems(ll_Index).Checked) Then
                    lb_AllSelected = False
                Else
                    lb_EnableComp = True
                End If
            Next ll_Index
        End With
        
        Call InitCompList
        
        With Me
            .SelectAllWorkbooks.Value = lb_AllSelected
            .SelectAllComponents.Enabled = lb_EnableComp
            .CodeComponents.Enabled = lb_EnableComp
            
            .SelectClassModules.Value = (.SelectClassModules.Value And .SelectClassModules.Enabled)
            .SelectOthers.Value = (.SelectOthers.Value And .SelectOthers.Enabled)
            .SelectStandardModules.Value = (.SelectStandardModules.Value And .SelectStandardModules.Enabled)
            .SelectUserforms.Value = (.SelectUserforms.Value And .SelectUserforms.Enabled)
            .SelectWbkDocuments.Value = (.SelectWbkDocuments.Value And .SelectWbkDocuments.Enabled)
            .SelectWsDocuments.Value = (.SelectWsDocuments.Value And .SelectWsDocuments.Enabled)
        End With
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub WorkbookList_ItemClick(ByVal item As MSComctlLib.ListItem)
    item.Checked = Not item.Checked
    Call WorkbookList_ItemCheck(item)
End Sub

Private Sub WorkbookList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.WorkbookList
        .SortKey = ColumnHeader.Index - 1
        
        If (.sortOrder = lvwAscending) Then
            .sortOrder = lvwDescending
        Else
            .sortOrder = lvwAscending
        End If
        
        .Sorted = True
    End With
End Sub

Private Sub WorkbookList_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1 ' Do no authorize label edit
End Sub

' ************************************* Workbook Checkbox*********************************************
Private Sub SelectAllWorkbooks_Click()
    Dim ll_Index As Long
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        If (Me.WorkbookList.ListItems.Count = 0) Then
            Me.SelectAllWorkbooks.Value = False
        Else
            With Me.WorkbookList
                For ll_Index = 1 To .ListItems.Count
                    .ListItems(ll_Index).Checked = Me.SelectAllWorkbooks.Value
                Next ll_Index
                
                Call InitCompList
            End With
        End If
        
        With Me
            .SelectAllComponents.Enabled = .SelectAllWorkbooks.Value
            .CodeComponents.Enabled = .SelectAllWorkbooks.Value
            
            .SelectClassModules.Value = (.SelectClassModules.Value And .SelectClassModules.Enabled)
            .SelectOthers.Value = (.SelectOthers.Value And .SelectOthers.Enabled)
            .SelectStandardModules.Value = (.SelectStandardModules.Value And .SelectStandardModules.Enabled)
            .SelectUserforms.Value = (.SelectUserforms.Value And .SelectUserforms.Enabled)
            .SelectWbkDocuments.Value = (.SelectWbkDocuments.Value And .SelectWbkDocuments.Enabled)
            .SelectWsDocuments.Value = (.SelectWsDocuments.Value And .SelectWsDocuments.Enabled)
        End With
        
        mb_UsfEvents = True
    End If
End Sub

' ************************************* Components listview *********************************************
Private Sub CodeComponents_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Dim ll_Index As Long
    Dim lb_AllSelected As Boolean
    Dim lb_ClassModule As Boolean
    Dim lb_Other As Boolean
    Dim lb_StdModule As Boolean
    Dim lb_Usf As Boolean
    Dim lb_Wbk As Boolean
    Dim lb_Ws As Boolean
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        lb_AllSelected = True
        lb_ClassModule = True
        lb_Other = True
        lb_StdModule = True
        lb_Usf = True
        lb_Wbk = True
        lb_Ws = True
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                If (Not .ListItems(ll_Index).Checked) Then
                    lb_AllSelected = False
                    
                    Select Case .ListItems(ll_Index).ListSubItems(2).Text
                        Case "Standard Module"
                            lb_StdModule = False
                        Case "Class Module"
                            lb_ClassModule = False
                        Case "Workbook Document"
                            lb_Wbk = False
                        Case "Worksheet Document"
                            lb_Ws = False
                        Case "UserForm"
                            lb_Usf = False
                        Case "ActiveX Designer", "Unknown"
                            lb_Other = False
                    End Select
                End If
            Next ll_Index
        End With
        
        Me.SelectAllComponents.Value = lb_AllSelected
        Me.SelectClassModules.Value = (lb_ClassModule And Me.SelectClassModules.Enabled)
        Me.SelectOthers.Value = (lb_Other And Me.SelectOthers.Enabled)
        Me.SelectStandardModules.Value = (lb_StdModule And Me.SelectStandardModules.Enabled)
        Me.SelectUserforms.Value = (lb_Usf And Me.SelectUserforms.Enabled)
        Me.SelectWbkDocuments.Value = (lb_Wbk And Me.SelectWbkDocuments.Enabled)
        Me.SelectWsDocuments.Value = (lb_Ws And Me.SelectWsDocuments.Enabled)
        
        Call UpdateProblemList
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub CodeComponents_ItemClick(ByVal item As MSComctlLib.ListItem)
    item.Checked = Not item.Checked
    Call CodeComponents_ItemCheck(item)
End Sub

Private Sub CodeComponents_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.CodeComponents
        .SortKey = ColumnHeader.Index - 1
        
        If (.sortOrder = lvwAscending) Then
            .sortOrder = lvwDescending
        Else
            .sortOrder = lvwAscending
        End If
        
        .Sorted = True
    End With
End Sub

Private Sub CodeComponents_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1 ' Do no authorize label edit
End Sub

' ************************************* Components Checkbox *********************************************
Private Sub SelectAllComponents_Click()
    Dim ll_Index As Long
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                .ListItems(ll_Index).Checked = Me.SelectAllComponents.Value
            Next ll_Index
        End With
        
        With Me.SelectAllComponents
            Me.SelectClassModules.Value = (.Value And Me.SelectClassModules.Enabled)
            Me.SelectOthers.Value = (.Value And Me.SelectOthers.Enabled)
            Me.SelectStandardModules.Value = (.Value And Me.SelectStandardModules.Enabled)
            Me.SelectUserforms.Value = (.Value And Me.SelectUserforms.Enabled)
            Me.SelectWbkDocuments.Value = (.Value And Me.SelectWbkDocuments.Enabled)
            Me.SelectWsDocuments.Value = (.Value And Me.SelectWsDocuments.Enabled)
        End With
        
        mb_UsfEvents = True
    End If
End Sub

' ************************************* Rules listview *********************************************
Private Sub CodeRules_ItemCheck(ByVal item As MSComctlLib.ListItem)
    Dim ll_Index As Long
    Dim lb_AllHautes As Boolean
    Dim lb_AllMoyennes As Boolean
    Dim lb_AllBasses As Boolean
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        If (Me.SelectAllCodeRules.Value And (item.Checked)) Then
            Me.SelectAllCodeRules.Value = False
            
            Select Case item.ListSubItems(2).Text
                Case HAUTE
                    Me.SelectHautes.Value = False
                Case MOYENNE
                    Me.SelectMoyennes.Value = False
                Case BASSE
                    Me.SelectBasses.Value = False
            End Select
        Else
            lb_AllHautes = True
            lb_AllMoyennes = True
            lb_AllBasses = True
            
            With Me.CodeRules
                For ll_Index = 1 To .ListItems.Count
                    If (Not .ListItems(ll_Index).Checked) Then
                        Me.SelectAllCodeRules.Value = False
                        
                        Select Case .ListItems(ll_Index).ListSubItems(2).Text
                            Case HAUTE
                                lb_AllHautes = False
                            Case MOYENNE
                                lb_AllMoyennes = False
                            Case BASSE
                                lb_AllBasses = False
                        End Select
                    End If
                Next ll_Index
                
                If (lb_AllHautes And lb_AllMoyennes And lb_AllBasses) Then
                    Me.SelectAllCodeRules.Value = True
                End If
                
                Me.SelectHautes.Value = lb_AllHautes
                Me.SelectMoyennes.Value = lb_AllMoyennes
                Me.SelectBasses.Value = lb_AllBasses
            End With
        End If
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub CodeRules_ItemClick(ByVal item As MSComctlLib.ListItem)
    item.Checked = Not item.Checked
    Call CodeRules_ItemCheck(item)
End Sub

Private Sub CodeRules_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1 ' Do no authorize label edit
End Sub

' ************************************* All Rules Checkbox*********************************************
Private Sub SelectAllCodeRules_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeRules
            For ll_Index = 1 To .ListItems.Count
                .ListItems(ll_Index).Checked = Me.SelectAllCodeRules.Value
            Next ll_Index
            
            Me.SelectBasses.Value = Me.SelectAllCodeRules.Value
            Me.SelectMoyennes.Value = Me.SelectAllCodeRules.Value
            Me.SelectHautes.Value = Me.SelectAllCodeRules.Value
        End With
        
        mb_UsfEvents = True
    End If
End Sub

' ************************************* Hautes Checkbox*********************************************
Private Sub SelectHautes_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeRules
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = HAUTE) Then
                    .ListItems(ll_Index).Checked = Me.SelectHautes.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllCodeRules.Value = (Me.SelectBasses.Value And Me.SelectMoyennes.Value And Me.SelectHautes.Value)
        
        mb_UsfEvents = True
    End If
End Sub

' ************************************* Moyennes Checkbox*********************************************
Private Sub SelectMoyennes_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeRules
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = MOYENNE) Then
                    .ListItems(ll_Index).Checked = Me.SelectMoyennes.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllCodeRules.Value = (Me.SelectBasses.Value And Me.SelectMoyennes.Value And Me.SelectHautes.Value)
        
        mb_UsfEvents = True
    End If
End Sub

' ************************************* Basses Checkbox*********************************************
Private Sub SelectBasses_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeRules
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = BASSE) Then
                    .ListItems(ll_Index).Checked = Me.SelectBasses.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllCodeRules.Value = (Me.SelectBasses.Value And Me.SelectMoyennes.Value And Me.SelectHautes.Value)
        
        mb_UsfEvents = True
    End If
End Sub

' ************************************* Type Component Checkboxes*********************************************
Private Sub SelectClassModules_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = "Class Module") Then
                    .ListItems(ll_Index).Checked = Me.SelectClassModules.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllComponents.Value = ((Me.SelectClassModules.Value Or (Not Me.SelectClassModules.Enabled)) And _
                (Me.SelectOthers.Value Or (Not Me.SelectOthers.Value)) And _
                (Me.SelectStandardModules.Value Or (Not Me.SelectStandardModules.Enabled)) And _
                (Me.SelectUserforms.Value Or (Not Me.SelectUserforms.Enabled)) And _
                (Me.SelectWbkDocuments.Value Or (Not Me.SelectWbkDocuments.Enabled)) And _
                (Me.SelectWsDocuments.Value Or (Not Me.SelectWsDocuments.Enabled)))
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub SelectOthers_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                If ((.ListItems(ll_Index).ListSubItems(2).Text = "ActiveX Designer") Or (.ListItems(ll_Index).ListSubItems(2).Text = "Unknown")) Then
                    .ListItems(ll_Index).Checked = Me.SelectOthers.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllComponents.Value = ((Me.SelectClassModules.Value Or (Not Me.SelectClassModules.Enabled)) And _
                (Me.SelectOthers.Value Or (Not Me.SelectOthers.Value)) And _
                (Me.SelectStandardModules.Value Or (Not Me.SelectStandardModules.Enabled)) And _
                (Me.SelectUserforms.Value Or (Not Me.SelectUserforms.Enabled)) And _
                (Me.SelectWbkDocuments.Value Or (Not Me.SelectWbkDocuments.Enabled)) And _
                (Me.SelectWsDocuments.Value Or (Not Me.SelectWsDocuments.Enabled)))
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub SelectStandardModules_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = "Standard Module") Then
                    .ListItems(ll_Index).Checked = Me.SelectStandardModules.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllComponents.Value = ((Me.SelectClassModules.Value Or (Not Me.SelectClassModules.Enabled)) And _
                (Me.SelectOthers.Value Or (Not Me.SelectOthers.Value)) And _
                (Me.SelectStandardModules.Value Or (Not Me.SelectStandardModules.Enabled)) And _
                (Me.SelectUserforms.Value Or (Not Me.SelectUserforms.Enabled)) And _
                (Me.SelectWbkDocuments.Value Or (Not Me.SelectWbkDocuments.Enabled)) And _
                (Me.SelectWsDocuments.Value Or (Not Me.SelectWsDocuments.Enabled)))
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub SelectUserforms_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = "UserForm") Then
                    .ListItems(ll_Index).Checked = Me.SelectUserforms.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllComponents.Value = ((Me.SelectClassModules.Value Or (Not Me.SelectClassModules.Enabled)) And _
                (Me.SelectOthers.Value Or (Not Me.SelectOthers.Value)) And _
                (Me.SelectStandardModules.Value Or (Not Me.SelectStandardModules.Enabled)) And _
                (Me.SelectUserforms.Value Or (Not Me.SelectUserforms.Enabled)) And _
                (Me.SelectWbkDocuments.Value Or (Not Me.SelectWbkDocuments.Enabled)) And _
                (Me.SelectWsDocuments.Value Or (Not Me.SelectWsDocuments.Enabled)))
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub SelectWbkDocuments_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = "Workbook Document") Then
                    .ListItems(ll_Index).Checked = Me.SelectWbkDocuments.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllComponents.Value = ((Me.SelectClassModules.Value Or (Not Me.SelectClassModules.Enabled)) And _
                (Me.SelectOthers.Value Or (Not Me.SelectOthers.Value)) And _
                (Me.SelectStandardModules.Value Or (Not Me.SelectStandardModules.Enabled)) And _
                (Me.SelectUserforms.Value Or (Not Me.SelectUserforms.Enabled)) And _
                (Me.SelectWbkDocuments.Value Or (Not Me.SelectWbkDocuments.Enabled)) And _
                (Me.SelectWsDocuments.Value Or (Not Me.SelectWsDocuments.Enabled)))
        
        mb_UsfEvents = True
    End If
End Sub

Private Sub SelectWsDocuments_Click()
    Dim ll_Index As Integer
    
    If (mb_UsfEvents) Then
        mb_UsfEvents = False
        
        With Me.CodeComponents
            For ll_Index = 1 To .ListItems.Count
                If (.ListItems(ll_Index).ListSubItems(2).Text = "Worksheet Document") Then
                    .ListItems(ll_Index).Checked = Me.SelectWsDocuments.Value
                End If
            Next ll_Index
        End With
        
        Me.SelectAllComponents.Value = ((Me.SelectClassModules.Value Or (Not Me.SelectClassModules.Enabled)) And _
                (Me.SelectOthers.Value Or (Not Me.SelectOthers.Value)) And _
                (Me.SelectStandardModules.Value Or (Not Me.SelectStandardModules.Enabled)) And _
                (Me.SelectUserforms.Value Or (Not Me.SelectUserforms.Enabled)) And _
                (Me.SelectWbkDocuments.Value Or (Not Me.SelectWbkDocuments.Enabled)) And _
                (Me.SelectWsDocuments.Value Or (Not Me.SelectWsDocuments.Enabled)))
        
        mb_UsfEvents = True
    End If
End Sub

' ************************************* Problems listview *********************************************
Private Sub CodeProblems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.CodeProblems
        .SortKey = ColumnHeader.Index - 1
        
        If (.sortOrder = lvwAscending) Then
            .sortOrder = lvwDescending
        Else
            .sortOrder = lvwAscending
        End If
        
        .Sorted = True
    End With
End Sub

Private Sub CodeProblems_DblClick()
    Dim ll_Index    As Long
    Dim ll_Ligne    As Long
    Dim lo_CMod     As Object
    
    With Me.CodeProblems
        If (.ListItems.Count <> 0) Then
            ll_Index = CLng(.SelectedItem.Index)
            
            If (mi_prevItemIndex <> -1) Then
                .ListItems(mi_prevItemIndex).ForeColor = RGB(0, 0, 0)
                
                For ll_Ligne = 1 To 7
                    .ListItems(mi_prevItemIndex).ListSubItems(ll_Ligne).ForeColor = RGB(0, 0, 0)
                Next ll_Ligne
            End If
            
            mi_prevItemIndex = ll_Index
            .SelectedItem.ForeColor = RGB(0, 255, 0)
            
            For ll_Ligne = 1 To 7
                .SelectedItem.ListSubItems(ll_Ligne).ForeColor = RGB(0, 255, 0)
            Next ll_Ligne
            
            Call Me.Repaint
            
            Set currentWorkbook = Application.Workbooks(.ListItems(ll_Index).Text)
            Set lo_CMod = currentWorkbook.VBProject.VBComponents(.ListItems(ll_Index).ListSubItems(1).Text).CodeModule
            
            ll_Ligne = CLng(Trim(.ListItems(ll_Index).ListSubItems(4)))
            
            Call lo_CMod.CodePane.Show
            Call lo_CMod.CodePane.SetSelection(ll_Ligne, 1, ll_Ligne, 1020)
        End If
    End With
End Sub

Private Sub CodeProblems_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1 ' Do no authorize label edit
End Sub

'****************************************************************************************************
'****************************************** Init methods *****************************************
'****************************************************************************************************
Private Sub InitWorkbookList()
    Dim lwbk_Wb         As Workbook
    Dim li_SelectItems  As Integer
    
    With Me.WorkbookList
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Workbook Name", 180 * md_ratioWidth
        
        For Each lwbk_Wb In Application.Workbooks
            .ListItems.Add , , lwbk_Wb.Name
        Next lwbk_Wb
        
        .Sorted = True
        .SortKey = 0
        .CheckBoxes = True
        
        For li_SelectItems = 1 To .ListItems.Count
            .ListItems(li_SelectItems).Checked = False
        Next li_SelectItems
        
        Set .SelectedItem = Nothing
    End With
End Sub

Private Sub InitCompList()
    Dim lb_ClassModule      As Boolean
    Dim lb_StandardModule   As Boolean
    Dim lb_Others           As Boolean
    Dim lb_Usf              As Boolean
    Dim lb_Workbook         As Boolean
    Dim lb_Worksheet        As Boolean
    Dim li_SelectItems      As Integer
    Dim lo_VBC              As VBComponent
    
    With Me.CodeComponents
        .ListItems.Clear
        
        With .ColumnHeaders
            .Clear
            .Add , , "Workbook Name", 180 * md_ratioWidth
            .Add , , "Component", 180 * md_ratioWidth
            .Add , , "Type", 110 * md_ratioWidth
        End With
        
        lb_StandardModule = False
        lb_ClassModule = False
        lb_Workbook = False
        lb_Worksheet = False
        lb_Usf = False
        lb_Others = False
        
        For li_SelectItems = 1 To Me.WorkbookList.ListItems.Count
            If (Me.WorkbookList.ListItems(li_SelectItems).Checked) Then
                Set currentWorkbook = Application.Workbooks(Me.WorkbookList.ListItems(li_SelectItems).Text)
                
                For Each lo_VBC In currentWorkbook.VBProject.VBComponents
                    .ListItems.Add , , currentWorkbook.Name
                    .ListItems(.ListItems.Count).ListSubItems.Add , , lo_VBC.Name
                    .ListItems(.ListItems.Count).ListSubItems.Add , , GetCompType(lo_VBC)
                    
                    Select Case GetCompType(lo_VBC)
                        Case "Standard Module"
                            lb_StandardModule = True
                        Case "Class Module"
                            lb_ClassModule = True
                        Case "Workbook Document"
                            lb_Workbook = True
                        Case "Worksheet Document"
                            lb_Worksheet = True
                        Case "UserForm"
                            lb_Usf = True
                        Case "ActiveX Designer", "Unknown"
                            lb_Others = True
                    End Select
                Next lo_VBC
            End If
        Next li_SelectItems
        
        .Sorted = True
        .SortKey = 0
        .CheckBoxes = True
        
        For li_SelectItems = 1 To .ListItems.Count
            .ListItems(li_SelectItems).Checked = False
        Next li_SelectItems
        
        Me.SelectClassModules.Enabled = lb_ClassModule
        Me.SelectStandardModules.Enabled = lb_StandardModule
        Me.SelectOthers.Enabled = lb_Others
        Me.SelectUserforms.Enabled = lb_Usf
        Me.SelectWbkDocuments.Enabled = lb_Workbook
        Me.SelectWsDocuments.Enabled = lb_Worksheet
        
        Set .SelectedItem = Nothing
    End With
    
    Me.SelectAllComponents.Value = False
End Sub

Private Sub InitProblemList()
    With Me.CodeProblems
        .ListItems.Clear
        
        'Définit le nombre de colonnes et Entêtes
        With .ColumnHeaders
            'Supprime les anciens entêtes
            .Clear
            'Ajoute 3 colonnes en spécifiant le nom de l'entête et la largeur des colonnes
            .Add , , "Workbook", 180 * md_ratioWidth
            .Add , , "Component", 180 * md_ratioWidth
            .Add , , "Method", 180 * md_ratioWidth
            .Add , , "Object", 180 * md_ratioWidth
            .Add , , "Line", 36 * md_ratioWidth
            .Add , , "Description", 280 * md_ratioWidth
            .Add , , "Solution", 280 * md_ratioWidth
            .Add , , "Criticity", 52 * md_ratioWidth
        End With
    End With
End Sub

Private Sub InitRuleList()
    Dim li_SelectItems As Integer
    
    With Me.CodeRules
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Rule Name", 120 * md_ratioWidth
        .ColumnHeaders.Add , , "Rule Description", 280 * md_ratioWidth
        .ColumnHeaders.Add , , "Criticity", 52 * md_ratioWidth
        
        Call AddRule("checkVarUse", "Unused Var", "Look for unused variables, constants", HAUTE)
        Call AddRule("checkMethodUse", "Unused Method", "Look for unused methods and functions", HAUTE)
        Call AddRule("checkParamUse", "Unused Param", "Look for unused parameters in methods and functions", HAUTE)
        Call AddRule("checkByref", "Passing parameter", "Look for missing ""ByRef"", ""ByVal"", ""ParamArray"" for parameters", HAUTE)
        Call AddRule("checkType", "Implicit type", "Look for var, const, param and function with implicit type declaration", HAUTE)
        Call AddRule("checkMissingOption", "Check Option Explicit", "Look for components where ""Option Explicit"" is missing or at wrong emplacement", HAUTE)
        Call AddRule("checkResume", "Check Resume", "Look for ""Resume"" instruction without ""On Error""", HAUTE)
        Call AddRule("checkEnd", "Check End", "Look for ""End"" instruction alone", HAUTE)
        Call AddRule("checkGoTo", "Check GoTo", "Look for ""GoTo"" instructions without ""On Error""", HAUTE)
        Call AddRule("checkInLine", "Inline instructions", "Look for multiple instructions on a single line", HAUTE)
        Call AddRule("checkStringConcat", "String concatenation", "Look for string concatenantion using ""+"" instead of ""&""", HAUTE)
        Call AddRule("checkImplicitWs", "Check implicit worksheet", "Look for ""Range"" and ""Cells"" with implicit call of worksheet", HAUTE)
        Call AddRule("checkImplicitWbk", "Check implicit workbook", "Look for ""Sheets"" and ""Worksheets"" with implicit call of workbook", MOYENNE)
        Call AddRule("checkSquareBrackets", "Check Square Brackets", "Look for use of ""["" (square brackets)", MOYENNE)
        Call AddRule("checkBoolean", "Boolean overload", "Look for overloaded boolean comparisons and instructions", MOYENNE)
        Call AddRule("checkScope", "Implicit scope", "Look for declarations with implicit scope", MOYENNE)
        Call AddRule("checkEmptyMethod", "Empty method", "Look for empty methods", MOYENNE)
        Call AddRule("checkDebug", "Check Debug", "Look for presence of ""Debug"" in code", MOYENNE)
        Call AddRule("checkBoolParen", "Boolean Parenthesis", "Look for missing parenthesis in boolean conditions", MOYENNE)
        Call AddRule("checkMultiDim", "Inline declarations", "Look for multiple declarations on a single line", BASSE)
        Call AddRule("checkNameConst", "Const naming convention", "Look for constants not complying with naming convention", BASSE)
        Call AddRule("checkNameLength", "Name lengths", "Look for names with length upper than 20 characters", BASSE)
        Call AddRule("checkNameMethod", "Method naming convention", "Look for methods not complying with naming convention", BASSE)
        Call AddRule("checkNameVar", "Var naming convention", "Look for variables not complying with naming convention", BASSE)
        Call AddRule("checkMissingNext", "Implicit Next", "Look for ""Next"" instructions with implicit variable use", BASSE)
        Call AddRule("checKCommentMulti", "Multiline Comment", "Look for comments ending with "" _""", BASSE)
        Call AddRule("checkTodo", "Check TODO", "Look for ""TODO"" comments", BASSE)
        
        .CheckBoxes = True
        
        For li_SelectItems = 1 To .ListItems.Count
            .ListItems(li_SelectItems).Checked = False
        Next li_SelectItems
        
        Set .SelectedItem = Nothing
    End With
End Sub

'****************************************************************************************************
'************************************* Control linked methods ************************************
'****************************************************************************************************
Private Sub UpdateProblemList()
    Dim li_SelectItem   As Integer
    Dim li_Pb           As Integer
    
    If (Len(ms_ProblemArray(1, 1)) > 0) Then
        With Me.CodeProblems
            .ListItems.Clear
            .Sorted = False
            
            With Me.CodeComponents
                For li_SelectItem = 1 To .ListItems.Count
                    If (.ListItems(li_SelectItem).Checked) Then
                        For li_Pb = 1 To UBound(ms_ProblemArray)
                            If ((ms_ProblemArray(li_Pb, 1) = .ListItems(li_SelectItem).Text) And (ms_ProblemArray(li_Pb, 2) = .ListItems(li_SelectItem).ListSubItems(1))) Then
                                Set currentWorkbook = Application.Workbooks(ms_ProblemArray(li_Pb, 1))
                                Call AddIssue(ms_ProblemArray(li_Pb, 2), ms_ProblemArray(li_Pb, 5), ms_ProblemArray(li_Pb, 6), ms_ProblemArray(li_Pb, 7), ms_ProblemArray(li_Pb, 8), ms_ProblemArray(li_Pb, 4))
                            End If
                        Next li_Pb
                    End If
                Next li_SelectItem
            End With
            
            .Sorted = True
            .SortKey = 4
            .sortOrder = lvwDescending
        End With
    End If
End Sub

Private Sub AddRule(ByVal ps_Key As String, ByVal ps_RuleName As String, ByVal ps_RuleDesc As String, ByVal ps_Criticity As String)
    With Me.CodeRules
        .ListItems.Add , ps_Key, ps_RuleName
        .ListItems(.ListItems.Count).ListSubItems.Add , , ps_RuleDesc
        .ListItems(.ListItems.Count).ListSubItems.Add , , ps_Criticity
    End With
End Sub

Private Sub GetRules()
    With Me.CodeRules
        checkBoolean = .ListItems("checkBoolean").Checked
        checkByref = .ListItems("checkByref").Checked
        checKCommentMulti = .ListItems("checKCommentMulti").Checked
        checkDebug = .ListItems("checkDebug").Checked
        checkEmptyMethod = .ListItems("checkEmptyMethod").Checked
        checkEnd = .ListItems("checkEnd").Checked
        checkGoTo = .ListItems("checkGoTo").Checked
        checkBoolParen = .ListItems("checkBoolParen").Checked
        checkImplicitWbk = .ListItems("checkImplicitWbk").Checked
        checkImplicitWs = .ListItems("checkImplicitWs").Checked
        checkInLine = .ListItems("checkInLine").Checked
        checkMissingNext = .ListItems("checkMissingNext").Checked
        checkMissingOption = .ListItems("checkMissingOption").Checked
        checkMultiDim = .ListItems("checkMultiDim").Checked
        checkNameConst = .ListItems("checkNameConst").Checked
        checkNameLength = .ListItems("checkNameLength").Checked
        checkNameMethod = .ListItems("checkNameMethod").Checked
        checkNameVar = .ListItems("checkNameVar").Checked
        checkParamUse = .ListItems("checkParamUse").Checked
        checkMethodUse = .ListItems("checkMethodUse").Checked
        checkResume = .ListItems("checkResume").Checked
        checkScope = .ListItems("checkScope").Checked
        checkSquareBrackets = .ListItems("checkSquareBrackets").Checked
        checkStringConcat = .ListItems("checkStringConcat").Checked
        checkTodo = .ListItems("checkTodo").Checked
        checkType = .ListItems("checkType").Checked
        checkVarUse = .ListItems("checkVarUse").Checked
    End With
End Sub

Private Sub InitDicoMethods()
    Dim lv_Each As Variant
    
    Set gdic_UsfMethods = New Scripting.Dictionary
    Set gdic_WbkMethods = New Scripting.Dictionary
    Set gdic_WsMethods = New Scripting.Dictionary
    
    For Each lv_Each In Split(USF_METHODS, ",")
        gdic_UsfMethods.Add "USERFORM_" & UCase(lv_Each), vbNullString
    Next lv_Each
    
    For Each lv_Each In Split(WBK_METHODS, ",")
        gdic_WbkMethods.Add "WORKBOOK_" & UCase(lv_Each), vbNullString
    Next lv_Each
    
    For Each lv_Each In Split(WS_METHODS, ",")
        gdic_WsMethods.Add "WORKSHEET_" & UCase(lv_Each), vbNullString
    Next lv_Each
End Sub
