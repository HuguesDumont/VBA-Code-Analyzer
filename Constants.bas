Attribute VB_Name = "Constants"
Option Explicit

Public Const WBK_METHODS As String = "Open,Activate,BeforeClose,SheetActivate,SheetDeactivate,SheetChange,SheetSelectionChange,BeforeSave,Deactivate,AfterSave,WindowDeactivate,WindowActivate," & _
                                    "SheetBeforeRightClick,SheetBeforeDoubleClick,SheetCalculate,NewSheet,WindowResize,Sync,SheetPivotTableUpdate,SheetPivotTableChangeSync," & _
                                    "SheetPivotTableBeforeDiscardChanges,SheetPivotTableBeforeCommitChanges,SheetPivotTableBeforeAllocateChanges,SheetPivotTableAfterValueChange," & _
                                    "SheetFollowHyperlink,RowsetComplete,PivotTableOpenConnection,PivotTableCloseConnection," & _
                                    "NewChart,BeforeXmlImport,BeforeXmlExport,BeforePrint,AfterXmlImport,AfterXmlExport,AddinUninstall,AddinInstall"
                                    
Public Const WS_METHODS As String = "Change,SelectionChange,Activate,Deactivate,BeforeRightClick,BeforeDoubleClick,Calculate,PivotTableUpdate,PivotTableChangeSync,PivotTableBeforeDiscardChanges," & _
                                    "PivotTableBeforeCommitChanges,PivotTableBeforeAllocateChanges,PivotTableAfterValueChange,FollowHyperlink"

Public Const USF_METHODS As String = "Initialize,Activate,Deactivate,Terminate,QueryClose,MouseUp,MouseMove,MouseDown,KeyUp,KeyPress,KeyDown,Click,DblClick," & _
                                    "Resize,Scroll,BeforeDropOrPaste,BeforeDragOver,AddControl,RemoveControl,Zoom,Layout,Error"
                                    
Public Const HAUTE As String = "HAUTE"
Public Const MOYENNE As String = "MOYENNE"
Public Const BASSE As String = "BASSE"

Public Const PRIVATE_SCOPE As String = "Private "
Public Const PUBLIC_SCOPE As String = "Public "

Public Const VARIABLE As String = "Var"
Public Const CONSTANTE As String = "Const"
Public Const PARAMETER As String = "Param"
Public Const METHOD As String = "Method"

Public Const IMPLICIT_SCOPE As String = "Scope is implicit"
Public Const EXPLICIT_PUBLIC As String = "Explicitly declare scope as ""Public"""
Public Const EXPLICIT_PRIVATE As String = "Explicitly declare scope as ""Private"""

Public Const IMPLICIT_PASSING As String = "Implicit use of ByRef/ByVal/ParamArray"
Public Const EXPLICIT_PASSING As String = "Explicitly indicate passing argument method (ByRef/ByVal/ParamArray)"

Public Const IMPLICIT_RETURN As String = "Implicit return type declaration for Function/Property"
Public Const EXPLICIT_RETURN As String = "Explicitly declare return type"

Public Const IMPLICIT_TYPE As String = "Type declaration is implicit for "
Public Const EXPLICIT_TYPE As String = "Explicitly indicate type for "

Public Const IMPLICIT_WORKBOOK As String = "Call of ""Sheets""/""Worksheets"" is implicit"
Public Const EXPLICIT_WORKBOOK As String = "Explicitly indicate calling workbook ""Sheets""/""Worksheets"""

Public Const IMPLICIT_WORKSHEET As String = "Call of ""Range""/""Cells"" is implicit"
Public Const EXPLICIT_WORKSHEET As String = "Explicitly indicate calling Worksheet ""Range""/""Cells"""

Public Const NOTUSED As String = " is not used"
Public Const DELETE As String = "Delete "

Public Const INLINE As String = "Several instructions/declaration on a single line"
Public Const MULTILINE As String = "Each instruction/declaration should be done on a separate line"

Public Const MISSING_NEXT As String = """Next"" of For Loop doesn't refer to a variable"
Public Const ADD_NEXT As String = "Explictly indicate the looping var after ""Next"""

Public Const MISSING_OPTION As String = """Option Explicit"" is missing"
Public Const ADD_OPTION As String = "Add ""Option Explicit"" at first line of component"
Public Const OFFSET_OPTION As String = """Option explicit"" is present but not on first line of component"
Public Const MOVE_OPTION As String = "Move ""Option Explicit"" to first line of component"

Public Const CONCAT_PLUS As String = "String concatenation shouldn't use ""+"" symbol"
Public Const CONCAT_AND As String = "Use ""&"" symbol for string concatenation or convert string before summing"

Public Const EMPTY_METHOD As String = "Method is empty"

Public Const DEBUG_IN As String = "Debug shouldn't stay in project except for debug"
Public Const DEBUG_REMOVE As String = "Remove debug instruction"

Public Const SQUARE_BRACKETS As String = "Square brackets should never be used"
Public Const NO_SQUARE_BRACKETS As String = "Directly use what's in brackets instead of evaluating it"

Public Const PARENTHESIS_MISSING As String = "Boolean condition isn't encapsulated in parenthesis (ex : ""If x = 1 Then"")"
Public Const PARENTHESIS_SOL As String = "Boolean condition should always be encapsulated (ex : ""If (x = 1) Then"")"

Public Const RESUME_ISSUE As String = """Resume"" instruction shouldn't be used without ""On Error"""
Public Const RESUME_SOL As String = "Reformat code to avoid using of ""Resume"" instruction"

Public Const END_ISSUE As String = """End"" instruction alone should never be used"
Public Const END_SOL As String = "Reformat code to avoid using of ""End"" instruction"

Public Const GOTO_ISSUE As String = """GoTo"" instruction shouldn't be used (Except with ""On Error"")"
Public Const GOTO_SOL As String = "Refactor code to avoid use of ""GoTo"" instruction"

Public Const TODO_ISSUE As String = """TODO"" instruction should be removed when code is implemented"
Public Const TODO_SOL As String = "Remove the ""TODO"" comment"

Public Const NEGATED_BOOL As String = "Booleans shouldn't be negated using ""Not"""
Public Const NEGATION_BOOL As String = "Directly use boolean negation (Not True==>False / Not False==>True)"
Public Const CONDITIONAL_BOOL As String = "Comparison of boolean (var = True/False) value shouldn't be used in condition or affectation"
Public Const SIMPLE_BOOL As String = "Directly use value"

Public Const COMMENT_MULTI As String = "Comment continues previous line using "" _"" instead of starting with ""'"""
Public Const COMMENT_ONELINE As String = "Comment should always start with ""'"" and never end with "" _"""

Public Const CONST_NAMING_ISSUE As String = "Constants should comply with a naming convention"
Public Const CONST_NAMING As String = "Constants should respect the naming convention : ^([A-Z]{2}[A-Z_0-9]*[A-Z0-9]+)$"

Public Const VAR_NAMING_ISSUE As String = "Variables should comply with a naming convention"
Public Const VAR_NAMING As String = "Variables should respect the naming convention : ^([a-z][a-zA-Z_0-9]*[a-zA-Z0-9]+)$"

Public Const METHOD_NAMING_ISSUE As String = "Methods should comply with a naming convention"
Public Const METHOD_NAMING As String = "Methods should respect the naming convention : ^([A-Z][a-zA-Z_0-9]*[a-zA-Z0-9]+)$"

Public Const NAME_LENGTH_ISSUE As String = "Method, constant or variable is higher thant 20 characters"
Public Const NAME_LENGTH_SOL As String = "Methods, constants and variables should be lesser than 20 characters"
