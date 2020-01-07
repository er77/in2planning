Attribute VB_Name = "mSVCSVGeneralLIB"
' Attribute VB_Name = "SmartViewVBA"
' Copyright (c) 1992, 2017, Oracle and/or its affiliates. All Rights Reserved.
'
' RESTRICTED RIGHTS LEGEND:
' Use, duplication, or disclosure by the Government is subject to
' restrictions as set forth in subparagraph (c)(1)(ii) of the Rights
' in Technical Data and Computer Software clause at DFARS 252.227-7013,
' or in the Commercial Computer Software Restricted Rights clause at
' FAR 52.227-19, as applicable.
'
' Oracle Corporation
' 500 Oracle Parkway; Redwood Shores, CA, 94065 USA
'
' Function Smartview VBA Declaration.


'**************************************************************************
'  Global Constants for outline query types
'**************************************************************************

Global Const HYP_CHILDREN = 1
Global Const HYP_DESCENDANTS = 2
Global Const HYP_BOTTOMLEVEL = 3
Global Const HYP_SIBLINGS = 4
Global Const HYP_SAMELEVEL = 5
Global Const HYP_SAMEGENERATION = 6
Global Const HYP_PARENT = 7
Global Const HYP_DIMENSION = 8
Global Const HYP_NAMEDGENERATION = 9
Global Const HYP_NAMEDLEVEL = 10
Global Const HYP_SEARCH = 11
Global Const HYP_WILDSEARCH = 12
Global Const HYP_USERATTRIBUTE = 13
Global Const HYP_ANCESTORS = 14
Global Const HYP_DTSMEMBER = 15
Global Const HYP_DIMUSERATTRIBUTES = 16


'**************************************************************************
'  Global Constants for outline member query
'**************************************************************************

Global Const HYP_COMMENT = 1
Global Const HYP_FORMULA = 2
Global Const HYP_UDA = 3
Global Const HYP_ATTRIBUTE = 4


'**************************************************************************
'  Global Constants for outline member query options
'**************************************************************************

Global Const HYP_MEMBERSONLY = 1
Global Const HYP_ALIASESONLY = 2
Global Const HYP_MEMBERSANDALIASES = 4

'**************************************************************************
'  Global Constants for version info ids
'**************************************************************************

Global Const BUILD_VERSION = "VERSION"
Global Const BUILD_NUMBER = "BUILD NO"
Global Const BUILD_DATE = "BUILD DATE"
Global Const PRODUCT_ID = "PRODUCT ID"

'**************************************************************************
'  Global Constants for supported data providers
'**************************************************************************

Global Const HYP_ANALYTIC_SERVICES = "Analytic Provider Services"
Global Const HYP_FINANCIAL_MANAGEMENT = "Hyperion Financial Management"
Global Const HYP_ESSBASE = "Essbase"
Global Const HYP_PLANNING = "Planning"
Global Const HYP_OBIEE = "OBIEE"
Global Const HYP_ENTERPRISE = "Hyperion Enterprise"
Global Const HYP_RA = "Hyperion Smart View Provider for Hyperion Reporting and Analysis"

'**************************************************************************
' Global Constants for Member Information Property Name
'**************************************************************************
Global Const HYP_MI_NAME = "Name"
Global Const HYP_MI_DIM = "Dim"
Global Const HYP_MI_LEVEL = "Level"
Global Const HYP_MI_GENERATION = "Generation"
Global Const HYP_MI_PARENT_MEMBER_NAME = "ParentMbrName"
Global Const HYP_MI_CHILD_MEMBER_NAME = "ChildMbrName"
Global Const HYP_MI_PREVIOUS_MEMBER_NAME = "PrevMbrName"
Global Const HYP_MI_NEXT_MEMBER_NAME = "NextMbrName"
Global Const HYP_MI_CONSOLIDATION = "Consolidation"
Global Const HYP_MI_IS_TWO_PASS_CAL_MEMBER = "IsTwoPassCalcMbr"
Global Const HYP_MI_IS_EXPENSE_MEMBER = "IsExpenseMbr"
Global Const HYP_MI_CURRENCY_CONVERSION_TYPE = "CurrencyConversionType"
Global Const HYP_MI_CURRENCY_CATEGORY = "CurrencyCategory"
Global Const HYP_MI_TIME_BALANCE_OPTION = "TimeBalanceOption"
Global Const HYP_MI_TIME_BALANCE_SKIP_OPTION = "TimeBalanceSkipOption"
Global Const HYP_MI_SHARE_OPTION = "ShareOption"
Global Const HYP_MI_STORAGE_CATEGORY = "StorageCategory"
Global Const HYP_MI_CHILD_COUNT = "ChildCount"
Global Const HYP_MI_ATTRIBUTED = "Attributed"
Global Const HYP_MI_RELATIONAL_DESCENDANT_PRESENT = "RelDescendantPresent"
Global Const HYP_MI_RELATIONAL_PARTITION_ENABLED = "RelPartitionEnabled"
Global Const HYP_MI_DEFAULT_ALIAS = "DefaultAlias"
Global Const HYP_MI_HIERARCHY_TYPE = "HierarchyType"
Global Const HYP_MI_DIM_SOLVE_ORDER = "DimSolveOrder"
Global Const HYP_MI_IS_DUPLICATE_NAME = "IsDuplicateName"
Global Const HYP_MI_UNIQUE_NAME = "UniqueName"
Global Const HYP_MI_ORIGINAL_MEMBER = "OrigMember"
Global Const HYP_MI_IS_FLOW_TYPE = "IsFlowType"
Global Const HYP_MI_AGGREGATE_LEVEL = "AggLevel"
Global Const HYP_MI_FORMAT_STRING = "FormatString"
Global Const HYP_MI_ATTRIBUTE_DIMENSIONS = "AttributeDims"
Global Const HYP_MI_ATTRIBUTE_MEMBERS = "AttributeMbrs"
Global Const HYP_MI_ATTRIBUTE_TYPES = "AttributeTypes"
Global Const HYP_MI_ALIAS_NAMES = "AliasNames"
Global Const HYP_MI_ALIAS_TABLES = "AliasTables"
Global Const HYP_MI_FORMULA = "Formula"
Global Const HYP_MI_COMMENT = "Comment"
Global Const HYP_MI_LAST_FORMULA = "LastFormula"
Global Const HYP_MI_UDAS = "Udas"
Global Const vCurrPasswordLine = "MyVery1SecretPa@sswordAn$dEncry^ptionLine"

'**************************************************************************
' Global Constants for Doc Type returned from HypListDocuments
'**************************************************************************
Global Const HYP_LIST_DOC_FORM = "DOC_FORM"
Global Const HYP_LIST_DOC_FOLDER = "DOC_FOLDER"

'**************************************************************************
'   Enumeration of Linked Object Type
'**************************************************************************
Enum TYPE_OF_LRO
    CELL_NOTE_LRO = 1
    FILE_LRO = 2
    URL_LRO = 3
End Enum

'**************************************************************************
'  Enumeration of sheet content types
'**************************************************************************

Enum TYPE_OF_CONTENTS_IN_SHEET
    EMPTY_SHEET
    ADHOC_SHEET
    FORM_SHEET
    INTERACTIVE_REPORT_SHEET
End Enum


'**************************************************************************
' Enumeration of Smart View Error Codes
'**************************************************************************

Enum SmartViewErrors
 SS_ERR_ERROR = 4
 SS_NO_GRID_ON_SHEET_BUT_FUNCTIONS_SUBMITTED = 2
 SS_SHEET_NOT_CONNECTED_BUT_FUNCTIONS_SUBMITTED = 1
 SS_OK = 0
 SS_INIT_ERR = -1
 SS_TERM_ERR = -2
 SS_NOT_INIT = -3
 SS_NOT_CONNECTED = -4
 SS_NOT_LOCKED = -5
 SS_INVALID_SSTABLE = -6
 SS_INVALID_SSDATA = -7
 SS_NOUNDO_INFO = -8
 SS_CANCELED = -9
 SS_GLOBALOPTS = -10
 SS_SHEETOPTS = -11
 SS_NOTENABLED = -12
 SS_NO_MEMORY = -13
 SS_DIALOG_ERROR = -14
 SS_INVALID_PARAM = -15
 SS_CALCULATING = -16
 SS_SQL_IN_PROGRESS = -17
 SS_FORMULAPRESERVE = -18
 SS_INTERNALSSERROR = -19
 SS_INVALID_SHEET = -20
 SS_NOACTIVESHEET = -21
 SS_NOTCALCULATING = -22
 SS_INVALIDSELECTION = -23
 SS_INVALIDTOKEN = -24
 SS_CASCADENOTALLOWED = -25
 SS_NOMACROS = -26
 SS_NOREADONLYMACROS = -27
 SS_READONLYSS = -28
 SS_NOSQLACCESS = -29
 SS_MENUALREADYREMOVED = -30
 SS_MENUALREADYADDED = -31
 SS_NOSPREADSHEETACCESS = -32
 SS_NOHANDLES = -33
 SS_NOPREVCONNECTION = -34
 SS_LROERROR = -35
 SS_LROWINAPPACCESSERR = -36
 SS_DATANAVINITERR = -37
 SS_PARAMSETNOTALLOWED = -38
 SS_SHEET_PROTECTED = -39
 SS_CALCSCRIPT_NOTFOUND = -40
 SS_NOSUPPORT_PROVIDER = -41
 SS_INVALID_ALIAS = -42
 SS_CONN_NOT_FOUND = -43
 SS_APS_CONN_NOT_FOUND = -44
 SS_APS_NOT_CONNECTED = -45
 SS_APS_CANT_CONNECT = -46
 SS_CONN_ALREADY_EXISTS = -47
 SS_APS_URL_NOT_SAVED = -48
 SS_MIGRATION_OF_CONN_NOT_ALLOWED = -49
 SS_CONN_MGR_NOT_INITIALIZED = -50
 SS_FAILED_TO_GET_APS_OVERRIDE_PROPERTY = -51
 SS_FAILED_TO_SET_APS_OVERRIDE_PROPERTY = -52
 SS_FAILED_TO_GET_APS_URL = -53
 SS_APS_DISCONNECT_FAILED = -54
 SS_OPERATION_FAILED = -55
 SS_CANNOT_ASSOCIATE_SHEET_WITH_CONNECTION = -56
 SS_REFRESH_SHEET_NEEDED = -57
 SS_NO_GRID_OBJECT_ON_SHEET = -58
 SS_NO_CONNECTION_ASSOCIATED = -59
 SS_NON_DATA_CELL_PASSED = -60
 SS_DATA_CELL_IS_NOT_WRITABLE = -61
 SS_NO_SVC_CONTENT_ON_SHEET = -62
 SS_FAILED_TO_GET_OFFICE_OBJECT = -63
 SS_OP_FAILED_AS_CHART_IS_SELECTED = -64
 SS_EXCEL_IN_EDIT_MODE = -65
 SS_SHEET_NON_SMARTVIEW_COMPATIBLE = -66
 SS_APP_NOT_STANDALONE = -67
 SS_SMART_VIEW_DISABLED = -68
 SS_VBA_DEPRECATED = -69
 SS_OPERATION_NOT_SUPPORTED_IN_MULTIGRID_MODE = -70
    SS_INVALID_MEMBER = -71
SS_NO_SV_NAME_RANGE = -72
    SS_AMBIGUOUS_MENU = -73
End Enum

'**************************************************************************
'  Enumeration of options index to be used for HypGetOption/HypSetOption
'**************************************************************************

Enum HYP_SVC_OPTIONS_INDEX
    HSV_ZOOMIN = 1
    HSV_INCLUDE_SELECTION
    HSV_WITHIN_SELECTEDGROUP
    HSV_REMOVE_UNSELECTEDGROUP
    HSV_INDENTATION
    HSV_SUPPRESSROWS_MISSING
    HSV_SUPPRESSROWS_ZEROS
    HSV_SUPPRESSROWS_UNDERSCORE
    HSV_SUPPRESSROWS_NOACCESS
    HSV_SUPPRESSROWS_REPEATEDMEMBERS
    HSV_SUPPRESSROWS_INVALID
    HSV_ANCESTOR_POSITION
    HSV_MISSING_LABEL
    HSV_NOACCESS_LABEL
    HSV_CELL_STATUS
    HSV_MEMBER_DISPLAY
    HSV_INVALID_LABEL
    HSV_SUBMITZERO
    HSV_MOVEESSBASEMEMBERFORMULAONZOOM
    HSV_PRESERVE_ESSBASECOMMENT_UNKNOWNMEMBERS
    HSV_PRESERVE_FORMULA_COMMENT
    HSV_22
    HSV_FORMULA_FILL
    HSV_PRESERVE_FORMULA_ONPOVCHANGE
    HSV_EXCEL_FORMATTING = 30
    HSV_RETAIN_NUMERIC_FORMATTING
    HSV_THOUSAND_SEPARATOR
    HSV_NAVIGATE_WITHOUTDATA
    HSV_ENABLE_FORMATSTRING
    HSV_ENHANCED_COMMENT_HANDLING
    HSV_ADJUSTCOLUMNWIDTH
    HSV_DECIMALPLACES
    HSV_SCALE
    HSV_MOVEFORMATS_ON_ADHOC
    HSV_DISPLAY_INVALIDDATA
    HSV_SUPPRESSCOLUMNS_MISSING
    HSV_SUPPRESSCOLUMNS_ZEROS
    HSV_SUPPRESSCOLUMNS_NOACCESS
    HSV_SUPPRESS_MISSINGBLOCKS
    HSV_REPEATMEMBERS_IN_FORMS
    HSV_DOUBLECLICK_FOR_ADHOC = 101
    HSV_UNDO_ENABLE
    HSV_103
    HSV_LOGMESSAGE_DISPLAY
    HSV_ROUTE_LOGMESSAGE_TO_FILE
    HSV_CLEAR_LOG_ON_NEXTLAUNCH
    HSV_REDUCE_EXCEL_FILESIZE
    HSV_ENABLE_RIBBON_CONTEXT
    HSV_DISPLAY_HOMEPANEL_ONSTARTUP
    HSV_SHOW_COMMENTDIALOG_ON_REFRESH
    HSV_NUMBER_OF_UNDO_ACTION
    HSV_NUMBER_OF_MRU_ITEMS
    HSV_ROUTE_LOGMESSAGE_FILE_LOCATION
    HSV_DISABLE_SMARTVIEW_IN_OUTLOOK
    HSV_DISPLAY_SMARTVIEW_SHORTCUT_MENU_ONLY
    HSV_DISPLAY_DRILL_THROUGH_REPORT_TOOLTIP
    HSV_SHOW_PROGRESSINFORMATION
    HSV_PROGRESSINFO_TIMEDELAY
    HSV_ENABLE_PROFILING
    HSV_120 'reserved for refreshed linked workbook
    HSV_REFRESH_SELECTED_DEPENDENT_FUNCTIONS
    HSV_IMPROVE_METADATASTORAGE
End Enum


Enum DIMENSION_TYPE
    ROW_DIM = 0
    COL = 1
    POV = 2
    Page = 3
    USERVAR = 5
End Enum

'**************************************************************************
'  Enumeration of form attributes returned from HypListDocuments
'**************************************************************************
Enum FORM_ATTRIBUTES
    NO_ATTRIBUTE = -1
    HFM_BASIC_FORM = 0
    ADHOC_ENABLED = 8
    COMPOSITE_FORM = 16
    SMART_FORM = 128
    SAVED_ADHOC_GRID = 40
    SAVED_ADHOC_EXCLUSIVE_GRID = 104
    SMART_FORM_ADHOC_ENABLED = 136
End Enum

' This check is for 64 bit version of office.
#If VBA7 Then

'**************************************************************************
'  Menu Functions
'**************************************************************************

Public Declare PtrSafe Function HypMenuVAbout Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVAdjust Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVBusinessRules Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVCalculation Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVCellText Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVCollapse Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVConnect Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVCopyDataPoints Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVExpand Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVFunctionBuilder Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVInstruction Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVKeepOnly Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVMemberSelection Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVOptions Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVPasteDataPoints Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVPivot Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVPOVManager Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVQueryDesigner Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVRedo Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVRefresh Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVRefreshAll Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVRefreshOfflineDefinition Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVRemoveOnly Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVRulesOnForm Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVRunReport Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVSelectForm Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVShowHelpHtml Lib "HsAddin" (ByVal vtHelpPage As Variant) As Long
Public Declare PtrSafe Function HypMenuVSubmitData Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVSupportingDetails Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVSyncBack Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVTakeOffline Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVUndo Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVVisualizeinExcel Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVZoomIn Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVZoomOut Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVMigrate Lib "HsAddin" (ByVal vtOption As Variant, ByRef vtOutput As Variant) As Long
Public Declare PtrSafe Function HypMenuVLRO Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVCascadeSameWorkbook Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVCascadeNewWorkbook Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypMenuVMemberInformation Lib "HsAddin" () As Long
Public Declare PtrSafe Function HypExecuteMenu Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtMenuName As Variant) As Long
Public Declare PtrSafe Function HypHideRibbonMenu Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ParamArray vtMenus() As Variant) As Long
Public Declare PtrSafe Function HypHideRibbonMenuReset Lib "HsAddin" (ByVal vtSheetName As Variant) As Long


'**************************************************************************
'  General Functions
'**************************************************************************
                                                                                                                                                                                                                                                             

Type DOC_Info
    numDocs As Long
    docTypes As Variant
    docNames As Variant
    docDescriptions As Variant
    docPlanTypes As Variant
    docAttributes As Variant
End Type

Public Declare PtrSafe Function HypListDocuments Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtConnInfo As Variant, ByVal vtCompletePath As Variant, ByRef vtDocs As DOC_Info) As Long

Public Declare PtrSafe Function HypListApplications Lib "HsAddin" (ByVal vtURL As Variant, ByVal vtServerName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByRef vtApplications As Variant, ByRef vtDescriptions As Variant) As Long

Public Declare PtrSafe Function HypListDatabases Lib "HsAddin" (ByVal vtURL As Variant, ByVal vtServerName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtApplication As Variant, ByRef vtDatabases As Variant) As Long

Public Declare PtrSafe Function HypGetSheetInfo Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtItemNames As Variant, ByRef vtItemValues As Variant) As Long

Public Declare PtrSafe Function HypSetSSO Lib "HsAddin" (ByVal vtSSO As Variant) As Long
                                                  
Public Declare PtrSafe Function HypCopyMetaData Lib "HsAddin" (ByVal vtSourceSheetName As Variant, _
                                                           ByVal vtDestinationSheetName As Variant) As Long

Public Declare PtrSafe Function HypDeleteMetaData Lib "HsAddin" (ByVal vtDispObject As Variant, _
                                                             ByVal vtbWorkbook As Variant, _
                                                             ByVal vtbClearMetadataOnAllSheetsWithinWorkbook As Variant) As Long

Public Declare PtrSafe Function HypGetSubstitutionVariable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                      ByVal vtApplicationName As Variant, _
                                                                      ByVal vtDatabaseName As Variant, _
                                                                      ByVal vtVariableName As Variant, _
                                                                      ByRef vtVariableNames As Variant, _
                                                                      ByRef vtVariableValues As Variant) As Long

Public Declare PtrSafe Function HypIsDataModified Lib "HsAddin" (ByVal vtSheetName As Variant) As Boolean

Public Declare PtrSafe Function HypIsFreeForm Lib "HsAddin" (ByVal vtSheetName As Variant) As Boolean

Public Declare PtrSafe Function HypIsSmartViewContentPresent Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByRef vtTypeOfContentsInSheet As TYPE_OF_CONTENTS_IN_SHEET) As Boolean

Public Declare PtrSafe Function HypPreserveFormatting Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtSelectionRange As Variant) As Long

Public Declare PtrSafe Function HypRedo Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare PtrSafe Function HypRemovePreservedFormats Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByVal vtbRemoveAllCapturedFormats As Variant, _
                                                                    ByVal vtSelectionRange As Variant) As Long

Public Declare PtrSafe Function HypSetAliasTable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtAliasTableName As Variant) As Long

Public Declare PtrSafe Function HypSetMenu Lib "HsAddin" (ByVal bSetMenu As Boolean) As Long

Public Declare PtrSafe Function HypShowPov Lib "HsAddin" (ByVal bShowPov As Boolean) As Long

Public Declare PtrSafe Function HypSetSubstitutionVariable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByVal vtApplicationName As Variant, _
                                                                    ByVal vtDatabaseName As Variant, _
                                                                    ByVal vtVariableName As Variant, _
                                                                    ByVal vtVariableValue As Variant) As Long

Public Declare PtrSafe Function HypUndo Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare PtrSafe Function HypShowPanel Lib "HsAddin" (ByVal bShow As Boolean) As Long

Public Declare PtrSafe Function HypGetLastError Lib "HsAddin" (ByRef vtErrorCode As Variant, ByRef vtErrorMessage As Variant, ByRef vtErrorDescription As Variant) As Long

Public Declare PtrSafe Function HypGetVersion Lib "HsAddin" (ByVal vtID As Variant, _
                                                     ByRef vtValueList As Variant, ByVal vtVersionInfoFileCommand As Variant) As Long

Public Declare PtrSafe Function HypGetDatabaseNote Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtDBNote As Variant) As Long


'**************************************************************************
'  Connection Functions
'**************************************************************************

Public Declare PtrSafe Function HypConnect Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                  ByVal vtUserName As Variant, _
                                                  ByVal vtPassword As Variant, _
                                                  ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypUIConnect Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                  ByVal vtUserName As Variant, _
                                                  ByVal vtPassword As Variant, _
                                                  ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypConnected Lib "HsAddin" (ByVal vtSheetName As Variant) As Variant

Public Declare PtrSafe Function HypConnectionExists Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Variant

Public Declare PtrSafe Function HypCreateConnection Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtUserName As Variant, _
                                                           ByVal vtPassword As Variant, _
                                                           ByVal vtProvider As Variant, _
                                                           ByVal vtProviderURL As Variant, _
                                                           ByVal vtServerName As Variant, _
                                                           ByVal vtApplicationName As Variant, _
                                                           ByVal vtDatabaseName As Variant, _
                                                           ByVal vtFriendlyName As Variant, _
                                                           ByVal vtDescription As Variant) As Long
                                                           
Public Declare PtrSafe Function HypCreateConnectionEx Lib "HsAddin" (ByVal vtProviderType As Variant, _
                                                             ByVal vtServerName As Variant, _
                                                             ByVal vtApplicationName As Variant, _
                                                             ByVal vtDatabaseName As Variant, _
                                                             ByVal vtFormName As Variant, _
                                                             ByVal vtProviderURL As Variant, _
                                                             ByVal vtFriendlyName As Variant, _
                                                             ByVal vtUserName As Variant, _
                                                             ByVal vtPassword As Variant, _
                                                             ByVal vtDescription As Variant, _
                                                             ByVal vtReserved1 As Variant, _
                                                             ByVal vtReserved2 As Variant) As Long

Public Declare PtrSafe Function HypDisconnect Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal bLogoutUser As Boolean) As Long

Public Declare PtrSafe Function HypDisconnectAll Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypDisconnectEx Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypGetSharedConnectionsURL Lib "HsAddin" (ByRef vtSharedConnURL As Variant) As Long

Public Declare PtrSafe Function HypInvalidateSSO Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypIsConnectedToSharedConnections Lib "HsAddin" () As Variant

Public Declare PtrSafe Function HypRemoveConnection Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypResetFriendlyName Lib "HsAddin" (ByVal vtOldFriendlyName As Variant, _
                                                                ByVal vtNewFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypSetActiveConnection Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypSetAsDefault Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypSetConnAliasTable Lib "HsAddin" (ByVal vtFriendlyName As Variant, _
                                                            ByVal vtAliasTableName As Variant) As Long

Public Declare PtrSafe Function HypSetSharedConnectionsURL Lib "HsAddin" (ByVal vtSharedConnURL As Variant) As Long

Public Declare PtrSafe Function HypModifyConnection Lib "HsAddin" (ByVal vtDocumentName As Variant, _
                                                        ByVal vtSheetName As Variant, _
                                                        ByVal vtGridName As Variant, _
                                                        ByVal vtServer As Variant, _
                                                        ByVal vtURL As Variant, _
                                                        ByVal vtApp As Variant, _
                                                        ByVal vtDB As Variant, _
                                                        ByVal vtConnParam As Variant) As Long

Public Declare PtrSafe Function HypModifyRangeGridName Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtGridName As Variant, _
                                                        ByVal vtNewGridName) As Long

'**************************************************************************
'  Ad-Hoc Functions
'**************************************************************************

Public Declare PtrSafe Function HypExecuteQuery Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtMDXQuery As Variant) As Long

Public Declare PtrSafe Function HypKeepOnly Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                   ByVal vtSelection As Variant) As Long

Public Declare PtrSafe Function HypPivot Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                ByVal vtStart As Variant, _
                                                ByVal vtEnd As Variant) As Long
                                                
Public Declare PtrSafe Function HypPivotToGrid Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtDimensionName As Variant, _
                                                      ByVal vtSelection As Variant) As Long

Public Declare PtrSafe Function HypPivotToPOV Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal vtSelection As Variant) As Long

Public Declare PtrSafe Function HypRemoveOnly Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal vtSelection As Variant) As Long

Public Declare PtrSafe Function HypRetrieve Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare PtrSafe Function HypRetrieveRange Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtRange As Variant, _
                                                        ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypCreateRangeGrid Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtRange As Variant, _
                                                        ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypRetrieveNameRange Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtGridName As Variant) As Long

Public Declare PtrSafe Function HypGetNameRangeList Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtFriendlyName As Variant, _
                                                        ByRef vtNameList As Variant) As Long

Public Declare PtrSafe Function HypRetrieveAllWorkbooks Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypSubmitData Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare PtrSafe Function HypSubmitSelectedRangeWithoutRefresh Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                          ByVal vtSubmitBlankCellsAsMissing As Variant, _
                                                                          ByVal vtRefreshGridAfterSubmit As Variant, _
                                                                          ByVal vtUseWholeSheet As Variant) As Long

Public Declare PtrSafe Function HypSubmitSelectedDataCells Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                          ByVal vtDataRange As Variant, _
                                                                          ByVal vtSubmitBlankCellsAsMissingInFreeFormGrid As Variant) As Long

Public Declare PtrSafe Function HypZoomIn Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                 ByVal vtSelection As Variant, _
                                                 ByVal vtLevel As Variant, _
                                                 ByVal vtAcross As Variant) As Long

Public Declare PtrSafe Function HypZoomOut Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                  ByVal vtSelection As Variant) As Long


Public Declare PtrSafe Function HypPerformAdhocOnForm Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtFormName As Variant) As Long


'**************************************************************************
'  Form Functions
'**************************************************************************

Public Declare PtrSafe Function HypOpenForm Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                   ByVal vtFolderPath As Variant, _
                                                   ByVal vtFormName As Variant, _
                                                   ByRef vtDimensionList() As Variant, _
                                                   ByRef vtMemberList() As Variant) As Long


'**************************************************************************
'  Cell Functions
'**************************************************************************

Type LRO_Info
    lNumLRO As Long
    lNumDim As Long
    LROList As Variant
End Type


Public Declare PtrSafe Function HypCell Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                               ParamArray MemberList() As Variant) As Variant

Public Declare PtrSafe Function HypFreeDataPoint Lib "HsAddin" (ByRef vtInfo As Variant) As Long

Public Declare PtrSafe Function HypGetCellRangeForMbrCombination Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                        ByRef vtDimNames() As Variant, _
                                                                        ByRef vtMbrNames() As Variant, _
                                                                        ByRef vtCellIntersectionRange As Variant) As Long

Public Declare PtrSafe Function HypGetDataPoint Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtCell As Variant) As Variant

Public Declare PtrSafe Function HypGetDimMbrsForDataCell Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtCellRange As Variant, _
                                                                ByRef vtServerName As Variant, _
                                                                ByRef vtAppName As Variant, _
                                                                ByRef vtCubeName As Variant, _
                                                                ByRef vtFormName As Variant, _
                                                                ByRef vtDimensionNames As Variant, _
                                                                ByRef vtMemberNames As Variant) As Long

Public Declare PtrSafe Function HypIsCellWritable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtCellRange As Variant) As Boolean

Public Declare PtrSafe Function HypSetCellsDirty Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtRange As Variant) As Long


Public Declare PtrSafe Function HypDeleteAllLROs Lib "HsAddin" (ByVal vtSheetName As Variant, _
                            ByVal vtSelectionRange As Variant) As Long

Public Declare PtrSafe Function HypDeleteLROs Lib "HsAddin" (ByVal vtSheetName As Variant, _
                             ByVal vtSelectionRange As Variant, _
                             ByRef vtLROIDs() As Variant) As Long

Public Declare PtrSafe Function HypAddLRO Lib "HsAddin" (ByVal vtSheetName As Variant, _
                         ByVal vtSelectionRange As Variant, _
                         ByVal vtlType As Variant, _
                         ByVal vtName As Variant, _
                             ByVal vtDescription As Variant) As Long

Public Declare PtrSafe Function HypUpdateLRO Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                ByVal vtSelectionRange As Variant, _
                                ByVal vtID As Variant, _
                            ByVal vtlType As Variant, _
                                ByVal vtName As Variant, _
                            ByVal vtDescription As Variant) As Long


Public Declare PtrSafe Function HypListLROs Lib "HsAddin" (ByVal vtSheetName As Variant, _
                               ByVal vtSelectionRange As Variant, _
                               ByRef vtID As LRO_Info) As Long

Public Declare PtrSafe Function HypRetrieveLRO Lib "HsAddin" (ByVal vtSheetName As Variant, _
                              ByVal vtSelectionRange As Variant, _
                              ByVal vtID As Variant, _
                              ByRef vtName As Variant, _
                              ByRef vtDescription As Variant) As Long


Public Declare PtrSafe Function HypGetDrillThroughReports Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtSelectionRange As Variant, _
                                                      ByRef vtIDs As Variant, _
                                                      ByRef vtNames As Variant, _
                                                      ByRef vtURLs As Variant, _
                                                      ByRef vtURLTemplates As Variant, _
                                                      ByRef vtTypes As Variant) As Long
                                                      
                                                      
Public Declare PtrSafe Function HypExecuteDrillThroughReport Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtSelectionRange As Variant, _
                                                      ByVal vtID As Variant, _
                                                      ByVal vtName As Variant, _
                                                      ByVal vtURL As Variant, _
                                                      ByVal vtURLTemplate As Variant, _
                                                      ByVal vtType As Variant) As Long



'**************************************************************************
'  POV Functions
'**************************************************************************

Public Declare PtrSafe Function HypGetPagePOVChoices Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtDimensionName As Variant, _
                                                            ByRef vtMbrNameChoices As Variant, _
                                                            ByRef vtMbrDescChoices As Variant) As Long

Public Declare PtrSafe Function HypSetBackgroundPOV Lib "HsAddin" (ByVal vtFriendlyName As Variant, _
                                                           ParamArray MemberList() As Variant) As Long

Public Declare PtrSafe Function HypSetPages Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                   ParamArray MemberList() As Variant) As Long

Public Declare PtrSafe Function HypSetPOV Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                 ParamArray MemberList() As Variant) As Long

Public Declare PtrSafe Function HypSetMembers Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtDimensionName As Variant, _
                                                            ParamArray MemberList() As Variant) As Long
Public Declare PtrSafe Function HypGetPOV Lib "HsAddin" (ByVal vtSheetName, _
                                                 ByRef vtDimensionNames As Variant, _
                                                 ByRef vtMemberNames As Variant, ByRef vtType As Variant) As Long
Public Declare PtrSafe Function HypGetDimensions Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                 ByRef vtMemberNames As Variant, ByRef vtType As Variant) As Long
Public Declare PtrSafe Function HypGetMembers Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtDimensionName As Variant, _
                                                            ByRef vtMbrNameChoices As Variant, _
                                                            ByRef vtMbrDescChoices As Variant) As Long
                                                            
Public Declare PtrSafe Function HypSetDimensions Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                              ByRef vtDimNames() As Variant, _
                                                              ByRef vtTypes() As Variant) As Long

Public Declare PtrSafe Function HypGetBackgroundPOV Lib "HsAddin" (ByVal vtFriendlyName As Variant, _
                                                        ByRef vtDimensionNames As Variant, _
                                                        ByRef vtMemberNames As Variant) As Long
                            
Public Declare PtrSafe Function HypGetActiveMember Lib "HsAddin" (ByVal vtDimName As Variant, _
                                                        ByRef vtMember As Variant) As Long
                            
Public Declare PtrSafe Function HypSetActiveMember Lib "HsAddin" (ByVal vtDimName As Variant, _
                                                        ByVal vtMember As Variant) As Long

'**************************************************************************
'  Calculation Script / Business Rule Functions
'**************************************************************************

Public Declare PtrSafe Function HypDeleteCalc Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal vtApplicationName As Variant, _
                                                     ByVal vtDatabaseName As Variant, _
                                                     ByVal vtCalcScript As Variant) As Long

Public Declare PtrSafe Function HypExecuteCalcScript Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtCalcScript As Variant, _
                                                            ByVal vtSynchronous As Variant) As Long

Public Declare PtrSafe Function HypExecuteCalcScriptString Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtCalcScript As Variant, _
                                                            ByVal vtSubVars As Variant) As Long

Public Declare PtrSafe Function HypGetCalcScript Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtName As Variant, _
                                                            ByVal vtType As Variant, _
                                                            ByRef vtCalcScript As Variant) As Long

Public Declare PtrSafe Function HypExecuteCalcScriptEx2 Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtCalcScript As Variant) As Long
                                                                
Public Declare PtrSafe Function HypExecuteCalcScriptEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                              ByVal vtCubeName As Variant, _
                                                              ByVal vtBRName As Variant, _
                                                              ByVal vtBRType As Variant, _
                                                              ByVal vtbBRHasPrompts As Variant, _
                                                              ByVal vtbBRNeedPageInfo As Variant, _
                                                              ByRef vtRTPNames() As Variant, _
                                                              ByRef vtRTPValues() As Variant, _
                                                              ByVal vtbShowRTPDlg As Variant, _
                                                              ByVal vtbRuleOnForm As Variant, _
                                                              ByRef vtBRRanSuccessfully As Variant, _
                                                              ByRef vtCubeName As Variant, _
                                                              ByRef vtBRName As Variant, _
                                                              ByRef vtBRType As Variant, _
                                                              ByRef vtbBRHasPrompts As Variant, _
                                                              ByRef vtbBRNeedPageInfo As Variant, _
                                                              ByRef vtbBRHidePrompts As Variant, _
                                                              ByRef vtRTPNamesUsed As Variant, _
                                                              ByRef vtRTPValuesUsed As Variant) As Long

Public Declare PtrSafe Function HypListCalcScripts Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByRef scriptArray As Variant) As Long

Public Declare PtrSafe Function HypListCalcScriptsEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtbRuleOnForm As Variant, _
                                                            ByRef vtCubeNames As Variant, _
                                                            ByRef vtBRNames As Variant, _
                                                            ByRef vtBRTypes As Variant, _
                                                            ByRef vtBRHasPrompts As Variant, _
                                                            ByRef vtBRNeedsPageInfo As Variant, _
                                                            ByRef vtBRHidePrompts As Variant) As Long
                                                         

'**************************************************************************
'  Calculate / Consolidate / Translate Functions
'**************************************************************************

Public Declare PtrSafe Function HypCalculate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                    ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypCalculateContribution Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypConsolidate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypConsolidateAll Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypConsolidateAllWithData Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                 ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypForceCalculate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypForceCalculateContribution Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                     ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypForceTranslate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtRange As Variant) As Long

Public Declare PtrSafe Function HypTranslate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                    ByVal vtRange As Variant) As Long


'**************************************************************************
'  Member Query Functions
'**************************************************************************

Public Declare PtrSafe Function HypFindMember Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtMemberName As Variant, _
                                                         ByVal vtAliasTable As Variant, _
                                                         ByRef vtDimensionName As Variant, _
                                                         ByRef vtAliasName As Variant, _
                                                         ByRef vtGenerationName As Variant, _
                                                         ByRef vtLevelName As Variant) As Long

Public Declare PtrSafe Function HypFindMemberEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtAliasTable As Variant, _
                                                           ByRef vtDimensionName As Variant, _
                                                           ByRef vtAliasName As Variant, _
                                                           ByRef vtGenerationName As Variant, _
                                                           ByRef vtLevelName As Variant) As Long

Public Declare PtrSafe Function HypGetAncestor Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                          ByVal vtMemberName As Variant, _
                                                          ByVal vtLayerType As Variant, _
                                                          ByVal intLayerNumber As Integer, _
                                                          ByRef vtAncestor As Variant) As Long

Public Declare PtrSafe Function HypGetChildren Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                          ByVal vtMemberName As Variant, _
                                                          ByVal intChildCount As Integer, _
                                                          ByRef vtChildNameArray As Variant) As Long

Public Declare PtrSafe Function HypGetParent Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtMemberName As Variant, _
                                                        ByRef vtParentName As Variant) As Long

Public Declare PtrSafe Function HypIsAttribute Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                          ByVal vtDimensionName As Variant, _
                                                          ByVal vtMemberName As Variant, _
                                                          ByVal vtUDAString As Variant) As Variant

Public Declare PtrSafe Function HypIsDescendant Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtDescendantName As Variant) As Boolean

Public Declare PtrSafe Function HypIsAncestor Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtAncestorName As Variant) As Variant

Public Declare PtrSafe Function HypIsExpense Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtDimensionName As Variant, _
                                                        ByVal vtMemberName As Variant) As Variant

Public Declare PtrSafe Function HypIsParent Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtMemberName As Variant, _
                                                       ByVal ParentName As Variant) As Boolean

Public Declare PtrSafe Function HypIsChild Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtParentName As Variant, _
                                                       ByVal vtChildName As Variant) As Variant


Public Declare PtrSafe Function HypIsUDA Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                    ByVal vtDimensionName As Variant, _
                                                    ByVal vtMemberName As Variant, _
                                                    ByVal vtUDAString As Variant) As Variant

Public Declare PtrSafe Function HypOtlGetMemberInfo Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                               ByVal vtDimensionName As Variant, _
                                                               ByVal vtMemberName As Variant, _
                                                               ByVal vtPredicate As Variant, _
                                                               ByRef vtMemberArray As Variant) As Long

Public Declare PtrSafe Function HypQueryMembers Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtPredicate As Variant, _
                                                           ByVal vtOption As Variant, _
                                                           ByVal vtDimensionName As Variant, _
                                                           ByVal vtInput1 As Variant, _
                                                           ByVal vtInput2 As Variant, _
                                                           ByRef vtMemberArray As Variant) As Long


Public Declare PtrSafe Function HypGetMemberInformation Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                               ByVal vtMemberName As Variant, _
                                                               ByVal vtPropertyName As Variant, _
                                                               ByRef vtPropertyValue As Variant, _
                                   ByRef vtPropertyValueStrings As Variant) As Long


Public Declare PtrSafe Function HypGetMemberInformationEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                 ByVal vtMemberName As Variant, _
                                                                 ByRef vtPropertyNames As Variant, _
                                                                 ByRef vtPropertyValues As Variant, _
                                     ByRef vtPropertyValueStrings As Variant) As Long

'**************************************************************************
'  Option Functions
'**************************************************************************

Public Declare PtrSafe Function HypGetGlobalOption Lib "HsAddin" (ByVal vtItem As Long) As Variant

Public Declare PtrSafe Function HypGetSheetOption Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtItem As Variant) As Variant

Public Declare PtrSafe Function HypGetOption Lib "HsAddin" (ByVal vtItem As Variant, ByRef vtRet As Variant, ByVal vtSheetName As Variant) As Long

Public Declare PtrSafe Function HypSetGlobalOption Lib "HsAddin" (ByVal vtItem As Long, _
                                                          ByVal vtGlobalOption As Variant) As Long

Public Declare PtrSafe Function HypSetSheetOption Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtItem As Variant, _
                                                         ByVal vtOption As Variant) As Long

Public Declare PtrSafe Function HypSetOption Lib "HsAddin" (ByVal vtItem As Variant, _
                                                         ByVal vtOption As Variant, ByVal vtSheetName As Variant) As Long


Public Declare PtrSafe Function HypDeleteAllMRUItems Lib "HsAddin" () As Long


'**************************************************************************
'  Dynamic Link Functions
'**************************************************************************

Public Declare PtrSafe Function HypDisplayToLinkView Lib "HsAddin" (ByVal vtDocumentType As Variant, _
                                                            ByVal vtDocumentPath As Variant) As Long

Public Declare PtrSafe Function HypGetColCount Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypGetColItems Lib "HsAddin" (ByVal vtColID As Variant, _
                                                      ByRef vtDimensionName As Variant, _
                                                      ByRef vtMemberNames As Variant) As Long

Public Declare PtrSafe Function HypGetConnectionInfo Lib "HsAddin" (ByRef vtServerName As Variant, _
                                                            ByRef vtUserName As Variant, _
                                                            ByRef vtPassword As Variant, _
                                                            ByRef vtApplicationName As Variant, _
                                                            ByRef vtDatabaseName As Variant, _
                                                            ByRef vtFriendlyName As Variant, _
                                                            ByRef vtURL As Variant, _
                                                            ByRef vtProviderType As Variant) As Long

Public Declare PtrSafe Function HypGetLinkMacro Lib "HsAddin" (ByRef vtMacroName As Variant) As Long

Public Declare PtrSafe Function HypGetPOVCount Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypGetPOVItems Lib "HsAddin" (ByRef vtDimensionNames As Variant, _
                                                      ByRef vtPOVNames As Variant) As Long

Public Declare PtrSafe Function HypGetRowCount Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypGetRowItems Lib "HsAddin" (ByVal rowID As Variant, _
                                                      ByRef vtDimensionName As Variant, _
                                                      ByRef vtMemberNames As Variant) As Long

Public Declare PtrSafe Function HypGetSourceGrid Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByRef vtGrid As Variant) As Long

Public Declare PtrSafe Function HypSetColItems Lib "HsAddin" (ByVal vtColID As Variant, _
                                                      ByVal vtDimensionName As Variant, _
                                                      ParamArray MemberList() As Variant) As Long

Public Declare PtrSafe Function HypSetConnectionInfo Lib "HsAddin" (ByVal vtServerName As Variant, _
                                                            ByVal vtUserName As Variant, _
                                                            ByVal vtPassword As Variant, _
                                                            ByVal vtApplicationName As Variant, _
                                                            ByVal vtDatabaseName As Variant, _
                                                            ByVal vtFriendlyName As Variant, _
                                                            ByVal vtURL As Variant, _
                                                            ByVal vtProviderType As Variant) As Long

Public Declare PtrSafe Function HypSetLinkMacro Lib "HsAddin" (ByVal vtMacroName As Variant) As Long

Public Declare PtrSafe Function HypSetPOVItems Lib "HsAddin" (ParamArray MemberList() As Variant) As Long

Public Declare PtrSafe Function HypSetRowItems Lib "HsAddin" (ByVal vtRowID As Variant, _
                                                      ByVal vtDimensionName As Variant, _
                                                      ParamArray MemberList() As Variant) As Long

Public Declare PtrSafe Function HypUseLinkMacro Lib "HsAddin" (ByVal bUse As Boolean) As Long


'**************************************************************************
'  Deprecated Functions
'**************************************************************************

Public Declare PtrSafe Function HypCaptureFormatting Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtSelectionRange As Variant) As Long

Public Declare PtrSafe Function HypRemoveCapturedFormats Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByVal vtbRemoveAllCapturedFormats As Variant, _
                                                                    ByVal vtSelectionRange As Variant) As Long

Public Declare PtrSafe Function HypConnectToAPS Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypDisconnectFromAPS Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypGetCurrentAPSURL Lib "HsAddin" (ByRef vtAPSURL As Variant) As Long

Public Declare PtrSafe Function HypGetOverrideFlag Lib "HsAddin" (ByRef vtOverride As Boolean) As Long

Public Declare PtrSafe Function HypIsConnectedToAPS Lib "HsAddin" () As Long

Public Declare PtrSafe Function HypMigrateConnectionToDataSourceMgr Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare PtrSafe Function HypSetCurrentUserAPSURL Lib "HsAddin" (ByVal vtAPSURL As Variant) As Long

Public Declare PtrSafe Function HypSetOverrideFlag Lib "HsAddin" (ByVal vtOverride As Boolean) As Long

Public Declare PtrSafe Function HypMenuVVisualizeinHVE Lib "HsAddin" () As Long

'**************************************************************************
'**************************************************************************

'**************************************************************************
' ADVANCED MDX QUERY SECTION
'**************************************************************************

'**************************************************************************
' Type Declarations
'**************************************************************************

Type MDX_CELL
 CellValue As Double
 CellStatus As Long
End Type

Type MDX_PROPERTY
 PropertyName As Variant
 PropertyValue As Variant
 PropertyType As Variant
End Type

Type MDX_MEMBER
 MemberName As Variant
 NumClusters As Long
 NumProps As Long
 PropInfo() As MDX_PROPERTY
End Type

Type MDX_DIMENSION
 DimensionName As Variant
 NumProps As Long
 NumMembers As Long
 PropsInfo() As MDX_PROPERTY
 MemberInfo() As MDX_MEMBER
End Type

Type MDX_CLUSTER
 DimensionInfo() As MDX_DIMENSION
 TupleCount As Long
End Type

Type MDX_AXIS
 DimensionsCount As Long
 TuplesCount As Long
 ClustersCount As Long
 DimensionInfo() As MDX_DIMENSION
 ClusterInfo() As MDX_CLUSTER
End Type

Type MDX_AXES_NATIVE
 AxisCount As Long
 CellCount As Long
 AxisInfo As Variant
 CellInfo As Variant
End Type

Type MDX_AXES
 AxisCount As Long
 CellCount As Long
 AxisInfo() As MDX_AXIS
 CellInfo() As MDX_CELL
End Type

'**************************************************************************
' MDX Query Function
'**************************************************************************

Public Declare PtrSafe Function HypExecuteMDXEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtQuery As Variant, _
                                                       ByVal vtBoolHideData As Variant, _
                                                       ByVal vtBoolDataLess As Variant, _
                                                       ByVal vtBoolNeedStatus As Variant, _
                                                       ByVal vtMbrIDType As Variant, _
                                                       ByVal vtAliasTable As Variant, _
                                                       ByRef outResult As MDX_AXES_NATIVE) As Long 'Essbase

' For 32 bit version of office
#Else


'**************************************************************************
'  Menu Functions
'**************************************************************************

Public Declare Function HypMenuVAbout Lib "HsAddin" () As Long
Public Declare Function HypMenuVAdjust Lib "HsAddin" () As Long
Public Declare Function HypMenuVBusinessRules Lib "HsAddin" () As Long
Public Declare Function HypMenuVCalculation Lib "HsAddin" () As Long
Public Declare Function HypMenuVCellText Lib "HsAddin" () As Long
Public Declare Function HypMenuVCollapse Lib "HsAddin" () As Long
Public Declare Function HypMenuVConnect Lib "HsAddin" () As Long
Public Declare Function HypMenuVCopyDataPoints Lib "HsAddin" () As Long
Public Declare Function HypMenuVExpand Lib "HsAddin" () As Long
Public Declare Function HypMenuVFunctionBuilder Lib "HsAddin" () As Long
Public Declare Function HypMenuVInstruction Lib "HsAddin" () As Long
Public Declare Function HypMenuVKeepOnly Lib "HsAddin" () As Long
Public Declare Function HypMenuVMemberSelection Lib "HsAddin" () As Long
Public Declare Function HypMenuVOptions Lib "HsAddin" () As Long
Public Declare Function HypMenuVPasteDataPoints Lib "HsAddin" () As Long
Public Declare Function HypMenuVPivot Lib "HsAddin" () As Long
Public Declare Function HypMenuVPOVManager Lib "HsAddin" () As Long
Public Declare Function HypMenuVQueryDesigner Lib "HsAddin" () As Long
Public Declare Function HypMenuVRedo Lib "HsAddin" () As Long
Public Declare Function HypMenuVRefresh Lib "HsAddin" () As Long
Public Declare Function HypMenuVRefreshAll Lib "HsAddin" () As Long
Public Declare Function HypMenuVRefreshOfflineDefinition Lib "HsAddin" () As Long
Public Declare Function HypMenuVRemoveOnly Lib "HsAddin" () As Long
Public Declare Function HypMenuVRulesOnForm Lib "HsAddin" () As Long
Public Declare Function HypMenuVRunReport Lib "HsAddin" () As Long
Public Declare Function HypMenuVSelectForm Lib "HsAddin" () As Long
Public Declare Function HypMenuVShowHelpHtml Lib "HsAddin" (ByVal vtHelpPage As Variant) As Long
Public Declare Function HypMenuVSubmitData Lib "HsAddin" () As Long
Public Declare Function HypMenuVSupportingDetails Lib "HsAddin" () As Long
Public Declare Function HypMenuVSyncBack Lib "HsAddin" () As Long
Public Declare Function HypMenuVTakeOffline Lib "HsAddin" () As Long
Public Declare Function HypMenuVUndo Lib "HsAddin" () As Long
Public Declare Function HypMenuVVisualizeinExcel Lib "HsAddin" () As Long
Public Declare Function HypMenuVZoomIn Lib "HsAddin" () As Long
Public Declare Function HypMenuVZoomOut Lib "HsAddin" () As Long
Public Declare Function HypMenuVMigrate Lib "HsAddin" (ByVal vtOption As Variant, ByRef vtOutput As Variant) As Long
Public Declare Function HypMenuVLRO Lib "HsAddin" () As Long
Public Declare Function HypMenuVCascadeSameWorkbook Lib "HsAddin" () As Long
Public Declare Function HypMenuVCascadeNewWorkbook Lib "HsAddin" () As Long
Public Declare Function HypMenuVMemberInformation Lib "HsAddin" () As Long
Public Declare Function HypExecuteMenu Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtMenuName As Variant) As Long
Public Declare Function HypHideRibbonMenu Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ParamArray vtMenus() As Variant) As Long
Public Declare Function HypHideRibbonMenuReset Lib "HsAddin" (ByVal vtSheetName As Variant) As Long


'**************************************************************************
'  General Functions
'**************************************************************************
                                                                                                                                                                                                                                                     

Type DOC_Info
    numDocs As Long
    docTypes As Variant
    docNames As Variant
    docDescriptions As Variant
    docPlanTypes As Variant
    docAttributes As Variant
End Type

Public Declare Function HypListDocuments Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtConnInfo As Variant, ByVal vtCompletePath As Variant, ByRef vtDocs As DOC_Info) As Long

Public Declare Function HypListApplications Lib "HsAddin" (ByVal vtURL As Variant, ByVal vtServerName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByRef vtApplications As Variant, ByRef vtDescriptions As Variant) As Long

Public Declare Function HypListDatabases Lib "HsAddin" (ByVal vtURL As Variant, ByVal vtServerName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtApplication As Variant, ByRef vtDatabases As Variant) As Long

Public Declare Function HypGetSheetInfo Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtItemNames As Variant, ByRef vtItemValues As Variant) As Long

Public Declare Function HypSetSSO Lib "HsAddin" (ByVal vtSSO As Variant) As Long

Public Declare Function HypCopyMetaData Lib "HsAddin" (ByVal vtSourceSheetName As Variant, _
                                                           ByVal vtDestinationSheetName As Variant) As Long

Public Declare Function HypDeleteMetaData Lib "HsAddin" (ByVal vtDispObject As Variant, _
                                                             ByVal vtbWorkbook As Variant, _
                                                             ByVal vtbClearMetadataOnAllSheetsWithinWorkbook As Variant) As Long

Public Declare Function HypGetSubstitutionVariable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                      ByVal vtApplicationName As Variant, _
                                                                      ByVal vtDatabaseName As Variant, _
                                                                      ByVal vtVariableName As Variant, _
                                                                      ByRef vtVariableNames As Variant, _
                                                                      ByRef vtVariableValues As Variant) As Long

Public Declare Function HypIsDataModified Lib "HsAddin" (ByVal vtSheetName As Variant) As Boolean

Public Declare Function HypIsFreeForm Lib "HsAddin" (ByVal vtSheetName As Variant) As Boolean

Public Declare Function HypIsSmartViewContentPresent Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByRef vtTypeOfContentsInSheet As TYPE_OF_CONTENTS_IN_SHEET) As Boolean

Public Declare Function HypPreserveFormatting Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtSelectionRange As Variant) As Long

Public Declare Function HypRedo Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare Function HypRemovePreservedFormats Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByVal vtbRemoveAllCapturedFormats As Variant, _
                                                                    ByVal vtSelectionRange As Variant) As Long

Public Declare Function HypSetAliasTable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtAliasTableName As Variant) As Long

Public Declare Function HypSetMenu Lib "HsAddin" (ByVal bSetMenu As Boolean) As Long

Public Declare Function HypShowPov Lib "HsAddin" (ByVal bShowPov As Boolean) As Long

Public Declare Function HypSetSubstitutionVariable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByVal vtApplicationName As Variant, _
                                                                    ByVal vtDatabaseName As Variant, _
                                                                    ByVal vtVariableName As Variant, _
                                                                    ByVal vtVariableValue As Variant) As Long

Public Declare Function HypUndo Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare Function HypShowPanel Lib "HsAddin" (ByVal bShow As Boolean) As Long

Public Declare Function HypGetLastError Lib "HsAddin" (ByRef vtErrorCode As Variant, ByRef vtErrorMessage As Variant, ByRef vtErrorDescription As Variant) As Long

Public Declare Function HypGetVersion Lib "HsAddin" (ByVal vtID As Variant, _
                                                     ByRef vtValueList As Variant, ByVal vtVersionInfoFileCommand As Variant) As Long

Public Declare Function HypGetDatabaseNote Lib "HsAddin" (ByVal vtSheetName As Variant, ByRef vtDBNote As Variant) As Long


'**************************************************************************
'  Connection Functions
'**************************************************************************

Public Declare Function HypConnect Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                  ByVal vtUserName As Variant, _
                                                  ByVal vtPassword As Variant, _
                                                  ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypUIConnect Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                  ByVal vtUserName As Variant, _
                                                  ByVal vtPassword As Variant, _
                                                  ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypConnected Lib "HsAddin" (ByVal vtSheetName As Variant) As Variant

Public Declare Function HypConnectionExists Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Variant

Public Declare Function HypCreateConnection Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtUserName As Variant, _
                                                           ByVal vtPassword As Variant, _
                                                           ByVal vtProvider As Variant, _
                                                           ByVal vtProviderURL As Variant, _
                                                           ByVal vtServerName As Variant, _
                                                           ByVal vtApplicationName As Variant, _
                                                           ByVal vtDatabaseName As Variant, _
                                                           ByVal vtFriendlyName As Variant, _
                                                           ByVal vtDescription As Variant) As Long
                                                           
Public Declare Function HypCreateConnectionEx Lib "HsAddin" (ByVal vtProviderType As Variant, _
                                                             ByVal vtServerName As Variant, _
                                                             ByVal vtApplicationName As Variant, _
                                                             ByVal vtDatabaseName As Variant, _
                                                             ByVal vtFormName As Variant, _
                                                             ByVal vtProviderURL As Variant, _
                                                             ByVal vtFriendlyName As Variant, _
                                                             ByVal vtUserName As Variant, _
                                                             ByVal vtPassword As Variant, _
                                                             ByVal vtDescription As Variant, _
                                                             ByVal vtReserved1 As Variant, _
                                                             ByVal vtReserved2 As Variant) As Long

Public Declare Function HypDisconnect Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal bLogoutUser As Boolean) As Long

Public Declare Function HypDisconnectAll Lib "HsAddin" () As Long

Public Declare Function HypDisconnectEx Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypGetSharedConnectionsURL Lib "HsAddin" (ByRef vtSharedConnURL As Variant) As Long

Public Declare Function HypInvalidateSSO Lib "HsAddin" () As Long

Public Declare Function HypIsConnectedToSharedConnections Lib "HsAddin" () As Variant

Public Declare Function HypRemoveConnection Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypResetFriendlyName Lib "HsAddin" (ByVal vtOldFriendlyName As Variant, _
                                                                ByVal vtNewFriendlyName As Variant) As Long

Public Declare Function HypSetActiveConnection Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypSetAsDefault Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypSetConnAliasTable Lib "HsAddin" (ByVal vtFriendlyName As Variant, _
                                                            ByVal vtAliasTableName As Variant) As Long

Public Declare Function HypSetSharedConnectionsURL Lib "HsAddin" (ByVal vtSharedConnURL As Variant) As Long

Public Declare Function HypModifyConnection Lib "HsAddin" (ByVal vtDocumentName As Variant, _
                                                        ByVal vtSheetName As Variant, _
                                                        ByVal vtGridName As Variant, _
                                                        ByVal vtServer As Variant, _
                                                        ByVal vtURL As Variant, _
                                                        ByVal vtApp As Variant, _
                                                        ByVal vtDB As Variant, _
                                                        ByVal vtConnParam As Variant) As Long

Public Declare Function HypModifyRangeGridName Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtGridName As Variant, _
                                                        ByVal vtNewGridName) As Long

'**************************************************************************
'  Ad-Hoc Functions
'**************************************************************************

Public Declare Function HypExecuteQuery Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtMDXQuery As Variant) As Long

Public Declare Function HypKeepOnly Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                   ByVal vtSelection As Variant) As Long

Public Declare Function HypPivot Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                ByVal vtStart As Variant, _
                                                ByVal vtEnd As Variant) As Long
                                                
Public Declare Function HypPivotToGrid Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtDimensionName As Variant, _
                                                      ByVal vtSelection As Variant) As Long

Public Declare Function HypPivotToPOV Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal vtSelection As Variant) As Long

Public Declare Function HypRemoveOnly Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal vtSelection As Variant) As Long

Public Declare Function HypRetrieve Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare Function HypRetrieveRange Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtRange As Variant, _
                                                        ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypCreateRangeGrid Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtRange As Variant, _
                                                        ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypRetrieveNameRange Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtGridName As Variant) As Long

Public Declare Function HypGetNameRangeList Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtFriendlyName As Variant, _
                                                        ByRef vtNameList As Variant) As Long

Public Declare Function HypRetrieveAllWorkbooks Lib "HsAddin" () As Long

Public Declare Function HypSubmitData Lib "HsAddin" (ByVal vtSheetName As Variant) As Long

Public Declare Function HypSubmitSelectedRangeWithoutRefresh Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                          ByVal vtSubmitBlankCellsAsMissing As Variant, _
                                                                          ByVal vtRefreshGridAfterSubmit As Variant, _
                                                                          ByVal vtUseWholeSheet As Variant) As Long

Public Declare Function HypSubmitSelectedDataCells Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                          ByVal vtDataRange As Variant, _
                                                                          ByVal vtSubmitBlankCellsAsMissingInFreeFormGrid As Variant) As Long

Public Declare Function HypZoomIn Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                 ByVal vtSelection As Variant, _
                                                 ByVal vtLevel As Variant, _
                                                 ByVal vtAcross As Variant) As Long

Public Declare Function HypZoomOut Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                  ByVal vtSelection As Variant) As Long


Public Declare Function HypPerformAdhocOnForm Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtFormName As Variant) As Long


'**************************************************************************
'  Form Functions
'**************************************************************************

Public Declare Function HypOpenForm Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                   ByVal vtFolderPath As Variant, _
                                                   ByVal vtFormName As Variant, _
                                                   ByRef vtDimensionList() As Variant, _
                                                   ByRef vtMemberList() As Variant) As Long


'**************************************************************************
'  Cell Functions
'**************************************************************************

Type LRO_Info
    lNumLRO As Long
    lNumDim As Long
    LROList As Variant
End Type


Public Declare Function HypCell Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                               ParamArray MemberList() As Variant) As Variant

Public Declare Function HypFreeDataPoint Lib "HsAddin" (ByRef vtInfo As Variant) As Long

Public Declare Function HypGetCellRangeForMbrCombination Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                        ByRef vtDimNames() As Variant, _
                                                                        ByRef vtMbrNames() As Variant, _
                                                                        ByRef vtCellIntersectionRange As Variant) As Long

Public Declare Function HypGetDataPoint Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtCell As Variant) As Variant

Public Declare Function HypGetDimMbrsForDataCell Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtCellRange As Variant, _
                                                                ByRef vtServerName As Variant, _
                                                                ByRef vtAppName As Variant, _
                                                                ByRef vtCubeName As Variant, _
                                                                ByRef vtFormName As Variant, _
                                                                ByRef vtDimensionNames As Variant, _
                                                                ByRef vtMemberNames As Variant) As Long

Public Declare Function HypIsCellWritable Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtCellRange As Variant) As Boolean

Public Declare Function HypSetCellsDirty Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtRange As Variant) As Long


Public Declare Function HypDeleteAllLROs Lib "HsAddin" (ByVal vtSheetName As Variant, _
                            ByVal vtSelectionRange As Variant) As Long

Public Declare Function HypDeleteLROs Lib "HsAddin" (ByVal vtSheetName As Variant, _
                             ByVal vtSelectionRange As Variant, _
                             ByRef vtLROIDs() As Variant) As Long

Public Declare Function HypAddLRO Lib "HsAddin" (ByVal vtSheetName As Variant, _
                         ByVal vtSelectionRange As Variant, _
                         ByVal vtlType As Variant, _
                         ByVal vtName As Variant, _
                             ByVal vtDescription As Variant) As Long

Public Declare Function HypUpdateLRO Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                ByVal vtSelectionRange As Variant, _
                                ByVal vtID As Variant, _
                            ByVal vtlType As Variant, _
                                ByVal vtName As Variant, _
                            ByVal vtDescription As Variant) As Long


Public Declare Function HypListLROs Lib "HsAddin" (ByVal vtSheetName As Variant, _
                               ByVal vtSelectionRange As Variant, _
                               ByRef vtID As LRO_Info) As Long

Public Declare Function HypRetrieveLRO Lib "HsAddin" (ByVal vtSheetName As Variant, _
                              ByVal vtSelectionRange As Variant, _
                              ByVal vtID As Variant, _
                              ByRef vtName As Variant, _
                              ByRef vtDescription As Variant) As Long


Public Declare Function HypGetDrillThroughReports Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtSelectionRange As Variant, _
                                                      ByRef vtIDs As Variant, _
                                                      ByRef vtNames As Variant, _
                                                      ByRef vtURLs As Variant, _
                                                      ByRef vtURLTemplates As Variant, _
                                                      ByRef vtTypes As Variant) As Long
                                                      
                                                      
Public Declare Function HypExecuteDrillThroughReport Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtSelectionRange As Variant, _
                                                      ByVal vtID As Variant, _
                                                      ByVal vtName As Variant, _
                                                      ByVal vtURL As Variant, _
                                                      ByVal vtURLTemplate As Variant, _
                                                      ByVal vtType As Variant) As Long



'**************************************************************************
'  POV Functions
'**************************************************************************

Public Declare Function HypGetPagePOVChoices Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtDimensionName As Variant, _
                                                            ByRef vtMbrNameChoices As Variant, _
                                                            ByRef vtMbrDescChoices As Variant) As Long

Public Declare Function HypSetBackgroundPOV Lib "HsAddin" (ByVal vtFriendlyName As Variant, _
                                                           ParamArray MemberList() As Variant) As Long

Public Declare Function HypSetPages Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                   ParamArray MemberList() As Variant) As Long

Public Declare Function HypSetPOV Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                 ParamArray MemberList() As Variant) As Long

Public Declare Function HypSetMembers Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtDimensionName As Variant, _
                                                            ParamArray MemberList() As Variant) As Long
Public Declare Function HypGetPOV Lib "HsAddin" (ByVal vtSheetName, _
                                                 ByRef vtDimensionNames As Variant, _
                                                 ByRef vtMemberNames As Variant, ByRef vtType As Variant) As Long
Public Declare Function HypGetDimensions Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                 ByRef vtMemberNames As Variant, ByRef vtType As Variant) As Long
Public Declare Function HypGetMembers Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtDimensionName As Variant, _
                                                            ByRef vtMbrNameChoices As Variant, _
                                                            ByRef vtMbrDescChoices As Variant) As Long
                                                            
Public Declare Function HypSetDimensions Lib "HsAddin" (ByVal vtSheetName, _
                                                              ByRef vtDimNames() As Variant, _
                                                              ByRef vtTypes() As Variant) As Long

Public Declare Function HypGetBackgroundPOV Lib "HsAddin" (ByVal vtFriendlyName As Variant, _
                                                        ByRef vtDimensionNames As Variant, _
                                                        ByRef vtMemberNames As Variant) As Long
                            
Public Declare Function HypGetActiveMember Lib "HsAddin" (ByVal vtDimName As Variant, _
                                                        ByRef vtMember As Variant) As Long
                            
Public Declare Function HypSetActiveMember Lib "HsAddin" (ByVal vtDimName As Variant, _
                                                        ByVal vtMember As Variant) As Long

'**************************************************************************
'  Calculation Script / Business Rule Functions
'**************************************************************************

Public Declare Function HypDeleteCalc Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                     ByVal vtApplicationName As Variant, _
                                                     ByVal vtDatabaseName As Variant, _
                                                     ByVal vtCalcScript As Variant) As Long

Public Declare Function HypExecuteCalcScript Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtCalcScript As Variant, _
                                                            ByVal vtSynchronous As Variant) As Long

Public Declare Function HypExecuteCalcScriptString Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtCalcScript As Variant, _
                                                            ByVal vtSubVars As Variant) As Long

Public Declare Function HypGetCalcScript Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtName As Variant, _
                                                            ByVal vtType As Variant, _
                                                            ByRef vtCalcScript As Variant) As Long

Public Declare Function HypExecuteCalcScriptEx2 Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtCalcScript As Variant) As Long

                                                                
Public Declare Function HypExecuteCalcScriptEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                              ByVal vtCubeName As Variant, _
                                                              ByVal vtBRName As Variant, _
                                                              ByVal vtBRType As Variant, _
                                                              ByVal vtbBRHasPrompts As Variant, _
                                                              ByVal vtbBRNeedPageInfo As Variant, _
                                                              ByRef vtRTPNames() As Variant, _
                                                              ByRef vtRTPValues() As Variant, _
                                                              ByVal vtbShowRTPDlg As Variant, _
                                                              ByVal vtbRuleOnForm As Variant, _
                                                              ByRef vtBRRanSuccessfully As Variant, _
                                                              ByRef vtCubeName As Variant, _
                                                              ByRef vtBRName As Variant, _
                                                              ByRef vtBRType As Variant, _
                                                              ByRef vtbBRHasPrompts As Variant, _
                                                              ByRef vtbBRNeedPageInfo As Variant, _
                                                              ByRef vtbBRHidePrompts As Variant, _
                                                              ByRef vtRTPNamesUsed As Variant, _
                                                              ByRef vtRTPValuesUsed As Variant) As Long

Public Declare Function HypListCalcScripts Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByRef scriptArray As Variant) As Long

Public Declare Function HypListCalcScriptsEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                            ByVal vtbRuleOnForm As Variant, _
                                                            ByRef vtCubeNames As Variant, _
                                                            ByRef vtBRNames As Variant, _
                                                            ByRef vtBRTypes As Variant, _
                                                            ByRef vtBRHasPrompts As Variant, _
                                                            ByRef vtBRNeedsPageInfo As Variant, _
                                                            ByRef vtBRHidePrompts As Variant) As Long
                                                         

'**************************************************************************
'  Calculate / Consolidate / Translate Functions
'**************************************************************************

Public Declare Function HypCalculate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                    ByVal vtRange As Variant) As Long

Public Declare Function HypCalculateContribution Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtRange As Variant) As Long

Public Declare Function HypConsolidate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                      ByVal vtRange As Variant) As Long

Public Declare Function HypConsolidateAll Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtRange As Variant) As Long

Public Declare Function HypConsolidateAllWithData Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                 ByVal vtRange As Variant) As Long

Public Declare Function HypForceCalculate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtRange As Variant) As Long

Public Declare Function HypForceCalculateContribution Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                     ByVal vtRange As Variant) As Long

Public Declare Function HypForceTranslate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtRange As Variant) As Long

Public Declare Function HypTranslate Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                    ByVal vtRange As Variant) As Long


'**************************************************************************
'  Member Query Functions
'**************************************************************************

Public Declare Function HypFindMember Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtMemberName As Variant, _
                                                         ByVal vtAliasTable As Variant, _
                                                         ByRef vtDimensionName As Variant, _
                                                         ByRef vtAliasName As Variant, _
                                                         ByRef vtGenerationName As Variant, _
                                                         ByRef vtLevelName As Variant) As Long

Public Declare Function HypFindMemberEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtAliasTable As Variant, _
                                                           ByRef vtDimensionName As Variant, _
                                                           ByRef vtAliasName As Variant, _
                                                           ByRef vtGenerationName As Variant, _
                                                           ByRef vtLevelName As Variant) As Long

Public Declare Function HypGetAncestor Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                          ByVal vtMemberName As Variant, _
                                                          ByVal vtLayerType As Variant, _
                                                          ByVal intLayerNumber As Integer, _
                                                          ByRef vtAncestor As Variant) As Long

Public Declare Function HypGetChildren Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                          ByVal vtMemberName As Variant, _
                                                          ByVal intChildCount As Integer, _
                                                          ByRef vtChildNameArray As Variant) As Long

Public Declare Function HypGetParent Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtMemberName As Variant, _
                                                        ByRef vtParentName As Variant) As Long

Public Declare Function HypIsAttribute Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                          ByVal vtDimensionName As Variant, _
                                                          ByVal vtMemberName As Variant, _
                                                          ByVal vtUDAString As Variant) As Variant

Public Declare Function HypIsDescendant Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtDescendantName As Variant) As Boolean

Public Declare Function HypIsAncestor Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtAncestorName As Variant) As Variant

Public Declare Function HypIsExpense Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByVal vtDimensionName As Variant, _
                                                        ByVal vtMemberName As Variant) As Variant

Public Declare Function HypIsParent Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtMemberName As Variant, _
                                                       ByVal ParentName As Variant) As Boolean

Public Declare Function HypIsChild Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtParentName As Variant, _
                                                       ByVal vtChildName As Variant) As Variant


Public Declare Function HypIsUDA Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                    ByVal vtDimensionName As Variant, _
                                                    ByVal vtMemberName As Variant, _
                                                    ByVal vtUDAString As Variant) As Variant

Public Declare Function HypOtlGetMemberInfo Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                               ByVal vtDimensionName As Variant, _
                                                               ByVal vtMemberName As Variant, _
                                                               ByVal vtPredicate As Variant, _
                                                               ByRef vtMemberArray As Variant) As Long

Public Declare Function HypQueryMembers Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                           ByVal vtMemberName As Variant, _
                                                           ByVal vtPredicate As Variant, _
                                                           ByVal vtOption As Variant, _
                                                           ByVal vtDimensionName As Variant, _
                                                           ByVal vtInput1 As Variant, _
                                                           ByVal vtInput2 As Variant, _
                                                           ByRef vtMemberArray As Variant) As Long


Public Declare Function HypGetMemberInformation Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                               ByVal vtMemberName As Variant, _
                                                               ByVal vtPropertyName As Variant, _
                                                               ByRef vtPropertyValue As Variant, _
                                   ByRef vtPropertyValueStrings As Variant) As Long


Public Declare Function HypGetMemberInformationEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                 ByVal vtMemberName As Variant, _
                                                                 ByRef vtPropertyNames As Variant, _
                                                                 ByRef vtPropertyValues As Variant, _
                                     ByRef vtPropertyValueStrings As Variant) As Long

'**************************************************************************
'  Option Functions
'**************************************************************************

Public Declare Function HypGetGlobalOption Lib "HsAddin" (ByVal vtItem As Long) As Variant

Public Declare Function HypGetSheetOption Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtItem As Variant) As Variant

Public Declare Function HypGetOption Lib "HsAddin" (ByVal vtItem As Variant, ByRef vtRet As Variant, ByVal vtSheetName As Variant) As Long

Public Declare Function HypSetGlobalOption Lib "HsAddin" (ByVal vtItem As Long, _
                                                          ByVal vtGlobalOption As Variant) As Long

Public Declare Function HypSetSheetOption Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                         ByVal vtItem As Variant, _
                                                         ByVal vtOption As Variant) As Long

Public Declare Function HypSetOption Lib "HsAddin" (ByVal vtItem As Variant, _
                                                         ByVal vtOption As Variant, ByVal vtSheetName As Variant) As Long


Public Declare Function HypDeleteAllMRUItems Lib "HsAddin" () As Long


'**************************************************************************
'  Dynamic Link Functions
'**************************************************************************

Public Declare Function HypDisplayToLinkView Lib "HsAddin" (ByVal vtDocumentType As Variant, _
                                                            ByVal vtDocumentPath As Variant) As Long

Public Declare Function HypGetColCount Lib "HsAddin" () As Long

Public Declare Function HypGetColItems Lib "HsAddin" (ByVal vtColID As Variant, _
                                                      ByRef vtDimensionName As Variant, _
                                                      ByRef vtMemberNames As Variant) As Long

Public Declare Function HypGetConnectionInfo Lib "HsAddin" (ByRef vtServerName As Variant, _
                                                            ByRef vtUserName As Variant, _
                                                            ByRef vtPassword As Variant, _
                                                            ByRef vtApplicationName As Variant, _
                                                            ByRef vtDatabaseName As Variant, _
                                                            ByRef vtFriendlyName As Variant, _
                                                            ByRef vtURL As Variant, _
                                                            ByRef vtProviderType As Variant) As Long

Public Declare Function HypGetLinkMacro Lib "HsAddin" (ByRef vtMacroName As Variant) As Long

Public Declare Function HypGetPOVCount Lib "HsAddin" () As Long

Public Declare Function HypGetPOVItems Lib "HsAddin" (ByRef vtDimensionNames As Variant, _
                                                      ByRef vtPOVNames As Variant) As Long

Public Declare Function HypGetRowCount Lib "HsAddin" () As Long

Public Declare Function HypGetRowItems Lib "HsAddin" (ByVal rowID As Variant, _
                                                      ByRef vtDimensionName As Variant, _
                                                      ByRef vtMemberNames As Variant) As Long

Public Declare Function HypGetSourceGrid Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                        ByRef vtGrid As Variant) As Long

Public Declare Function HypSetColItems Lib "HsAddin" (ByVal vtColID As Variant, _
                                                      ByVal vtDimensionName As Variant, _
                                                      ParamArray MemberList() As Variant) As Long

Public Declare Function HypSetConnectionInfo Lib "HsAddin" (ByVal vtServerName As Variant, _
                                                            ByVal vtUserName As Variant, _
                                                            ByVal vtPassword As Variant, _
                                                            ByVal vtApplicationName As Variant, _
                                                            ByVal vtDatabaseName As Variant, _
                                                            ByVal vtFriendlyName As Variant, _
                                                            ByVal vtURL As Variant, _
                                                            ByVal vtProviderType As Variant) As Long

Public Declare Function HypSetLinkMacro Lib "HsAddin" (ByVal vtMacroName As Variant) As Long

Public Declare Function HypSetPOVItems Lib "HsAddin" (ParamArray MemberList() As Variant) As Long

Public Declare Function HypSetRowItems Lib "HsAddin" (ByVal vtRowID As Variant, _
                                                      ByVal vtDimensionName As Variant, _
                                                      ParamArray MemberList() As Variant) As Long

Public Declare Function HypUseLinkMacro Lib "HsAddin" (ByVal bUse As Boolean) As Long


'**************************************************************************
'  Deprecated Functions
'**************************************************************************

Public Declare Function HypCaptureFormatting Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                ByVal vtSelectionRange As Variant) As Long

Public Declare Function HypRemoveCapturedFormats Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                                    ByVal vtbRemoveAllCapturedFormats As Variant, _
                                                                    ByVal vtSelectionRange As Variant) As Long

Public Declare Function HypConnectToAPS Lib "HsAddin" () As Long

Public Declare Function HypDisconnectFromAPS Lib "HsAddin" () As Long

Public Declare Function HypGetCurrentAPSURL Lib "HsAddin" (ByRef vtAPSURL As Variant) As Long

Public Declare Function HypGetOverrideFlag Lib "HsAddin" (ByRef vtOverride As Boolean) As Long

Public Declare Function HypIsConnectedToAPS Lib "HsAddin" () As Long

Public Declare Function HypMigrateConnectionToDataSourceMgr Lib "HsAddin" (ByVal vtFriendlyName As Variant) As Long

Public Declare Function HypSetCurrentUserAPSURL Lib "HsAddin" (ByVal vtAPSURL As Variant) As Long

Public Declare Function HypSetOverrideFlag Lib "HsAddin" (ByVal vtOverride As Boolean) As Long

Public Declare Function HypMenuVVisualizeinHVE Lib "HsAddin" () As Long

'**************************************************************************
'**************************************************************************

'**************************************************************************
' ADVANCED MDX QUERY SECTION
'**************************************************************************

'**************************************************************************
' Type Declarations
'**************************************************************************

Type MDX_CELL
 CellValue As Double
 CellStatus As Long
End Type

Type MDX_PROPERTY
 PropertyName As Variant
 PropertyValue As Variant
 PropertyType As Variant
End Type

Type MDX_MEMBER
 MemberName As Variant
 NumClusters As Long
 NumProps As Long
 PropInfo() As MDX_PROPERTY
End Type

Type MDX_DIMENSION
 DimensionName As Variant
 NumProps As Long
 NumMembers As Long
 PropsInfo() As MDX_PROPERTY
 MemberInfo() As MDX_MEMBER
End Type

Type MDX_CLUSTER
 DimensionInfo() As MDX_DIMENSION
 TupleCount As Long
End Type

Type MDX_AXIS
 DimensionsCount As Long
 TuplesCount As Long
 ClustersCount As Long
 DimensionInfo() As MDX_DIMENSION
 ClusterInfo() As MDX_CLUSTER
End Type

Type MDX_AXES_NATIVE
 AxisCount As Long
 CellCount As Long
 AxisInfo As Variant
 CellInfo As Variant
End Type

Type MDX_AXES
 AxisCount As Long
 CellCount As Long
 AxisInfo() As MDX_AXIS
 CellInfo() As MDX_CELL
End Type

'**************************************************************************
' MDX Query Function
'**************************************************************************

Public Declare Function HypExecuteMDXEx Lib "HsAddin" (ByVal vtSheetName As Variant, _
                                                       ByVal vtQuery As Variant, _
                                                       ByVal vtBoolHideData As Variant, _
                                                       ByVal vtBoolDataLess As Variant, _
                                                       ByVal vtBoolNeedStatus As Variant, _
                                                       ByVal vtMbrIDType As Variant, _
                                                       ByVal vtAliasTable As Variant, _
                                                       ByRef outResult As MDX_AXES_NATIVE) As Long 'Essbase



#End If

'**************************************************************************
'  For converting C++ based MDX structure to a VB compliant MDX structure
'  **To be used with HypExecuteMDXEx only**
'**************************************************************************

Sub GetVBCompatibleMDXStructure(ByRef inStruct As MDX_AXES_NATIVE, ByRef outStruct As MDX_AXES)

outStruct.AxisCount = inStruct.AxisCount
outStruct.CellCount = inStruct.CellCount

'Process Cell Info
If inStruct.CellCount Then
    Dim vtCellStruct As Variant
                  
    ReDim outStruct.CellInfo(inStruct.CellCount - 1)
    For iCellCount = 0 To inStruct.CellCount - 1
    vtCellStruct = inStruct.CellInfo(iCellCount)
    outStruct.CellInfo(iCellCount).CellStatus = vtCellStruct(0)
    outStruct.CellInfo(iCellCount).CellValue = vtCellStruct(1)
    Next
End If
'End Processing Cell Info

'Process Axis Info
If inStruct.AxisCount Then
    ReDim outStruct.AxisInfo(inStruct.AxisCount - 1)
    Dim vtAxisStruct As Variant
                   
    For iAxisCount = 0 To inStruct.AxisCount - 1
         vtAxisStruct = inStruct.AxisInfo(iAxisCount)
         outStruct.AxisInfo(iAxisCount).DimensionsCount = vtAxisStruct(0)
         outStruct.AxisInfo(iAxisCount).TuplesCount = vtAxisStruct(1)
         outStruct.AxisInfo(iAxisCount).ClustersCount = vtAxisStruct(2)
         
         'Add dimensions Info under Axis
          If outStruct.AxisInfo(iAxisCount).DimensionsCount Then
            ReDim outStruct.AxisInfo(iAxisCount).DimensionInfo(outStruct.AxisInfo(iAxisCount).DimensionsCount - 1)
            Dim vtAllDims As Variant
            Dim vtDimStruct As Variant
            vtAllDims = vtAxisStruct(3)
                         
            For iDimCount = 0 To outStruct.AxisInfo(iAxisCount).DimensionsCount - 1
                 vtDimStruct = vtAllDims(iDimCount)
                 outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).DimensionName = vtDimStruct(0)
                 outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).NumMembers = vtDimStruct(1)
                 outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).NumProps = vtDimStruct(2)
       
                'Start --- Add Properties Info Under Each Dimension
                 If outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).NumProps Then
                    ReDim outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).PropsInfo(outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).NumProps - 1)
                    Dim vtAllProps As Variant
                    Dim vtPropStruct As Variant
                    vtAllProps = vtDimStruct(3)
                                  
                    For iCountProp = 0 To outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).NumProps - 1
                        vtPropStruct = vtAllProps(iCountProp)
                        outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).PropsInfo(iCountProp).PropertyName = vtPropStruct(0)
                        outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).PropsInfo(iCountProp).PropertyType = vtPropStruct(1)
                        outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimCount).PropsInfo(iCountProp).PropertyValue = Null 'Not sent
                    Next
                 End If
               'End ----- Add Properties Info under each Dimension
             Next
          End If
         'End Dimensions Info under Axis
    
         'Add Cluster Info under Axis
        If outStruct.AxisInfo(iAxisCount).ClustersCount Then
            ReDim outStruct.AxisInfo(iAxisCount).ClusterInfo(outStruct.AxisInfo(iAxisCount).ClustersCount - 1)
            Dim vtAllClusters As Variant
            Dim vtClusterStruct As Variant
            vtAllClusters = vtAxisStruct(4)
                             
            For iClusterCount = 0 To outStruct.AxisInfo(iAxisCount).ClustersCount - 1
            vtClusterStruct = vtAllClusters(iClusterCount)
            outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).TupleCount = vtClusterStruct(1)
            
            'Add Dimensions info under cluster
            If outStruct.AxisInfo(iAxisCount).DimensionsCount Then
                ReDim outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(outStruct.AxisInfo(iAxisCount).DimensionsCount - 1)
                Dim vtAllDimsUnderCluster As Variant
                Dim vtDimUnderCluster As Variant
                vtAllDimsUnderCluster = vtClusterStruct(0)
                                          
                For iDimsUnderClusterCount = 0 To outStruct.AxisInfo(iAxisCount).DimensionsCount - 1
                     vtDimUnderCluster = vtAllDimsUnderCluster(iDimsUnderClusterCount)
                     outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).NumMembers = vtDimUnderCluster(2)
                    
                    'Add members Under Cluster->Dimensions
                     If outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).NumMembers Then
                        ReDim outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).MemberInfo(outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).NumMembers - 1)
                        Dim vtAllMembersUnderClusterDim As Variant
                        Dim vtMemberUnderClusterDim As Variant
                        vtAllMembersUnderClusterDim = vtDimUnderCluster(4)
                                                    
                        For iMemUnderClusterDimCount = 0 To outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).NumMembers - 1
                            vtMemberUnderClusterDim = vtAllMembersUnderClusterDim(iMemUnderClusterDimCount)
                            outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).MemberInfo(iMemUnderClusterDimCount).MemberName = vtMemberUnderClusterDim(0)
                            outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).MemberInfo(iMemUnderClusterDimCount).NumClusters = vtMemberUnderClusterDim(1)
                            
                            'Add Properties Info
                            If outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimsUnderClusterCount).NumProps Then
                                ReDim outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).MemberInfo(iMemUnderClusterDimCount).PropInfo(outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimsUnderClusterCount).NumProps - 1)
                                Dim vtAllPropsUnderCluster As Variant
                                Dim vtPropUnderCluster As Variant
                                vtAllPropsUnderCluster = vtMemberUnderClusterDim(2)
                                                          
                                For iPropCountUnderCluster = 0 To outStruct.AxisInfo(iAxisCount).DimensionInfo(iDimsUnderClusterCount).NumProps - 1
                                    vtPropUnderCluster = vtAllPropsUnderCluster(iPropCountUnderCluster)
                                    outStruct.AxisInfo(iAxisCount).ClusterInfo(iClusterCount).DimensionInfo(iDimsUnderClusterCount).MemberInfo(iMemUnderClusterDimCount).PropInfo(iPropCountUnderCluster).PropertyValue = vtPropUnderCluster(2)
                                Next
                            End If
                            'End Properties Info
                         Next
                    End If
                   'End --- Add members Under Cluster -->Dimensions
                 Next
            End If
        'End Dimensions Info under cluster
        Next
    End If
 Next
 End If
 'End Cluster Info Under Axis
 ' End Processing Axis Info
End Sub


'**************************************************************************
' Error Code Message Function
'**************************************************************************

Function GetReturnCodeMessage(sts As Long) As String

    Select Case sts
        Case SmartViewErrors.SS_ERR_ERROR
            GetReturnCodeMessage = "General Error"
        Case SmartViewErrors.SS_NO_GRID_ON_SHEET_BUT_FUNCTIONS_SUBMITTED
            GetReturnCodeMessage = "No Grid on Sheet but Functions Submitted"
        Case SmartViewErrors.SS_SHEET_NOT_CONNECTED_BUT_FUNCTIONS_SUBMITTED
            GetReturnCodeMessage = "Sheet Not Connected but Functions Submitted"
        Case SmartViewErrors.SS_OK
            GetReturnCodeMessage = "OK"
        Case SmartViewErrors.SS_INIT_ERR
            GetReturnCodeMessage = "Initialization Error"
        Case SmartViewErrors.SS_TERM_ERR
            GetReturnCodeMessage = "Termination Error"
        Case SmartViewErrors.SS_NOT_INIT
            GetReturnCodeMessage = "Not Initialized"
        Case SmartViewErrors.SS_NOT_CONNECTED
            GetReturnCodeMessage = "Not Connected"
        Case SmartViewErrors.SS_NOT_LOCKED
            GetReturnCodeMessage = "Not Locked"
        Case SmartViewErrors.SS_INVALID_SSTABLE
            GetReturnCodeMessage = "Invalid Spreadsheet Table"
        Case SmartViewErrors.SS_INVALID_SSDATA
            GetReturnCodeMessage = "Invalid Spreadsheet Data"
        Case SmartViewErrors.SS_NOUNDO_INFO
            GetReturnCodeMessage = "No Undo Information Exists"
        Case SmartViewErrors.SS_CANCELED
            GetReturnCodeMessage = "Operation Has Been Cancelled"
        Case SmartViewErrors.SS_GLOBALOPTS
            GetReturnCodeMessage = "Global Options Error"
        Case SmartViewErrors.SS_SHEETOPTS
            GetReturnCodeMessage = "Sheet Options Error"
        Case SmartViewErrors.SS_NOTENABLED
            GetReturnCodeMessage = "Undo Is Not Enabled"
        Case SmartViewErrors.SS_NO_MEMORY
            GetReturnCodeMessage = "Not Enough Memory"
        Case SmartViewErrors.SS_DIALOG_ERROR
            GetReturnCodeMessage = "Appropriate Dialog Could Not Be Displayed"
        Case SmartViewErrors.SS_INVALID_PARAM
            GetReturnCodeMessage = "Function Contains an Invalid Parameter"
        Case SmartViewErrors.SS_CALCULATING
            GetReturnCodeMessage = "Calculation In Progress"
        Case SmartViewErrors.SS_SQL_IN_PROGRESS
            GetReturnCodeMessage = "SQL In Progress"
        Case SmartViewErrors.SS_FORMULAPRESERVE
            GetReturnCodeMessage = "Operation Is Not Allowed Because Spreadsheet Is In Formula Preservation Mode"
        Case SmartViewErrors.SS_INTERNALSSERROR
            GetReturnCodeMessage = "Operation Cannot Take Place On The Specified Sheet"
        Case SmartViewErrors.SS_INVALID_SHEET
            GetReturnCodeMessage = "Current Sheet Cannot Be Determined"
        Case SmartViewErrors.SS_NOACTIVESHEET
            GetReturnCodeMessage = "No Active Sheet Is Selected"
        Case SmartViewErrors.SS_NOTCALCULATING
            GetReturnCodeMessage = "Calculation Cannot Be Cancelled Because No Calculation Is Running"
        Case SmartViewErrors.SS_INVALIDSELECTION
            GetReturnCodeMessage = "Selection Parameter Is Invalid"
        Case SmartViewErrors.SS_INVALIDTOKEN
            GetReturnCodeMessage = "Invalid Token"
        Case SmartViewErrors.SS_CASCADENOTALLOWED
            GetReturnCodeMessage = "Cascade List File Cannot Be Created"
        Case SmartViewErrors.SS_NOMACROS
            GetReturnCodeMessage = "Spreadsheet Macros Cannot Be Run Due To Licensing Agreement"
        Case SmartViewErrors.SS_NOREADONLYMACROS
            GetReturnCodeMessage = "Spreadsheet Macros Which Update The Database Cannot Be Run Due To Licensing Agreement"
        Case SmartViewErrors.SS_READONLYSS
            GetReturnCodeMessage = "Database Cannot Be Updated Because You Have A Read Only License"
        Case SmartViewErrors.SS_NOSQLACCESS
            GetReturnCodeMessage = "No SQL Access"
        Case SmartViewErrors.SS_MENUALREADYREMOVED
            GetReturnCodeMessage = "Menu Already Removed"
        Case SmartViewErrors.SS_MENUALREADYADDED
            GetReturnCodeMessage = "Menu Already Added"
        Case SmartViewErrors.SS_NOSPREADSHEETACCESS
            GetReturnCodeMessage = "No Spreadsheet Access"
        Case SmartViewErrors.SS_NOHANDLES
            GetReturnCodeMessage = "No Handles"
        Case SmartViewErrors.SS_NOPREVCONNECTION
            GetReturnCodeMessage = "No Previous Connection"
        Case SmartViewErrors.SS_LROERROR
            GetReturnCodeMessage = "LRO Error"
        Case SmartViewErrors.SS_LROWINAPPACCESSERR
            GetReturnCodeMessage = "LRO Windows Application Access Error"
        Case SmartViewErrors.SS_DATANAVINITERR
            GetReturnCodeMessage = "Data Navigation Initialization Error"
        Case SmartViewErrors.SS_PARAMSETNOTALLOWED
            GetReturnCodeMessage = "Parameter Set Not Allowed"
        Case SmartViewErrors.SS_SHEET_PROTECTED
            GetReturnCodeMessage = "Specified Spreadsheet Is Protected"
        Case SmartViewErrors.SS_CALCSCRIPT_NOTFOUND
            GetReturnCodeMessage = "Calculation Script Not Found"
        Case SmartViewErrors.SS_NOSUPPORT_PROVIDER
            GetReturnCodeMessage = "Provider Not Supported"
        Case SmartViewErrors.SS_INVALID_ALIAS
            GetReturnCodeMessage = "Invalid Alias"
        Case SmartViewErrors.SS_CONN_NOT_FOUND
            GetReturnCodeMessage = "Connection Not Found"
        Case SmartViewErrors.SS_APS_CONN_NOT_FOUND
            GetReturnCodeMessage = "APS Connection Not Found"
        Case SmartViewErrors.SS_APS_NOT_CONNECTED
            GetReturnCodeMessage = "APS Not Connected"
        Case SmartViewErrors.SS_APS_CANT_CONNECT
            GetReturnCodeMessage = "APS Cannot Connect"
        Case SmartViewErrors.SS_CONN_ALREADY_EXISTS
            GetReturnCodeMessage = "Connection Already Exists"
        Case SmartViewErrors.SS_APS_URL_NOT_SAVED
            GetReturnCodeMessage = "APS URL Not Saved"
        Case SmartViewErrors.SS_MIGRATION_OF_CONN_NOT_ALLOWED
            GetReturnCodeMessage = "Migration of Connection Not Allowed"
        Case SmartViewErrors.SS_CONN_MGR_NOT_INITIALIZED
            GetReturnCodeMessage = "Connection Manager Not Initialized"
        Case SmartViewErrors.SS_FAILED_TO_GET_APS_OVERRIDE_PROPERTY
            GetReturnCodeMessage = "Failed to Get APS Override Property"
        Case SmartViewErrors.SS_FAILED_TO_SET_APS_OVERRIDE_PROPERTY
            GetReturnCodeMessage = "Failed to Set APS Override Property"
        Case SmartViewErrors.SS_FAILED_TO_GET_APS_URL
            GetReturnCodeMessage = "Failed to Get APS URL"
        Case SmartViewErrors.SS_APS_DISCONNECT_FAILED
            GetReturnCodeMessage = "APS Disconnect Failed"
        Case SmartViewErrors.SS_OPERATION_FAILED
            GetReturnCodeMessage = "Operation Failed"
        Case SmartViewErrors.SS_CANNOT_ASSOCIATE_SHEET_WITH_CONNECTION
            GetReturnCodeMessage = "Cannot Associate Sheet with Connection"
        Case SmartViewErrors.SS_REFRESH_SHEET_NEEDED
            GetReturnCodeMessage = "Refresh Sheet Needed"
        Case SmartViewErrors.SS_NO_GRID_OBJECT_ON_SHEET
            GetReturnCodeMessage = "No Grid Object On Sheet"
        Case SmartViewErrors.SS_NO_CONNECTION_ASSOCIATED
            GetReturnCodeMessage = "No Connection Associated"
        Case SmartViewErrors.SS_NON_DATA_CELL_PASSED
            GetReturnCodeMessage = "Non Data Cell Passed"
        Case SmartViewErrors.SS_DATA_CELL_IS_NOT_WRITABLE
            GetReturnCodeMessage = "Data Cell Is Not Writeable"
        Case SmartViewErrors.SS_NO_SVC_CONTENT_ON_SHEET
            GetReturnCodeMessage = "No Smart View Content on Sheet"
        Case SmartViewErrors.SS_FAILED_TO_GET_OFFICE_OBJECT
            GetReturnCodeMessage = "Failed to Get Office Object"
        Case SmartViewErrors.SS_OP_FAILED_AS_CHART_IS_SELECTED
            GetReturnCodeMessage = "Operation Failed as Chart is Selected"
        Case SmartViewErrors.SS_EXCEL_IN_EDIT_MODE
            GetReturnCodeMessage = "Excel in Edit Mode"
        Case SmartViewErrors.SS_SHEET_NON_SMARTVIEW_COMPATIBLE
            GetReturnCodeMessage = "Sheet Not Smart View Compatible"
        Case SmartViewErrors.SS_APP_NOT_STANDALONE
            GetReturnCodeMessage = "Application Not Stand Alone"
        Case SmartViewErrors.SS_SMART_VIEW_DISABLED
            GetReturnCodeMessage = "Smart View Disabled"
        Case SmartViewErrors.SS_VBA_DEPRECATED
            GetReturnCodeMessage = "Function Has Been Deprecated"
    Case SmartViewErrors.SS_OPERATION_NOT_SUPPORTED_IN_MULTIGRID_MODE
        GetReturnCodeMessage = "Operation is not supported in the Multigrid mode."
    Case SmartViewErrors.SS_INVALID_MEMBER
            GetReturnCodeMessage = "Invalid Member"
        Case SmartViewErrors.SS_NO_SV_NAME_RANGE
            GetReturnCodeMessage = "No Smart View Named Range On Sheet"
        Case SmartViewErrors.SS_AMBIGUOUS_MENU
            GetReturnCodeMessage = "could not resolve menu name"
        Case Else
            GetReturnCodeMessage = "Undefined Error Code " & sts
    End Select

End Function








