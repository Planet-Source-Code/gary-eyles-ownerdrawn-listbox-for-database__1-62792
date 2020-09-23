VERSION 5.00
Begin VB.UserControl Datalist 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4035
   KeyPreview      =   -1  'True
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox vScroll2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   3600
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   2880
      Left            =   2640
      Picture         =   "DataList.ctx":0000
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   1350
      Left            =   2160
      Picture         =   "DataList.ctx":3042
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   1920
      Picture         =   "DataList.ctx":38C6
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Image4 
      Height          =   1320
      Left            =   1560
      Picture         =   "DataList.ctx":3E08
      Top             =   1080
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   1020
      Left            =   1200
      Picture         =   "DataList.ctx":47CA
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   2160
      Left            =   960
      Picture         =   "DataList.ctx":515C
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   4080
      Left            =   600
      Picture         =   "DataList.ctx":5C5E
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Datalist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Property Variables:
Dim m_IsEditable As Boolean
Dim m_UseColumnColor As Boolean
Dim m_VerticalHeader As Boolean
Dim m_XPtheme As Boolean
Dim m_GridlineColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_ScrollbarBARcolor As OLE_COLOR
Dim m_ScrollbarBackColor As OLE_COLOR
Dim m_ColumnFixedWidth As Boolean
Dim m_HeaderTextColor As OLE_COLOR
Dim m_SelColor1 As OLE_COLOR
Dim m_SelColor2 As OLE_COLOR
Dim m_Headercolor As OLE_COLOR
Dim m_MultiSelect As Boolean
Dim m_ShowFocusRect As Boolean
Dim m_ShowCheck As Boolean
Dim m_Picture As Picture
Dim m_Selected As Long
Dim m_HeaderHeight As Long
Dim m_ShowHeader As Boolean
Dim m_ShowGrid As Boolean
Dim m_RowHeight As Long
Dim m_ColumnCount As Long
Dim m_TheData As Data
'Event Declarations:
Event HeaderClick(colIndex)
Event ItemSelected(tIndex As Long)
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Default Property Values:
Const m_def_IsEditable = 0
Const m_def_UseColumnColor = 0
Const m_def_VerticalHeader = 0
Const m_def_XPtheme = 1
Const m_def_GridlineColor = 14474460
Const m_def_BackColor = vbWindowBackground
Const m_def_ScrollbarBARcolor = 13158600
Const m_def_ScrollbarBackColor = vb3DFace
Const m_def_ColumnFixedWidth = 0
Const m_def_HeaderTextColor = 0
Const m_def_SelColor1 = 255
Const m_def_SelColor2 = 16777215
Const m_def_Headercolor = vb3DFace
Const m_def_MultiSelect = 0
Const m_def_ShowFocusRect = 0
Const m_def_ShowCheck = 0
Const m_def_Selected = -1
Const m_def_HeaderHeight = 22
Const m_def_ShowHeader = 0
Const m_def_ShowGrid = 0
Const m_def_RowHeight = 20
Const m_def_ColumnCount = 0
Dim NumFill As Long
Dim ColWidth() As Long
Dim ColColor() As Long
Dim aCheck() As Boolean
Dim aSelect() As Boolean
Dim CheckDown As Boolean
Dim HasFocus As Boolean
Dim sold As Long
Dim gValue As Double
Dim gMin As Long
Dim gMax As Long
Dim mdHover As Long
Dim HeadHover As Long
Dim WithEvents aTrack As clsTracking
Attribute aTrack.VB_VarHelpID = -1
Dim tData As Recordset

'' //---------------------------------------------------------------------------------------
'' //   "Window" (i.e., non-client) Parts & States
'' //
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeWindowParts
    WP_CAPTION = 1
    WP_SMALLCAPTION = 2
    WP_MINCAPTION = 3
    WP_SMALLMINCAPTION = 4
    WP_MAXCAPTION = 5
    WP_SMALLMAXCAPTION = 6
    WP_FRAMELEFT = 7
    WP_FRAMERIGHT = 8
    WP_FRAMEBOTTOM = 9
    WP_SMALLFRAMELEFT = 10
    WP_SMALLFRAMERIGHT = 11
    WP_SMALLFRAMEBOTTOM = 12
    '' //---- window frame buttons ----
    WP_SYSBUTTON = 13
    WP_MDISYSBUTTON = 14
    WP_MINBUTTON = 15
    WP_MDIMINBUTTON = 16
    WP_MAXBUTTON = 17
    WP_CLOSEBUTTON = 18
    WP_SMALLCLOSEBUTTON = 19
    WP_MDICLOSEBUTTON = 20
    WP_RESTOREBUTTON = 21
    WP_MDIRESTOREBUTTON = 22
    WP_HELPBUTTON = 23
    WP_MDIHELPBUTTON = 24
    '' //---- scrollbars
    WP_HORZSCROLL = 25
    WP_HORZTHUMB = 26
    WP_VERTSCROLL = 27
    WP_VERTTHUMB = 28
    '' //---- dialog ----
    WP_DIALOG = 29
    '' //---- hit-test templates ---
    WP_CAPTIONSIZINGTEMPLATE = 30
    WP_SMALLCAPTIONSIZINGTEMPLATE = 31
    WP_FRAMELEFTSIZINGTEMPLATE = 32
    WP_SMALLFRAMELEFTSIZINGTEMPLATE = 33
    WP_FRAMERIGHTSIZINGTEMPLATE = 34
    WP_SMALLFRAMERIGHTSIZINGTEMPLATE = 35
    WP_FRAMEBOTTOMSIZINGTEMPLATE = 36
    WP_SMALLFRAMEBOTTOMSIZINGTEMPLATE = 37
End Enum

Public Enum UxThemeFrameStates
    FS_ACTIVE = 1
    FS_INACTIVE = 2
End Enum

Public Enum UxThemeCaptionStates
    CS_ACTIVE = 1
    CS_INACTIVE = 2
    CS_DISABLED = 3
End Enum
    
Public Enum UxThemeMaxCaptionStates
    MXCS_ACTIVE = 1
    MXCS_INACTIVE = 2
    MXCS_DISABLED = 3
End Enum

Public Enum UxThemeMinCaptionStates
    MNCS_ACTIVE = 1
    MNCS_INACTIVE = 2
    MNCS_DISABLED = 3
End Enum

Public Enum UxThemeHorzScrollStates
    HSS_NORMAL = 1
    HSS_HOT = 2
    HSS_PUSHED = 3
    HSS_DISABLED = 4
End Enum

Public Enum UxThemeHorzThumbStates
    HTS_NORMAL = 1
    HTS_HOT = 2
    HTS_PUSHED = 3
    HTS_DISABLED = 4
End Enum

Public Enum UxThemeVertScrollStates
    VSS_NORMAL = 1
    VSS_HOT = 2
    VSS_PUSHED = 3
    VSS_DISABLED = 4
End Enum

Public Enum UxThemeVertThumbStates
    VTS_NORMAL = 1
    VTS_HOT = 2
    VTS_PUSHED = 3
    VTS_DISABLED = 4
End Enum

Public Enum UxThemeSysButtonStates
    SBS_NORMAL = 1
    SBS_HOT = 2
    SBS_PUSHED = 3
    SBS_DISABLED = 4
End Enum

Public Enum UxThemeMinButtonStates
    MINBS_NORMAL = 1
    MINBS_HOT = 2
    MINBS_PUSHED = 3
    MINBS_DISABLED = 4
End Enum

Public Enum UxThemeMaxButtonStates
    MAXBS_NORMAL = 1
    MAXBS_HOT = 2
    MAXBS_PUSHED = 3
    MAXBS_DISABLED = 4
End Enum

Public Enum UxThemeRestoreButtonStates
    RBS_NORMAL = 1
    RBS_HOT = 2
    RBS_PUSHED = 3
    RBS_DISABLED = 4
End Enum

Public Enum UxThemeHelpButtonStates
    HBS_NORMAL = 1
    HBS_HOT = 2
    HBS_PUSHED = 3
    HBS_DISABLED = 4
End Enum

Public Enum UxThemeCloseButtonStates
    CBS_NORMAL = 1
    CBS_HOT = 2
    CBS_PUSHED = 3
    CBS_DISABLED = 4
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Button" Parts & States
'' //--------------------------------------------------------------------------------------
Public Enum UxThemeButtonParts
    BP_PUSHBUTTON = 1
    BP_RADIOBUTTON = 2
    bp_checkbox = 3
    BP_GROUPBOX = 4
    BP_USERBUTTON = 5
End Enum

Public Enum UxThemePushButtonStates
    PBS_NORMAL = 1
    PBS_HOT = 2
    PBS_PRESSED = 3
    PBS_DISABLED = 4
    PBS_DEFAULTED = 5
End Enum

Public Enum UxThemeRadioButtonStates
    RBS_UNCHECKEDNORMAL = 1
    RBS_UNCHECKEDHOT = 2
    RBS_UNCHECKEDPRESSED = 3
    RBS_UNCHECKEDDISABLED = 4
    RBS_CHECKEDNORMAL = 5
    RBS_CHECKEDHOT = 6
    RBS_CHECKEDPRESSED = 7
    RBS_CHECKEDDISABLED = 8
End Enum

Public Enum UxThemeCheckBoxStates
    cbs_uncheckednormal = 1
    CBS_UNCHECKEDHOT = 2
    CBS_UNCHECKEDPRESSED = 3
    cbs_uncheckeddisabled = 4
    cbs_checkednormal = 5
    CBS_CHECKEDHOT = 6
    CBS_CHECKEDPRESSED = 7
    cbs_checkeddisabled = 8
    CBS_MIXEDNORMAL = 9
    CBS_MIXEDHOT = 10
    CBS_MIXEDPRESSED = 11
    CBS_MIXEDDISABLED = 12
End Enum

Public Enum UxThemeGroupBoxStates
    GBS_NORMAL = 1
    GBS_DISABLED = 2
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Rebar" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeRebarParts
    RP_GRIPPER = 1
    RP_GRIPPERVERT = 2
    RP_BAND = 3
    RP_CHEVRON = 4
    RP_CHEVRONVERT = 5
End Enum

Public Enum UxThemeChevronStates
    CHEVS_NORMAL = 1
    CHEVS_HOT = 2
    CHEVS_PRESSED = 3
End Enum


'' //---------------------------------------------------------------------------------------
'' //   "Toolbar" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeToolBarParts
    TP_BUTTON = 1
    TP_DROPDOWNBUTTON = 2
    TP_SPLITBUTTON = 3
    TP_SPLITBUTTONDROPDOWN = 4
    TP_SEPARATOR = 5
    TP_SEPARATORVERT = 6
End Enum

Public Enum UxThemeToolBarStates
    TS_NORMAL = 1
    TS_HOT = 2
    TS_PRESSED = 3
    TS_DISABLED = 4
    TS_CHECKED = 5
    TS_HOTCHECKED = 6
End Enum

'' //---------------------------------------------------------------------------------------
'' //   "Status" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeStatusParts
    SP_PANE = 1
    SP_GRIPPERPANE = 2
    SP_GRIPPER = 3
End Enum

'' //---------------------------------------------------------------------------------------
'' //   "Menu" Parts & States
'' //---------------------------------------------------------------------------------------
Public Enum UxThemeMenuParts
    MP_MENUITEM = 1
    MP_MENUDROPDOWN = 2
    MP_MENUBARITEM = 3
    MP_MENUBARDROPDOWN = 4
    MP_CHEVRON = 5
    MP_SEPARATOR = 6
End Enum

Public Enum UxThemeMenuStates
    MS_NORMAL = 1
    MS_SELECTED = 2
    MS_DEMOTED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "ListView" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeLISTVIEWParts
   LVP_LISTITEM = 1
   LVP_LISTGROUP = 2
   LVP_LISTDETAIL = 3
   LVP_LISTSORTEDDETAIL = 4
   LVP_EMPTYTEXT = 5
End Enum

Public Enum UxThemeLISTITEMStates
   LIS_NORMAL = 1
   LIS_HOT = 2
   LIS_SELECTED = 3
   LIS_DISABLED = 4
   LIS_SELECTEDNOTFOCUS = 5
End Enum

' //---------------------------------------------------------------------------------------
' //   "Header" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeHEADERParts
   hp_headeritem = 1
   HP_HEADERITEMLEFT = 2
   HP_HEADERITEMRIGHT = 3
   HP_HEADERSORTARROW = 4
End Enum

Public Enum UxThemeHEADERITEMStates
   his_normal = 1
   his_hot = 2
   HIS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMLEFTStates
   HILS_NORMAL = 1
   HILS_HOT = 2
   HILS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERITEMRIGHTStates
   HIRS_NORMAL = 1
   HIRS_HOT = 2
   HIRS_PRESSED = 3
End Enum

Public Enum UxThemeHEADERSORTARROWStates
   HSAS_SORTEDUP = 1
   HSAS_SORTEDDOWN = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Progress" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemePROGRESSParts
   PP_BAR = 1
   PP_BARVERT = 2
   PP_CHUNK = 3
   PP_CHUNKVERT = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tab" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UsxThemeTABParts
   TABP_TABITEM = 1
   TABP_TABITEMLEFTEDGE = 2
   TABP_TABITEMRIGHTEDGE = 3
   TABP_TABITEMBOTHEDGE = 4
   TABP_TOPTABITEM = 5
   TABP_TOPTABITEMLEFTEDGE = 6
   TABP_TOPTABITEMRIGHTEDGE = 7
   TABP_TOPTABITEMBOTHEDGE = 8
   TABP_PANE = 9
   TABP_BODY = 10
End Enum

Public Enum UxThemeTABITEMStates
   TIS_NORMAL = 1
   TIS_HOT = 2
   TIS_SELECTED = 3
   TIS_DISABLED = 4
   TIS_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMLEFTEDGEStates
   TILES_NORMAL = 1
   TILES_HOT = 2
   TILES_SELECTED = 3
   TILES_DISABLED = 4
   TILES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMRIGHTEDGEStates
   TIRES_NORMAL = 1
   TIRES_HOT = 2
   TIRES_SELECTED = 3
   TIRES_DISABLED = 4
   TIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTABITEMBOTHEDGESStates
   TIBES_NORMAL = 1
   TIBES_HOT = 2
   TIBES_SELECTED = 3
   TIBES_DISABLED = 4
   TIBES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMStates
   TTIS_NORMAL = 1
   TTIS_HOT = 2
   TTIS_SELECTED = 3
   TTIS_DISABLED = 4
   TTIS_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMLEFTEDGEStates
   TTILES_NORMAL = 1
   TTILES_HOT = 2
   TTILES_SELECTED = 3
   TTILES_DISABLED = 4
   TTILES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMRIGHTEDGEStates
   TTIRES_NORMAL = 1
   TTIRES_HOT = 2
   TTIRES_SELECTED = 3
   TTIRES_DISABLED = 4
   TTIRES_FOCUSED = 5
End Enum

Public Enum UxThemeTOPTABITEMBOTHEDGESStates
   TTIBES_NORMAL = 1
   TTIBES_HOT = 2
   TTIBES_SELECTED = 3
   TTIBES_DISABLED = 4
   TTIBES_FOCUSED = 5
End Enum

' //---------------------------------------------------------------------------------------
' //   "Trackbar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTRACKBARParts
   TKP_TRACK = 1
   TKP_TRACKVERT = 2
   TKP_THUMB = 3
   TKP_THUMBBOTTOM = 4
   TKP_THUMBTOP = 5
   TKP_THUMBVERT = 6
   TKP_THUMBLEFT = 7
   TKP_THUMBRIGHT = 8
   TKP_TICS = 9
   TKP_TICSVERT = 10
End Enum

Public Enum UxThemeTRACKBARStates
   TKS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKStates
   TRS_NORMAL = 1
End Enum

Public Enum UxThemeTRACKVERTStates
   TRVS_NORMAL = 1
End Enum

Public Enum UxThemeTHUMBStates
   TUS_NORMAL = 1
   TUS_HOT = 2
   TUS_PRESSED = 3
   TUS_FOCUSED = 4
   TUS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBBOTTOMStates
   TUBS_NORMAL = 1
   TUBS_HOT = 2
   TUBS_PRESSED = 3
   TUBS_FOCUSED = 4
   TUBS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBTOPStates
   TUTS_NORMAL = 1
   TUTS_HOT = 2
   TUTS_PRESSED = 3
   TUTS_FOCUSED = 4
   TUTS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBVERTStates
   TUVS_NORMAL = 1
   TUVS_HOT = 2
   TUVS_PRESSED = 3
   TUVS_FOCUSED = 4
   TUVS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBLEFTStates
   TUVLS_NORMAL = 1
   TUVLS_HOT = 2
   TUVLS_PRESSED = 3
   TUVLS_FOCUSED = 4
   TUVLS_DISABLED = 5
End Enum

Public Enum UxThemeTHUMBRIGHTStates
   TUVRS_NORMAL = 1
   TUVRS_HOT = 2
   TUVRS_PRESSED = 3
   TUVRS_FOCUSED = 4
   TUVRS_DISABLED = 5
End Enum

Public Enum UxThemeTICSStates
   TSS_NORMAL = 1
End Enum

Public Enum UxThemeTICSVERTStates
   TSVS_NORMAL = 1
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tooltips" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTOOLTIPParts
   TTP_STANDARD = 1
   TTP_STANDARDTITLE = 2
   TTP_BALLOON = 3
   TTP_BALLOONTITLE = 4
   TTP_CLOSE = 5
End Enum

Public Enum UxThemeCLOSEStates
   TTCS_NORMAL = 1
   TTCS_HOT = 2
   TTCS_PRESSED = 3
End Enum

Public Enum UxThemeSTANDARDStates
   TTSS_NORMAL = 1
   TTSS_LINK = 2
End Enum

Public Enum UxThemeBALLOONStates
   TTBS_NORMAL = 1
   TTBS_LINK = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "TreeView" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTREEVIEWParts
   TVP_TREEITEM = 1
   TVP_GLYPH = 2
   TVP_BRANCH = 3
End Enum

Public Enum UxThemeTREEITEMStates
   TREIS_NORMAL = 1
   TREIS_HOT = 2
   TREIS_SELECTED = 3
   TREIS_DISABLED = 4
   TREIS_SELECTEDNOTFOCUS = 5
End Enum

Public Enum UxThemeGLYPHStates
   GLPS_CLOSED = 1
   GLPS_OPENED = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Spin" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSPINStates
   SPNP_UP = 1
   SPNP_DOWN = 2
   SPNP_UPHORZ = 3
   SPNP_DOWNHORZ = 4
End Enum

Public Enum UxThemeUPStates
   UPS_NORMAL = 1
   UPS_HOT = 2
   UPS_PRESSED = 3
   UPS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNStates
   DNS_NORMAL = 1
   DNS_HOT = 2
   DNS_PRESSED = 3
   DNS_DISABLED = 4
End Enum

Public Enum UxThemeUPHORZStates
   UPHZS_NORMAL = 1
   UPHZS_HOT = 2
   UPHZS_PRESSED = 3
   UPHZS_DISABLED = 4
End Enum

Public Enum UxThemeDOWNHORZStates
   DNHZS_NORMAL = 1
   DNHZS_HOT = 2
   DNHZS_PRESSED = 3
   DNHZS_DISABLED = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Page" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemePageParts
   PGRP_UP = 1
   PGRP_DOWN = 2
   PGRP_UPHORZ = 3
   PGRP_DOWNHORZ = 4
End Enum

' //--- Pager uses same states as Spin ---

' //---------------------------------------------------------------------------------------
' //   "Scrollbar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSCROLLBARParts
   sbp_arrowbtn = 1
   SBP_THUMBBTNHORZ = 2
   sbp_thumbbtnvert = 3
   SBP_LOWERTRACKHORZ = 4
   SBP_UPPERTRACKHORZ = 5
   SBP_LOWERTRACKVERT = 6
   SBP_UPPERTRACKVERT = 7
   SBP_GRIPPERHORZ = 8
   sbp_grippervert = 9
   SBP_SIZEBOX = 10
End Enum



Public Enum UxThemeARROWBTNStates
   abs_upnormal = 1
   abs_uphot = 2
   abs_uppressed = 3
   abs_updisabled = 4
   abs_downnormal = 5
   abs_downhot = 6
   abs_downpressed = 7
   abs_downdisabled = 8
   ABS_LEFTNORMAL = 9
   ABS_LEFTHOT = 10
   ABS_LEFTPRESSED = 11
   ABS_LEFTDISABLED = 12
   ABS_RIGHTNORMAL = 13
   ABS_RIGHTHOT = 14
   ABS_RIGHTPRESSED = 15
   ABS_RIGHTDISABLED = 16
End Enum

Public Enum UxThemeSCROLLBARStates
   scrbs_normal = 1
   scrbs_hot = 2
   scrbs_pressed = 3
   scrbs_disabled = 4
End Enum

Public Enum UxThemeSIZEBOXStates
   SZB_RIGHTALIGN = 1
   SZB_LEFTALIGN = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "Edit" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeEDITParts
   EP_EDITTEXT = 1
   EP_CARET = 2
End Enum

Public Enum UxThemeEDITTEXTStates
   ETS_NORMAL = 1
   ETS_HOT = 2
   ETS_SELECTED = 3
   ETS_DISABLED = 4
   ETS_FOCUSED = 5
   ETS_READONLY = 6
   ETS_ASSIST = 7
End Enum

' //---------------------------------------------------------------------------------------
' //   "ComboBox" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeComboBoxParts
   CP_DROPDOWNBUTTON = 1
End Enum

Public Enum UxThemeComboBoxStates
   CBXS_NORMAL = 1
   CBXS_HOT = 2
   CBXS_PRESSED = 3
   CBXS_DISABLED = 4
End Enum

' //---------------------------------------------------------------------------------------
' //   "Taskbar Clock" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeCLOCKParts
   CLP_TIME = 1
End Enum

Public Enum UxThemeCLOCKStates
   CLS_NORMAL = 1
End Enum

' //---------------------------------------------------------------------------------------
' //   "Tray Notify" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTRAYNOTIFYParts
   TNP_BACKGROUND = 1
   TNP_ANIMBACKGROUND = 2
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTASKBARParts
   TBP_BACKGROUNDBOTTOM = 1
   TBP_BACKGROUNDRIGHT = 2
   TBP_BACKGROUNDTOP = 3
   TBP_BACKGROUNDLEFT = 4
   TBP_SIZINGBARBOTTOM = 5
   TBP_SIZINGBARRIGHT = 6
   TBP_SIZINGBARTOP = 7
   TBP_SIZINGBARLEFT = 8
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBand" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeTASKBANDParts
   TDP_GROUPCOUNT = 1
   TDP_FLASHBUTTON = 2
   TDP_FLASHBUTTONGROUPMENU = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "StartPanel" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeSTARTPANELParts
   SPP_USERPANE = 1
   SPP_MOREPROGRAMS = 2
   SPP_MOREPROGRAMSARROW = 3
   SPP_PROGLIST = 4
   SPP_PROGLISTSEPARATOR = 5
   SPP_PLACESLIST = 6
   SPP_PLACESLISTSEPARATOR = 7
   SPP_LOGOFF = 8
   SPP_LOGOFFBUTTONS = 9
   SPP_USERPICTURE = 10
   SPP_PREVIEW = 11
End Enum

Public Enum UxThemeMOREPROGRAMSARROWStates
   SPS_NORMAL = 1
   SPS_HOT = 2
   SPS_PRESSED = 3
End Enum

Public Enum UxThemeLOGOFFBUTTONSStates
   SPLS_NORMAL = 1
   SPLS_HOT = 2
   SPLS_PRESSED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "ExplorerBar" Parts & States
' //---------------------------------------------------------------------------------------
Public Enum UxThemeEXPLORERBARParts
   EBP_HEADERBACKGROUND = 1
   EBP_HEADERCLOSE = 2
   EBP_HEADERPIN = 3
   EBP_IEBARMENU = 4
   EBP_NORMALGROUPBACKGROUND = 5
   EBP_NORMALGROUPCOLLAPSE = 6
   EBP_NORMALGROUPEXPAND = 7
   EBP_NORMALGROUPHEAD = 8
   EBP_SPECIALGROUPBACKGROUND = 9
   EBP_SPECIALGROUPCOLLAPSE = 10
   EBP_SPECIALGROUPEXPAND = 11
   EBP_SPECIALGROUPHEAD = 12
End Enum

Public Enum UxThemeHEADERCLOSEStates
   EBHC_NORMAL = 1
   EBHC_HOT = 2
   EBHC_PRESSED = 3
End Enum

Public Enum UxThemeHEADERPINStates
   EBHP_NORMAL = 1
   EBHP_HOT = 2
   EBHP_PRESSED = 3
   EBHP_SELECTEDNORMAL = 4
   EBHP_SELECTEDHOT = 5
   EBHP_SELECTEDPRESSED = 6
End Enum

Public Enum UxThemeIEBARMENUStates
   EBM_NORMAL = 1
   EBM_HOT = 2
   EBM_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPCOLLAPSEStates
   EBNGC_NORMAL = 1
   EBNGC_HOT = 2
   EBNGC_PRESSED = 3
End Enum

Public Enum UxThemeNORMALGROUPEXPANDStates
   EBNGE_NORMAL = 1
   EBNGE_HOT = 2
   EBNGE_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPCOLLAPSEStates
   EBSGC_NORMAL = 1
   EBSGC_HOT = 2
   EBSGC_PRESSED = 3
End Enum

Public Enum UxThemeSPECIALGROUPEXPANDStates
   EBSGE_NORMAL = 1
   EBSGE_HOT = 2
   EBSGE_PRESSED = 3
End Enum

' //---------------------------------------------------------------------------------------
' //   "TaskBand" Parts & States
' //---------------------------------------------------------------------------------------

'Required Enums
Public Enum UxThemeMENUBANDParts
   MDP_NEWAPPBUTTON = 1
   MDP_SEPERATOR = 2
End Enum

Public Enum UxThemeMENUBANDStates
   MDS_NORMAL = 1
   MDS_HOT = 2
   MDS_PRESSED = 3
   MDS_DISABLED = 4
   MDS_CHECKED = 5
   MDS_HOTCHECKED = 6
End Enum

Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Private Const DT_SINGLELINE = &H20
'Private Const DT_VCENTER = &H4
'Private Const DT_END_ELLIPSIS = &H8000&

Private isMouseDown As Boolean
Private ColSize As Long
Private colSizeN As Long
Private mdScroll As Long

Private Const DIB_RGB_ColS      As Long = 0
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Public Enum GradientDirectionEnum
    [Fill_None] = 0
    [Fill_Horizontal] = 1
    [Fill_HorizontalMiddleOut] = 2
    [Fill_Vertical] = 3
    [Fill_VerticalMiddleOut] = 4
    [Fill_DownwardDiagonal] = 5
    [Fill_UpwardDiagonal] = 6
End Enum

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFacename As String * 33
End Type

Private Enum THEMESIZE
    TS_MIN             '// minimum size
    TS_TRUE            '// size without stretching
    TS_DRAW            '// size that theme mgr will use to draw part
End Enum

Private Type POINT
   x As Long
   Y As Long
End Type

Private Type SIZE
   cX As Long
   cY As Long
End Type

Private Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

'Open a hTheme, Needed at the begginning of the drawing
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
'Close the hTeme Handle
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
'Draw the background of the control or section
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
'Draw the parent background (for transparent and semitransparent controls with blending over the parent object
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal hdc As Long, prc As RECT) As Long
'Get the rect of the control where theme should be applyed
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As RECT, pContentRect As RECT) As Long
'Draw the theme text on the control.
Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlag As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
'Draw the themed Icon works With the Imagelist Object
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal hIml As Long, ByVal iImageIndex As Long) As Long
'Returns the default size of a theme data, in a THEMESIZE variable
Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, prc As RECT, ByVal eSize As THEMESIZE, psz As SIZE) As Long
'Returns the extent of the thewt drawn with the theme style
Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As DrawTextFlags, pBoundingRect As RECT, pExtentRect As RECT) As Long
'Returns true If the selected theme part is defined in the current theme
Private Declare Function IsThemePartDefined Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long) As Long


Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'DrawEdge Constants
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' Flags for DrawFrameControl
Const DFC_CAPTION = 1
Const DFC_MENU = 2
Const DFC_SCROLL = 3
Const DFC_BUTTON = 4
Const DFCS_BUTTONCHECK = &H0
Const DFCS_CAPTIONCLOSE = &H0
Const DFCS_CAPTIONMIN = &H1
Const DFCS_CAPTIONMAX = &H2
Const DFCS_CAPTIONRESTORE = &H3
Const DFCS_CAPTIONHELP = &H4
Const DFCS_MENUARROW = &H0
Const DFCS_MENUCHECK = &H1
Const DFCS_MENUBULLET = &H2
Const DFCS_MENUARROWRIGHT = &H4
Const DFCS_SCROLLUP = &H0
Const DFCS_SCROLLDOWN = &H1
Const DFCS_SCROLLLEFT = &H2
Const DFCS_SCROLLRIGHT = &H3
Const DFCS_SCROLLCOMBOBOX = &H5
Const DFCS_SCROLLSIZEGRIP = &H8
Const DFCS_SCROLLSIZEGRIPRIGHT = &H10

Const DFCS_BUTTONRADIOIMAGE = &H1
Const DFCS_BUTTONRADIOMASK = &H2
Const DFCS_BUTTONRADIO = &H4
Const DFCS_BUTTON3STATE = &H8
Const DFCS_BUTTONPUSH = &H10

Private Const DFCS_CHECKED = &H400
Private Const DFCS_FLAT = &H4000
Private Const DFCS_HOT = &H1000
Private Const DFCS_INACTIVE = &H100
Private Const DFCS_MONO = &H8000
Private Const DFCS_PUSHED = &H200
Private Const DFCS_TRANSPARENT = &H800

Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

'Size for the check
Private Const CheckW = 30

Private Const PM_REMOVE = &H1

Dim DblX As Single
Dim DblY As Single
Dim DblB As Integer

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

'Private Type BITMAP
'      bmType As Integer
'      bmWidth As Integer
'      bmHeight As Integer
'      bmWidthBytes As Integer
'      bmPlanes As String * 1
'      bmBitsPixel As String * 1
'      bmBits As Long
'End Type
   
Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
   
   'Declare Function BitBlt Lib "GDI" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Long) As Integer
'Private Declare Function SetBkColor Lib "GDI" (ByVal hdc As Integer, ByVal cColor As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "GDI" (ByVal hdc As Integer) As Integer
'Private Declare Function DeleteDC Lib "GDI" (ByVal hdc As Integer) As Integer
'Private Declare Function CreateBitmap Lib "GDI" (ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal cbPlanes As Integer, ByVal cbBits As Integer, lpvBits As Any) As Integer
'Private Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
'Private Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
'Private Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
'Private Declare Function GetObj Lib "GDI" Alias "GetObject" (ByVal hObject As Integer, ByVal nCount As Integer, bmp As Any) As Integer

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Sub aTrack_cLostFocus()
    Text1.Visible = False
End Sub

Private Sub aTrack_MouseLeave()
    'Debug.Print "LEAVING "; Rnd
    HeadHover = 0
    mdHover = 0
    Call DrawList
    vScroll2_Paint
End Sub

Private Function DrawTheme(thdc As Long, sClass As String, ByVal iPart As Long, ByVal iState As Long, m_btnRect As RECT) As Boolean
    'hTheme handle
    Dim hTheme As Long
    'Temp variable for
    Dim lResult As Long
    'If a error occurs then or we are not running XP or the visual style is windows Classic
    On Error GoTo NoXP
    'Get out hTheme Handle
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
    'Did we get a theme handle?
    If hTheme Then
        'Yes! draw the control background
        lResult = DrawThemeBackground(hTheme, thdc, iPart, iState, m_btnRect, m_btnRect)
        'If drawing was successful, return true, or false If not.
        DrawTheme = IIf(lResult, False, True)
    Else
        'No, we couldn't get a hTheme, drawing failed
        DrawTheme = False
    End If
    'Exit the function now
    Exit Function
NoXP:
    'An Error was detected, drawing Failed
    DrawTheme = False
End Function

Sub TransparentBlt2(dsthDC As Long, srcBmp As Long, x As Integer, _
                   Y As Integer, aWidth As Integer, _
                   aHeight As Integer, srcX As Long, srcY As Long, TransColor As Long)
    
    Dim maskDC As Long      'DC for the mask
    Dim tempDC As Long      'DC for temporary data
    Dim hMaskBmp As Long    'Bitmap for mask
    Dim hTempBmp As Long    'Bitmap for temporary data
    
    Dim Success As Long
    Dim bmp As BITMAP
    Dim hSrcPrevBmp As Long
    Dim srchDC As Long
    Dim srcDC As Long
    Dim hPrevBmp As Long
    
    
'    Success = GetObjectAPI(srcBmp, Len(bmp), bmp)
    
    'First, create some DC's. These are our gateways to associated
    'bitmaps in RAM
    maskDC = CreateCompatibleDC(dsthDC)
    tempDC = CreateCompatibleDC(dsthDC)
    
    srcDC = CreateCompatibleDC(dsthDC)
    hSrcPrevBmp = SelectObject(srcDC, srcBmp)     'Select bitmap in DC
    srchDC = srcDC

    'Then, we need the bitmaps. Note that we create a monochrome
    'bitmap here!
    'This is a trick we use for creating a mask fast enough.
    hMaskBmp = CreateBitmap(aWidth, aHeight, 1, 1, ByVal 0&)
    hTempBmp = CreateCompatibleBitmap(dsthDC, aWidth, aHeight)

    'Then we can assign the bitmaps to the DCs
'    hMaskBmp = SelectObject(maskDC, hMaskBmp)
'    hTempBmp = SelectObject(tempDC, hTempBmp)
Call SelectObject(maskDC, hMaskBmp)
Call SelectObject(tempDC, hTempBmp)

    'Now we can create a mask. First, we set the background color
    'to the transparent color; then we copy the image into the
    'monochrome bitmap.
    'When we are done, we reset the background color of the
    'original source.
    TransColor = SetBkColor(srchDC, TransColor)
    BitBlt maskDC, 0, 0, aWidth, aHeight, srchDC, srcX, srcY, vbSrcCopy
    TransColor = SetBkColor(srchDC, TransColor)

    'The first we do with the mask is to MergePaint it into the
    'destination.
    'This will punch a WHITE hole in the background exactly were
    'we want the graphics to be painted in.
    BitBlt tempDC, 0, 0, aWidth, aHeight, maskDC, 0, 0, vbSrcCopy
    BitBlt dsthDC, x, Y, aWidth, Height, tempDC, 0, 0, vbMergePaint

    'Now we delete the transparent part of our source image. To do
    'this, we must invert the mask and MergePaint it into the
    'source image. The transparent area will now appear as WHITE.
    BitBlt maskDC, 0, 0, aWidth, aHeight, maskDC, 0, 0, vbNotSrcCopy
    BitBlt tempDC, 0, 0, aWidth, aHeight, srchDC, srcX, srcY, vbSrcCopy
    BitBlt tempDC, 0, 0, aWidth, aHeight, maskDC, 0, 0, vbMergePaint

    'Both target and source are clean. All we have to do is to AND
    'them together!
    BitBlt dsthDC, x, Y, aWidth, aHeight, tempDC, 0, 0, vbSrcAnd

    'Now all we have to do is to clean up after us and free system
    'resources..
    hPrevBmp = SelectObject(srcDC, hSrcPrevBmp) 'Select orig object
    Success = DeleteDC(srcDC)
    DeleteObject (hSrcPrevBmp)
    
    DeleteObject (hMaskBmp)
    DeleteObject (hTempBmp)
    DeleteDC (maskDC)
    DeleteDC (tempDC)
End Sub


Private Sub TransparentBlt(dest As Long, ByVal srcBmp As Long, ByVal destX As Integer, ByVal destY As Integer, ByVal destWidth As Integer, ByVal destHeight As Long, ByVal TransColor As Long)
      Dim destScale As Long
      Dim srcDC As Long  'source bitmap (color)
      Dim saveDC As Long 'backup copy of source bitmap
      Dim maskDC As Long 'mask bitmap (monochrome)
      Dim invDC As Long  'inverse of mask bitmap (monochrome)
      Dim resultDC As Long 'combination of source bitmap & background
      Dim bmp As BITMAP 'description of the source bitmap
      Dim hResultBmp As Long 'Bitmap combination of source & background
      Dim hSaveBmp As Long 'Bitmap stores backup copy of source bitmap
      Dim hMaskBmp As Long 'Bitmap stores mask (monochrome)
      Dim hInvBmp As Long  'Bitmap holds inverse of mask (monochrome)
      Dim hPrevBmp As Long 'Bitmap holds previous bitmap selected in DC
      Dim hSrcPrevBmp As Long  'Holds previous bitmap in source DC
      Dim hSavePrevBmp As Long 'Holds previous bitmap in saved DC
      Dim hDestPrevBmp As Long 'Holds previous bitmap in destination DC
      Dim hMaskPrevBmp As Long 'Holds previous bitmap in the mask DC
      Dim hInvPrevBmp As Long 'Holds previous bitmap in inverted mask DC
      Dim OrigColor As Long 'Holds original background color from source DC
      Dim Success As Integer 'Stores result of call to Windows API
        
        'Retrieve bitmap to get width (bmp.bmWidth) & height (bmp.bmHeight)
        Success = GetObjectAPI(srcBmp, Len(bmp), bmp)
        srcDC = CreateCompatibleDC(dest)    'Create DC to hold stage
        saveDC = CreateCompatibleDC(dest)   'Create DC to hold stage
        maskDC = CreateCompatibleDC(dest)   'Create DC to hold stage
        invDC = CreateCompatibleDC(dest)    'Create DC to hold stage
        resultDC = CreateCompatibleDC(dest) 'Create DC to hold stage
        'Create monochrome bitmaps for the mask-related bitmaps:
        hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
        hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
        'Create color bitmaps for final result & stored copy of source
        hResultBmp = CreateCompatibleBitmap(dest, bmp.bmWidth, bmp.bmHeight)
        hSaveBmp = CreateCompatibleBitmap(dest, bmp.bmWidth, bmp.bmHeight)
        
        hSrcPrevBmp = SelectObject(srcDC, srcBmp)     'Select bitmap in DC
        hSavePrevBmp = SelectObject(saveDC, hSaveBmp) 'Select bitmap in DC
        hMaskPrevBmp = SelectObject(maskDC, hMaskBmp) 'Select bitmap in DC
        hInvPrevBmp = SelectObject(invDC, hInvBmp)    'Select bitmap in DC
        hDestPrevBmp = SelectObject(resultDC, hResultBmp) 'Select bitmap
        
        Success = BitBlt(saveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, vbSrcCopy) 'Make backup of source bitmap to restore later
        'Create mask: set background color of source to transparent color.
        OrigColor = SetBkColor(srcDC, TransColor)
        'Debug.Print OrigColor, TransColor
        Success = BitBlt(maskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, vbSrcCopy)
        TransColor = SetBkColor(srcDC, OrigColor)
        'Create inverse of mask to AND w/ source & combine w/ background.
        Success = BitBlt(invDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, vbNotSrcCopy) ' vbNotSrcCopy)
        
        'Copy background bitmap to result & create final transparent bitmap
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, dest, destX, destY, vbSrcCopy) 'vbSrcCopy)
        
        'AND mask bitmap w/ result DC to punch hole in the background by
        'painting black area for non-transparent portion of source bitmap.
        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, vbSrcAnd) ' vbSrcAnd)
        
        
        'AND inverse mask w/ source bitmap to turn off bits associated
        'with transparent area of source bitmap by making it black.
BitBlt maskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, vbNotSrcCopy
        Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, vbSrcAnd) 'vbSrcAnd)7
        'Success = BitBlt(SrcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, vbSrcCopy) 'vbSrcAnd)
'Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, invDC, 0, 0, vbNotSrcCopy) 'vbSrcAnd)
        'XOR result w/ source bitmap to make background show through.
'        Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, SrcDC, 0, 0, vbSrcPaint) 'vbSrcPaint)
'Success = BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, invDC, 0, 0, vbSrcPaint) 'vbSrcPaint)

'        Success = BitBlt(dest, destX, destY, bmp.bmWidth, bmp.bmHeight, resultDC, 0, 0, vbSrcCopy) 'Display transparent bitmap on backgrnd
        Success = BitBlt(dest, destX, destY, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, vbSrcCopy) 'vbSrcAnd)
        
        Success = BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, saveDC, 0, 0, vbSrcCopy) 'Restore backup of bitmap.

    'BitBlt maskDC, 0, 0, Width, Height, srchDC, 0, 0, vbSrcCopy
    'BitBlt tempDC, 0, 0, Width, Height, maskDC, 0, 0, vbSrcCopy
    'BitBlt dsthDC, X, Y, Width, Height, tempDC, 0, 0, vbMergePaint
    'BitBlt maskDC, 0, 0, Width, Height, maskDC, 0, 0, vbNotSrcCopy
    'BitBlt tempDC, 0, 0, Width, Height, srchDC, 0, 0, vbSrcCopy
    'BitBlt tempDC, 0, 0, Width, Height, maskDC, 0, 0, vbMergePaint
    'BitBlt dsthDC, X, Y, Width, Height, tempDC, 0, 0, vbSrcAnd

        hPrevBmp = SelectObject(srcDC, hSrcPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(saveDC, hSavePrevBmp) 'Select orig object
        hPrevBmp = SelectObject(resultDC, hDestPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(maskDC, hMaskPrevBmp) 'Select orig object
        hPrevBmp = SelectObject(invDC, hInvPrevBmp) 'Select orig object
        Success = DeleteObject(hSaveBmp)   'Deallocate system resources.
        Success = DeleteObject(hMaskBmp)   'Deallocate system resources.
        Success = DeleteObject(hInvBmp)    'Deallocate system resources.
        Success = DeleteObject(hResultBmp) 'Deallocate system resources.
        Success = DeleteDC(srcDC)          'Deallocate system resources.
        Success = DeleteDC(saveDC)         'Deallocate system resources.
        Success = DeleteDC(invDC)          'Deallocate system resources.
        Success = DeleteDC(maskDC)         'Deallocate system resources.
        Success = DeleteDC(resultDC)       'Deallocate system resources.
   End Sub


Public Sub aTrack_ScrollUp()
'Debug.Print "Up"
If gValue > gMin Then
    Text1.Visible = False
    gValue = gValue - 1
    Call DrawList
    vScroll2_Paint
End If
End Sub

Public Sub aTrack_ScrollDown()
'Debug.Print "Down"
If gValue < gMax Then
    Text1.Visible = False
    gValue = gValue + 1
    Call DrawList
    vScroll2_Paint
End If
End Sub

Private Sub DrawHeader()
Dim oldFore As Long
Dim rct As RECT
Dim cCol As Long
Dim f As LOGFONT, hPrevFont As Long, hFont As Long, fontname As String
Dim xCor As Long
Dim fText As String
Dim dtSuccess As Boolean


If m_ShowHeader = True Then
    oldFore = ForeColor
    ForeColor = m_HeaderTextColor
    If m_XPtheme = False Then
        Line (0, 0)-(ScaleWidth, m_HeaderHeight), m_Headercolor, BF
    Else
        UserControl.PaintPicture Image6.Picture, 0, 0, ScaleWidth - VScroll2.ScaleWidth, m_HeaderHeight, 0, 0, 16, 15, vbSrcCopy
    End If
    
        If m_ShowCheck = True Then
            xCor = CheckW
            If m_ShowHeader = True Then
                rct.left = 1
                rct.Right = CheckW
                rct.tOp = 1
                rct.Bottom = m_HeaderHeight
            If m_XPtheme = False Then
                DrawEdge hdc, rct, BDR_RAISEDINNER, BF_RECT
            Else
                If Enabled = True Then
                    dtSuccess = False
                    dtSuccess = DrawTheme(hdc, "Header", hp_headeritem, his_normal, rct)
                    If dtSuccess = False Then
                        UserControl.PaintPicture Image6.Picture, rct.left, 0, rct.Right, m_HeaderHeight, 0, 0, 18, 15, vbSrcCopy
                        UserControl.PaintPicture Image6.Picture, rct.left, rct.Bottom - 3, rct.Right, 3, 0, 15, 18, 3, vbSrcCopy
                        'UserControl.PaintPicture Image6.Picture, rct.Right - 2, 0, 2, m_HeaderHeight, 16, 0, 2, 18, vbSrcCopy
                    End If
                End If
            End If
            End If
        End If
    For cCol = 0 To m_TheData.Recordset.Fields.Count - 1
        If ColWidth(cCol) > 0 Then
            rct.tOp = 1
            rct.Bottom = m_HeaderHeight
            rct.left = xCor + 1
            rct.Right = rct.left + ColWidth(cCol)
            fText = m_TheData.Recordset.Fields(cCol).Name
            
            'Debug.Print m_thedata.recordset.Fields(cCol).Type
            If m_XPtheme = False Then
                Line (rct.left, 0)-(rct.Right, rct.Bottom), m_Headercolor, BF
                DrawEdge hdc, rct, BDR_RAISEDINNER, BF_RECT
            Else
                If Enabled = True Then
                If HeadHover - 1 = cCol Then
                    dtSuccess = False
                    dtSuccess = DrawTheme(hdc, "Header", hp_headeritem, his_hot, rct)
                    If dtSuccess = False Then
                        UserControl.PaintPicture Image6.Picture, rct.left, 0, rct.Right - rct.left, m_HeaderHeight - 5, 0, 18, 18, 9, vbSrcCopy
                        TransparentBlt2 UserControl.hdc, Image6.Picture, CInt(rct.left), m_HeaderHeight - 18, 9, 18, 0, 18, RGB(255, 0, 0)
                        UserControl.PaintPicture Image6.Picture, rct.left + 9, m_HeaderHeight - 18, rct.Right - rct.left - 18, 18, 9, 18, 1, 18, vbSrcCopy
                        TransparentBlt2 UserControl.hdc, Image6.Picture, CInt(rct.Right) - 9, m_HeaderHeight - 18, 9, 18, 9, 18, RGB(255, 0, 0)
                    End If
                Else
                    dtSuccess = False
                    dtSuccess = DrawTheme(hdc, "Header", hp_headeritem, his_normal, rct)
                    If dtSuccess = False Then
                        UserControl.PaintPicture Image6.Picture, rct.left, 0, ColWidth(cCol), m_HeaderHeight - 3, 0, 0, 16, 15, vbSrcCopy
                        UserControl.PaintPicture Image6.Picture, rct.left, rct.Bottom - 3, rct.Right, 3, 0, 15, 18, 3, vbSrcCopy
                        UserControl.PaintPicture Image6.Picture, rct.Right - 2, 0, 2, m_HeaderHeight - 3, 16, 0, 2, 15, vbSrcCopy
                    End If
                End If
                End If
            End If
            
'            If m_UseColumnColor = True Then
'                'Debug.Print cCol, ColColor(cCol)
'                Line (rct.left, m_HeaderHeight)-(rct.Right, ScaleHeight), ColColor(cCol), BF
'            End If
            
            rct.left = rct.left + 5
            rct.Right = rct.Right - 10
                
    'Debug.Print "1ST ", rct.left, rct.Right, rct.tOp, rct.Bottom
    
    If m_VerticalHeader = True Then
        f.lfEscapement = 10 * Val(-90) 'rotation angle, in tenths
        fontname = "Arial" + Chr$(0) 'null terminated
        f.lfFacename = fontname
        f.lfHeight = (FontSize * -20) / Screen.TwipsPerPixelY
        hFont = CreateFontIndirect(f)
        hPrevFont = SelectObject(UserControl.hdc, hFont)
        'DrawText UserControl.hdc, fText, Len(fText), rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or &H800
        CurrentX = rct.left + ColWidth(cCol) / 2 + TextHeight(fText) / 2 - 2
        CurrentY = 5
        
        If TextWidth(fText) + 10 > m_HeaderHeight Then
        Do
            fText = Mid(fText, 1, Len(fText) - 3) & ".."
            If Len(fText) = 2 Then fText = ""
        Loop Until TextWidth(fText) + 10 < (m_HeaderHeight) Or fText = ""
        End If
        If TextHeight(fText) + 4 > ColWidth(cCol) Then
            fText = ""
        End If
        
        Print fText
    'Debug.Print "2ND ", rct.left, rct.Right, rct.tOp, rct.Bottom
    Else
        DrawText UserControl.hdc, fText, Len(fText), rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER Or &H800
    End If
            
    hFont = SelectObject(UserControl.hdc, hPrevFont)
    DeleteObject hFont
    
    
            xCor = xCor + ColWidth(cCol)
        End If
    Next
End If

ForeColor = oldFore

    rct.left = 0
    rct.tOp = 0
    rct.Right = ScaleWidth
    rct.Bottom = ScaleHeight
    DrawEdge hdc, rct, BDR_SUNKENOUTER, BF_RECT
End Sub

Public Sub DrawList()
'Debug.Print m_TheData.Recordset.Fields(0), m_RowHeight
Dim nD As Long
Dim aBook As Variant
Dim cCol As Long
Dim xCor As Long
Dim fText As String
Dim rct As RECT
Dim checkw2 As Long
Dim cSize As Long
Dim dtSuccess As Boolean
On Error GoTo errCor

If m_ShowCheck = True Then
    checkw2 = CheckW
End If
'Dim ttCheck As alCheck
'Set ttCheck = New alCheck

'UserControl.AutoRedraw = True
    Cls
    UserControl.BackColor = m_BackColor
    rct.left = 0
    rct.tOp = 0
    rct.Right = ScaleWidth
    rct.Bottom = ScaleHeight
    DrawEdge hdc, rct, BDR_SUNKENOUTER, BF_RECT
    'Debug.Print VScroll1.Value, VScroll1.Max

Dim xx As Long
Dim yy As Long
Dim pWidth As Long
Dim pHeight As Long
Dim oldFore As Long

vScroll2_Paint

If Not m_Picture Is Nothing Then
    pWidth = ScaleX(m_Picture.Width, vbHimetric, vbPixels)
    pHeight = ScaleY(m_Picture.Height, vbHimetric, vbPixels)
    If pWidth = 0 Then Exit Sub
    For yy = 0 To Int(ScaleHeight / pHeight)
        For xx = 0 To Int(ScaleWidth / pWidth)
            UserControl.PaintPicture m_Picture, xx * pWidth, yy * pHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0), , , , , , , vbSrcCopy
        Next
    Next
End If

    If m_TheData Is Nothing Then Exit Sub
      
    Set tData = m_TheData.Recordset.Clone
'    aBook = m_TheData.Recordset.Bookmark
    tData.AbsolutePosition = gValue
    'Debug.Print tData.Index
Call DrawHeader


    If m_UseColumnColor = True Then
xCor = checkw2
For cCol = 0 To tData.Fields.Count - 1
    rct.left = xCor
    rct.Right = rct.left + ColWidth(cCol)
        Line (rct.left, m_HeaderHeight + 1)-(rct.Right, ScaleHeight), ColColor(cCol), BF
    xCor = xCor + ColWidth(cCol)
Next
    End If

Do
    
    xCor = checkw2
   
    If nD + gValue = m_Selected And Enabled = True Then
        'Line (1, nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0))-(ScaleWidth - vScroll2.ScaleWidth - 2, nD * m_RowHeight + m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0)), m_SelColor1, BF
        Call FillGradient(UserControl.hdc, 0, nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0), ScaleWidth - VScroll2.ScaleWidth, m_RowHeight, m_SelColor1, m_SelColor2, Fill_Horizontal, False)
    End If
    
    If UBound(aSelect) <> tData.RecordCount Then ReDim Preserve aSelect(tData.RecordCount)
    If aSelect(nD + gValue) = True And Enabled = True Then
        Call FillGradient(UserControl.hdc, 0, nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0), ScaleWidth - VScroll2.ScaleWidth, m_RowHeight, m_SelColor1, m_SelColor2, Fill_Horizontal, False)
    End If
    
    For cCol = 0 To tData.Fields.Count - 1
        If cCol = 0 Then
            CurrentX = 5
        Else
            CurrentX = xCor
        End If
    
    CurrentY = nD * m_RowHeight
    fText = IIf(IsNull(tData.Fields(cCol)), "", tData.Fields(cCol))
    
    
    If tData.Fields(cCol).Type <> 1 Then
        rct.tOp = nD * m_RowHeight
        rct.Bottom = rct.tOp + m_RowHeight
        rct.left = xCor + 5
        rct.Right = rct.left + ColWidth(cCol) - 10
        If m_ShowHeader = True Then
            rct.tOp = rct.tOp + m_HeaderHeight
            rct.Bottom = rct.Bottom + m_HeaderHeight
        End If
        DrawText UserControl.hdc, fText, Len(fText), rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER Or &H800
        If Enabled = False Then
            ForeColor = QBColor(15)
            'rct.left = rct.left + 1
            'rct.Right = rct.Right + 1
            rct.tOp = rct.tOp + 1
            rct.Bottom = rct.Bottom + 1
            'DrawText UserControl.hdc, fText, Len(fText), rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER Or &H800
            ForeColor = oldFore
        End If
    Else
        rct.left = xCor + 5
        rct.Right = rct.left + ColWidth(cCol) - 10
        rct.tOp = (nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0)) + (m_RowHeight / 2) - 7
        rct.Bottom = rct.tOp + 14
        If m_XPtheme = True Then
            If ColWidth(cCol) < 16 Then
                cSize = ColWidth(cCol)
            Else
                cSize = 16
            End If
            If m_RowHeight < cSize Then
                cSize = m_RowHeight
            End If
            If cSize > 2 Then
                dtSuccess = False
                If Me.Enabled = True Then
                    dtSuccess = DrawTheme(hdc, "Button", bp_checkbox, IIf(CBool(fText) = True, cbs_checkednormal, cbs_uncheckednormal), rct)
                Else
                    dtSuccess = DrawTheme(hdc, "Button", bp_checkbox, IIf(CBool(fText) = True, cbs_checkeddisabled, cbs_uncheckeddisabled), rct)
                End If
               'Debug.Print "DRAWTHEME"
                If dtSuccess = False Then
                     UserControl.PaintPicture Image7.Picture, rct.left - 5 + ColWidth(cCol) / 2 - cSize / 2, rct.tOp, cSize, cSize, 0, 0 + IIf(CBool(fText) = True, 64, 0) + IIf(Enabled = False, 48, 0), 16, 16, vbSrcCopy
                End If
            End If
        Else
                Call DrawFrameControl(UserControl.hdc, rct, DFC_BUTTON, DFCS_BUTTONCHECK Or IIf(CBool(fText) = True, DFCS_CHECKED, 0) Or DFCS_FLAT Or IIf(Me.Enabled = False, DFCS_INACTIVE, 0))
            'Debug.Print "DRAWNORMAL"
        End If
    End If
    
    If m_ShowGrid = True Then
        Line (xCor + ColWidth(cCol), nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0))-(xCor + ColWidth(cCol), nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0) + m_RowHeight), m_GridlineColor
    End If
    
    xCor = xCor + ColWidth(cCol)
    Next
    
    If m_ShowGrid = True Then
        Line (0, nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0))-(ScaleWidth, nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0)), m_GridlineColor
    End If
    
    If m_ShowCheck = True Then
        rct.left = 1
        rct.Right = CheckW - 2
        rct.tOp = (nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0)) + (m_RowHeight / 2) - 7
        rct.Bottom = rct.tOp + 14
        If UBound(aCheck) <> tData.RecordCount Then ReDim Preserve aCheck(tData.RecordCount)
        
        If m_XPtheme = False Then
            Call DrawFrameControl(UserControl.hdc, rct, DFC_BUTTON, DFCS_BUTTONCHECK Or IIf(aCheck(nD + gValue) = True, DFCS_CHECKED, 0) Or DFCS_FLAT Or IIf(Me.Enabled = False, DFCS_INACTIVE, 0))
        Else
            If CheckW < 16 Then
                cSize = CheckW
            Else
               cSize = 16
            End If
            If m_RowHeight < cSize Then
                cSize = m_RowHeight
            End If
            If cSize > 2 Then
                dtSuccess = False
                If Me.Enabled = True Then
                    dtSuccess = DrawTheme(hdc, "Button", bp_checkbox, IIf(aCheck(nD + gValue) = True, cbs_checkednormal, cbs_uncheckednormal), rct)
                Else
                    dtSuccess = DrawTheme(hdc, "Button", bp_checkbox, IIf(aCheck(nD + gValue) = True, cbs_checkeddisabled, cbs_uncheckeddisabled), rct)
                End If
                
                If dtSuccess = False Then
                    UserControl.PaintPicture Image7.Picture, rct.left + CheckW / 2 - cSize / 2, rct.tOp, cSize, cSize, 0, 0 + IIf(aCheck(nD + gValue) = True, 64, 0) + IIf(Enabled = False, 48, 0), 16, 16, vbSrcCopy
                End If
            End If
        End If
        
    End If
    
    'Debug.Print m_Selected, nD + gValue, HasFocus
    If nD + gValue = m_Selected And m_ShowFocusRect = True And HasFocus = True Then
        rct.left = 2
        rct.Right = ScaleWidth - VScroll2.ScaleWidth - 4
        rct.tOp = nD * m_RowHeight + IIf(m_ShowHeader = True, m_HeaderHeight, 0) + 2 '+ m_RowHeight
        rct.Bottom = rct.tOp + m_RowHeight - 4 '+ 100
        DrawFocusRect UserControl.hdc, rct
    End If
    
    nD = nD + 1
    tData.MoveNext
Loop Until nD = NumFill Or tData.EOF

    rct.left = 0
    rct.tOp = 0
    rct.Right = ScaleWidth
    rct.Bottom = ScaleHeight
    DrawEdge hdc, rct, BDR_SUNKENOUTER, BF_RECT

If Enabled = False Then
    UserControl.DrawMode = 9
    Line (0, m_HeaderHeight)-(ScaleWidth - VScroll2.ScaleWidth, ScaleHeight), RGB(230, 230, 230), BF
    UserControl.DrawMode = 13
End If

'tData.Bookmark = aBook
UserControl.Refresh

Exit Sub
errCor:
Debug.Print "Drawlist error "; Rnd
End Sub

Public Function SortSql(colIndex As Long) As String
Dim cc As Long
Dim tmpStr As String
Dim astr As Long
astr = InStr(1, m_TheData.RecordSource, "FROM", vbTextCompare)
If astr > 0 Then
    'Debug.Print tmpStr, astr
    tmpStr = Mid(m_TheData.RecordSource, astr + 6, Len(m_TheData.RecordSource) - astr - 6)
    astr = InStr(1, tmpStr, "ORDER", vbTextCompare)
    tmpStr = Mid(tmpStr, 1, astr - 3)
    m_TheData.RecordSource = tmpStr
    'Debug.Print tmpStr, astr
    'Exit Function
End If

tmpStr = ""
For cc = 0 To m_TheData.Recordset.Fields.Count - 1
    tmpStr = tmpStr & "[" & m_TheData.RecordSource & "].[" & m_TheData.Recordset.Fields(cc).Name & "], "
Next
tmpStr = Mid(tmpStr, 1, Len(tmpStr) - 2)
tmpStr = "SELECT " & tmpStr & " FROM [" & m_TheData.RecordSource & "] ORDER BY " & "[" & m_TheData.RecordSource & "].[" & m_TheData.Recordset.Fields(colIndex).Name & "]"
'Debug.Print tmpStr
SortSql = tmpStr
End Function

Public Function GetCheck(tIndex As Long) As Boolean
    If tIndex < -1 And tIndex <= UBound(aCheck) Then
        GetCheck = aCheck(tIndex)
    End If
End Function

Public Sub SetCheck(tIndex As Long, isChecked As Boolean)
    If tIndex < -1 And tIndex <= UBound(aCheck) Then
        aCheck(tIndex) = isChecked
    End If
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procedure : FillGradient
' Auther    : Jim Jose
' Input     : Hdc + Parameters
' OutPut    : None
' Purpose   : Middleout Gradients with Carls's DIB solution
'-------------------------------------------------------------------------------------------------------------------------

Private Sub FillGradient(ByVal hdc As Long, _
                         ByVal x As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum, _
                         Optional Right2Left As Boolean = True)
                         
Dim tmpCol  As Long
  
    ' Exit if needed
    If GradientDirection = Fill_None Then Exit Sub
    
    ' Right-To-Left
    If Right2Left Then
        tmpCol = Col1
        Col1 = Col2
        Col2 = tmpCol
    End If
    
    Select Case GradientDirection
        Case Fill_HorizontalMiddleOut
            DIBGradient hdc, x, Y, Width / 2, Height, Col1, Col2, Fill_Horizontal
            DIBGradient hdc, x + Width / 2 - 1, Y, Width / 2, Height, Col2, Col1, Fill_Horizontal

        Case Fill_VerticalMiddleOut
            DIBGradient hdc, x, Y, Width, Height / 2, Col1, Col2, Fill_Vertical
            DIBGradient hdc, x, Y + Height / 2 - 1, Width, Height / 2, Col2, Col1, Fill_Vertical

        Case Else
            DIBGradient hdc, x, Y, Width, Height, Col1, Col2, GradientDirection
    End Select
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procedure : DIBGradient
' Auther    : Carls P.V.
' Input     : Hdc + Parameters
' OutPut    : None
' Purpose   : DIB solution for fast gradients
'-------------------------------------------------------------------------------------------------------------------------

Private Sub DIBGradient(ByVal hdc As Long, _
                         ByVal x As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum)

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
    '-- A minor check
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    '-- Decompose Cols
    Col1 = Col1 And &HFFFFFF
    R1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    G1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    B1 = Col1 Mod &H100&
    Col2 = Col2 And &HFFFFFF
    R2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    G2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    B2 = Col2 Mod &H100&
    
    '-- Get Col distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-Cols array
    Select Case GradientDirection
        Case [Fill_Horizontal]
            ReDim lGrad(0 To Width - 1)
        Case [Fill_Vertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-Cols
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [Fill_Horizontal]
        
            For j = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next j
        
        Case [Fill_Vertical]
        
            For j = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(j)
                Next i
                iOffset = iOffset + Scan
            Next j
            
        Case [Fill_DownwardDiagonal]
            
            iOffset = jEnd * Scan
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = j
            Next j
            
        Case [Fill_UpwardDiagonal]
            
            iOffset = 0
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = j
            Next j
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hdc, x, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)

End Sub

Sub RefreshInfo()
    On Error Resume Next
    If m_RowHeight > 0 Then
        NumFill = Int(IIf(m_ShowHeader = False, ScaleHeight, ScaleHeight - m_HeaderHeight) / m_RowHeight)
        If NumFill < 1 Then NumFill = 1
    End If
    If UBound(aCheck) <> m_TheData.Recordset.RecordCount Then
        ReDim Preserve aCheck(m_TheData.Recordset.RecordCount)
    End If
    If UBound(aSelect) <> m_TheData.Recordset.RecordCount Then
        ReDim Preserve aSelect(m_TheData.Recordset.RecordCount)
    End If
    gMax = m_TheData.Recordset.RecordCount - NumFill
End Sub

Public Function GetWidth(ColNum As Long) As Long
    GetWidth = ColWidth(ColNum)
End Function

Public Sub SetWidth(ColNum As Long, tColWidth As Long)
    ColWidth(ColNum) = tColWidth
End Sub

Public Function GetColColor(ColNum As Long) As Long
    GetColColor = ColColor(ColNum)
End Function

Public Sub SetColColor(ColNum As Long, tColColor As Long)
    ColColor(ColNum) = tColColor
End Sub


'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    UserControl.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim dselX As Long
Dim dselY As Long
Dim dComma As Long
Dim aBook

On Error Resume Next
If KeyAscii = 13 Then
    dComma = InStr(1, Text1.Tag, ",", vbTextCompare)
    dselX = CLng(Mid(Text1.Tag, 1, dComma))
    dselY = CLng(Mid(Text1.Tag, dComma + 1, Len(Text1.Tag) - dComma))
    
    aBook = m_TheData.Recordset.Bookmark
    m_TheData.Recordset.AbsolutePosition = dselY
    'tStr = m_TheData.Recordset.Fields(dselX - 1)
    If m_TheData.Recordset.Fields(dselX - 1).Type <> 1 Then
        m_TheData.Recordset.Edit
        m_TheData.Recordset.Fields(dselX - 1) = Text1.Text
        m_TheData.Recordset.Update
    End If
    m_TheData.Recordset.Bookmark = aBook
    Text1.Visible = False
    KeyAscii = 0
    Call DrawList
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Visible = False
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
Dim cc As Long
Dim cad As Long
Dim cadf As Boolean
Dim checkw2 As Long
Dim tHead As Long
Dim dselX As Long
Dim dselY As Long
Dim aBook
Dim tStr As String

If m_ShowCheck = True Then checkw2 = CheckW
    
If m_ShowCheck = True And DblX < CheckW And DblY > IIf(m_ShowHeader = True, m_HeaderHeight, 0) Then
    CheckDown = True
End If

If m_ShowHeader = True Then
    tHead = m_HeaderHeight
End If
    
dselY = Int((DblY - tHead) / m_RowHeight) + gValue
If m_ShowHeader = True And DblY < m_HeaderHeight Then dselY = -1
    
If m_ShowCheck = True Then cad = checkw2
    For cc = 0 To UBound(ColWidth)
        cad = cad + ColWidth(cc)
        If DblX > cad - ColWidth(cc) And DblX < cad And ColSize = -1 Then
            dselX = cc + 1
            Exit For
        End If
    Next
    
    'Debug.Print DblX, DblY, DblB, dSelx, dSely, tStr
If dselX > 0 And m_IsEditable = True And dselY > -1 Then
    aBook = m_TheData.Recordset.Bookmark
    m_TheData.Recordset.AbsolutePosition = dselY
    If IsNull(m_TheData.Recordset.Fields(dselX - 1)) Then
        tStr = ""
    Else
        tStr = m_TheData.Recordset.Fields(dselX - 1)
    End If
    If m_TheData.Recordset.Fields(dselX - 1).Type = 1 Then
        m_TheData.Recordset.Edit
        m_TheData.Recordset.Fields(dselX - 1) = Not CBool(tStr)
        m_TheData.Recordset.Update
    Else
        Text1.tOp = (dselY - gValue) * m_RowHeight + tHead
        Text1.Height = m_RowHeight
        'If cc = 0 Then
        '    Text1.left = checkw2
        'Else
            Text1.left = cad - ColWidth(cc)
        'End If
        Text1.Width = ColWidth(cc)
        Text1.Text = tStr
        Text1.Tag = dselX & "," & dselY
        Text1.Visible = True
        Text1.SetFocus
        Text1.SelLength = Len(Text1.Text)
    End If
    'Debug.Print DblX, DblY, DblB, dSelx, dSely, tStr
    m_TheData.Recordset.Bookmark = aBook
    Call DrawList
End If



    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    HasFocus = True
    Call DrawList
End Sub

Private Sub UserControl_Initialize()
    Set aTrack = New clsTracking
    gMax = 3
    ColSize = -1
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
If KeyCode = vbKeyUp Then
    aTrack_ScrollUp
ElseIf KeyCode = vbKeyDown Then
    aTrack_ScrollDown
End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    HasFocus = False
    Call DrawList
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
isMouseDown = True
Text1.Visible = False

Dim cc As Long
Dim cad As Long
Dim cadf As Boolean
Dim checkw2 As Long
If m_ShowCheck = True Then checkw2 = CheckW
    
If m_ShowCheck = True And x < CheckW And Y > IIf(m_ShowHeader = True, m_HeaderHeight, 0) Then
    CheckDown = True
End If

    Dim tHead As Long
    If m_ShowHeader = True Then
        tHead = m_HeaderHeight
    End If
    
Dim dSel As Long
If Not GetKeyState(vbKeyShift) >= 0 And m_MultiSelect = True Then
    dSel = Int((Y - tHead) / m_RowHeight) + gValue
    For cc = m_Selected To dSel Step IIf(dSel > m_Selected, 1, -1)
            If aSelect(m_Selected) = False Or cc <> m_Selected Then
                aSelect(cc) = Not aSelect(cc)
            End If
    Next
    Call DrawList
End If
sold = m_Selected

If m_ShowHeader = True And m_ColumnFixedWidth = False Then
    If Y < m_HeaderHeight Then
        cad = checkw2
        For cc = 0 To UBound(ColWidth)
            cad = cad + ColWidth(cc)
            If x > cad - 2 And x < cad + 2 Then
                UserControl.MousePointer = 6
                ColSize = cc
                colSizeN = x
                cadf = True
            End If
        Next
        If cadf = False Then
            UserControl.MousePointer = 0
        End If
    Else
        UserControl.MousePointer = 0
    End If
End If

    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim cc As Long
Dim cad As Long

If mdScroll <> 0 Then
    mdScroll = 0
    vScroll2_Paint
End If

If isMouseDown = True Then
    If ColSize > -1 Then
        'Debug.Print X - colSizeN, ColSize, ColWidth(ColSize), X
        ColWidth(ColSize) = IIf(ColWidth(ColSize) + x - colSizeN < 0, 0, ColWidth(ColSize) + x - colSizeN)
        If ColWidth(ColSize) + x - colSizeN > 0 Then
            colSizeN = x
        End If
        'Debug.Print "Drawlist 1"
        Call DrawList
        'Call DrawHeader
        Exit Sub
    End If
    
    Dim tHead As Long
    
    If m_ShowHeader = True Then
        tHead = m_HeaderHeight
    End If
    
    If Y >= tHead And Int((Y - tHead) / m_RowHeight) < NumFill Then
        m_Selected = Int((Y - tHead) / m_RowHeight) + gValue
        If GetKeyState(vbKeyControl) >= 0 Then
            ReDim aSelect(m_TheData.Recordset.RecordCount)
        Else
            If m_MultiSelect = True Then
            If sold > -1 Then
                aSelect(sold) = True
            End If
            End If
        End If
        'Debug.Print "Drawlist 2"
        Call DrawList
    End If
End If

'Dim cad As Long
Dim cadf As Boolean
Dim checkw2 As Long
If m_ShowCheck = True Then checkw2 = CheckW

If m_ShowHeader = True Then
    cad = checkw2
    'HeadHover = 0
        If Y < m_HeaderHeight Then
        For cc = 0 To UBound(ColWidth)
            cad = cad + ColWidth(cc)
            If x > cad - ColWidth(cc) And x < cad And ColSize = -1 Then
                'Debug.Print cc, Rnd
                If cc + 1 <> HeadHover Then
                    'Debug.Print "Drawlist 3", cc + 1, HeadHover, Rnd
                    HeadHover = cc + 1
                    'Debug.Print HeadHover
                    'Call DrawList
                    Call DrawHeader
                End If
                Exit For
            End If
        Next
        'If x > cad Then
        '    Debug.Print "to long"; Rnd
        'End If
        If x < checkw2 Or x > cad Then
            HeadHover = 0
            Call DrawHeader
        End If
        cad = 0
        Else
            If HeadHover <> 0 Then
            HeadHover = 0
            'Debug.Print "Drawlist 4"
            'Call DrawList
            Call DrawHeader
            End If
        End If
        
    
    cad = checkw2
    If Y < m_HeaderHeight And m_ColumnFixedWidth = False Then
        For cc = 0 To UBound(ColWidth)
            cad = cad + ColWidth(cc)
            If x > cad - 2 And x < cad + 2 Then
                UserControl.MousePointer = 6
                HeadHover = 0
                cadf = True
                Exit For
            End If
        Next
        If cadf = False Then
            UserControl.MousePointer = 0
        End If
    Else
        UserControl.MousePointer = 0
    End If
Else
    UserControl.MousePointer = 0
End If
        
        RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim tHead As Long
    Dim sold As Long
    
DblX = x
DblY = Y
DblB = Button

    If m_ShowHeader = True Then
        tHead = m_HeaderHeight
    End If
    
    Dim cad As Long
    Dim cc As Long
    If m_ShowHeader = True Then
    If Y < m_HeaderHeight Then
        cad = CheckW
        For cc = 0 To UBound(ColWidth)
            If x > cad And x < cad + ColWidth(cc) Then
                RaiseEvent HeaderClick(cc)
            End If
            cad = cad + ColWidth(cc)
        Next
    End If
    End If
    
    If Y >= tHead And Int((Y - tHead) / m_RowHeight) < NumFill And isMouseDown = True And ColSize = -1 Then
        sold = m_Selected
        m_Selected = Int((Y - tHead) / m_RowHeight) + gValue
        If GetKeyState(vbKeyControl) >= 0 Then
            If GetKeyState(vbKeyShift) >= 0 Then
                ReDim aSelect(m_TheData.Recordset.RecordCount)
            End If
        Else
            If m_MultiSelect = True Then
            If aSelect(Int((Y - tHead) / m_RowHeight) + gValue) = True Then
                m_Selected = -1
            End If
            If sold > -1 Then
                aSelect(sold) = True
            End If
            aSelect(Int((Y - tHead) / m_RowHeight) + gValue) = Not aSelect(Int((Y - tHead) / m_RowHeight) + gValue)
            End If
        End If
        Call DrawList
        RaiseEvent ItemSelected(m_Selected)
    End If
    
        
    If CheckDown = True Then
        aCheck(Int((Y - tHead) / m_RowHeight) + gValue) = Not aCheck(Int((Y - tHead) / m_RowHeight) + gValue)
        CheckDown = False
        Call DrawList
    End If
    isMouseDown = False
    ColSize = -1
        RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=15,0,0,0
Public Property Get theData() As Object
    Set theData = m_TheData
End Property

Public Property Set theData(ByVal New_TheData As Control)
    Set m_TheData = New_TheData
    PropertyChanged "TheData"
    
    m_ColumnCount = m_TheData.Recordset.Fields.Count
    PropertyChanged "ColumnCount"
    
    Dim aBook As Variant
    aBook = m_TheData.Recordset.Bookmark
    m_TheData.Recordset.MoveLast
    'vScroll.Max = m_TheData.Recordset.RecordCount - NumFill
    m_TheData.Recordset.Bookmark = aBook
    'Debug.Print "Count ", m_TheData.Recordset.AbsolutePosition
    ReDim ColWidth(m_TheData.Recordset.Fields.Count - 1)
    ReDim ColColor(m_TheData.Recordset.Fields.Count - 1)
    Dim cc As Long
    For cc = 1 To UBound(ColWidth) + 1
        ColWidth(cc - 1) = (ScaleWidth - VScroll2.ScaleWidth - CheckW - 5) / (UBound(ColWidth) + 1) '* cc
    Next
    Call RefreshInfo
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_ColumnCount = m_def_ColumnCount
    m_RowHeight = m_def_RowHeight
    m_ShowGrid = m_def_ShowGrid
    m_ShowHeader = m_def_ShowHeader
    m_HeaderHeight = m_def_HeaderHeight
    m_Selected = m_def_Selected
    Set m_Picture = LoadPicture("")
    m_ShowCheck = m_def_ShowCheck
    m_ShowFocusRect = m_def_ShowFocusRect
    m_MultiSelect = m_def_MultiSelect
    m_Headercolor = m_def_Headercolor
    m_SelColor1 = m_def_SelColor1
    m_SelColor2 = m_def_SelColor2
    m_HeaderTextColor = m_def_HeaderTextColor
    m_ColumnFixedWidth = m_def_ColumnFixedWidth
    m_ScrollbarBackColor = m_def_ScrollbarBackColor
    m_ScrollbarBARcolor = m_def_ScrollbarBARcolor
    m_BackColor = m_def_BackColor
    m_GridlineColor = m_def_GridlineColor
    m_XPtheme = m_def_XPtheme
    m_VerticalHeader = m_def_VerticalHeader
    m_UseColumnColor = m_def_UseColumnColor
    m_IsEditable = m_def_IsEditable
End Sub

Private Sub UserControl_Paint()
Call DrawList
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Set aTrack = New clsTracking
aTrack.hwnd = UserControl.hwnd
aTrack.ScrollHwnd = VScroll2.hwnd
aTrack.TextHwnd = Text1.hwnd

If Ambient.UserMode Then
    
    Dim trk As tagTRACKMOUSEEVENT
    trk.cbSize = 16
    trk.dwFlags = TME_LEAVE Or TME_HOVER
    trk.dwHoverTime = 10
    trk.hwndTrack = UserControl.hwnd
    
    Hook aTrack
End If

'    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set m_TheData = PropBag.ReadProperty("TheData", Nothing)
    m_ColumnCount = PropBag.ReadProperty("ColumnCount", m_def_ColumnCount)
    m_RowHeight = PropBag.ReadProperty("RowHeight", m_def_RowHeight)
    m_ShowGrid = PropBag.ReadProperty("ShowGrid", m_def_ShowGrid)
    m_ShowHeader = PropBag.ReadProperty("ShowHeader", m_def_ShowHeader)
    m_HeaderHeight = PropBag.ReadProperty("HeaderHeight", m_def_HeaderHeight)
    m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_ShowCheck = PropBag.ReadProperty("ShowCheck", m_def_ShowCheck)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 0)
    m_ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", m_def_ShowFocusRect)
    m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
    m_Headercolor = PropBag.ReadProperty("Headercolor", m_def_Headercolor)
    m_SelColor1 = PropBag.ReadProperty("SelColor1", m_def_SelColor1)
    m_SelColor2 = PropBag.ReadProperty("SelColor2", m_def_SelColor2)
    m_HeaderTextColor = PropBag.ReadProperty("HeaderTextColor", m_def_HeaderTextColor)
    m_ColumnFixedWidth = PropBag.ReadProperty("ColumnFixedWidth", m_def_ColumnFixedWidth)
    m_ScrollbarBackColor = PropBag.ReadProperty("ScrollbarBackColor", m_def_ScrollbarBackColor)
    m_ScrollbarBARcolor = PropBag.ReadProperty("ScrollbarBARcolor", m_def_ScrollbarBARcolor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_GridlineColor = PropBag.ReadProperty("GridlineColor", m_def_GridlineColor)
    m_XPtheme = PropBag.ReadProperty("XPtheme", m_def_XPtheme)
    m_VerticalHeader = PropBag.ReadProperty("VerticalHeader", m_def_VerticalHeader)
    m_UseColumnColor = PropBag.ReadProperty("UseColumnColor", m_def_UseColumnColor)
    m_IsEditable = PropBag.ReadProperty("IsEditable", m_def_IsEditable)
End Sub

Private Sub UserControl_Resize()
VScroll2.Move ScaleWidth - VScroll2.ScaleWidth - 1, 1, VScroll2.Width, ScaleHeight - 2
'vScroll2.Move 0, 1, 17, ScaleHeight - 2
Call RefreshInfo
Call DrawList
End Sub

Private Sub UserControl_Show()
    Call DrawList
End Sub

Private Sub UserControl_Terminate()
Set m_TheData = Nothing
Set tData = Nothing
UnHook aTrack
Set aTrack = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("TheData", m_TheData, Nothing)
    Call PropBag.WriteProperty("ColumnCount", m_ColumnCount, m_def_ColumnCount)
    Call PropBag.WriteProperty("RowHeight", m_RowHeight, m_def_RowHeight)
    Call PropBag.WriteProperty("ShowGrid", m_ShowGrid, m_def_ShowGrid)
    Call PropBag.WriteProperty("ShowHeader", m_ShowHeader, m_def_ShowHeader)
    Call PropBag.WriteProperty("HeaderHeight", m_HeaderHeight, m_def_HeaderHeight)
    Call PropBag.WriteProperty("Selected", m_Selected, m_def_Selected)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("ShowCheck", m_ShowCheck, m_def_ShowCheck)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 0)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowFocusRect, m_def_ShowFocusRect)
    Call PropBag.WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
    Call PropBag.WriteProperty("Headercolor", m_Headercolor, m_def_Headercolor)
    Call PropBag.WriteProperty("SelColor1", m_SelColor1, m_def_SelColor1)
    Call PropBag.WriteProperty("SelColor2", m_SelColor2, m_def_SelColor2)
    Call PropBag.WriteProperty("HeaderTextColor", m_HeaderTextColor, m_def_HeaderTextColor)
    Call PropBag.WriteProperty("ColumnFixedWidth", m_ColumnFixedWidth, m_def_ColumnFixedWidth)
    Call PropBag.WriteProperty("ScrollbarBackColor", m_ScrollbarBackColor, m_def_ScrollbarBackColor)
    Call PropBag.WriteProperty("ScrollbarBARcolor", m_ScrollbarBARcolor, m_def_ScrollbarBARcolor)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("GridlineColor", m_GridlineColor, m_def_GridlineColor)
    Call PropBag.WriteProperty("XPtheme", m_XPtheme, m_def_XPtheme)
    Call PropBag.WriteProperty("VerticalHeader", m_VerticalHeader, m_def_VerticalHeader)
    Call PropBag.WriteProperty("UseColumnColor", m_UseColumnColor, m_def_UseColumnColor)
    Call PropBag.WriteProperty("IsEditable", m_IsEditable, m_def_IsEditable)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get ColumnCount() As Long
Attribute ColumnCount.VB_MemberFlags = "400"
    ColumnCount = m_ColumnCount
End Property

Public Property Let ColumnCount(ByVal New_ColumnCount As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_ColumnCount = New_ColumnCount
    PropertyChanged "ColumnCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,20
Public Property Get RowHeight() As Long
    RowHeight = m_RowHeight
End Property

Public Property Let RowHeight(ByVal New_RowHeight As Long)
    m_RowHeight = New_RowHeight
    PropertyChanged "RowHeight"
    Call RefreshInfo
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowGrid() As Boolean
    ShowGrid = m_ShowGrid
End Property

Public Property Let ShowGrid(ByVal New_ShowGrid As Boolean)
    m_ShowGrid = New_ShowGrid
    PropertyChanged "ShowGrid"
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowHeader() As Boolean
    ShowHeader = m_ShowHeader
End Property

Public Property Let ShowHeader(ByVal New_ShowHeader As Boolean)
    m_ShowHeader = New_ShowHeader
    PropertyChanged "ShowHeader"
    Call RefreshInfo
    Call DrawList
    'Call DrawHeader
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,20
Public Property Get HeaderHeight() As Long
    HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal New_HeaderHeight As Long)
    m_HeaderHeight = New_HeaderHeight
    PropertyChanged "HeaderHeight"
    Call DrawList
    'Call DrawHeader
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,-1
Public Property Get Selected() As Integer
    Selected = m_Selected
End Property

Public Property Let Selected(ByVal New_Selected As Integer)
    m_Selected = New_Selected
    PropertyChanged "Selected"
    
    If New_Selected <= gValue Or New_Selected >= gValue + NumFill Then
    If New_Selected > gMax Then
        gValue = gMax
    Else
        gValue = New_Selected
    End If
    End If
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowCheck() As Boolean
    ShowCheck = m_ShowCheck
End Property

Public Property Let ShowCheck(ByVal New_ShowCheck As Boolean)
    m_ShowCheck = New_ShowCheck
    PropertyChanged "ShowCheck"
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    m_MultiSelect = New_MultiSelect
    PropertyChanged "MultiSelect"
    If New_MultiSelect = False Then ReDim aSelect(m_TheData.Recordset.RecordCount)
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Headercolor() As OLE_COLOR
    Headercolor = m_Headercolor
End Property

Public Property Let Headercolor(ByVal New_Headercolor As OLE_COLOR)
    m_Headercolor = New_Headercolor
    PropertyChanged "Headercolor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SelColor1() As OLE_COLOR
    SelColor1 = m_SelColor1
End Property

Public Property Let SelColor1(ByVal New_SelColor1 As OLE_COLOR)
    m_SelColor1 = New_SelColor1
    PropertyChanged "SelColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SelColor2() As OLE_COLOR
    SelColor2 = m_SelColor2
End Property

Public Property Let SelColor2(ByVal New_SelColor2 As OLE_COLOR)
    m_SelColor2 = New_SelColor2
    PropertyChanged "SelColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HeaderTextColor() As OLE_COLOR
    HeaderTextColor = m_HeaderTextColor
End Property

Public Property Let HeaderTextColor(ByVal New_HeaderTextColor As OLE_COLOR)
    m_HeaderTextColor = New_HeaderTextColor
    PropertyChanged "HeaderTextColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ColumnFixedWidth() As Boolean
    ColumnFixedWidth = m_ColumnFixedWidth
End Property

Public Property Let ColumnFixedWidth(ByVal New_ColumnFixedWidth As Boolean)
    m_ColumnFixedWidth = New_ColumnFixedWidth
    PropertyChanged "ColumnFixedWidth"
End Property

Private Sub vScroll2_GotFocus()
HasFocus = True
End Sub

Private Sub vScroll2_LostFocus()
HasFocus = False
End Sub

Sub ResetHover()
mdHover = 0
vScroll2_Paint
End Sub

Private Sub vScroll2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim NextTime As Long
Dim gDraw As Long
Dim gBox As Double
Dim gTop As Double
gDraw = VScroll2.ScaleHeight - VScroll2.ScaleWidth * 2
gBox = gDraw / (gMax + 1 - gMin)
If gBox < 20 Then gBox = 20
gTop = ((gValue * (gDraw - gBox)) / (gMax - gMin)) + VScroll2.ScaleWidth
    NextTime = timeGetTime + 500

If Button = 1 Then
    If Y <= VScroll2.ScaleWidth Then
        mdScroll = 1
        Call vScroll2_Paint
    ElseIf Y >= VScroll2.ScaleHeight - VScroll2.ScaleWidth Then
        mdScroll = 2
        Call vScroll2_Paint
    ElseIf Y > VScroll2.ScaleWidth And Y < gTop Then
        mdScroll = 3
        Call vScroll2_Paint
    ElseIf Y > gTop + gBox And Y < VScroll2.ScaleHeight - VScroll2.ScaleWidth Then
        mdScroll = 4
        Call vScroll2_Paint
    ElseIf Y >= gTop And Y < gTop + gBox Then
        mdHover = 0
        mdScroll = 5
    Else
        mdScroll = 0
    End If
End If

    Do While mdScroll > 0
    
    If mdScroll = 1 Or mdScroll = 3 Then
        If gValue > gMin Then
            'gValue = gValue - 1
            'Call DrawList
            aTrack_ScrollUp
        End If
    ElseIf mdScroll = 2 Or mdScroll = 4 Then
        If gValue < gMax Then
            'gValue = gValue + 1
            'Call DrawList
            aTrack_ScrollDown
        End If
    End If
        Do While timeGetTime() < NextTime
            vScroll2_Paint
            DoEvents
        Loop
        NextTime = timeGetTime() + 20
    Loop
End Sub

Private Sub vScroll2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim gDraw As Long
Dim gBox As Double
Dim gTop As Double
gDraw = VScroll2.ScaleHeight - VScroll2.ScaleWidth * 2
gBox = gDraw / (gMax + 1 - gMin)
If gBox < 20 Then gBox = 20
gTop = ((gValue * (gDraw - gBox)) / (gMax - gMin)) + VScroll2.ScaleWidth
 
If Y > gTop And Y < gTop + gBox And Button = 0 Then
    mdHover = 5
    Call vScroll2_Paint
ElseIf Y < VScroll2.ScaleWidth Then
    mdHover = 1
    Call vScroll2_Paint
ElseIf Y > VScroll2.ScaleHeight - VScroll2.ScaleWidth Then
    mdHover = 2
    Call vScroll2_Paint
Else
    mdHover = 0
    Call vScroll2_Paint
End If
 
    
Dim pos As Double
If mdScroll = 5 Then
    pos = ((Y - VScroll2.ScaleWidth * 1.25) * (gMax + 1 - gMin)) / gDraw
    pos = Round(pos, 0)
    If pos >= gMax Then pos = gMax
    If pos <= gMin Then pos = gMin
    If pos <> gValue Then
        gValue = pos
        Call vScroll2_Paint
        Call DrawList
    End If
End If
End Sub

Private Sub vScroll2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
mdScroll = 0
Call vScroll2_Paint
End Sub

Private Sub vScroll2_Paint()
Dim tmpRct As RECT
Dim gDraw As Double
Dim gBox As Double
Dim gTop As Double
Dim dtSuccess As Boolean
Dim rct As RECT

gDraw = VScroll2.ScaleHeight - VScroll2.ScaleWidth * 2
gBox = gDraw / (gMax - gMin)
If gBox < 20 Then gBox = 20
gTop = ((gValue * (gDraw - gBox)) / (gMax - gMin)) + VScroll2.ScaleWidth

VScroll2.BackColor = m_ScrollbarBackColor
'vScroll2.BackColor = 0
VScroll2.Cls
'vScroll2.BackColor = QBColor(15)


If m_XPtheme = False Then
VScroll2.FontSize = 10
'FillGradient vScroll2.hdc, 0, 0, vScroll2.ScaleWidth, vScroll2.ScaleHeight, RGB(200, 200, 200), RGB(255, 255, 255), Fill_VerticalMiddleOut, False

If mdScroll = 3 Then
    VScroll2.Line (0, VScroll2.ScaleWidth)-(VScroll2.ScaleWidth, gTop), 0, BF
ElseIf mdScroll = 4 Then
    VScroll2.Line (0, gTop + gBox)-(ScaleWidth, VScroll2.ScaleHeight - VScroll2.ScaleWidth), 0, BF
End If

tmpRct.tOp = 0
tmpRct.Bottom = VScroll2.ScaleWidth
tmpRct.left = 0
tmpRct.Right = VScroll2.ScaleWidth
'DrawFrameControl vScroll2.hdc, tmpRct, DFC_SCROLL, DFCS_SCROLLUP Or DFCS_FLAT Or IIf(mdScroll = 1, DFCS_PUSHED Or DFCS_MONO, 0)
VScroll2.CurrentX = VScroll2.ScaleWidth / 2 - VScroll2.TextWidth("5") / 2
VScroll2.CurrentY = VScroll2.ScaleWidth / 2 - VScroll2.TextHeight("5") / 2

VScroll2.Print "5"
DrawEdge VScroll2.hdc, tmpRct, IIf(mdScroll = 1, BDR_RAISEDOUTER, BDR_RAISEDINNER), BF_RECT

tmpRct.tOp = VScroll2.Height - VScroll2.ScaleWidth
tmpRct.Bottom = VScroll2.Height
'DrawFrameControl vScroll2.hdc, tmpRct, DFC_SCROLL, DFCS_SCROLLDOWN Or DFCS_FLAT Or IIf(mdScroll = 2, DFCS_PUSHED Or DFCS_MONO, 0)
VScroll2.CurrentX = VScroll2.ScaleWidth / 2 - VScroll2.TextWidth("6") / 2
VScroll2.CurrentY = VScroll2.ScaleWidth / 2 - VScroll2.TextHeight("6") / 2 + gDraw + VScroll2.ScaleWidth
VScroll2.Print "6"
DrawEdge VScroll2.hdc, tmpRct, IIf(mdScroll = 2, BDR_RAISEDOUTER, BDR_RAISEDINNER), BF_RECT

'gMax = 4

VScroll2.Line (0, gTop)-(VScroll2.ScaleWidth, gTop + gBox - 1), m_ScrollbarBARcolor, BF
tmpRct.tOp = gTop
tmpRct.Bottom = tmpRct.tOp + gBox
DrawEdge VScroll2.hdc, tmpRct, IIf(mdScroll = 5, BDR_RAISEDOUTER, BDR_RAISEDINNER), BF_RECT


Else
    VScroll2.PaintPicture Image3.Picture, 0, 0, VScroll2.ScaleWidth, VScroll2.ScaleHeight, 0, 0, 17, 34, vbSrcCopy
    
    If mdScroll = 3 Then
        VScroll2.PaintPicture Image3.Picture, 0, 0, VScroll2.ScaleWidth, gTop, 0, 34, 17, 17, vbSrcCopy
    ElseIf mdScroll = 4 Then
        VScroll2.PaintPicture Image3.Picture, 0, gTop + gBox, VScroll2.ScaleWidth, VScroll2.ScaleHeight - gTop + gBox, 0, 34, 17, 17, vbSrcCopy
    End If
    
    If Enabled = True Then
        
        dtSuccess = False
        SetRect rct, 0, 0, 17, 17
        If mdHover = 1 And mdScroll = 0 Then
            dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_arrowbtn, abs_uphot, rct)
            If dtSuccess = False Then
                VScroll2.PaintPicture Image1.Picture, 0, 0, 17, 17, 0, 17, 17, 17, vbSrcCopy
            End If
        Else
            dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_arrowbtn, IIf(mdScroll = 2, abs_uppressed, abs_upnormal), rct)
            If dtSuccess = False Then
                VScroll2.PaintPicture Image1.Picture, 0, 0, 17, 17, 0, IIf(mdScroll = 1, 34, 0), 17, 17, vbSrcCopy
            End If
        End If
        If dtSuccess = False Then
            TransparentBlt2 VScroll2.hdc, Image2.Picture, 4, 4, 9, 9, 0, 0, RGB(255, 0, 255)
        End If
        
        dtSuccess = False
        SetRect rct, 0, VScroll2.ScaleHeight - 17, 17, VScroll2.ScaleHeight
        If mdHover = 2 And mdScroll = 0 Then
            dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_arrowbtn, abs_downhot, rct)
            If dtSuccess = False Then
                VScroll2.PaintPicture Image1.Picture, 0, VScroll2.ScaleHeight - 17, 17, 17, 0, 17, 17, 17, vbSrcCopy
            End If
        Else
            dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_arrowbtn, IIf(mdScroll = 2, abs_downpressed, abs_downnormal), rct)
            If dtSuccess = False Then
                VScroll2.PaintPicture Image1.Picture, 0, VScroll2.ScaleHeight - 17, 17, 17, 0, IIf(mdScroll = 2, 34, 0), 17, 17, vbSrcCopy
            End If
        End If
        If dtSuccess = False Then
            TransparentBlt2 VScroll2.hdc, Image2.Picture, 4, VScroll2.ScaleHeight - 13, 9, 9, 0, 36, RGB(255, 0, 255)
        End If

        
    Else
        dtSuccess = False
        SetRect rct, 0, 0, 17, 17
        dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_arrowbtn, abs_updisabled, rct)
        If dtSuccess = False Then
            VScroll2.PaintPicture Image1.Picture, 0, 0, 17, 17, 0, 51, 17, 17, vbSrcCopy
            TransparentBlt2 VScroll2.hdc, Image2.Picture, 4, 4, 9, 9, 0, 0, RGB(255, 0, 255)
        End If
        
        dtSuccess = False
        SetRect rct, 0, VScroll2.ScaleHeight - 17, 17, VScroll2.ScaleHeight
        dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_arrowbtn, abs_downdisabled, rct)
        If dtSuccess = False Then
            VScroll2.PaintPicture Image1.Picture, 0, VScroll2.ScaleHeight - 17, 17, 17, 0, 51, 17, 17, vbSrcCopy
            TransparentBlt2 VScroll2.hdc, Image2.Picture, 4, VScroll2.ScaleHeight - 13, 9, 9, 0, 36, RGB(255, 0, 255)
        End If
    End If
    
    SetRect rct, 1, gTop, 17, gTop + gBox
    If Enabled = True Then
    If mdHover = 5 Then
        dtSuccess = False
        dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_thumbbtnvert, scrbs_hot, rct)
        If dtSuccess = False Then
            VScroll2.PaintPicture Image4.Picture, 1, gTop, 15, gBox, 0, 3 + 22, 15, 16, vbSrcCopy
            VScroll2.PaintPicture Image4.Picture, 1, gTop, 15, 3, 0, 0 + 22, 15, 3, vbSrcCopy
            VScroll2.PaintPicture Image4.Picture, 1, gTop + gBox - 3, 15, 3, 0, 19 + 22, 15, 3, vbSrcCopy
        End If
    Else
        dtSuccess = False
        dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_thumbbtnvert, IIf(mdScroll = 5, scrbs_pressed, scrbs_normal), rct)
        If dtSuccess = False Then
            VScroll2.PaintPicture Image4.Picture, 1, gTop, 15, gBox, 0, 3 + IIf(mdScroll = 5, 44, 0), 15, 16, vbSrcCopy
            VScroll2.PaintPicture Image4.Picture, 1, gTop, 15, 3, 0, 0 + IIf(mdScroll = 5, 44, 0), 15, 3, vbSrcCopy
            VScroll2.PaintPicture Image4.Picture, 1, gTop + gBox - 3, 15, 3, 0, 19 + IIf(mdScroll = 5, 44, 0), 15, 3, vbSrcCopy
        End If
    End If
    
    dtSuccess = False
    dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_grippervert, IIf(mdScroll = 5, scrbs_pressed, scrbs_normal), rct)
    
    If dtSuccess = False Then
        TransparentBlt2 VScroll2.hdc, Image5.Picture, 4, Int(gBox / 2) - 4 + gTop, 8, 8, 0, 0, RGB(255, 0, 255)
    End If
    
    Else
        dtSuccess = False
        dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_thumbbtnvert, scrbs_disabled, rct)
        If dtSuccess = False Then
            VScroll2.PaintPicture Image4.Picture, 1, gTop, 15, gBox, 0, 3 + 66, 15, 16, vbSrcCopy
            VScroll2.PaintPicture Image4.Picture, 1, gTop, 15, 3, 0, 0 + 66, 15, 3, vbSrcCopy
            VScroll2.PaintPicture Image4.Picture, 1, gTop + gBox - 3, 15, 3, 0, 19 + 66, 15, 3, vbSrcCopy
        End If
        
        dtSuccess = DrawTheme(VScroll2.hdc, "scrollbar", sbp_grippervert, scrbs_disabled, rct)
        'TransparentBlt2 vScroll2.hdc, Image5.Picture, 4, Int(gBox / 2) - 4 + gTop, 8, 8, 0, 24, RGB(255, 0, 255)
    End If
    
End If

VScroll2.Refresh
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ScrollbarBackColor() As OLE_COLOR
    ScrollbarBackColor = m_ScrollbarBackColor
End Property

Public Property Let ScrollbarBackColor(ByVal New_ScrollbarBackColor As OLE_COLOR)
    m_ScrollbarBackColor = New_ScrollbarBackColor
    PropertyChanged "ScrollbarBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ScrollbarBARcolor() As OLE_COLOR
    ScrollbarBARcolor = m_ScrollbarBARcolor
End Property

Public Property Let ScrollbarBARcolor(ByVal New_ScrollbarBARcolor As OLE_COLOR)
    m_ScrollbarBARcolor = New_ScrollbarBARcolor
    PropertyChanged "ScrollbarBARcolor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get GridlineColor() As OLE_COLOR
    GridlineColor = m_GridlineColor
End Property

Public Property Let GridlineColor(ByVal New_GridlineColor As OLE_COLOR)
    m_GridlineColor = New_GridlineColor
    PropertyChanged "GridlineColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get XPtheme() As Boolean
    XPtheme = m_XPtheme
End Property

Public Property Let XPtheme(ByVal New_XPtheme As Boolean)
    m_XPtheme = New_XPtheme
    PropertyChanged "XPtheme"
    vScroll2_Paint
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get VerticalHeader() As Boolean
    VerticalHeader = m_VerticalHeader
End Property

Public Property Let VerticalHeader(ByVal New_VerticalHeader As Boolean)
    m_VerticalHeader = New_VerticalHeader
    PropertyChanged "VerticalHeader"
    'Call DrawList
    Call DrawHeader
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseColumnColor() As Boolean
    UseColumnColor = m_UseColumnColor
End Property

Public Property Let UseColumnColor(ByVal New_UseColumnColor As Boolean)
    m_UseColumnColor = New_UseColumnColor
    PropertyChanged "UseColumnColor"
    Call DrawList
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get IsEditable() As Boolean
    IsEditable = m_IsEditable
End Property

Public Property Let IsEditable(ByVal New_IsEditable As Boolean)
    m_IsEditable = New_IsEditable
    PropertyChanged "IsEditable"
End Property

