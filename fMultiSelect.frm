VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMultiSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Multi-Node Selection Prototype"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDialog 
      Caption         =   "To&ggle Select"
      Height          =   330
      Index           =   6
      Left            =   3255
      TabIndex        =   12
      Top             =   1680
      Width           =   1170
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "T&oggle"
      Height          =   330
      Index           =   5
      Left            =   3255
      TabIndex        =   11
      Top             =   1260
      Width           =   1170
   End
   Begin VB.ListBox lstDialog 
      Height          =   3570
      Left            =   4620
      TabIndex        =   7
      Top             =   420
      Width           =   2325
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Transfer &Mode:"
      Height          =   960
      Left            =   3255
      TabIndex        =   8
      ToolTipText     =   "Transfer in displayed Node sequence or order of selection..."
      Top             =   4095
      Width           =   3690
      Begin VB.OptionButton optDialog 
         Caption         =   "Enumerate Tag Collection"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   630
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optDialog 
         Caption         =   "Enumerate Treeview Nodes"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   315
         Width           =   2430
      End
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Transfer ->>"
      Height          =   330
      Index           =   4
      Left            =   3255
      TabIndex        =   5
      Top             =   3570
      Width           =   1170
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Select &All"
      Height          =   330
      Index           =   1
      Left            =   3255
      TabIndex        =   2
      Top             =   525
      Width           =   1170
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "Clear A&ll"
      Height          =   330
      Index           =   3
      Left            =   3255
      TabIndex        =   4
      Top             =   2835
      Width           =   1170
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Clear"
      Height          =   330
      Index           =   2
      Left            =   3255
      TabIndex        =   3
      Top             =   2415
      Width           =   1170
   End
   Begin VB.CommandButton cmdDialog 
      Caption         =   "&Select"
      Height          =   330
      Index           =   0
      Left            =   3255
      TabIndex        =   1
      Top             =   105
      Width           =   1170
   End
   Begin MSComctlLib.TreeView tvwDialog 
      Height          =   4950
      Left            =   0
      TabIndex        =   0
      Top             =   105
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   8731
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblDialog 
      Caption         =   "Selected Nodes:"
      Height          =   225
      Left            =   4620
      TabIndex        =   6
      Top             =   105
      Width           =   2325
   End
End
Attribute VB_Name = "fMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fMultiSelect
' Author:       Graeme Grant
' Date:         24/01/2002
' Version:      01.00.00
' Description:  Prototype for TreeView Mult-Node selection
' Edit History: 01.00.00 24/01/2002 Initial Release
'
'===========================================================================

Option Explicit

Private Const cSELFORECOLOR As Long = &HFFFF&
Private Const cSELBACKCOLOR As Long = &HC0&

Private Enum eCommand
    [Select] = 0
    [Select All] = 1
    [Clear] = 2
    [Clear All] = 3
    [Transfer] = 4
    [Toggle] = 5
    [Toggle Selection] = 6
End Enum

'
'## TreeView nodes can have different fore & back colors plus a Bold state.
'   For Muli-Node selection to work, we need to capture and store this
'   information for each node. We can't use a collection of nodes due to
'   only pointers to objects are stored and not seperate new objects.
'   Therefore a specialised collection is required.
'
'   I haven't used a type'd array due to the overhead of management. Therefore
'   a collection class of variants has been used. I've chosen variants over
'   explicit properties (i.e. NodeKey, ForeColor, BackColor, Bold) to make
'   the class more generic for future projects - any type of
'   variable/object and any number of elements per collection item can stored.
'
'## Tag Element IDs used
Private Enum eTags                      'Private Type tTags
    [Node Key] = 0                      '    lNodeKey   as Long
    [ForeColor] = 1                     '    lForeColor as OLE_COLOR
    [BackColor] = 2                     '    lBackColor as OLE_COLOR
    [Bold] = 3                          '    bBold      as Boolean
End Enum                                'End Type

Private moTags    As cTags              'Private moTags() as tTags
Private moSelNode As MSComctlLib.Node

'===========================================================================
' Form Control Events
'
Private Sub cmdDialog_Click(Index As Integer)

    Dim oNode    As MSComctlLib.Node
    Dim lPtr     As Long
    Dim lSelItem As Variant
    
    On Error GoTo ErrorHandler
    lPtr = ObjPtr(moSelNode)
    Select Case Index
        Case [Select]
            With moSelNode
                moTags.Add lPtr, .Key, .ForeColor, .BackColor, .Bold
                .ForeColor = cSELFORECOLOR
                .BackColor = cSELBACKCOLOR
                .Bold = True
            End With

        Case [Select All]
            tvwDialog.Visible = False
            For Each oNode In tvwDialog.Nodes
                With oNode
                    moTags.Add ObjPtr(oNode), .Key, .ForeColor, .BackColor, .Bold
                    .ForeColor = cSELFORECOLOR
                    .BackColor = cSELBACKCOLOR
                    .Bold = True
                End With
            Next

        Case [Toggle]
            With moSelNode
                If moTags.Exist(lPtr) Then
                    .ForeColor = moTags(lPtr, [ForeColor])
                    .BackColor = moTags(lPtr, [BackColor])
                    .Bold = moTags(lPtr, [Bold])
                    moTags.Remove lPtr
                Else
                    moTags.Add lPtr, .Key, .ForeColor, .BackColor, .Bold
                    .ForeColor = cSELFORECOLOR
                    .BackColor = cSELBACKCOLOR
                    .Bold = True
                End If
        End With

        Case [Toggle Selection]
            tvwDialog.Visible = False
            For Each oNode In tvwDialog.Nodes
                lPtr = ObjPtr(oNode)
                With oNode
                    If moTags.Exist(lPtr) Then
                        .ForeColor = moTags(lPtr, [ForeColor])
                        .BackColor = moTags(lPtr, [BackColor])
                        .Bold = moTags(lPtr, [Bold])
                        moTags.Remove lPtr
                    Else
                        moTags.Add lPtr, .Key, .ForeColor, .BackColor, .Bold
                        .ForeColor = cSELFORECOLOR
                        .BackColor = cSELBACKCOLOR
                        .Bold = True
                    End If
                End With
            Next

        Case [Clear]
            With moSelNode
                If moTags.Exist(lPtr) Then
                    .ForeColor = moTags(lPtr, [ForeColor])
                    .BackColor = moTags(lPtr, [BackColor])
                    .Bold = moTags(lPtr, [Bold])
                    moTags.Remove lPtr
                End If
            End With

        Case [Clear All]
            tvwDialog.Visible = False
            With moTags
                For Each oNode In tvwDialog.Nodes
                    lPtr = ObjPtr(oNode)
                    If .Exist(lPtr) Then
                        oNode.ForeColor = .Element(lPtr, [ForeColor])
                        oNode.BackColor = .Element(lPtr, [BackColor])
                        oNode.Bold = .Element(lPtr, [Bold])
                        .Remove lPtr
                    End If
                Next
            End With
            lstDialog.Clear

        Case [Transfer]
            If moTags.Count Then
                With lstDialog
                    .Clear
                    lblDialog.FontBold = True
                    If optDialog(1).Value Then
                        For Each lSelItem In moTags
                            .AddItem tvwDialog.Nodes(lSelItem([Node Key])).Text
                        Next
                    Else
                        For Each oNode In tvwDialog.Nodes
                            lPtr = ObjPtr(oNode)
                            If moTags.Exist(lPtr) Then
                                .AddItem tvwDialog.Nodes(moTags(lPtr, [Node Key])).Text
                            End If
                        Next
                    End If
                End With
            Else
                lblDialog.FontBold = False
                MsgBox "No items selected in TreeView.", vbCritical + vbExclamation, "WARNING!"
            End If

    End Select
    tvwDialog.SetFocus

ErrorHandler:
    tvwDialog.Visible = True

End Sub

Private Sub tvwDialog_NodeClick(ByVal Node As MSComctlLib.Node)
    Set moSelNode = Node
    Node.Selected = False
End Sub

'===========================================================================
' Form Events
'
Private Sub Form_Load()

    With tvwDialog
        .Style = tvwTreelinesPlusMinusPictureText
        .LineStyle = tvwRootLines
        .Indentation = 10
        .FullRowSelect = False
        .HideSelection = False
        .HotTracking = True
    End With
    pInitData
    With tvwDialog.Nodes(1)
        .Selected = True
        .Selected = False
    End With
End Sub

'===========================================================================
' Internal Functions
'
Private Sub pInitData()

    Set moTags = New cTags

    With tvwDialog.Nodes
        .Add , , "A", "Palmyra Atoll"
        .Add , , "B", "Kazakhstan"
        .Add , , "C", "Sao Tome And Principe"
        .Add , , "D", "Faeroe Islands"
        .Add , , "E", "Luxembourg"

        Dim lLoop As Long
        .Add , , "X1", "Node Item 1"
        For lLoop = 2 To 50
            .Add tvwDialog.Nodes("X" + CStr(lLoop - 1)), tvwChild, "X" + CStr(lLoop), "Node Item " + CStr(lLoop)
        Next
    End With
    
    Randomize (Timer)
    Dim oNode As MSComctlLib.Node
    For Each oNode In tvwDialog.Nodes
        pColorize oNode
        oNode.Expanded = True
    Next

End Sub

Private Sub pColorize(Node As MSComctlLib.Node)
    With Node
        Select Case CInt(Rnd(1) * 4)
            Case 0
                .ForeColor = &H80&
            Case 1
                .ForeColor = &H80FF&
            Case 2
                .ForeColor = &H8000&
            Case 3
                .ForeColor = &H8000&
                .Bold = True
            Case 4
                .ForeColor = vbWindowText
        End Select
    End With
End Sub
