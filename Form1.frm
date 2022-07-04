VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmIceCreamParlor 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Registry Routine Example"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Write Remote Registry"
      Height          =   915
      Left            =   105
      TabIndex        =   40
      Top             =   6900
      Width           =   3825
      Begin VB.CommandButton cmdWriteRemote 
         Caption         =   "Write"
         Height          =   240
         Left            =   135
         TabIndex        =   43
         Top             =   270
         Width           =   870
      End
      Begin VB.TextBox txtNewStartPage 
         Height          =   285
         Left            =   135
         TabIndex        =   42
         Text            =   "http:\\www.PlanetSourceCode.com"
         Top             =   555
         Width           =   3585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Internet Explorer Start Page"
         Height          =   240
         Left            =   1095
         TabIndex        =   44
         Top             =   285
         Width           =   2580
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Read Remote Registry"
      Height          =   915
      Left            =   105
      TabIndex        =   37
      ToolTipText     =   "Reads Internet Explorer Start Page"
      Top             =   5970
      Width           =   3825
      Begin VB.TextBox txtStartPage 
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   555
         Width           =   3570
      End
      Begin VB.CommandButton cmdReadRemote 
         Caption         =   "Read"
         Height          =   255
         Left            =   105
         TabIndex        =   38
         ToolTipText     =   "Click To Select Remote Computer"
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Internet Explorer Start Page"
         Height          =   240
         Left            =   975
         TabIndex        =   41
         Top             =   300
         Width           =   2580
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   1500
      Left            =   3375
      TabIndex        =   33
      ToolTipText     =   "Enumerate Registry Names and Values EXAMPLE"
      Top             =   1815
      Width           =   5625
      Begin VB.CommandButton cmdEnumValues 
         Caption         =   "Enumerate Values"
         Height          =   600
         Left            =   3990
         TabIndex        =   34
         Top             =   480
         Width           =   1080
      End
      Begin ComctlLib.ListView lstviewEnumValues 
         Height          =   1200
         Left            =   75
         TabIndex        =   35
         Top             =   180
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   2117
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Value Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Default is HKLM\SOFTWARE"
      Height          =   1845
      Left            =   75
      TabIndex        =   30
      ToolTipText     =   "Registry Enumerate Keys EXAMPLE"
      Top             =   2115
      Width           =   3150
      Begin VB.CommandButton cmdEnumKeys 
         Caption         =   "Enumerate Keys"
         Height          =   300
         Left            =   645
         TabIndex        =   32
         Top             =   1425
         Width           =   1860
      End
      Begin VB.ListBox lstEnumKeys 
         Height          =   1035
         Left            =   225
         TabIndex        =   31
         Top             =   345
         Width           =   2700
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Sundae's Name (saves on exit)"
      Height          =   1005
      Left            =   5640
      TabIndex        =   27
      ToolTipText     =   "REG_SZ / Registry Delete Value EXAMPLE"
      Top             =   795
      Width           =   3360
      Begin VB.CommandButton cmdDeleteValue 
         Caption         =   "Delete Value"
         Height          =   330
         Left            =   1920
         TabIndex        =   29
         Top             =   375
         Width           =   1110
      End
      Begin VB.TextBox txtSundaeName 
         Height          =   315
         Left            =   195
         TabIndex        =   28
         Text            =   "Type a Name Here"
         Top             =   390
         Width           =   1680
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create/Delete SubKey"
      Height          =   1935
      Left            =   4200
      TabIndex        =   21
      ToolTipText     =   "Create/Delete Key Example"
      Top             =   3975
      Width           =   2070
      Begin VB.CommandButton cmdDeleteKey 
         Caption         =   "Delete Key"
         Height          =   330
         Left            =   465
         TabIndex        =   24
         Top             =   1515
         Width           =   1110
      End
      Begin VB.TextBox txtNewKey 
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   405
         Width           =   1815
      End
      Begin VB.CommandButton cmdCreateKey 
         Caption         =   "Create Key"
         Height          =   330
         Left            =   480
         TabIndex        =   22
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   26
         Top             =   780
         Width           =   1710
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "KeyName"
         Height          =   255
         Left            =   135
         TabIndex        =   25
         Top             =   210
         Width           =   795
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REG_EXPAND_SZ Example"
      Height          =   1935
      Left            =   90
      TabIndex        =   16
      ToolTipText     =   "REG_EXPAND_SZ EXAMPLE"
      Top             =   3975
      Width           =   3825
      Begin VB.TextBox txtExpandActual 
         Height          =   315
         Left            =   75
         TabIndex        =   36
         ToolTipText     =   "Actual Text in key"
         Top             =   630
         Width           =   3660
      End
      Begin VB.CommandButton cmdSetKey 
         Caption         =   "Set Key"
         Height          =   360
         Left            =   75
         TabIndex        =   19
         Top             =   1500
         Width           =   810
      End
      Begin VB.CommandButton cmdTempDrive 
         Caption         =   "Retrieve Key"
         Height          =   480
         Left            =   75
         TabIndex        =   18
         Top             =   960
         Width           =   960
      End
      Begin VB.TextBox txtExpanded 
         Height          =   315
         Left            =   75
         TabIndex        =   17
         ToolTipText     =   "Expanded Value"
         Top             =   270
         Width           =   3660
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Press 'Set Key' to set with %SYSTEMROOT%\TEMP Then press Retrieve to see what happens"
         Height          =   825
         Left            =   1650
         TabIndex        =   20
         Top             =   1035
         Width           =   2085
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Flavors"
      Height          =   645
      Left            =   75
      TabIndex        =   12
      ToolTipText     =   "REG_MULTI_SZ /Add Value EXAMPLE"
      Top             =   45
      Width           =   8340
      Begin VB.ComboBox cmbFlavors 
         Height          =   315
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   195
         Width           =   3540
      End
      Begin VB.TextBox txtNewFlavor 
         Height          =   285
         Left            =   3645
         TabIndex        =   14
         Top             =   210
         Width           =   3330
      End
      Begin VB.CommandButton cmdAddFlavor 
         Caption         =   "Add Flavor"
         Height          =   375
         Left            =   7005
         TabIndex        =   13
         Top             =   165
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Toppings"
      Height          =   1005
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "REG_BINARY EXAMPLE"
      Top             =   795
      Width           =   3360
      Begin VB.CheckBox chkToppings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bannanas"
         Height          =   240
         Index           =   5
         Left            =   1845
         TabIndex        =   10
         Top             =   690
         Width           =   1125
      End
      Begin VB.CheckBox chkToppings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cherries"
         Height          =   240
         Index           =   4
         Left            =   1845
         TabIndex        =   9
         Top             =   435
         Width           =   1020
      End
      Begin VB.CheckBox chkToppings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Caramel"
         Height          =   240
         Index           =   3
         Left            =   1845
         TabIndex        =   8
         Top             =   195
         Width           =   945
      End
      Begin VB.CheckBox chkToppings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hot Fudge"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   705
         Width           =   2505
      End
      Begin VB.CheckBox chkToppings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nuts"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   465
         Width           =   2505
      End
      Begin VB.CheckBox chkToppings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sprinkles"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   2505
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ice Cream Type"
      Height          =   1260
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "REG_DWORD EXAMPLE"
      Top             =   795
      Width           =   1800
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Custard"
         Height          =   255
         Index           =   3
         Left            =   105
         TabIndex        =   11
         Top             =   930
         Width           =   1590
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frogurt"
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   3
         Top             =   675
         Width           =   1590
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Soft"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   450
         Width           =   1590
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hard"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmIceCreamParlor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reg As RegistryRoutines
Attribute Reg.VB_VarHelpID = -1
Dim Toppings As Variant
Dim BlankArray(5) As Byte
Dim Flavors As Variant
Dim MainKeyRoot As String
Dim MainSubKey As String





Private Sub Form_Load()
 Dim i As Integer
 Set Reg = New RegistryRoutines
 
    MainKeyRoot = "Software\Dons Ice Cream Parlor\Sundae Maker"
    MainSubKey = "Settings"
 
    'Default Values for the REG_MULTI_SZ Key
    Flavors = "Vanilla" & vbNullChar & "Chocolate" & vbNullChar & _
              "Strawberry" & vbNullChar & "Mint Chocolate Chip" & _
              vbNullChar & "Cookie Dough" & vbNullChar & "Cinnamon" _
              & vbNullChar & "Pistachio" & vbNullChar & "Chocolate Fudge" _
              & vbNullChar & "Orange" & vbNullChar & "Bubble Gum"
   
    Reg.hkey = HKEY_LOCAL_MACHINE
    Reg.KeyRoot = MainKeyRoot
    Reg.Subkey = MainSubKey
    If Not Reg.KeyExists Then Reg.CreateKey 'Uses "Settings" as the key to create
    
    'Get Form's saved Settings from Registry
    Me.Top = Reg.GetRegistryValue("Top", Me.Top)
    Me.Left = Reg.GetRegistryValue("Left", Me.Left)
    Me.Width = Reg.GetRegistryValue("Width", Me.Width)
    Me.Height = Reg.GetRegistryValue("Height", Me.Height)
    optType(Reg.GetRegistryValue("Type", 0)).Value = True
    txtSundaeName.Text = Reg.GetRegistryValue("Name", " ")
    Toppings = Reg.GetRegistryValue("Toppings", BlankArray())
    Flavors = Reg.GetRegistryValue("Flavors", Flavors)
    Flavors = Split(Flavors, vbNullChar)
    
    For i = 0 To UBound(Flavors) 'put data in the text box
      cmbFlavors.AddItem Flavors(i)
    Next i
      cmbFlavors.ListIndex = 0
            
    For i = 0 To UBound(Toppings) 'put data in the text box
      ReDim Preserve Toppings(UBound(Toppings))
      chkToppings(i).Value = Toppings(i)
    Next i
        
End Sub

Private Sub chkToppings_Click(index As Integer)
    Toppings(index) = chkToppings(index).Value
End Sub

Private Sub cmdAddFlavor_Click()
    cmbFlavors.AddItem txtNewFlavor.Text
End Sub

Private Sub cmdCreateKey_Click()
    'Sub called will ignore current reg.subkey and create a new using text
    If Not Reg.CreateKey(txtNewKey.Text) Then lblStatus.Caption = "Key Created"
End Sub

Private Sub cmdDeleteKey_Click()
    'Uses current reg.hkey and reg.keyroot
     If Not Reg.DeleteKey(txtNewKey.Text) Then lblStatus.Caption = "Key Deleted"
End Sub

Private Sub cmdDeleteValue_Click()
    'Uses current reg.hkey and reg.keyroot and reg.subkey
    If Not Reg.DeleteValue("Name") Then txtSundaeName.Text = ""
End Sub

Private Sub cmdSetKey_Click()
    Reg.SetRegistryValue "TEMP", "%SYSTEMDRIVE%\TEMP", REG_EXPAND_SZ
End Sub

Private Sub cmdTempDrive_Click()
    txtExpanded.Text = Reg.GetRegistryValue("TEMP", "?")
End Sub

Private Sub cmdReadRemote_Click()
    txtStartPage.Text = Reg.ReadRemoteRegistryValue(GetBrowseNetworkWorkstation(Me.hWnd), HKEY_CURRENT_USER, "Start Page", "SOFTWARE\Microsoft\Internet Explorer\Main")
End Sub

Private Sub cmdWriteRemote_Click()
Dim ret As Boolean
ret = Reg.WriteRemoteRegistryValue(GetBrowseNetworkWorkstation(Me.hWnd), HKEY_CURRENT_USER, "Start Page", txtNewStartPage.Text, REG_SZ, "Software\Microsoft\Internet Explorer\Main")
If ret = False Then MsgBox "Failed to Write Key", vbInformation & vbOKOnly, "Registry Error"
End Sub
Private Sub cmdEnumKeys_Click()
    Dim KeyCollection As Collection
    Dim Object As Variant
        Set KeyCollection = Reg.EnumRegistryKeys(HKEY_LOCAL_MACHINE, "Software")
      
        For Each Object In KeyCollection
            lstEnumKeys.AddItem Object
        Next
        Set KeyCollection = Nothing
End Sub

Private Sub cmdEnumValues_Click()
    Dim KeyCollection As Collection
    Dim Object As Variant
    Dim itmx As ListItem
    
    
        Set KeyCollection = Reg.EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\Dons Ice Cream Parlor\Sundae Maker\Settings")
    
        For Each Object In KeyCollection
        Set itmx = lstviewEnumValues.ListItems.Add(, , Object(0))
            itmx.SubItems(1) = Object(1)
            Set itmx = Nothing
        Next
        Set KeyCollection = Nothing
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim Flavorlist As String
    'Save Form Settings to Registry
    Reg.hkey = HKEY_LOCAL_MACHINE
    Reg.KeyRoot = MainKeyRoot
    Reg.Subkey = MainSubKey
    Reg.SetRegistryValue "Top", Me.Top, REG_DWORD
    Reg.SetRegistryValue "Left", Me.Left, REG_DWORD
    Reg.SetRegistryValue "Height", Me.Height, REG_DWORD
    Reg.SetRegistryValue "Width", Me.Width, REG_DWORD
    Select Case True
        Case optType(0).Value
            Reg.SetRegistryValue "Type", 0, REG_DWORD
        Case optType(1).Value
            Reg.SetRegistryValue "Type", 1, REG_DWORD
        Case optType(2).Value
            Reg.SetRegistryValue "Type", 2, REG_DWORD
        Case optType(3).Value
            Reg.SetRegistryValue "Type", 3, REG_DWORD
    End Select
    Reg.SetRegistryValue "Toppings", Toppings, REG_BINARY
    'Create null delimited string of flavors to save into Registry
    For i = 0 To cmbFlavors.ListCount
        Flavorlist = Flavorlist & cmbFlavors.List(i) & vbNullChar
    Next i
    Reg.SetRegistryValue "Flavors", Flavorlist, REG_MULTI_SZ
    Reg.SetRegistryValue "Name", txtSundaeName, REG_SZ
    Set Reg = Nothing
End Sub


