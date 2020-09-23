VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simplest Address Database"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   4080
      TabIndex        =   9
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox cboName 
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtIcNo 
         Height          =   285
         Left            =   870
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   870
         TabIndex        =   2
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtTel 
         Height          =   285
         Left            =   870
         TabIndex        =   1
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ICQ :"
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Phone :"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   1380
         Width           =   555
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dont forget to reference 'Microsoft DAO 3.6 Object Library'
'using the reference menu in this Project

'I've tried to keep things as simple as possible
'to make it easy to understand. Yes I know there's
'probably better ways of achieving similar results
'- this is an introduction to databases

'Variables used to control Database
Dim mDB As Database
Dim mTb As TableDef
Dim mFld As Field
Dim mRS As Recordset
Private Sub cboName_Click()
'Open up a database
Set mDB = OpenDatabase(App.Path + "\contacts.mdb")
With mDB
'Read the database - the group of records we're
'interested in is "Contacts"
    Set mRS = .OpenRecordset("Contacts")
    With mRS
    'If there are no records lets bail out now
        If .RecordCount <> 0 Then
            'Otherwise read the first one
            'Got to start somewhere !
            .MoveFirst
            'Is this what we're looking for
            If !Name = cboName.Text Then
                'Yes - fill our text boxes
                txtIcNo.Text = !ICQ
                txtAddress.Text = !Address
                txtTel.Text = !Phone
            Else
                'No - keep looking
                For dby = 1 To .RecordCount - 1
                    .MoveNext
                    If .EOF Then Exit For
                    If !Name = cboName.Text Then
                        'Found it - fill our text boxes
                        txtIcNo.Text = !ICQ
                        txtAddress.Text = !Address
                        txtTel.Text = !Phone
                        Exit For
                    End If
                Next dby
            End If
        End If
        'Finished with that group of records
    End With
    'In fact we're finished with the whole damn lot
    .Close
End With
'This little arrangement deals with the fact that
'databases dont like having empty fields.
'So when we saved this data, if the textbox
'was empty we saved it as "-" instead.
'Now to show the user we return it to "".
If txtIcNo.Text = "-" Then txtIcNo.Text = ""
If txtAddress.Text = "-" Then txtAddress.Text = ""
If txtTel = "-" Then txtTel = ""

End Sub




Private Sub cmdDelete_Click()
'Open up a database
Set mDB = OpenDatabase(App.Path + "\contacts.mdb")
With mDB
'Open the group of records we're interested in
    Set mRS = .OpenRecordset("Contacts")
    With mRS
        If .RecordCount <> 0 Then
            .MoveFirst
            If !Name = cboName.Text Then
            'Here it is - KILL
                .Delete
            Else
                For dby = 1 To .RecordCount - 1
                    .MoveNext
                    If .EOF Then Exit For
                    If !Name = cboName.Text Then
                        'Here it is - KILL
                        .Delete
                        Exit For
                    End If
                Next dby
            End If
        End If
    End With
    .Close
End With
LoadDB
End Sub

Private Sub cmdNew_Click()
Dim temp As String
'Using an input box with a combobox set to style 2
'stops people messing with the name property
temp = InputBox("Enter a new name.")
If temp = "" Then Exit Sub
cboName.AddItem temp
'Make sure we're looking at the new name
cboName.ListIndex = cboName.ListCount - 1
txtIcNo.Text = ""
txtAddress.Text = ""
txtTel = ""
'Open up a database
Set mDB = OpenDatabase(App.Path + "\contacts.mdb")
With mDB
    'Open the group of records we're interested in
    Set mRS = .OpenRecordset("Contacts")
    With mRS
        'Add the new data
        .AddNew
        !Name = cboName.Text
        !ICQ = "-"
        !Address = "-"
        !Phone = "-"
        .Update
    End With
    .Close
End With
End Sub
Public Sub BuildDB()
'If we already have a database dont bother
If FileExists(App.Path + "\contacts.mdb") Then Exit Sub
'Otherwise better get our Microsoft DLL to build
'a new one to our specifications
Set mDB = DBEngine.Workspaces(0).CreateDatabase(App.Path + "\contacts.mdb", dbLangGeneral)
Set mTb = mDB.CreateTableDef("Contacts")
Set mFld = mTb.CreateField("Name", dbText, 100)
mTb.Fields.Append mFld
Set mFld = mTb.CreateField("ICQ", dbText, 100)
mTb.Fields.Append mFld
Set mFld = mTb.CreateField("Address", dbText, 100)
mTb.Fields.Append mFld
Set mFld = mTb.CreateField("Phone", dbText, 100)
mTb.Fields.Append mFld
mDB.TableDefs.Append mTb
End Sub

Function FileExists(ByVal Filename As String) As Integer
'Just used to find out if the database is there
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(Filename)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
                End
            End If
    End Select
End Function

Private Sub cmdSave_Click()
'If any of the textboxes are empty then
'temporarily fill them so we dont have any
'empty fields to upset the database
If txtIcNo.Text = "" Then txtIcNo.Text = "-"
If txtAddress.Text = "" Then txtAddress.Text = "-"
If txtTel = "" Then txtTel = "-"
'Open the database
Set mDB = OpenDatabase(App.Path + "\contacts.mdb")
With mDB
    'Open our group of records
    Set mRS = .OpenRecordset("Contacts")
    With mRS
        If .RecordCount <> 0 Then
            .MoveFirst
            If !Name = cboName.Text Then
                'if this is the right one then
                'add the new data
                .Edit
                !ICQ = txtIcNo.Text
                !Address = txtAddress
                !Phone = txtTel
                .Update
            Else
                For dby = 1 To .RecordCount - 1
                    .MoveNext
                    If .EOF Then Exit For
                    If !Name = cboName.Text Then
                    'if this is the right one then
                    'add the new data
                        .Edit
                        !ICQ = txtIcNo.Text
                        !Address = txtAddress
                        !Phone = txtTel
                        .Update
                        Exit For
                    End If
                Next dby
            End If
        End If
    End With
    .Close
End With
'Return any textboxes that we altered back to empty
If txtIcNo.Text = "-" Then txtIcNo.Text = ""
If txtAddress.Text = "-" Then txtAddress.Text = ""
If txtTel = "-" Then txtTel = ""

End Sub

Private Sub Form_Load()
Dim dby As Integer
BuildDB
LoadDB
End Sub

Public Sub LoadDB()
cboName.Clear
'Open the database
Set mDB = OpenDatabase(App.Path + "\contacts.mdb")
With mDB
    'Open our group of records
    Set mRS = .OpenRecordset("Contacts")
    With mRS
        If .RecordCount <> 0 Then
        'If there's any records fill the combobox
            .MoveFirst
            cboName.AddItem !Name
            For dby = 1 To .RecordCount - 1
                .MoveNext
                If .EOF Then Exit For
                cboName.AddItem !Name
            Next dby
        End If
    End With
    .Close
End With
'We just needed to fill the combobox in this sub
'because now - if there's anything in the combo
'we're going to make the comboindex change
'which will fire the combo click event
'which fills our textboxes
If cboName.ListCount > 0 Then cboName.ListIndex = 0

End Sub
