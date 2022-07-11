VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Examples:"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete File"
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "&Safe Data Input"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdGetValue 
      Caption         =   "&Get Value of a Key"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemoveKey 
      Caption         =   "&Remove Key"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddReg 
      Caption         =   "&Add to Registry"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Directory and FileName"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4680
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      Caption         =   "Key Value:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Key Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'© 2001 Blazon Resources and Creative Engineering
'www.odin.prohosting.com/blazon
'
'Limited ability to save to and retrieve data from the registry
'Ability to check to see if files exist and delete files from the computer
'
'Note:  All Keys saved to the registry are located in:
'HKEY_CURRENT_USER\Software\VB and VBA Program Settings|?UserDefined?|?UserDefined?|KeyName
'
'Explanation for lines used extensively
'
'   On Error GoTo ErrHandle     'If an error occurs the jump to "ErrHandle:"
'   ErrHandle:                  'Program skips to here if an error occurs

Private Sub cmdAddReg_Click()
'   Add the value of the key to the registry of the current computer
    Dim Main As String, SubCategory As String, KeyName As String, Value As Variant
    
    On Error GoTo ErrHandle
    
    Main = "Demonstration"          'The Main Directory
    SubCategory = "Visual Basic"    'The Sub-Folder
    KeyName = txtKey.Text           'The Key for the Variable(storage location name)
    Value = txtValue.Text           'Value to save
    
    'Call to add the values to the registry
    Call SetKey(Main, SubCategory, KeyName, Value)
    
Exit Sub
ErrHandle:
    MsgBox "Could not add to registry!"
End Sub

Private Sub SetKey(Main As String, SubCategory As String, KeyName As String, Value As Variant)
'   Actually save the value to the registry

'   Save to:    Main directory, Sub Folder, Key(Variable storage name), Value to save
    Call SaveSetting(Main, SubCategory, KeyName, Value)
End Sub

Private Sub cmdDelete_Click()
'   Delete the file from the computer, use this power for good only!
    Dim FileName As String
    
    FileName = txtFilename.Text 'File name and pathname
    
    Call DeleteFile(FileName)
End Sub
Private Sub DeleteFile(FileName As String)
'   Delete the file from the computer

    On Error GoTo ErrHandle

    Kill FileName   'Remove the file from the computer
    
Exit Sub
ErrHandle:
    MsgBox "Could not delete the file " & Chr(34) & FileName & Chr(34) & "."
End Sub
Private Sub cmdGetValue_Click()
'   Get the Value of Key from the registry
    Dim Main As String, SubCategory As String, KeyName As String, Value As Variant
    
    On Error GoTo ErrHandle
    
    Main = "Demonstration"          'The Main Directory
    SubCategory = "Visual Basic"    'The Sub-Folder
    KeyName = txtKey.Text           'The Key for the Variable(storage location name)
    
    txtValue.Text = GetKey(Main, SubCategory, KeyName)
    
Exit Sub
ErrHandle:
    MsgBox "Could not get from registry!"
End Sub
Private Function GetKey(Main As String, SubCategory As String, KeyName As String) As Variant
'   Get the Value of a key from the Registry

'   get the value   main directory, sub category, name of key, default value if key doesn't exist
    GetKey = GetSetting(Main, SubCategory, KeyName, "None")
End Function

Private Sub cmdInput_Click()
'   Safely input data from a file by checking to see if it exists
    Dim FileName As String
    
    FileName = txtFilename.Text 'The file's name & directory
    
    Call CheckExists(FileName)
End Sub
Private Sub CheckExists(FileName As String)
'   Check to see if the file exists
    On Error GoTo ErrHandle     'If the file does not exist
    
    Open FileName For Input As #1   'Open the file, if it does not exist, the program skips to 'ErrHandle:'
    Close #1    'The file needs to be closed
    
Exit Sub    'Do not continue and over-write data
ErrHandle:  'Program only comes here if the file did not exist
    If FileName = "" Then Exit Sub  'If there is no file then exit
    Open FileName For Output As #1
        'Save all data the file should have had here
    Close #1    'The file needs to be closed
End Sub
Private Sub cmdRemoveKey_Click()
'   Remove the key from the registry
    Dim Main As String, SubCategory As String, KeyName As String, Value As Variant
    
    On Error GoTo ErrHandle
    
    Main = "Demonstration"          'The Main Directory
    SubCategory = "Visual Basic"    'The Sub-Folder
    KeyName = txtKey.Text           'The Key to delete
    
    'Delete the key
    Call DeleteKey(Main, SubCategory, KeyName)
    
Exit Sub
ErrHandle:
    MsgBox "Could not remove from registry!"
End Sub
Private Sub DeleteKey(Main As String, SubCategory As String, KeyName As String)
'   Delete a key from the registry

    'Delete      Main Folder, SubCategory, Name of the Key
    Call DeleteSetting(Main, SubCategory, KeyName)
    
'*************************************************************
'NOTE:
'The statement:
'Call DeleteSetting(Main, SubCategory)
'or
'Call Deletesetting(main)
'
'Will Delete the directory itself!  A very useful technique indeed!
'*************************************************************

End Sub
