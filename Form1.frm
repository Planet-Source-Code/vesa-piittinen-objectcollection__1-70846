VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ObjectCollection sample"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Items As New ObjectCollection, Item As Object
    ' add some items
    Items.Add Command1, "a"
    Items.Add Command2, "b", 1
    Items.Add Command3, "c"
    ' result:
    ' 1 = "b" Command2
    ' 2 = "a" Command1
    ' 3 = "c" Command3
    
    ' now swap contents between items "a" and "b"
    Items.Swap "a", "b"
    ' result:
    ' 1 = "a" Command2
    ' 2 = "b" Command1
    ' 3 = "c" Command3
    
    ' now move "c" to be the first item
    Items.Index("c") = 1
    ' result:
    ' 1 = "c" Command3
    ' 2 = "a" Command2
    ' 3 = "b" Command1
    
    ' then loop through all items
    For Each Item In Items
        ' yeah, now you do not need to set pointers as keys yourself: you can have a key and still get a reference to the object with ObjPtr
        Item.Caption = Items.KeyByPtr(ObjPtr(Item)) & ") My index is " & Items.IndexByPtr(ObjPtr(Item))
    Next Item
    
    ' output the final result to immediate window
    Debug.Print "Keys: " & Join(Items.Keys)
End Sub
