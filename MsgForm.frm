VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MsgForm 
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "MsgForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MsgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub ShowMessage(ByVal txt As String)
    Label1.Caption = txt
    Me.Repaint
End Sub

