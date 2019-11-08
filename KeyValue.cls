VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Key As Variant
Public Value As Variant

Sub Init(Key As Variant, Value As Variant)
  Me.Key = Key
  If IsObject(Value) Then Set Me.Value = Value Else Me.Value = Value
End Sub

