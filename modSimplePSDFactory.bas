Attribute VB_Name = "modSimplePSDFactory"
Option Explicit

'==============================================================================
' Factory helper for the clean-room cSimplePSD class.
'
' RC6's New_c.SimplePSD(...) factory is implemented inside RC6 itself, so a
' project-local replacement cannot intercept that exact constructor call.
'
' This helper keeps the calling syntax nearly identical:
'
'   Dim PSD As cSimplePSD
'   Set PSD = SimplePSD(PSDFile)
'
' Everything after construction uses the same cSimplePSD public interface.
'==============================================================================

Public Function SimplePSD(Optional ByRef FileNameOrByteArray As Variant) As cSimplePSD
   Dim PSD As cSimplePSD

   Set PSD = New cSimplePSD

   If Not IsMissing(FileNameOrByteArray) Then
      PSD.Load FileNameOrByteArray
   End If

   Set SimplePSD = PSD
End Function
