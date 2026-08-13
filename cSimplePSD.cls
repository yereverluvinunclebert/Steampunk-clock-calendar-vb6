VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSimplePSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' cSimplePSD - clean-room, source-compatible reimplementation of RC6.cSimplePSD
'
' Intent:
'   A deliberately small PSD reader for the same sort of use as RC6.cSimplePSD:
'   extracting ordinary raster layers into cCairoSurface objects and rendering
'   the visible layer stack with Cairo.
'
' Supported subset:
'   - PSD version 1 (not PSB)
'   - RGB, 8 bits/channel
'   - Per-layer R/G/B, transparency and user-mask channels
'   - RAW, PackBits/RLE, ZIP and ZIP-with-prediction channel compression
'   - Pascal and Unicode ("luni") layer names
'   - Photoshop folder/group markers ("lsct"/"lsdk")
'   - Common Photoshop blend modes that have direct Cairo equivalents
'
' Deliberately not interpreted:
'   - Text/vector/smart-object rendering
'   - Adjustment/fill layer semantics
'   - Layer effects/styles
'   - Vector masks
'   - Clipping groups
'   - Group isolation/knockout and exact Photoshop group-opacity semantics
'   - CMYK/Lab/Indexed/16-bit/32-bit/PSB
'
' RC6 is used for:
'   - file I/O       New_c.FSO.ReadByteContent
'   - zlib           New_c.Crypt.ZLibDecompress
'   - Cairo surfaces Cairo.CreateSurface / cCairoSurface.BindToArray
'   - alpha          Cairo.PreMultiplyAlpha
'   - compositing    cCairoContext
'
' No Win32 API declarations are used.
'
' Performance notes (V4):
'   - large byte-array copies use RC6 New_c.MemCopy where practical
'   - zero-based Variant Byte() input is assigned by the VB runtime directly
'   - RLE decode loops use linear source/destination positions
'   - common alpha/no-mask surface packing uses a branch-free linear loop
'==============================================================================


Private Const PSD_SIGNATURE As String = "8BPS"
Private Const PSD_VERSION As Long = 1
Private Const PSD_COLOR_MODE_RGB As Long = 3
Private Const PSD_DEPTH_8 As Long = 8

Private Const CH_TRANSPARENCY As Long = -1
Private Const CH_USER_MASK As Long = -2
Private Const CH_REAL_USER_MASK As Long = -3
Private Const CH_RED As Long = 0
Private Const CH_GREEN As Long = 1
Private Const CH_BLUE As Long = 2

Private Const COMP_RAW As Long = 0
Private Const COMP_RLE As Long = 1
Private Const COMP_ZIP As Long = 2
Private Const COMP_ZIP_PREDICTION As Long = 3

Private Const SECTION_OTHER As Long = 0
Private Const SECTION_OPEN_FOLDER As Long = 1
Private Const SECTION_CLOSED_FOLDER As Long = 2
Private Const SECTION_BOUNDING_DIVIDER As Long = 3

'These values intentionally match RC6.eLayerFlags.
Private Const LF_PROTECTED As Long = 1
Private Const LF_VISIBLE As Long = 2
Private Const LF_OBSOLETE As Long = 4
Private Const LF_USEFUL As Long = 8
Private Const LF_IRRELEVANT As Long = 16

Private Const ERR_BASE As Long = vbObjectError + 28760

Private Type TPSDChannel
   ID As Integer
   DataLength As Long
   DataOffset As Long
End Type

Private Type TPSDLayer
   Top As Long
   Left As Long
   Bottom As Long
   Right As Long

   ChannelFirst As Long
   ChannelCount As Long

   BlendMode As String
   BlendOperator As Long
   Opacity As Byte
   Clipping As Byte
   RawFlags As Byte
   PublicFlags As Long

   Name As String
   Path As String
   SectionType As Long
   EffectiveVisible As Boolean

   HasMaskInfo As Boolean
   MaskTop As Long
   MaskLeft As Long
   MaskBottom As Long
   MaskRight As Long
   MaskDefault As Byte
   MaskFlags As Byte
End Type

Private m_Data() As Byte
Private m_DataSize As Long

Private m_Layers() As TPSDLayer
Private m_Channels() As TPSDChannel
Private m_LayerSurfaces() As cCairoSurface

Private m_LayerCount As Long
Private m_ChannelCount As Long
Private m_Width As Long
Private m_Height As Long

Private m_AllLayersSurface As cCairoSurface

'------------------------------------------------------------------------------
' Public RC6-compatible interface
'------------------------------------------------------------------------------

Public Property Get AllLayersSurface() As cCairoSurface
   If m_AllLayersSurface Is Nothing Then
      Set m_AllLayersSurface = RenderAllLayers()
   End If
   Set AllLayersSurface = m_AllLayersSurface
End Property

Public Property Get Height() As Long
   Height = m_Height
End Property

Public Property Get LayerAlphaPercent(ByVal Idx As Long) As Double
   CheckLayerIndex Idx

   'Despite the property name, RC6.cSimplePSD exposes this as a normalized
   '0..1 value (e.g. PSD opacity 117 -> 0.458823529411765), not 0..100.
   LayerAlphaPercent = CDbl(m_Layers(Idx).Opacity) / 255#
End Property

Public Property Get LayerBlendMode(ByVal Idx As Long) As String
   CheckLayerIndex Idx
   LayerBlendMode = m_Layers(Idx).BlendMode
End Property

Public Property Get LayerBlendOperator(ByVal Idx As Long) As cairo_operator_enm
   CheckLayerIndex Idx
   LayerBlendOperator = m_Layers(Idx).BlendOperator
End Property

Public Property Get LayerByteSize(ByVal Idx As Long) As Long
   CheckLayerIndex Idx

   If m_LayerSurfaces(Idx) Is Nothing Then Exit Property

   If m_LayerSurfaces(Idx).Height > 0 Then
      If CDbl(m_LayerSurfaces(Idx).Stride) * CDbl(m_LayerSurfaces(Idx).Height) > 2147483647# Then
         RaisePSD "Layer surface is too large for LayerByteSize."
      End If
      LayerByteSize = m_LayerSurfaces(Idx).Stride * m_LayerSurfaces(Idx).Height
   End If
End Property

Public Property Get LayerChannelCount(ByVal Idx As Long) As Long
   CheckLayerIndex Idx
   LayerChannelCount = m_Layers(Idx).ChannelCount
End Property

Public Property Get LayerClipping(ByVal Idx As Long) As Byte
   CheckLayerIndex Idx
   LayerClipping = m_Layers(Idx).Clipping
End Property

Public Property Get LayerDataPtr(ByVal Idx As Long) As Long
   CheckLayerIndex Idx
   If Not m_LayerSurfaces(Idx) Is Nothing Then
      LayerDataPtr = m_LayerSurfaces(Idx).DataPtr
   End If
End Property

Public Property Get LayerFlags(ByVal Idx As Long) As eLayerFlags
   CheckLayerIndex Idx

   'RC6 returns the PSD layer-record flags byte verbatim here, even though
   'the typelib declares eLayerFlags.  Differential testing against RC6
   'shows ordinary raster layers returning &H28 (40) and folder markers
   'returning &H38 (56) for the Panzer sample.
   LayerFlags = m_Layers(Idx).RawFlags
End Property

Public Property Get LayerHeight(ByVal Idx As Long) As Long
   CheckLayerIndex Idx
   LayerHeight = m_Layers(Idx).Bottom - m_Layers(Idx).Top
End Property

Public Property Get LayerName(ByVal Idx As Long) As String
   CheckLayerIndex Idx
   LayerName = m_Layers(Idx).Name
End Property

Public Property Get LayerPath(ByVal Idx As Long) As String
   CheckLayerIndex Idx
   LayerPath = m_Layers(Idx).Path
End Property

Public Property Get LayersCount() As Long
   LayersCount = m_LayerCount
End Property

Public Property Get LayerSurface(ByVal Idx As Long) As cCairoSurface
   CheckLayerIndex Idx
   Set LayerSurface = m_LayerSurfaces(Idx)
End Property

Public Property Get LayerWidth(ByVal Idx As Long) As Long
   CheckLayerIndex Idx
   LayerWidth = m_Layers(Idx).Right - m_Layers(Idx).Left
End Property

Public Property Get LayerX(ByVal Idx As Long) As Long
   CheckLayerIndex Idx
   LayerX = m_Layers(Idx).Left
End Property

Public Property Get LayerY(ByVal Idx As Long) As Long
   CheckLayerIndex Idx
   LayerY = m_Layers(Idx).Top
End Property

Public Property Get Width() As Long
   Width = m_Width
End Property

Public Sub Load(ByRef FileNameOrByteArray As Variant)
   Dim B() As Byte

   ResetState

   Select Case VarType(FileNameOrByteArray)
      Case vbString
         B = New_c.FSO.ReadByteContent(CStr(FileNameOrByteArray))
         CopyToInternalBuffer B

      Case (vbArray Or vbByte)
         CopyVariantByteArrayToInternalBuffer FileNameOrByteArray

      Case Else
         RaisePSD "Load expects a PSD filename or a Byte() array."
   End Select

   ParsePSD
   BuildLayerPaths
   BuildLayerSurfaces
   Set m_AllLayersSurface = RenderAllLayers()
End Sub

Public Function RenderAllLayers() As cCairoSurface
   Dim S As cCairoSurface
   Dim CC As cCairoContext
   Dim I As Long
   Dim A As Double

   If m_Width <= 0 Or m_Height <= 0 Then
      Set RenderAllLayers = Nothing
      Exit Function
   End If

   Set S = Cairo.CreateSurface(m_Width, m_Height)
   Set CC = S.CreateContext

   'Be explicit about the transparent initial state.
   CC.Operator = CAIRO_OPERATOR_CLEAR
   CC.Paint
   CC.Operator = CAIRO_OPERATOR_OVER

   'PSD layer records are stored background-to-foreground for the ordinary
   'layer stack. Folder marker records have no raster surface and are skipped.
   For I = 0 To m_LayerCount - 1
      If m_Layers(I).EffectiveVisible Then
         If Not m_LayerSurfaces(I) Is Nothing Then
            If m_Layers(I).Opacity <> 0 Then
               A = CDbl(m_Layers(I).Opacity) / 255#
               CC.Operator = m_Layers(I).BlendOperator
               CC.RenderSurfaceContent m_LayerSurfaces(I), _
                                       m_Layers(I).Left, _
                                       m_Layers(I).Top, _
                                       0, 0, _
                                       CAIRO_FILTER_GOOD, _
                                       A
            End If
         End If
      End If
   Next I

   CC.Operator = CAIRO_OPERATOR_OVER

   Set m_AllLayersSurface = S
   Set RenderAllLayers = S
End Function

'------------------------------------------------------------------------------
' Parse the PSD container
'------------------------------------------------------------------------------

Private Sub ParsePSD()
   Dim P As Long
   Dim Sig As String
   Dim Version As Long
   Dim DocChannels As Long
   Dim Depth As Long
   Dim ColorMode As Long
   Dim SectionLen As Long
   Dim LayerMaskEnd As Long
   Dim LayerInfoLen As Long
   Dim LayerInfoEnd As Long
   Dim SignedLayerCount As Long
   Dim I As Long
   Dim J As Long
   Dim C As Long
   Dim TotalChannels As Long
   Dim ExtraLen As Long
   Dim ExtraEnd As Long
   Dim MaskLen As Long
   Dim MaskEnd As Long
   Dim BlendRangesLen As Long
   Dim TagSig As String
   Dim TagKey As String
   Dim TagLen As Long
   Dim TagEnd As Long

   P = 0

   Sig = ReadFourCC(P)
   If Sig <> PSD_SIGNATURE Then RaisePSD "Not a PSD file (missing 8BPS signature)."

   Version = ReadU16BE(P)
   If Version <> PSD_VERSION Then
      RaisePSD "Only PSD version 1 is supported; PSB/version 2 is not supported."
   End If

   SkipBytes P, 6 'reserved

   DocChannels = ReadU16BE(P)
   If DocChannels < 1 Or DocChannels > 56 Then RaisePSD "Invalid PSD channel count."

   m_Height = ReadPositiveS32(P, "document height")
   m_Width = ReadPositiveS32(P, "document width")

   Depth = ReadU16BE(P)
   If Depth <> PSD_DEPTH_8 Then
      RaisePSD "This implementation currently supports only 8 bits/channel PSD files."
   End If

   ColorMode = ReadU16BE(P)
   If ColorMode <> PSD_COLOR_MODE_RGB Then
      RaisePSD "This implementation currently supports only RGB PSD files."
   End If

   'Color mode data.
   SectionLen = ReadLength32(P, "color mode data")
   SkipBytes P, SectionLen

   'Image resources.
   SectionLen = ReadLength32(P, "image resources")
   SkipBytes P, SectionLen

   'Layer and mask information.
   SectionLen = ReadLength32(P, "layer and mask information")
   If SectionLen = 0 Then
      m_LayerCount = 0
      Exit Sub
   End If

   LayerMaskEnd = CheckedEnd(P, SectionLen, "layer and mask information")

   LayerInfoLen = ReadLength32(P, "layer info")
   If LayerInfoLen = 0 Then
      P = LayerMaskEnd
      m_LayerCount = 0
      Exit Sub
   End If

   LayerInfoEnd = CheckedEnd(P, LayerInfoLen, "layer info")

   SignedLayerCount = ReadS16BE(P)
   If SignedLayerCount < 0 Then
      If SignedLayerCount = -32768 Then RaisePSD "Invalid layer count."
      m_LayerCount = -SignedLayerCount
   Else
      m_LayerCount = SignedLayerCount
   End If

   If m_LayerCount = 0 Then
      P = LayerInfoEnd
      Exit Sub
   End If

   ReDim m_Layers(0 To m_LayerCount - 1)
   ReDim m_LayerSurfaces(0 To m_LayerCount - 1)

   'First pass: layer records. Count all channel descriptors so that the
   'flat channel array can be allocated once.
   TotalChannels = 0

   For I = 0 To m_LayerCount - 1
      With m_Layers(I)
         .Top = ReadS32BE(P)
         .Left = ReadS32BE(P)
         .Bottom = ReadS32BE(P)
         .Right = ReadS32BE(P)

         If .Bottom < .Top Or .Right < .Left Then
            RaisePSD "Invalid layer rectangle at layer " & CStr(I) & "."
         End If

         C = ReadU16BE(P)
         If C < 0 Or C > 56 Then RaisePSD "Invalid channel count at layer " & CStr(I) & "."

         .ChannelFirst = TotalChannels
         .ChannelCount = C

         TotalChannels = TotalChannels + C
         If TotalChannels < 0 Then RaisePSD "Too many layer channels."
      End With

      'We cannot store the channel descriptors until the final array exists,
      'so temporarily advance past them here. Their bytes are parsed in a
      'second pass below.
      SkipBytes P, C * 6

      If ReadFourCC(P) <> "8BIM" Then
         RaisePSD "Invalid blend-mode signature at layer " & CStr(I) & "."
      End If

      With m_Layers(I)
         .BlendMode = ReadFourCC(P)
         .BlendOperator = BlendModeToCairoOperator(.BlendMode)
         .Opacity = ReadU8(P)
         .Clipping = ReadU8(P)
         .RawFlags = ReadU8(P)
         SkipBytes P, 1 'filler

         .PublicFlags = .RawFlags

         ExtraLen = ReadLength32(P, "layer extra data")
         ExtraEnd = CheckedEnd(P, ExtraLen, "layer extra data")

         'Layer mask data.
         MaskLen = ReadLength32(P, "layer mask data")
         MaskEnd = CheckedEnd(P, MaskLen, "layer mask data")

         If MaskLen >= 18 Then
            .HasMaskInfo = True
            .MaskTop = ReadS32BE(P)
            .MaskLeft = ReadS32BE(P)
            .MaskBottom = ReadS32BE(P)
            .MaskRight = ReadS32BE(P)
            .MaskDefault = ReadU8(P)
            .MaskFlags = ReadU8(P)
         End If
         P = MaskEnd

         'Blending ranges.
         BlendRangesLen = ReadLength32(P, "layer blending ranges")
         SkipBytes P, BlendRangesLen

         'Legacy Pascal name, then tagged extra blocks.
         .Name = ReadPascalString4(P)

         Do While P + 12 <= ExtraEnd
            TagSig = ReadFourCC(P)
            TagKey = ReadFourCC(P)

            If TagSig <> "8BIM" And TagSig <> "8B64" Then
               'There is no safe way to resynchronize an unknown record.
               Exit Do
            End If

            TagLen = ReadLength32(P, "additional layer info: " & TagKey)
            TagEnd = CheckedEnd(P, TagLen, "additional layer info: " & TagKey)

            Select Case TagKey
               Case "luni"
                  .Name = ReadUnicodeString(P, TagEnd)

               Case "lsct", "lsdk"
                  If TagLen >= 4 Then .SectionType = ReadLength32(P, "section-divider type")
            End Select

            P = TagEnd
            If (TagLen And 1) <> 0 Then
               If P < ExtraEnd Then SkipBytes P, 1
            End If
         Loop

         P = ExtraEnd
      End With
   Next I

   'Re-read the layer-record area to capture the channel ID/length descriptors.
   'Doing this avoids an array-of-arrays UDT and keeps the public class simple.
   ParseChannelDescriptors TotalChannels

   'P currently points to the first channel-image-data record because
   'ParseChannelDescriptors preserves the original post-record offset.
   P = ChannelDataStartOffset()

   For I = 0 To m_ChannelCount - 1
      m_Channels(I).DataOffset = P
      SkipBytes P, m_Channels(I).DataLength
   Next I

   If P > LayerInfoEnd Then RaisePSD "Layer channel data extends beyond the layer-info section."

   'The rest of layer/mask information is not needed by this simple renderer.
   P = LayerMaskEnd
End Sub

'Parse the channel descriptors by walking the layer records a second time.
'This keeps the main UDTs simple and avoids temporary per-layer collections.
Private Sub ParseChannelDescriptors(ByVal TotalChannels As Long)
   Dim P As Long
   Dim I As Long
   Dim J As Long
   Dim C As Long
   Dim K As Long
   Dim ExtraLen As Long

   If TotalChannels <= 0 Then
      m_ChannelCount = 0
      Exit Sub
   End If

   ReDim m_Channels(0 To TotalChannels - 1)
   m_ChannelCount = TotalChannels

   P = LayerRecordsStartOffset()
   K = 0

   For I = 0 To m_LayerCount - 1
      SkipBytes P, 16 'rectangle
      C = ReadU16BE(P)

      For J = 0 To C - 1
         m_Channels(K).ID = ReadS16BE(P)
         m_Channels(K).DataLength = ReadLength32(P, "layer channel data")
         K = K + 1
      Next J

      SkipBytes P, 12 'blend signature/key + opacity/clipping/flags/filler
      ExtraLen = ReadLength32(P, "layer extra data")
      SkipBytes P, ExtraLen
   Next I

   If K <> TotalChannels Then RaisePSD "Internal channel descriptor count mismatch."
End Sub

'The fixed file sections are cheap to walk and this avoids retaining parser-only
'offset state in the public object.
Private Function LayerRecordsStartOffset() As Long
   Dim P As Long
   Dim N As Long

   P = 0
   SkipBytes P, 4                 '8BPS
   SkipBytes P, 2                 'version
   SkipBytes P, 6                 'reserved
   SkipBytes P, 2                 'channels
   SkipBytes P, 4                 'height
   SkipBytes P, 4                 'width
   SkipBytes P, 2                 'depth
   SkipBytes P, 2                 'color mode

   N = ReadLength32(P, "color mode data")
   SkipBytes P, N

   N = ReadLength32(P, "image resources")
   SkipBytes P, N

   N = ReadLength32(P, "layer and mask information")
   N = ReadLength32(P, "layer info")
   SkipBytes P, 2                 'layer count

   LayerRecordsStartOffset = P
End Function

Private Function ChannelDataStartOffset() As Long
   Dim P As Long
   Dim I As Long
   Dim C As Long
   Dim ExtraLen As Long

   P = LayerRecordsStartOffset()

   For I = 0 To m_LayerCount - 1
      SkipBytes P, 16
      C = ReadU16BE(P)
      SkipBytes P, C * 6
      SkipBytes P, 12
      ExtraLen = ReadLength32(P, "layer extra data")
      SkipBytes P, ExtraLen
   Next I

   ChannelDataStartOffset = P
End Function

'------------------------------------------------------------------------------
' Layer hierarchy/path handling
'------------------------------------------------------------------------------

Private Sub BuildLayerPaths()
   Dim GroupNames() As String
   Dim GroupVisible() As Boolean
   Dim Depth As Long
   Dim I As Long
   Dim ParentVisible As Boolean
   Dim OwnVisible As Boolean

   If m_LayerCount = 0 Then Exit Sub

   ReDim GroupNames(0 To m_LayerCount - 1)
   ReDim GroupVisible(0 To m_LayerCount - 1)

   'RC6 uses slash-separated paths.  Named folder records and their matching
   '</Layer set>' divider both report the folder path with a trailing slash.
   '
   'Walking background-to-foreground means we encounter the named folder
   'record first, then its children, then its bounding divider.
   For I = m_LayerCount - 1 To 0 Step -1
      If Depth > 0 Then
         ParentVisible = GroupVisible(Depth - 1)
      Else
         ParentVisible = True
      End If

      OwnVisible = ((m_Layers(I).RawFlags And &H2) = 0)

      Select Case m_Layers(I).SectionType
         Case SECTION_OPEN_FOLDER, SECTION_CLOSED_FOLDER
            m_Layers(I).Path = JoinGroupPath(GroupNames, Depth, m_Layers(I).Name, True)
            m_Layers(I).EffectiveVisible = ParentVisible And OwnVisible

            GroupNames(Depth) = m_Layers(I).Name
            GroupVisible(Depth) = ParentVisible And OwnVisible
            Depth = Depth + 1

         Case SECTION_BOUNDING_DIVIDER
            'The divider belongs to the folder which is currently open.
            m_Layers(I).Path = JoinGroupPath(GroupNames, Depth, vbNullString, True)
            m_Layers(I).EffectiveVisible = False

            If Depth > 0 Then Depth = Depth - 1

         Case Else
            m_Layers(I).Path = JoinGroupPath(GroupNames, Depth, m_Layers(I).Name)
            m_Layers(I).EffectiveVisible = ParentVisible And OwnVisible
      End Select
   Next I
End Sub

Private Function JoinGroupPath(ByRef Names() As String, ByVal Depth As Long, _
                               ByVal Leaf As String, _
                               Optional ByVal FolderPath As Boolean = False) As String
   Dim I As Long
   Dim S As String

   For I = 0 To Depth - 1
      If Len(Names(I)) <> 0 Then
         If Len(S) <> 0 Then S = S & "/"
         S = S & Names(I)
      End If
   Next I

   If Len(Leaf) <> 0 Then
      If Len(S) <> 0 Then S = S & "/"
      S = S & Leaf
   End If

   If FolderPath And Len(S) <> 0 Then S = S & "/"

   JoinGroupPath = S
End Function

'------------------------------------------------------------------------------
' Decode raster layers into Cairo surfaces
'------------------------------------------------------------------------------

Private Sub BuildLayerSurfaces()
   Dim I As Long

   For I = 0 To m_LayerCount - 1
      If m_Layers(I).SectionType = SECTION_OTHER Then
         Set m_LayerSurfaces(I) = BuildLayerSurface(I)
      End If
   Next I
End Sub

Private Function BuildLayerSurface(ByVal LayerIndex As Long) As cCairoSurface
   Dim W As Long
   Dim H As Long
   Dim Red() As Byte
   Dim Green() As Byte
   Dim Blue() As Byte
   Dim Alpha() As Byte
   Dim Mask() As Byte
   Dim HasAlpha As Boolean
   Dim HasMask As Boolean
   Dim S As cCairoSurface
   Dim Pix() As Byte
   Dim Stride As Long
   Dim X As Long
   Dim Y As Long
   Dim SrcPos As Long
   Dim DstPos As Long
   Dim A As Long
   Dim M As Long
   Dim R As Long
   Dim G As Long
   Dim B As Long
   Dim MX As Long
   Dim MY As Long
   Dim MW As Long
   Dim MH As Long
   Dim DocX As Long
   Dim DocY As Long
   Dim MaskLeft As Long
   Dim MaskTop As Long

   W = m_Layers(LayerIndex).Right - m_Layers(LayerIndex).Left
   H = m_Layers(LayerIndex).Bottom - m_Layers(LayerIndex).Top

   If W <= 0 Or H <= 0 Then Exit Function

   'Adjustment/fill/group records can advertise RGB channels whose payload is
   'undefined or empty.  Those are intentionally not treated as raster layers.
   If Not LayerChannelHasPayload(LayerIndex, CH_RED) Then Exit Function
   If Not LayerChannelHasPayload(LayerIndex, CH_GREEN) Then Exit Function
   If Not LayerChannelHasPayload(LayerIndex, CH_BLUE) Then Exit Function

   Red = DecodeLayerChannel(LayerIndex, CH_RED, W, H)
   Green = DecodeLayerChannel(LayerIndex, CH_GREEN, W, H)
   Blue = DecodeLayerChannel(LayerIndex, CH_BLUE, W, H)

   HasAlpha = LayerChannelHasPayload(LayerIndex, CH_TRANSPARENCY)
   If HasAlpha Then Alpha = DecodeLayerChannel(LayerIndex, CH_TRANSPARENCY, W, H)

   HasMask = False

   If m_Layers(LayerIndex).HasMaskInfo Then
      If (m_Layers(LayerIndex).MaskFlags And &H2) = 0 Then
         If FindLayerChannel(LayerIndex, CH_USER_MASK) >= 0 Then
            MW = m_Layers(LayerIndex).MaskRight - m_Layers(LayerIndex).MaskLeft
            MH = m_Layers(LayerIndex).MaskBottom - m_Layers(LayerIndex).MaskTop

            If MW > 0 And MH > 0 Then
               Mask = DecodeLayerChannel(LayerIndex, CH_USER_MASK, MW, MH)
               HasMask = True

               If (m_Layers(LayerIndex).MaskFlags And &H1) <> 0 Then
                  MaskLeft = m_Layers(LayerIndex).Left + m_Layers(LayerIndex).MaskLeft
                  MaskTop = m_Layers(LayerIndex).Top + m_Layers(LayerIndex).MaskTop
               Else
                  MaskLeft = m_Layers(LayerIndex).MaskLeft
                  MaskTop = m_Layers(LayerIndex).MaskTop
               End If
            End If
         End If
      End If
   End If

   Set S = Cairo.CreateSurface(W, H)

   If Not S.BindToArray(Pix, True) Then
      RaisePSD "RC6 could not bind the Cairo surface to a Byte array."
   End If

   Stride = S.Stride

   'The overwhelmingly common case is an ordinary raster layer with an alpha
   'channel and no enabled user mask.  Keep that hot path linear and branch-free.
   'The previous code tested HasAlpha/HasMask and calculated mask/document
   'coordinates for every pixel even when no mask existed.
   If Not HasMask Then
      If HasAlpha Then
         If Stride = W * 4 Then
            DstPos = 0

            For SrcPos = 0 To W * H - 1
               Pix(DstPos) = Blue(SrcPos)
               Pix(DstPos + 1) = Green(SrcPos)
               Pix(DstPos + 2) = Red(SrcPos)
               Pix(DstPos + 3) = Alpha(SrcPos)
               DstPos = DstPos + 4
            Next SrcPos
         Else
            For Y = 0 To H - 1
               SrcPos = Y * W
               DstPos = Y * Stride

               For X = 0 To W - 1
                  Pix(DstPos) = Blue(SrcPos)
                  Pix(DstPos + 1) = Green(SrcPos)
                  Pix(DstPos + 2) = Red(SrcPos)
                  Pix(DstPos + 3) = Alpha(SrcPos)

                  SrcPos = SrcPos + 1
                  DstPos = DstPos + 4
               Next X
            Next Y
         End If
      Else
         If Stride = W * 4 Then
            DstPos = 0

            For SrcPos = 0 To W * H - 1
               Pix(DstPos) = Blue(SrcPos)
               Pix(DstPos + 1) = Green(SrcPos)
               Pix(DstPos + 2) = Red(SrcPos)
               Pix(DstPos + 3) = 255
               DstPos = DstPos + 4
            Next SrcPos
         Else
            For Y = 0 To H - 1
               SrcPos = Y * W
               DstPos = Y * Stride

               For X = 0 To W - 1
                  Pix(DstPos) = Blue(SrcPos)
                  Pix(DstPos + 1) = Green(SrcPos)
                  Pix(DstPos + 2) = Red(SrcPos)
                  Pix(DstPos + 3) = 255

                  SrcPos = SrcPos + 1
                  DstPos = DstPos + 4
               Next X
            Next Y
         End If
      End If
   Else
      'Only masked layers pay for document-relative mask coordinate handling.
      For Y = 0 To H - 1
         SrcPos = Y * W
         DstPos = Y * Stride
         DocY = m_Layers(LayerIndex).Top + Y

         For X = 0 To W - 1
            If HasAlpha Then
               A = Alpha(SrcPos)
            Else
               A = 255
            End If

            DocX = m_Layers(LayerIndex).Left + X
            MX = DocX - MaskLeft
            MY = DocY - MaskTop

            If MX >= 0 And MX < MW And MY >= 0 And MY < MH Then
               M = Mask(MY * MW + MX)
            Else
               M = m_Layers(LayerIndex).MaskDefault
            End If

            If (m_Layers(LayerIndex).MaskFlags And &H4) <> 0 Then M = 255 - M
            A = (A * M + 127) \ 255

            Pix(DstPos) = Blue(SrcPos)
            Pix(DstPos + 1) = Green(SrcPos)
            Pix(DstPos + 2) = Red(SrcPos)
            Pix(DstPos + 3) = CByte(A)

            SrcPos = SrcPos + 1
            DstPos = DstPos + 4
         Next X
      Next Y
   End If

   Cairo.PreMultiplyAlpha S.DataPtr, Stride * H
   S.ReleaseArray Pix
   Set BuildLayerSurface = S
End Function

Private Function DecodeLayerChannel(ByVal LayerIndex As Long, ByVal ChannelID As Long, _
                                    ByVal W As Long, ByVal H As Long) As Byte()
   Dim CI As Long

   CI = FindLayerChannel(LayerIndex, ChannelID)
   If CI < 0 Then RaisePSD "Requested layer channel is not present."

   DecodeLayerChannel = DecodeChannel(CI, W, H)
End Function

Private Function DecodeChannel(ByVal ChannelIndex As Long, ByVal W As Long, ByVal H As Long) As Byte()
   Dim P As Long
   Dim DataEnd As Long
   Dim Compression As Long
   Dim Expected As Long
   Dim B() As Byte

   Expected = CheckedPixelCount(W, H)
   If Expected = 0 Then
      DecodeChannel = B
      Exit Function
   End If

   P = m_Channels(ChannelIndex).DataOffset
   DataEnd = CheckedEnd(P, m_Channels(ChannelIndex).DataLength, "channel data")

   If DataEnd - P < 2 Then RaisePSD "Layer channel is too short."

   Compression = ReadU16BE(P)

   Select Case Compression
      Case COMP_RAW
         B = DecodeRaw(P, DataEnd, Expected)

      Case COMP_RLE
         B = DecodeRLE(P, DataEnd, W, H)

      Case COMP_ZIP
         B = DecodeZIP(P, DataEnd, W, H, False)

      Case COMP_ZIP_PREDICTION
         B = DecodeZIP(P, DataEnd, W, H, True)

      Case Else
         RaisePSD "Unsupported PSD channel compression method " & CStr(Compression) & "."
   End Select

   DecodeChannel = B
End Function

Private Function DecodeRaw(ByVal P As Long, ByVal DataEnd As Long, ByVal Expected As Long) As Byte()
   Dim B() As Byte

   If DataEnd - P < Expected Then RaisePSD "RAW layer channel is shorter than expected."

   ReDim B(0 To Expected - 1)

   'RC6 already exposes its own native memory-copy helper, so avoid a VB6 loop.
   New_c.MemCopy VarPtr(B(0)), VarPtr(m_Data(P)), Expected

   DecodeRaw = B
End Function

Private Function DecodeRLE(ByVal P As Long, ByVal DataEnd As Long, _
                           ByVal W As Long, ByVal H As Long) As Byte()
   Dim RowBytes() As Long
   Dim B() As Byte
   Dim Row As Long
   Dim RowEnd As Long
   Dim RowOutEnd As Long
   Dim N As Long
   Dim Count As Long
   Dim V As Byte
   Dim I As Long
   Dim OutPos As Long

   ReDim RowBytes(0 To H - 1)
   ReDim B(0 To CheckedPixelCount(W, H) - 1)

   If DataEnd - P < H * 2 Then RaisePSD "RLE channel is missing its row-length table."

   For Row = 0 To H - 1
      RowBytes(Row) = ReadU16BE(P)
   Next Row

   OutPos = 0

   For Row = 0 To H - 1
      RowEnd = CheckedEnd(P, RowBytes(Row), "RLE row")
      If RowEnd > DataEnd Then RaisePSD "RLE row extends beyond channel data."

      RowOutEnd = OutPos + W

      Do While P < RowEnd And OutPos < RowOutEnd
         N = m_Data(P)
         P = P + 1

         If N <= 127 Then
            Count = N + 1
            If Count > RowOutEnd - OutPos Then RaisePSD "PackBits literal run exceeds row width."
            If P + Count > RowEnd Then RaisePSD "PackBits literal run exceeds row data."

            'Keep source and destination positions linear.  The old loop
            'recomputed OutPos + X and incremented X on every copied byte.
            For I = 1 To Count
               B(OutPos) = m_Data(P)
               OutPos = OutPos + 1
               P = P + 1
            Next I

         ElseIf N >= 129 Then
            Count = 257 - N
            If Count > RowOutEnd - OutPos Then RaisePSD "PackBits repeated run exceeds row width."
            If P >= RowEnd Then RaisePSD "PackBits repeated run has no value byte."

            V = m_Data(P)
            P = P + 1

            For I = 1 To Count
               B(OutPos) = V
               OutPos = OutPos + 1
            Next I

         Else
            '128 is the PackBits no-op.
         End If
      Loop

      If OutPos <> RowOutEnd Then RaisePSD "PackBits row did not decode to the expected width."
      P = RowEnd
   Next Row

   DecodeRLE = B
End Function

Private Function DecodeZIP(ByVal P As Long, ByVal DataEnd As Long, _
                           ByVal W As Long, ByVal H As Long, _
                           ByVal WithPrediction As Boolean) As Byte()
   Dim Src() As Byte
   Dim Dst() As Byte
   Dim B() As Byte
   Dim SrcLen As Long
   Dim Expected As Long
   Dim DstLen As Long
   Dim X As Long
   Dim Y As Long
   Dim RowPos As Long
   Dim V As Long

   SrcLen = DataEnd - P
   Expected = CheckedPixelCount(W, H)

   If SrcLen <= 0 Then RaisePSD "ZIP layer channel contains no compressed data."

   ReDim Src(0 To SrcLen - 1)
   New_c.MemCopy VarPtr(Src(0)), VarPtr(m_Data(P)), SrcLen

   New_c.Crypt.ZLibDecompress Src, Dst

   DstLen = ByteArrayCount(Dst)
   If DstLen < Expected Then RaisePSD "ZIP layer channel decompressed to fewer bytes than expected."

   ReDim B(0 To Expected - 1)
   New_c.MemCopy VarPtr(B(0)), VarPtr(Dst(LBound(Dst))), Expected

   If WithPrediction Then
      'For 8-bit channels, Photoshop's ZIP prediction is a per-row delta:
      'each byte after the first is restored by adding the previous byte.
      For Y = 0 To H - 1
         RowPos = Y * W
         For X = 1 To W - 1
            V = CLng(B(RowPos + X)) + CLng(B(RowPos + X - 1))
            If V > 255 Then V = V - 256
            B(RowPos + X) = CByte(V)
         Next X
      Next Y
   End If

   DecodeZIP = B
End Function

Private Function FindLayerChannel(ByVal LayerIndex As Long, ByVal ChannelID As Long) As Long
   Dim I As Long
   Dim First As Long
   Dim Count As Long

   First = m_Layers(LayerIndex).ChannelFirst
   Count = m_Layers(LayerIndex).ChannelCount

   For I = First To First + Count - 1
      If m_Channels(I).ID = ChannelID Then
         FindLayerChannel = I
         Exit Function
      End If
   Next I

   FindLayerChannel = -1
End Function

Private Function LayerChannelHasPayload(ByVal LayerIndex As Long, ByVal ChannelID As Long) As Boolean
   Dim CI As Long

   CI = FindLayerChannel(LayerIndex, ChannelID)
   If CI < 0 Then Exit Function

   'A real pixel channel contains its two-byte compression word plus image data.
   'Shorter records are non-raster/undefined data and are deliberately ignored.
   LayerChannelHasPayload = (m_Channels(CI).DataLength > 2)
End Function

'------------------------------------------------------------------------------
' Blend-mode mapping
'------------------------------------------------------------------------------

Private Function BlendModeToCairoOperator(ByVal Key As String) As Long
   Select Case Key
      Case "norm", "pass", "diss"
         BlendModeToCairoOperator = CAIRO_OPERATOR_OVER

      Case "dark"
         BlendModeToCairoOperator = Cairo_OPERATOR_DARKEN

      Case "mul "
         BlendModeToCairoOperator = Cairo_OPERATOR_MULTIPLY

      Case "idiv"
         BlendModeToCairoOperator = Cairo_OPERATOR_COLOR_BURN

      Case "lite"
         BlendModeToCairoOperator = Cairo_OPERATOR_LIGHTEN

      Case "scrn"
         BlendModeToCairoOperator = Cairo_OPERATOR_SCREEN

      Case "div "
         BlendModeToCairoOperator = Cairo_OPERATOR_COLOR_DODGE

      Case "lddg"
         BlendModeToCairoOperator = Cairo_OPERATOR_ADD

      Case "over"
         BlendModeToCairoOperator = Cairo_OPERATOR_OVERLAY

      Case "sLit"
         BlendModeToCairoOperator = Cairo_OPERATOR_SOFT_LIGHT

      Case "hLit"
         BlendModeToCairoOperator = Cairo_OPERATOR_HARD_LIGHT

      Case "diff"
         BlendModeToCairoOperator = Cairo_OPERATOR_DIFFERENCE

      Case "smud"
         BlendModeToCairoOperator = Cairo_OPERATOR_EXCLUSION

      Case "hue "
         BlendModeToCairoOperator = Cairo_OPERATOR_HSL_HUE

      Case "sat "
         BlendModeToCairoOperator = Cairo_OPERATOR_HSL_SATURATION

      Case "colr"
         BlendModeToCairoOperator = Cairo_OPERATOR_HSL_COLOR

      Case "lum "
         BlendModeToCairoOperator = Cairo_OPERATOR_HSL_LUMINOSITY

      Case Else
         'Linear burn/darker color/vivid light/linear light/pin light/hard mix/
         'subtract/divide/etc. have no exact Cairo operator exposed by RC6.
         BlendModeToCairoOperator = CAIRO_OPERATOR_OVER
   End Select
End Function

Private Function ConvertPSDFlags(ByVal RawFlags As Byte) As Long
   Dim F As Long

   If (RawFlags And &H1) <> 0 Then F = F Or LF_PROTECTED

   'Despite the Adobe table wording, bit 1 is the stored "hidden" bit in
   'ordinary PSD layer records: zero means visible.
   If (RawFlags And &H2) = 0 Then F = F Or LF_VISIBLE

   If (RawFlags And &H4) <> 0 Then F = F Or LF_OBSOLETE
   If (RawFlags And &H8) <> 0 Then F = F Or LF_USEFUL
   If (RawFlags And &H10) <> 0 Then F = F Or LF_IRRELEVANT

   ConvertPSDFlags = F
End Function

'------------------------------------------------------------------------------
' Primitive big-endian readers
'------------------------------------------------------------------------------

Private Function ReadU8(ByRef P As Long) As Byte
   EnsureAvailable P, 1
   ReadU8 = m_Data(P)
   P = P + 1
End Function

Private Function ReadU16BE(ByRef P As Long) As Long
   EnsureAvailable P, 2
   ReadU16BE = CLng(m_Data(P)) * 256& + CLng(m_Data(P + 1))
   P = P + 2
End Function

Private Function ReadS16BE(ByRef P As Long) As Long
   Dim U As Long

   U = ReadU16BE(P)
   If U >= &H8000& Then U = U - &H10000
   ReadS16BE = U
End Function

Private Function ReadU32BEDouble(ByRef P As Long) As Double
   EnsureAvailable P, 4

   ReadU32BEDouble = _
       CDbl(m_Data(P)) * 16777216# + _
       CDbl(m_Data(P + 1)) * 65536# + _
       CDbl(m_Data(P + 2)) * 256# + _
       CDbl(m_Data(P + 3))

   P = P + 4
End Function

Private Function ReadS32BE(ByRef P As Long) As Long
   Dim U As Double

   U = ReadU32BEDouble(P)
   If U >= 2147483648# Then U = U - 4294967296#
   ReadS32BE = CLng(U)
End Function

Private Function ReadLength32(ByRef P As Long, ByVal What As String) As Long
   Dim U As Double

   U = ReadU32BEDouble(P)
   If U > 2147483647# Then RaisePSD What & " is too large for this VB6 implementation."
   ReadLength32 = CLng(U)
End Function

Private Function ReadPositiveS32(ByRef P As Long, ByVal What As String) As Long
   Dim V As Long

   V = ReadS32BE(P)
   If V <= 0 Then RaisePSD "Invalid " & What & "."
   ReadPositiveS32 = V
End Function

Private Function ReadFourCC(ByRef P As Long) As String
   EnsureAvailable P, 4

   ReadFourCC = Chr$(m_Data(P)) & _
                Chr$(m_Data(P + 1)) & _
                Chr$(m_Data(P + 2)) & _
                Chr$(m_Data(P + 3))

   P = P + 4
End Function

Private Function ReadPascalString4(ByRef P As Long) As String
   Dim StartPos As Long
   Dim N As Long
   Dim I As Long
   Dim Used As Long
   Dim Pad As Long
   Dim S As String

   StartPos = P
   N = ReadU8(P)

   EnsureAvailable P, N

   If N > 0 Then
      S = String$(N, vbNullChar)
      For I = 1 To N
         Mid$(S, I, 1) = Chr$(m_Data(P))
         P = P + 1
      Next I
   End If

   Used = P - StartPos
   Pad = (4 - (Used Mod 4)) Mod 4
   SkipBytes P, Pad

   ReadPascalString4 = S
End Function

Private Function ReadUnicodeString(ByRef P As Long, ByVal Limit As Long) As String
   Dim Count As Long
   Dim I As Long
   Dim U As Long
   Dim S As String

   If P + 4 > Limit Then Exit Function

   Count = ReadLength32(P, "Unicode layer-name length")
   If Count < 0 Then RaisePSD "Invalid Unicode layer-name length."

   If CDbl(Count) * 2# > CDbl(Limit - P) Then
      RaisePSD "Unicode layer name extends beyond its tagged block."
   End If

   If Count > 0 Then
      S = String$(Count, vbNullChar)

      For I = 1 To Count
         U = ReadU16BE(P)

         If U > 32767 Then
            Mid$(S, I, 1) = ChrW$(U - 65536)
         Else
            Mid$(S, I, 1) = ChrW$(U)
         End If
      Next I

      'A terminating UTF-16 NUL is sometimes included in the count.
      Do While Len(S) > 0 And Right$(S, 1) = vbNullChar
         S = Left$(S, Len(S) - 1)
      Loop
   End If

   ReadUnicodeString = S
End Function

Private Sub SkipBytes(ByRef P As Long, ByVal Count As Long)
   If Count < 0 Then RaisePSD "Negative byte count while parsing PSD."
   EnsureAvailable P, Count
   P = P + Count
End Sub

Private Sub EnsureAvailable(ByVal P As Long, ByVal Count As Long)
   If P < 0 Or Count < 0 Then RaisePSD "Invalid PSD read position."

   If CDbl(P) + CDbl(Count) > CDbl(m_DataSize) Then
      RaisePSD "Unexpected end of PSD data."
   End If
End Sub

Private Function CheckedEnd(ByVal P As Long, ByVal Count As Long, ByVal What As String) As Long
   Dim D As Double

   If Count < 0 Then RaisePSD "Negative length for " & What & "."

   D = CDbl(P) + CDbl(Count)
   If D > 2147483647# Then RaisePSD What & " exceeds VB6 addressable range."
   If D > CDbl(m_DataSize) Then RaisePSD What & " extends beyond the end of the PSD."

   CheckedEnd = CLng(D)
End Function

Private Function CheckedPixelCount(ByVal W As Long, ByVal H As Long) As Long
   Dim D As Double

   If W < 0 Or H < 0 Then RaisePSD "Invalid negative image dimension."

   D = CDbl(W) * CDbl(H)
   If D > 2147483647# Then RaisePSD "Layer is too large for a VB6 byte array."

   CheckedPixelCount = CLng(D)
End Function

'------------------------------------------------------------------------------
' Buffer/state helpers
'------------------------------------------------------------------------------

Private Sub CopyToInternalBuffer(ByRef B() As Byte)
   Dim N As Long
   Dim LB As Long

   N = ByteArrayCount(B)
   If N <= 0 Then RaisePSD "The PSD input is empty."

   LB = LBound(B)
   ReDim m_Data(0 To N - 1)

   'Use RC6's native copy rather than copying the PSD one byte at a time.
   New_c.MemCopy VarPtr(m_Data(0)), VarPtr(B(LB)), N

   m_DataSize = N
End Sub

Private Sub CopyVariantByteArrayToInternalBuffer(ByRef V As Variant)
   Dim LB As Long
   Dim UB As Long
   Dim N As Long
   Dim B() As Byte

   On Error GoTo BadArray

   LB = LBound(V)
   UB = UBound(V)
   N = UB - LB + 1

   If N <= 0 Then RaisePSD "The PSD input is empty."

   If LB = 0 Then
      'VB6 permits a Variant containing Byte() to be assigned to a typed Byte()
      'array.  The runtime performs this SAFEARRAY copy natively.
      m_Data = V
   Else
      'Preserve our parser's zero-based offset convention for unusual input
      'arrays with a non-zero lower bound.
      B = V
      ReDim m_Data(0 To N - 1)
      New_c.MemCopy VarPtr(m_Data(0)), VarPtr(B(LB)), N
   End If

   m_DataSize = N
   Exit Sub

BadArray:
   RaisePSD "Load received an invalid Byte() array."
End Sub

Private Function ByteArrayCount(ByRef B() As Byte) As Long
   On Error GoTo EmptyArray

   ByteArrayCount = UBound(B) - LBound(B) + 1
   Exit Function

EmptyArray:
   ByteArrayCount = 0
End Function

Private Sub ResetState()
   Dim I As Long

   Set m_AllLayersSurface = Nothing

   If m_LayerCount > 0 Then
      For I = 0 To m_LayerCount - 1
         Set m_LayerSurfaces(I) = Nothing
      Next I
   End If

   Erase m_LayerSurfaces
   Erase m_Layers
   Erase m_Channels
   Erase m_Data

   m_DataSize = 0
   m_LayerCount = 0
   m_ChannelCount = 0
   m_Width = 0
   m_Height = 0
End Sub

Private Sub CheckLayerIndex(ByVal Idx As Long)
   If Idx < 0 Or Idx >= m_LayerCount Then
      Err.Raise 9, TypeName(Me), "Layer index is out of range."
   End If
End Sub

Private Sub RaisePSD(ByVal Description As String)
   Err.Raise ERR_BASE, TypeName(Me), Description
End Sub


