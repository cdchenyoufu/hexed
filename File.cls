VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Chunks() As Chunk
Private mFilename As String

Private mLoaded As Boolean

Private Chunk1 As Chunk
Private Chunk2 As Chunk
Private ReSized As Boolean
Public size As Long

Public Property Get Filename() As String
    Filename = mFilename
End Property

Private Sub Class_Initialize()
   ' Set Chunks = New Collection
    mLoaded = False
    ReSized = False
End Sub

Public Sub LoadFromMemory(bArray As Variant)
    'dzzie chunk size const doesnt matter cause we load whatever length is
    'passed to us as one chunk regardless..
    
    Dim b() As Byte
    Dim ch As Chunk
    
    If TypeName(bArray) <> "Byte()" Then
        MsgBox "File.LoadFromMemory wrong data type to load!", vbExclamation
        Exit Sub
    End If
    
    mLoaded = True
    ReSized = False
    
    b() = bArray
    ReDim Chunks(1)
    Set ch = New Chunk
    ch.LoadFromMemory bArray
    ch.Offset = 0
    ch.size = UBound(b) + 1
    ch.Index = 1
    size = UBound(b)
    
    Set Chunks(1) = ch
    Set Chunk1 = ch
    
End Sub

Public Sub Load(Filename As String)
    Dim FileSize As Long
    Dim NumChunks As Long
    Dim Chunk As Chunk
    Dim i As Long
    
    FileSize = FileLen(Filename)
    If FileSize < 1 Then
        MsgBox "File is zero length"
        Exit Sub
    ElseIf FileSize > 20000000000# Then
        MsgBox "File is too big"
        Exit Sub
    End If
    mFilename = Filename
    
    size = FileSize - 1
    NumChunks = Int(size / ChunkSize) + 1
    ReDim Chunks(1 To NumChunks)
    For i = 0 To NumChunks - 1
        Set Chunk = New Chunk
        Chunk.Filename = mFilename
        Chunk.Offset = i * ChunkSize
        Chunk.size = ChunkSize
        Chunk.Index = i + 1
        Set Chunks(i + 1) = Chunk
    Next
    Chunks(NumChunks).size = FileSize - Chunks(NumChunks).Offset
    
    mLoaded = True
    ReSized = False
End Sub

Public Sub ActivateChunk(Pos As Long)
    Dim Chunk As Chunk
    Dim i As Long
    'bailout if not loaded
    If mLoaded = False Then Exit Sub
    
    
    If Not Chunk1 Is Nothing Then
        If Pos >= Chunk1.Offset And Pos <= Chunk1.size + Chunk1.Offset - 1 Then
            'chunk already loaded
            Exit Sub
        End If
    End If
    

    Dim Current As Long

    Current = GetChunkFromPos(Pos)
    Set Chunk1 = Chunks(Current)
    Chunk1.Load
    If Current < UBound(Chunks) Then
        Set Chunk2 = Chunks(Current + 1)
        Chunk2.Load
    Else
        Set Chunk2 = Nothing
    End If
End Sub

Private Function GetChunkFromPos(ByVal Pos As Long) As Long
    Dim Upper As Long
    Dim Lower As Long
    Dim Current As Long
    Dim hit As Boolean
    Dim Chunk As Chunk
 
    hit = False
    Lower = 1
    Upper = UBound(Chunks)
    
    Do While Not hit And Upper >= Lower
        Current = Int((Upper + Lower) / 2)
        Set Chunk = Chunks(Current)
        If Pos >= Chunk.Offset And Pos <= Chunk.Offset + Chunk.size Then
            hit = True
            GetChunkFromPos = Current   'return pos
            Exit Function
        ElseIf Pos > Chunk.Offset + Chunk.size Then
            Lower = Current + 1
        Else
            Upper = Current - 1
        End If
    Loop
    

End Function



Public Property Get Data(ByVal Index As Long) As Byte
    If mLoaded = False Then Exit Property
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Then
        Data = Chunk1.Data(Index)
    Else
        Data = Chunk2.Data(Index)
    End If
End Property

Public Function DataScreen(ByVal Index As Long, ByVal Count As Long) As Byte()
    'super optimized code to get screen data
    
    Dim buff() As Byte
    Dim Fillcount As Long
    Dim Offset As Long
    'buff
    ReDim buff(Count - 1)
    If mLoaded = False Then
        ReDim buff(Count - 1)
        DataScreen = buff
        Exit Function
    End If
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Then
        Offset = 0
        Chunk1.DataScreen Index, Count, Offset, Fillcount, buff()
        If Fillcount <> Count Then
            If Not Chunk2 Is Nothing Then
                Offset = Fillcount
                Chunk2.DataScreen Chunk2.Offset, Count - Fillcount, Offset, Fillcount, buff()
            End If
        End If
        DataScreen = buff
    Else
        If Not Chunk2 Is Nothing Then
            Offset = 0
            Chunk2.DataScreen Index, Count, Offset, Fillcount, buff()
        End If
        DataScreen = buff
    End If
End Function

Public Function StatusScreen(ByVal Index As Long, ByVal Count As Long) As Byte()
    'super optimized code to get screen data
    
    Dim buff() As Byte
    Dim Fillcount As Long
    Dim Offset As Long
    'buff
    ReDim buff(Count - 1)
    If mLoaded = False Then
        ReDim buff(Count - 1)
        StatusScreen = buff
        Exit Function
    End If
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Then
        Offset = 0
        Chunk1.StatusScreen Index, Count, Offset, Fillcount, buff()
        If Fillcount <> Count Then
            If Not Chunk2 Is Nothing Then
                Offset = Fillcount
                Chunk2.StatusScreen Chunk2.Offset, Count - Fillcount, Offset, Fillcount, buff()
            End If
        End If
        StatusScreen = buff
    Else
        If Not Chunk2 Is Nothing Then
            Offset = 0
            Chunk2.StatusScreen Index, Count, Offset, Fillcount, buff()
        End If
        StatusScreen = buff
    End If
End Function

Public Property Let Data(ByVal Index As Long, vData As Byte)
    If mLoaded = False Then Exit Property
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Then
        Chunk1.Data(Index) = vData
    Else
        Chunk2.Data(Index) = vData
    End If
End Property

Public Property Get Status(ByVal Index As Long) As Byte
    If mLoaded = False Then Exit Property
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Then
        Status = Chunk1.Status(Index)
    Else
        Status = Chunk2.Status(Index)
    End If
End Property

Public Property Let Status(ByVal Index As Long, vData As Byte)
    If mLoaded = False Then Exit Property
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Then
        Chunk1.Status(Index) = vData
    Else
        Chunk2.Status(Index) = vData
    End If
End Property


Public Sub InsertData(ByVal Index As Long, Data() As Byte)
    If mLoaded = False Then Exit Sub
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Or Chunk2 Is Nothing Then
        Chunk1.InsertData Index, Data
    Else
        Chunk2.InsertData Index, Data
    End If
    Call FixOffsets
    size = size + UBound(Data) + 1
    ReSized = True
End Sub

Public Sub InsertDataStatus(ByVal Index As Long, Data() As Byte, Status() As Byte)
    If mLoaded = False Then Exit Sub
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Or Chunk2 Is Nothing Then
        Chunk1.InsertDataStatus Index, Data, Status
    Else
        Chunk2.InsertDataStatus Index, Data, Status
    End If
    Call FixOffsets
    size = size + UBound(Data) + 1
    ReSized = True
End Sub

Public Sub DeleteData(ByVal Index As Long, ByVal Length As Long)
    '' fixa så man kan deletea över flera chunks
    ''
    Dim LastChunkID As Long
    Dim LastChunk As Chunk
    Dim tmpChunk As Chunk
    Dim Length1 As Long
    Dim Length2 As Long
    Dim Length3 As Long
    
    Dim i As Long
    
    
    If size < 0 Then Exit Sub
    
    If mLoaded = False Then Exit Sub
    If Index >= Chunk1.Offset And Index <= Chunk1.size + Chunk1.Offset - 1 Then
        If (Index + Length - 1) > (Chunk1.Offset + Chunk1.size - 1) Then
            LastChunkID = GetChunkFromPos(Index + Length - 1)
            '//lastchunk is the chunk that SELEND is located in
            Set LastChunk = Chunks(LastChunkID)
            'pos 7520 +14
    
            Length1 = (Chunk1.Offset + Chunk1.size) - Index
            Chunk1.DeleteData Index, Length1
            
            If LastChunk.Index = Chunk1.Index + 1 Then
                'overlapping one block only  - seems to work fine now =)
                Length2 = Length - Length1
                Chunk2.DeleteData Chunk2.Offset, Length2
                        
            Else
                'overlapping multiple blocks
                For i = Chunk1.Index + 1 To LastChunk.Index - 1
                    Set tmpChunk = Chunks(i)
                    Length3 = Length3 + tmpChunk.size
                    tmpChunk.DeleteData tmpChunk.Offset, tmpChunk.size
                Next
                
                Length2 = (Length - Length3) - Length1
                
                MsgBox "ohoj!!" & Length2
                LastChunk.DeleteData LastChunk.Offset, Length2
                
            End If
        Else
        'delete inside one chunk only
            Chunk1.DeleteData Index, Length
        End If
    Else
        If Not Chunk2 Is Nothing Then
            Chunk2.DeleteData Index, Length
        End If
    End If
    Call FixOffsets
    size = size - Length
    ReSized = True
End Sub

Private Sub FixOffsets()
    Dim Offset As Long
    Dim Chunk As Chunk
    Dim i As Long
    For i = 1 To UBound(Chunks)
        Set Chunk = Chunks(i)
        Chunk.Offset = Offset
        Offset = Chunk.Offset + Chunk.size
    Next
End Sub

Public Sub Save()
    If ReSized Then
        Call SaveReSized
    Else
        Call SaveNormal
    End If
End Sub

Private Sub SaveNormal()
    Dim i As Long
    For i = 1 To UBound(Chunks)
        Call Chunks(i).UnLoad
        If Chunks(i).IsInitialized Then
            'MsgBox "save chunk " & i
            Chunks(i).Save
            Set Chunks(i) = Nothing
        End If
    Next
    Erase Chunks
    Set Chunk1 = Nothing
    Set Chunk2 = Nothing
    Me.Load Me.Filename
    'save the file
End Sub

Private Sub SaveReSized()
    Dim i As Long
    Dim NewFilename  As String
    On Error Resume Next
        
    frmProgress.Show
    frmProgress.Refresh
    frmProgress.pb.Min = 1
    frmProgress.pb.Max = UBound(Chunks)
    Screen.MousePointer = vbHourglass
    NewFilename = App.path & "\resized.tmp"
    Kill NewFilename
    Err.Clear
    
    For i = 1 To UBound(Chunks)
        frmProgress.pb.Value = i
        Call Chunks(i).UnLoad
        Chunks(i).NewFilename = NewFilename
        Chunks(i).SaveNew
        Set Chunks(i) = Nothing
    Next
    
    frmProgress.Hide
    UnLoad frmProgress
    
    Erase Chunks
    Set Chunk1 = Nothing
    Set Chunk2 = Nothing

    Kill Me.Filename
    Err.Clear
    
    FileSystem.FileCopy NewFilename, Me.Filename
    Kill NewFilename
    
    Me.Load Me.Filename
    
    Screen.MousePointer = vbDefault
    'save the file
End Sub

