VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Chunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DataName As String
Public StatusName As String
Public Filename As String
Public NewFilename As String

Public size As Long
Private mData() As Byte
Private mStatus() As Byte
Private mLoaded As Boolean
Public Offset As Long
Private mInitialized As Boolean
Private mIsDirty As Boolean
Public Index As Long

Private Sub Class_Initialize()
   ' mLoaded = False
   ' Offset = 0
   ' Filename = "" 'get temp filename
   ' size = 0
End Sub

Public Sub LoadFromMemory(bArray As Variant) 'added dzzie for small files in viewer mode

    If TypeName(bArray) <> "Byte()" Then
        MsgBox "Chunk.LoadFromMemory wrong data type to load!", vbExclamation
        Exit Sub
    End If
    
    Offset = 0
    mData() = bArray
    size = UBound(mData) + 1
    ReDim mStatus(size)
    mLoaded = True
    mInitialized = True
    mIsDirty = False
    
End Sub


Public Sub DataScreen(ByVal Index As Long, ByVal Count As Long, ByVal ArrPos As Long, ByRef Fillcount As Long, ByRef buff() As Byte)
    Dim i As Long
    Dim j As Long
    Dim Max As Long
    j = Index - Offset
    Max = Count - 1
    If Index + Count > Me.Offset + Me.size Then
        Max = Me.Offset + Me.size - Index - 1
    End If
    For i = 0 To Max
        buff(i + ArrPos) = mData(j + i)
    Next
    Fillcount = Max + 1
End Sub

Public Sub StatusScreen(ByVal Index As Long, ByVal Count As Long, ByVal ArrPos As Long, ByRef Fillcount As Long, ByRef buff() As Byte)
    Dim i As Long
    Dim j As Long
    Dim Max As Long
    j = Index - Offset
    Max = Count - 1
    If Index + Count > Me.Offset + Me.size Then
        Max = Me.Offset + Me.size - Index - 1
    End If
    For i = 0 To Max
        buff(i + ArrPos) = mStatus(j + i)
    Next
    Fillcount = Max + 1
End Sub

Public Property Get Data(ByVal Index As Long) As Byte
    Data = mData(Index - Offset)
End Property

Public Property Let Data(ByVal Index As Long, vData As Byte)
    mData(Index - Offset) = vData
    mIsDirty = True
End Property

Public Property Get Status(ByVal Index As Long) As Byte
    Status = mStatus(Index - Offset)
End Property

Public Property Let Status(ByVal Index As Long, vData As Byte)
    mStatus(Index - Offset) = vData
End Property

Public Sub Load()
    Dim buffer() As Byte
    Dim id As Long
    id = FreeFile

    If mLoaded = True Then Exit Sub
    If mInitialized = False Then
        'this chunk has not been read yet
            ReDim buffer(size - 1)
            Open Filename For Binary As #id
            Get #id, Offset + 1, buffer
            Close #id
            
         '   MsgBox UBound(buffer)
            mData = buffer
            size = UBound(mData) + 1
            ReDim mStatus(size) ' clear and size the status buffer
            mIsDirty = False
        'do read from original file
            mInitialized = False 'the original filedata is in ram now
    Else
        'this chunk has a tempfile
            Open DataName For Binary As #id
            ReDim mData(LOF(id) - 1)
            ReDim mStatus(LOF(id) - 1)
            
            Get #id, 1, mData
            Close #id
            
            Open StatusName For Binary As #id
            Get #id, 1, mStatus
            Close #id
            
            mIsDirty = False
            'MsgBox "Loading chunk"
        'do read from temp file
    End If
    mLoaded = True  'yepp ,the data is in the ram now
End Sub

Public Sub UnLoad()
    If mIsDirty Then
        DataName = MakeName
        StatusName = MakeName
        Call SaveChunk
        mInitialized = True
    End If
    mLoaded = False
    Erase mData
    Erase mStatus
End Sub

Private Function MakeName() As String
    Dim i As Long
    Dim name As String
    Randomize
    For i = 0 To 5
        name = name & Chr(Rnd * 20 + 65)
    Next
    name = name & Int(Timer)
    name = App.path & "\" & name & ".tmp"
    MakeName = name
End Function
        
Private Sub SaveChunk()
    'save dirty post
    Dim id As Long
    On Error Resume Next
    'remove files if already existing
    Kill DataName
    Kill StatusName
    Err.Clear
    On Error GoTo 0
    id = FreeFile
    Open DataName For Binary As #id
        Put #id, , mData()
    Close #id
    id = FreeFile
    Open StatusName For Binary As #id
        Put #id, , mStatus()
    Close #id
    
End Sub

Private Sub Class_Terminate()
    'remove temp files
    If mInitialized = True Then
        On Error Resume Next
        Kill DataName
        Kill StatusName
        Err.Clear
        On Error GoTo 0
    End If
    
End Sub


Public Sub InsertData(ByVal Pos As Long, Data() As Byte)
    Dim tmp As String
    Dim ltmp As String
    Dim rtmp As String
    Dim insert As String
    Dim insertstatus As String
    
    '//safety .. bailout if pos is less than 0
    If Pos < 0 Then Exit Sub
    
    
    If UBound(mData) = -1 Then
        'if chunk is empty ... replace it with this data
        mData = Data
        insertstatus = String(UBound(mData) + 1, Chr(1))
        mStatus = StrConv(insertstatus, vbFromUnicode)
        mIsDirty = True
        size = UBound(mData) + 1
        Exit Sub
    End If
    
    
    Pos = Pos - Offset

    insert = StrConv(Data(), vbUnicode)
    tmp = StrConv(mData, vbUnicode)
    ltmp = Mid(tmp, 1, Pos)
    rtmp = Mid(tmp, Pos + 1)
    tmp = ltmp & insert & rtmp
    mData = StrConv(tmp, vbFromUnicode)

    insertstatus = String(Len(insert), Chr(1))
    tmp = StrConv(mStatus, vbUnicode)
    ltmp = Mid(tmp, 1, Pos)
    rtmp = Mid(tmp, Pos + 1)
    tmp = ltmp & insertstatus & rtmp
    mStatus = StrConv(tmp, vbFromUnicode)
    mIsDirty = True
    
    'increase chunksize
    size = size + UBound(Data) + 1
End Sub

Public Sub InsertDataStatus(ByVal Pos As Long, Data() As Byte, Status() As Byte)
    Dim tmp As String
    Dim ltmp As String
    Dim rtmp As String
    Dim insert As String
    Dim insertstatus As String
    Pos = Pos - Offset

    insert = StrConv(Data(), vbUnicode)
    tmp = StrConv(mData, vbUnicode)
    ltmp = Mid(tmp, 1, Pos)
    rtmp = Mid(tmp, Pos + 1)
    tmp = ltmp & insert & rtmp
    mData = StrConv(tmp, vbFromUnicode)

    insertstatus = StrConv(Status(), vbUnicode)
    tmp = StrConv(mStatus, vbUnicode)
    ltmp = Mid(tmp, 1, Pos)
    rtmp = Mid(tmp, Pos + 1)
    tmp = ltmp & insertstatus & rtmp
    mStatus = StrConv(tmp, vbFromUnicode)
    mIsDirty = True
    
    'increase chunksize
    size = size + UBound(Data) + 1
End Sub


Public Sub DeleteData(ByVal Pos As Long, ByVal Length As Long)
   Dim tmp As String
    Dim ltmp As String
    Dim rtmp As String
    Dim insert As String
    Dim insertstatus As String
    Pos = Pos - Offset

    If Pos - 1 + Length > size Then Exit Sub

    tmp = StrConv(mData, vbUnicode)
    ltmp = Mid(tmp, 1, Pos)
    rtmp = Mid(tmp, Pos + 1 + Length)
    tmp = ltmp & rtmp
    mData = StrConv(tmp, vbFromUnicode)

    tmp = StrConv(mStatus, vbUnicode)
    ltmp = Mid(tmp, 1, Pos)
    rtmp = Mid(tmp, Pos + 1 + Length)
    tmp = ltmp & rtmp
    mStatus = StrConv(tmp, vbFromUnicode)
    mIsDirty = True
    
    size = size - Length
End Sub

Public Property Get IsInitialized() As Boolean
    IsInitialized = mInitialized
End Property

Public Sub Save()
    Me.Load
    Dim id As Long
    id = FreeFile
    Open Filename For Binary As #id
    Put #id, Offset + 1, mData
    Close #id
    Me.UnLoad
End Sub

Public Sub SaveNew()
    Me.Load
    Dim id As Long
    id = FreeFile
    Open NewFilename For Binary As #id
    Put #id, Offset + 1, mData
    Close #id
    Me.UnLoad
End Sub

