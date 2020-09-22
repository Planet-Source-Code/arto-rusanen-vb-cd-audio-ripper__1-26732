Attribute VB_Name = "modAKRip32_2"
Option Explicit

' How many times we try to read CD before we report error...
Public Const RetriesCount = 3
Public Const NumOfFramesPerRead = 26

Public SelHA As Byte
Public SelTGT As Byte
Public SelLUN As Byte
Public CDHandle As Long
Public DiscToc As TOC

Public Function GetDriveInformation(DriveNo As Long) As CDREC
  Dim CD_List As CDLIST
  GetCDList CD_List
  GetDriveInformation = CD_List.cd(DriveNo)
End Function

Public Function InitCDDrive(DriveNo As Long) As Boolean
  Dim CD_Info As CDREC
  Dim CDH As GETCDHAND
  
  ' Get Info
  CD_Info = GetDriveInformation(DriveNo)
  
  SelHA = CD_Info.ha
  SelTGT = CD_Info.tgt
  SelLUN = CD_Info.lun
  
  CDH.Size = Len(CDH)
  CDH.ver = 1
  CDH.ha = SelHA
  CDH.tgt = SelTGT
  CDH.lun = SelLUN
  CDH.ReadType = CDR_ANY '      // set for autodetect
  
  ' Get Handle
  CDHandle = GetCDHandle(ByVal VarPtr(CDH))
  If CDHandle = 0 Then
    InitCDDrive = False
    Exit Function
  End If
  
  Call ModifyCDParms(CDHandle, CDP_MSF, False)
  
  ' Read Table of Contents
  If ReadTOC(CDHandle, ByVal VarPtr(DiscToc)) <> 1 Then
    InitCDDrive = False
    Exit Function
  End If
  
  InitCDDrive = True
End Function

Public Function RipTrack(addrStart As Long, addrEnd As Long, FileName As String)
  Dim StartAddr   As Long
  Dim EndAddr     As Long
  Dim NumFrames   As Long
  Dim Dummy       As PTRACKBUF
  Dim BufferPtr1  As Long
  Dim BufferPtr2  As Long
  Dim LLen        As Long
  Dim Retries     As Long
  Dim Status      As Long
  Dim NumWritten  As Long
  Dim OpenFile As clsFileIo
  
  NumFrames = NumOfFramesPerRead
   
  ' Convert Addresses
  StartAddr = addrStart
  EndAddr = addrEnd
  
  'Initialize buffer
  BufferPtr1 = GlobalAlloc(&H40, NumFrames * 2352 + Len(Dummy))
  BufferPtr2 = GlobalLock(BufferPtr1)
  
  Dummy.startFrame = 0
  Dummy.NumFrames = 0
  Dummy.maxLen = NumFrames * 2352
  Dummy.len = 0
  Dummy.Status = 0
  Dummy.startOffset = 0
  
  memcpy ByVal BufferPtr2, ByVal VarPtr(Dummy), Len(Dummy)
  
  Dim Temp As Long
  Temp = EndAddr - StartAddr
  LLen = EndAddr - StartAddr
  
  ' Open files
  Set OpenFile = New clsFileIo
  
  OpenFile.OpenFile FileName
  OpenFile.writeWavHeader LLen * 2352
  
  Dim TempCount As Byte
  
  ' Lets start rippin...
  Do While LLen
    ' Calculate how much we wanna rip from CD
    If LLen < NumFrames Then NumFrames = LLen
      
    Retries = RetriesCount
    Status = 0
    
    ' Try to read cd...
    Do While Retries > 0 And Status <> 1
      Dummy.NumFrames = NumFrames
      Dummy.startOffset = 0
      Dummy.len = 0
      Dummy.startFrame = StartAddr
      
      'Write info to buffer so that akrip knows what to read... :)
      memcpy ByVal BufferPtr2, ByVal VarPtr(Dummy), Len(Dummy)
      
      Status = ReadCDAudioLBA(CDHandle, BufferPtr2)
    Loop
    
    If Status = 1 Then
      ' Write buffer to disk
      ' Note: Don't write info to disk... Memory position is pointer + lenght of Dummy
      OpenFile.WriteBytes BufferPtr2 + Len(Dummy), NumFrames * 2352
    Else
      ' Doh.... This is bad.... and there is nothing we can do...
      MsgBox GetAKRipError
      Exit Do
    End If
    
    ' We have written this much bytes and blahblahblahblah.... :)
    NumWritten = NumWritten + NumFrames * 2352
    StartAddr = StartAddr + NumFrames
    LLen = LLen - NumFrames
    
    'Inform user where we go...
    Form1.Status.Caption = Format((Temp - LLen) / Temp * 100, "00.00") & " %"
    DoEvents
  Loop
  
  ' Delete buffer and close files
  GlobalFree BufferPtr2
  'OpenFile.writeWavHeader NumWritten
  OpenFile.CloseFile
  Set OpenFile = Nothing
  
End Function

' Lil helper... Hope I converted it right...
Public Function MSB2LONG(b() As Byte) As Long
  MSB2LONG = CLng(b(1)) * 256 * 256 * 256
  MSB2LONG = MSB2LONG + CLng(b(2)) * 256 * 256
  MSB2LONG = MSB2LONG + CLng(b(3)) * 256
  MSB2LONG = MSB2LONG + CLng(b(4))
End Function

' Close CD Drive
Public Function DeInitCDDrive() As Boolean
  DeInitCDDrive = CloseCDHandle(CDHandle)
End Function

' This one tells what went wrong....
Public Function GetAKRipError() As String
  Dim ErrNo As AKErrorEnum
  Err.Clear
  ErrNo = GetAspiLibError
  If Err Then Debug.Assert 0
  Select Case ErrNo
  Case ALERR_NOERROR: GetAKRipError = "No error..."
  Case ALERR_NOWNASPI: GetAKRipError = "Unable to load WNASPI32.DLL (95/98/NT/2000) or use SCSI passthrough (NT/2000)"
  Case ALERR_NOGETASPI32SUPP: GetAKRipError = "Could not load ASPI function GetASPI32SupportInfo. Only occurs when an ASPI manager (WNASPI32.DLL) is found, but cannot be correctly loaded by the system. Most often indicates that ASPI is improperly installed."
  Case ALERR_NOSENDASPICMD: GetAKRipError = "Could not load ASPI function SendASPI32Command. Only occurs when an ASPI manager (WNASPI32.DLL) is found, but cannot be correctly loaded by the system. Most often indicates that ASPI is improperly installed."
  Case ALERR_ASPI: GetAKRipError = "An error was returned by the ASPI manager. Use GetAspiLibAspiError to retrieve the actual ASPI error."
  Case ALERR_NOCDSELECTED: GetAKRipError = "Unused in the current implementation"
  Case ALERR_BUFTOOSMALL: GetAKRipError = "The buffer passed to ReadCDAudioLBA or ReadCDAudioLBAEx is too small for the requested number of frames."
  Case ALERR_INVHANDLE: GetAKRipError = "The handle to the CD-ROM unit is invalid."
  Case ALERR_NOMOREHAND: GetAKRipError = "All available slots for CD-ROM handles have been allocated."
  Case ALERR_BUFPTR: GetAKRipError = "Results from passing a bad buf parameter to GetCDId (ie. a NULL pointer)"
  Case ALERR_NOTACD: GetAKRipError = "The ha:tgt:lun values passed to GetCDHandle do not refer to a CD-ROM device."
  Case ALERR_LOCK: GetAKRipError = "Unable to obtain an exclusive lock on a CD handle."
  Case ALERR_DUPHAND: GetAKRipError = "Occurs when attempting to call GetCDHandle for a ha:tgt:lun value that has already been allocated."
  Case ALERR_INVPTR: GetAKRipError = "Invalid value for the LPGETCDHAND parameter to GetCDHandle"
  Case ALERR_INVPARM: GetAKRipError = "Invalid version or size specified in LPGETCDHAND parameter to GetCDHandle"
  Case ALERR_JITTER: GetAKRipError = "An automatic jitter adjust failed during a call to ReadCDAudioLBAEx"
  End Select
End Function

