VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2 undoc functions"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdvbaCopyBytesZero 
      Caption         =   "vbaCopyBytesZero"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdvbaCopyBytes 
      Caption         =   "vbaCopyBytes"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Undocumented Memory Functions
' -----------------------------
'
' There are 2 undocumented "memory functions" in msvbvm60.dll -> __vbaCopyBytes and __vbaCopyBytesZero.
'
' __vbaCopyBytes     -> this function copies bytes from one memory location to another memory location
' __vbaCopyBytesZero -> this function copies bytes from one memory location to another memory location and then source memory is filled with zeros
'
' Both functions have 3 parameters
' 1.param -> number of bytes to copy
' 2.param -> pointer to the destination memory
' 3.param -> pointer to the source memory
'
' That's all you need to know ;-)
' btw if you know ASM and want to know how __vbaCopyBytes and __vbaCopyBytesZero work, then check their ASM code below.
'
' Enjoy,
'   Libor
'

'
' ---> __vbaCopyBytes <---
'
'public __vbaCopyBytes
'.text:660D9B6B __vbaCopyBytes  proc near
'.text:660D9B6B
'.text:660D9B6B arg_0           = dword ptr  4              ;1.param - number of bytes to copy
'.text:660D9B6B arg_4           = dword ptr  8              ;2.param - pointer to the destination memory
'.text:660D9B6B arg_8           = dword ptr  0Ch            ;3.param - pointer to the source memory
'.text:660D9B6B
'.text:660D9B6B                 mov     ecx, [esp+arg_0]    ;ecx = 1.param
'.text:660D9B6F                 push    esi                 ;save esi
'.text:660D9B70                 mov     esi, [esp+4+arg_8]  ;esi = 3.param
'.text:660D9B74                 push    edi                 ;save edi
'.text:660D9B75                 mov     edi, [esp+8+arg_4]  ;edi = 2.param
'.text:660D9B79                 mov     eax, ecx            ;eax = ecx
'.text:660D9B7B                 mov     edx, edi            ;edx = edi
'.text:660D9B7D                 shr     ecx, 2              ;ecx = ecx/2
'.text:660D9B80                 rep movsd                   ;perform dword copy
'.text:660D9B82                 mov     ecx, eax            ;ecx = eax
'.text:660D9B84                 mov     eax, edx            ;eax = edx
'.text:660D9B86                 and     ecx, 3              ;ecx = ecx mod 4
'.text:660D9B89                 rep movsb                   ;perform byte copy
'.text:660D9B8B                 pop     edi                 ;restore edi
'.text:660D9B8C                 pop     esi                 ;restore esi
'.text:660D9B8D                 retn    0Ch                 ;return
'.text:660D9B8D
'.text:660D9B8D __vbaCopyBytes  endp
'

'
' ---> __vbaCopyBytesZero <---
'
'public __vbaCopyBytesZero
'.text:660D9B90 __vbaCopyBytesZero proc near
'.text:660D9B90
'.text:660D9B90 arg_0           = dword ptr  8              ;1.param - number of bytes to copy
'.text:660D9B90 arg_4           = dword ptr  0Ch            ;2.param - pointer to the destination memory
'.text:660D9B90 arg_8           = dword ptr  10h            ;3.param - pointer to the source memory
'.text:660D9B90
'.text:660D9B90                 push    ebp                 ;save ebp
'.text:660D9B91                 mov     ebp, esp            ;ebp = esp
'.text:660D9B93                 mov     ecx, [ebp+arg_0]    ;ecx = 1.param
'.text:660D9B96                 push    esi                 ;save edi
'.text:660D9B97                 mov     esi, [ebp+arg_8]    ;esi = 3.param
'.text:660D9B9A                 mov     eax, ecx            ;eax = ecx
'.text:660D9B9C                 push    edi                 ;save edi
'.text:660D9B9D                 mov     edi, [ebp+arg_4]    ;edi = 2.param
'.text:660D9BA0                 shr     ecx, 2              ;ecx = ecx/4
'.text:660D9BA3                 rep movsd                   ;perform dword copy
'.text:660D9BA5                 mov     ecx, eax            ;ecx = eax
'.text:660D9BA7                 and     ecx, 3              ;ecx = ecx mod 4
'.text:660D9BAA                 rep movsb                   ;perform byte copy
'.text:660D9BAC                 mov     edi, [ebp+arg_8]    ;edi = 3.param
'.text:660D9BAF                 mov     ecx, eax            ;ecx = eax
'.text:660D9BB1                 mov     edx, ecx            ;edx = ecx
'.text:660D9BB3                 xor     eax, eax            ;eax = 0
'.text:660D9BB5                 shr     ecx, 2              ;ecx = ecx/4
'.text:660D9BB8                 rep stosd                   ;store eax to [edi]
'.text:660D9BBA                 mov     ecx, edx            ;ecx = edx
'.text:660D9BBC                 and     ecx, 3              ;ecx = ecx mod 4
'.text:660D9BBF                 rep stosb                   ;store eax to [edi]
'.text:660D9BC1                 mov     eax, [ebp+arg_4]    ;eax = 2.param
'.text:660D9BC4                 pop     edi                 ;restore edi
'.text:660D9BC5                 pop     esi                 ;restore esi
'.text:660D9BC6                 pop     ebp                 ;restore ebp
'.text:660D9BC7                 retn    0Ch                 ;return
'.text:660D9BC7
'.text:660D9BC7 __vbaCopyBytesZero endp
'

'declare the undoc functions
Private Declare Sub vbaCopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, Destination As Any, Source As Any)
Private Declare Sub vbaCopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyBytesZero" (ByVal Length As Long, Destination As Any, Source As Any)

'test __vbaCopyBytesZero
Private Sub cmdvbaCopyBytesZero_Click()
Dim a(1 To 15) As Byte  'source
Dim b(1 To 15) As Byte  'destination
Dim s As String, i As Long

    'fill the first array
    For i = LBound(a) To UBound(a)
        a(i) = i
        s = s & vbCrLf & i & ". a = " & a(i) & " ; b = " & b(i)
    Next i

    MsgBox "Before __vbaCopyBytes" & s
    vbaCopyBytesZero LenB(a(1)) * 15, b(1), a(1)
    
    s = ""

    'check the second array
    For i = LBound(a) To UBound(a)
        s = s & vbCrLf & i & ". a = " & a(i) & " ; b = " & b(i)
    Next i
    
    MsgBox "After __vbaCopyBytesZero" & s
    
End Sub

'test __vbaCopyBytes
Private Sub cmdvbaCopyBytes_Click()
Dim a(1 To 15) As Byte  'source
Dim b(1 To 15) As Byte  'destination
Dim s As String, i As Long

    'fill the first array
    For i = LBound(a) To UBound(a)
        a(i) = i
        s = s & vbCrLf & i & ". a = " & a(i) & " ; b = " & b(i)
    Next i

    MsgBox "Before __vbaCopyBytes" & s
    vbaCopyBytes LenB(a(1)) * 15, b(1), a(1)
    s = ""

    'check the second array
    For i = LBound(a) To UBound(a)
        s = s & vbCrLf & i & ". a = " & a(i) & " ; b = " & b(i)
    Next i
    
    MsgBox "After __vbaCopyBytes" & s

End Sub
