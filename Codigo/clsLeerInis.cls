VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeerInis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.11.20
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

'Modulo para optimizar la lectura de Archivos en
'formato INI
'
'Creado por Alejandro Santos
'

Option Explicit

Public Enum tModoLeeIni

    MODO_CACHEAR = 1    'Usa el nuevo sistema de LeeInis-O-Matico
    MODO_API = 2        'Usa el viejo sistema de GetPrivateProfileString

End Enum

Private mSecciones As Collection
Private mArch      As String
Private mModo      As tModoLeeIni

Private Declare Function GetPrivateProfileString _
                Lib "kernel32" _
                Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                  ByVal lpKeyname As Any, _
                                                  ByVal lpdefault As String, _
                                                  ByVal lpreturnedstring As String, _
                                                  ByVal nsize As Long, _
                                                  ByVal lpfilename As String) As Long

Public Sub Abrir(Arch As String, Optional Modo As tModoLeeIni = MODO_CACHEAR)

    If Dir(Arch, vbNormal) = "" Then Exit Sub

    mModo = Modo
    mArch = Arch
    Set mSecciones = Nothing

    Select Case mModo

        Case MODO_CACHEAR
            Call AbrirCacheado(Arch)

        Case MODO_API

            ' ???
    End Select

End Sub

Public Function DarValor(ByVal Secc As String, ByVal Clave As String) As String

    Select Case mModo

        Case MODO_CACHEAR
            DarValor = DarValorCacheado(Secc, Clave)

        Case MODO_API
            DarValor = DarValorAPI(Secc, Clave)

    End Select

End Function

Private Sub AbrirCacheado(ByVal Arch As String)
    Dim ArchNum    As Integer
    Dim Linea      As String
    Dim Pos        As Long

    Dim Seccion    As String, ClaveValor() As String
    Dim ColSeccion As Collection

    ArchNum = FreeFile()

    Set mSecciones = Nothing
    Set mSecciones = New Collection

    On Error GoTo hayerror

    Open Arch For Input As #ArchNum

    On Error GoTo 0

    Do While Not EOF(ArchNum)
        Line Input #ArchNum, Linea

        If Left(Linea, 1) = "[" And Len(Linea) > 2 Then
            Pos = InStr(1, Linea, "]")

            If Pos > 2 Then
                Seccion = UCase(Trim(mid(Linea, 2, Pos - 2)))
            Else
                Seccion = ""

            End If

        ElseIf Seccion <> "" Then
            ClaveValor = Split(Linea, "=", 2)

            If UBound(ClaveValor) - LBound(ClaveValor) = 1 Then
                On Local Error Resume Next
                Set ColSeccion = mSecciones(Seccion)

                If Err Then 'La seccion no existe
                    Set ColSeccion = New Collection
                    mSecciones.Add ColSeccion, Seccion

                End If

                On Local Error GoTo 0
            
                'Malditos INIS, formato pedorrooooooo
                ClaveValor(LBound(ClaveValor)) = UCase(Trim(ClaveValor(LBound(ClaveValor))))
                ClaveValor(UBound(ClaveValor)) = Replace(ClaveValor(UBound(ClaveValor)), """", "")
            
                On Local Error Resume Next
                ColSeccion.Add ClaveValor(UBound(ClaveValor)), ClaveValor(LBound(ClaveValor))
                '            If Err Then
                '                ColSeccion.Remove ClaveValor(UBound(ClaveValor))
                '                ColSeccion.Add ClaveValor(UBound(ClaveValor)), ClaveValor(LBound(ClaveValor))
                '            End If
                On Local Error GoTo 0
            
                Set ColSeccion = Nothing

            End If

        End If

    Loop

    Close #ArchNum
    Exit Sub
hayerror:

End Sub

Private Function DarValorCacheado(ByVal Secc As String, ByVal Clave As String) As String
    Dim Ret As String

    DarValorCacheado = ""
    On Local Error Resume Next
    DarValorCacheado = mSecciones(UCase(Secc)).Item(UCase(Clave))
    On Local Error GoTo 0

End Function

Private Function DarValorAPI(ByVal Secc As String, ByVal Clave As String) As String

    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
  
    szReturn = ""
  
    sSpaces = Space(5000) ' This tells the computer how long the longest string can be

    GetPrivateProfileString Secc, Clave, szReturn, sSpaces, Len(sSpaces), mArch
  
    sSpaces = RTrim(sSpaces)
    DarValorAPI = Left$(sSpaces, Len(sSpaces) - 1)

End Function
