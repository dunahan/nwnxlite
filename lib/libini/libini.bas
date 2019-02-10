Attribute VB_Name = "libini"
'***************************************************************************
'                         libini.bas -  Visual Basic access functions
'                            -------------------
'   begin                : Fri Apr 21 2000
'   copyright            : (C) 2000 by Simon White
'   email                : s_a_white@email.com
'***************************************************************************

'***************************************************************************
'*                                                                         *
'*   This program is free software; you can redistribute it and/or modify  *
'*   it under the terms of the GNU General Public License as published by  *
'*   the Free Software Foundation; either version 2 of the License, or     *
'*   (at your option) any later version.                                   *
'*                                                                         *
'***************************************************************************
Option Explicit

' Open a existing file.  Create one if it dosen't exist.
' Returns 0 for error or a ini file descriptor for passing to
' the other functions.
Declare Function ini_new Lib "libini" (ByVal name As Any) As Long

' Open a existing file only.  Fail if one dosen't exist.
' Returns 0 for error or a ini file descriptor for passing to
' the other functions.
Declare Function ini_open Lib "libini" (ByVal name As Any) As Long

' Returns -1 on error.
Declare Function ini_close Lib "libini" (ByVal ini_fd As Long) As Integer

' Saves changes to the INI file.
' Returns -1 on error.
Declare Function ini_flush Lib "libini" (ByVal ini_fd As Long) As Integer

' To create a new key call ini_locateHeading, ini_locateKey.  Although these
' calls fail, what it actually inidcates is that these things don't exist.
' Check the return value from the write operation to see if the new key was
' created.

' All return -1 on error.
Declare Function ini_locateKey Lib "libini" _
    (ByVal ini_fd As Long, ByVal key As String) As Integer
Declare Function ini_locateHeading Lib "libini" _
    (ByVal ini_fd As Long, ByVal heading As String) As Integer

' To delete a heading use ini_locateHeading then ini_deleteHeading.
' To delete a key use ini_locateHeading, ini_locateKey then ini_deleteKey
' The above can be performed the same as Micrsoft by:
'     ini_locateHeading then ini_writeString, etc

' All return -1 on error.
Declare Function ini_deleteKey Lib "libini" (ByVal ini_fd As Long) As Integer
Declare Function ini_deleteHeading Lib "libini" _
    (ByVal ini_fd As Long) As Integer

Private Declare Function ini_dataLength Lib "libini" _
    (ByVal ini_fd As Long) As Integer

Private Declare Function ini_readStringFixed Lib "libini" _
    Alias "ini_readString" (ByVal ini_fd As Long, _
    ByVal str As String, ByVal size As Long) As Integer

Declare Function ini_writeString Lib "libini" (ByVal ini_fd As Long, _
    ByVal sString As String) As Integer

' All return -1 on error.
Declare Function ini_readInt Lib "libini" (ByVal ini_fd As Long, _
    ByRef value As Integer) As Integer
Declare Function ini_readLong Lib "libini" (ByVal ini_fd As Long, _
    ByRef value As Long) As Integer
Declare Function ini_readDouble Lib "libini" (ByVal ini_fd As Long, _
    ByRef value As Double) As Integer

Declare Function ini_writeInt Lib "libini" (ByVal ini_fd As Long, _
    ByVal value As Integer) As Integer
Declare Function ini_writeLong Lib "libini" (ByVal ini_fd As Long, _
    ByVal value As Long) As Integer
Declare Function ini_writeDouble Lib "libini" (ByVal ini_fd As Long, _
    ByVal value As Double) As Integer

' Use this to indicate the character sperating multiple key elements.  Only
' use this for keys which require it.  Disable it again by passing in 0& as
' the delimiter e.g.:
' key = data1, data2, data3
' ini_locateKey  (fd, "key")
' ini_listDelims (fd, ","); Enable
' ... do reads here.
' ini_listDelims (fd, 0&); Disable

' Returns -1 on error, or the number of data elements found after applying
' the delimiters
Declare Function ini_ListLength Lib "libini" (ByVal ini_fd As Long) As Integer

' All return -1 on error
' Sets the character(s) used to seperate data elements in the key
Declare Function ini_ListDelims Lib "libini" (ByVal ini_fd As Long, _
    ByVal delims As String) As Integer
    
' Set index of the data element to get on next read.  This will auto incrment
' when a read is performed.  Do a key_locate will automatically reset the index
' to 0 (first item).
Declare Function ini_ListIndex Lib "libini" (ByVal ini_fd As Long, _
    ByVal index As Long) As Integer

' This returns -1 on error or the number of characters read
Function ini_readString(ini_fd As Long, str As String) As Integer
    ' Create a fixed buffer to read string
    Dim length As Long
    Dim ret As Long
    Dim buffer As String
    length = ini_dataLength(ini_fd)
    If (length < 0) Then
        GoTo ini_readString_error
    End If

    length = length + 1: ' Reserve place for NULL (chr$(0))
    buffer = String(length, Chr$(0))
    ret = ini_readStringFixed(ini_fd, buffer, length)
    If (ret < 0) Then
        GoTo ini_readString_error
    End If

    ' Remove C String termination character (NULL)
    str = Left(buffer, ret)
    ini_readString = ret
Exit Function

ini_readString_error:
    ini_readString = -1
End Function
