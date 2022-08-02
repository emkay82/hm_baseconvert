' Author: emkay82
'
' Note:
' The qr-codes mentioned here contain both the SGTIN
' as well as the KEY. Unlike the readable KEY, the KEY
' in the qr-code is regarded as base-16.
'
' SGTIN starts at position 7* and is 24 characters long
' KEY starts at position 34* and is 32 characters long
'
' * assuming the first character = position 1
'
' EQ01SG                        DLK
'       0123456789ABCEFGHJKLMNPQ   0123456789ABCDEF0123456789ABCDEF
'                   ^                              ^
'                 SGTIN                       KEY base-16
'
' Change History:
' 19.07.2022 - emkay82: Initial release

Option Explicit



'--------------------------------------------------------------------
' Name:     ConvertHmIPKeyFromB32ToB16
'           ==========================
'
' Converts the readable key (base-32) of a HmIP device
' into the one used in the qr-code (base-16).
'
' Homematic uses a base-32 alphabet, which differs from RFC 4648.
' Since only the first 13 characters are identical to those of
' the base-16 alphabet, an intermediate step via a base smaller
' than 14 is required.
'
' Parameter:
'   value             HmIP key (base-32) (String)
'
' Returns:
'   String            HmIP key (base-16)
'--------------------------------------------------------------------

Function ConvertHmIPKeyFromB32ToB16(value)
  Dim result: result = BaseConvert(value, 32, 10, "0123456789ABCEFGHJKLMNPQRSTUWXYZ")
  result = BaseConvert(result, 10, 16, "0123456789ABCDEF")

  If Len(result) Mod 2 <> 0 Then
    result = "0" & result
  End If

  ConvertHmIPKeyFromB32ToB16 = result
End Function



'--------------------------------------------------------------------
' Name:     ConvertHmIPKeyFromB16ToB32
'           ==========================
'
' Converts the key used in the qr-code (base-16)
' into the a readable KEY (base-32).
'
' Homematic uses a base-32 alphabet, which differs from RFC 4648.
' Since only the first 13 characters are identical to those of
' the base-16 alphabet, an intermediate step via a base smaller
' than 14 is required.
'
' Parameter:
'   value             HmIP key (base-16) (String)
'
' Returns:
'   String            HmIP key (base-32)
'--------------------------------------------------------------------

Function ConvertHmIPKeyFromB16ToB32(value)
  Dim result: result = BaseConvert(value, 16, 10, "0123456789ABCDEF")
  result = BaseConvert(result, 10, 32, "0123456789ABCEFGHJKLMNPQRSTUWXYZ")

  If Len(result) Mod 2 <> 0 Then
    result = "0" & result
  End If

  ConvertHmIPKeyFromB16ToB32 = InsertSeparators(result, "-", "6, 12, 18, 24")
End Function



'--------------------------------------------------------------------
' Name:     BaseConvert
'           ===========
'
' Converts a value from any base to any base.
'
' Parameter:
'   value             Input in base of base_f (String)
'   base_f            Base of the source (Integer)
'   base_t            Desired target base (Integer)
'   alphabet          Base alphabet to be used (String)
'
' Returns:
'   String            Result in base of base_t
'--------------------------------------------------------------------

Function BaseConvert(value, base_f, base_t, alphabet)
  Dim collect: collect = Array(0)

  Dim j
  For j = 1 To Len(value)
    Dim tmp: tmp = InStr(alphabet, Mid(value, j, 1)) - 1

    Dim i: i = 0
    Do Until i = UBound(collect) + 1 And tmp = 0
      If i > UBound(collect) Then
        ReDim Preserve collect(i)
      End If

      Dim nxt: nxt = (collect(i) Or 0) * base_f + tmp
      collect(i) = nxt Mod base_t
      tmp = Int(nxt / base_t)

      i = i + 1
    Loop
  Next

  Dim result: result = ""
  For j = UBound(collect) To 0 Step -1
    result = result & Mid(alphabet, collect(j) + 1, 1)
  Next

  BaseConvert = result
End Function



'--------------------------------------------------------------------
' Name:     InsertSeparators
'           ================
'
' Inserts separators at specified positions.
'
' Parameter:
'   value             String to be modified (String)
'   character         Single character which is inserted as a seperator (String)
'   positions         Commaseparated values used for position definition (String)
'
' Returns:
'   String            Modified string with separators
'--------------------------------------------------------------------

Function InsertSeparators(value, character, positions)
  Dim tmp: tmp = Split(Replace(positions, " ", ""), ",")
  Dim result: result = value

  Dim i
  For i = 0 To UBound(tmp)
    result = Mid(result, 1, tmp(i) - 1) & character & Mid(result, tmp(i))
  Next

  InsertSeparators = result
End Function