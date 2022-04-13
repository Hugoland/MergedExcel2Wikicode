' By Hugo7 - hugoland.net
' You can reuse this code in your non-commercial projects.
' Released under Creative Commons BY NC (Attribution (credit) Non-Commercial)
' https://creativecommons.org/licenses/by-nc/2.0/

Option Explicit
Public Const crlf = vbNewLine

Public Sub excel2wikicode()
    Dim out As String
    Dim ligne As Long
    Dim col As Long
    Dim cell As Range
    Dim i As Long
    Dim starting_line As Long
    
    out = "{|class=""wikitable""" + crlf 'doubled double-quote is required to escape this character (like "" -> ")
    For Each cell In Selection
        ligne = cell.Row 'yeah, "ligne" is the French for "line", as "line" is not a correct name for a variable in VBA
        If (cell.Column < col And ligne <> starting_line) Then out = out & "|-" & crlf 'compares new value of col to it's previous val -- adds a Return wikicode statement
        col = cell.Column 'then we can update it
        starting_line = Split(Split(Selection.Address, ":")(0), "$")(2)
        
        If (Split(cell.MergeArea.Address, ":")(0) = Cells(ligne, col).Address) Then 'if cell is in a group of merged cells and if it's not the main cell of that group (top lef)
            ' If first line then different starting character in wikicode
            If (ligne = starting_line) Then
                out = out & "!"
            Else
                out = out & "|"
            End If
            
            ' Loops cells until it finds a cell which is not merged with the first (two FOR : going to the right (next n cols), then going down (next n lines))
            i = 0
            Do
                i = i + 1
                If (Split(Cells(ligne, col + i).MergeArea.Address, ":")(0) <> Cells(ligne, col).Address) Then Exit Do
            Loop
            out = out & "colspan=" & i & " "
            
            i = 0
            Do
                i = i + 1
                If (Split(Cells(ligne + i, col).MergeArea.Address, ":")(0) <> Cells(ligne, col).Address) Then Exit Do
            Loop
            
            'If (cell.Value = 0) Then  'uncomment these four commented lines if you want all non-zero cells having a background color (I needed that once so here it is... adapt as you wish)
                out = out & "rowspan=" & i & " " '/!\ never comment this line
            'Else
            '    out = out & "rowspan=" & i & " style=""background:#E0F2CE;""|"
            'End If
            
            'then adds style and value
            If (cell.Value = "-") Then cell.Value = "—" 'changing type of hyphen as cells containing only a hyphen is "|-" therefore breaks the whole wikitable
            out = out & "style=""background:" & hexcolor(cell) & """|" & cell.Value & crlf
        Else
            'do nothing !
        End If
    Next cell
    
    out = out & "|}"
    'MsgBox out
    
    'clears parameters that are not necessary when equals to 1 or if color is FFFFFF :
    out = Replace(out, "colspan=1 ", "")
    out = Replace(out, "rowspan=1 ", "")
    out = Replace(out, "style=""background:#FFFFFF""", "")
    out = Replace(out, "||", "|") 'if all above parameters are removed, it would leave a ||, so removing one |
    out = Replace(out, "!|", "!") 'same but for first line
    out = Replace(out, " |", "|") 'if style is removed but not col/rowspan, there's an extra space we may remove
    
    ActiveWorkbook.Sheets.Add
    With ActiveSheet.Range("A1")
        .Value = out
        .Columns.AutoFit
        .Rows.AutoFit
    End With
End Sub

Public Function hexcolor(cell As Range) As String
    Dim c As Long
    c = cell.Interior.Color
    
    'convert VBA Long color code to HEX color code for HTML support
    With WorksheetFunction
        hexcolor = "#" & add_zero(.Dec2Hex(c Mod 256)) & add_zero(.Dec2Hex(c \ 256 Mod 256)) & add_zero(.Dec2Hex(c \ 65536 Mod 256))
    End With
End Function

Private Function add_zero(val As String) As String
    'Adds a zero if necessary, as Dec2Hex function returns numbers on 1 digit for values hex < F (dec < 16) so this function does 1 -> 01 to F -> 0F
    If (Len(val) = 1) Then
        add_zero = "0" & val
    Else
        add_zero = val
    End If
End Function
