Attribute VB_Name = "basChess"
Option Explicit

Public Type Board
    X As Integer
    Y As Integer
End Type

Public Type ChessPiece
    Index As Integer
    Name As Integer
    side As Integer
    XY As Board
End Type

Public Pieces(31) As ChessPiece
Public Const TotalPieces = 32

Public KingIndex(-1 To 1) As Integer

Public Const NoPiece = 0
Public Const King = 1   ' General (King)
Public Const Castle = 2     ' Chariot (Rook)
Public Const Knight = 3     ' Horse
Public Const Rocket = 4     ' Cannon
Public Const Scholar = 5    ' Mandarin (Assistant)
Public Const Bishop = 6     ' Elephant
Public Const Pawn = 7   ' Soldier (Pawn)

Public Const TopSide = 1
Public Const BottomSide = -1

Public Const checkX = 0
Public Const checkY = 1
Public Const checkAcsend = 1
Public Const checkDesend = -1

Public Const Normal = 0
Public Const Killing = -1

Public ChessMaster(-1 To 1) As IAgentCtlCharacter
Public AgentAvail(-1 To 1) As Integer

Private Sub Main()
On Error Resume Next
    frmSplash.Show
    frmSplash.Refresh
    Load frmMain
    With frmMain
        .Agent1.Characters.Load "Merlin", "C:\WINDOWS\Msagent\CHARS\Merlin.acs"
        AgentAvail(TopSide) = IIf(Err, False, True)
        .Agent1.Characters.Load "Genie", "C:\WINDOWS\Msagent\CHARS\Genie.acs"
        AgentAvail(BottomSide) = IIf(Err, False, True)
    
        Set ChessMaster(TopSide) = .Agent1.Characters("Merlin")
        ChessMaster(TopSide).Description = TopSide
        Set ChessMaster(BottomSide) = .Agent1.Characters("Genie")
        ChessMaster(BottomSide).Description = BottomSide
        .Show
    End With
    
    Unload frmSplash
    
    With ChessMaster(TopSide)
        .Balloon.FontName = "Comic Sans MS"
        .MoveTo (frmMain.Left + frmMain.Width) / 16, frmMain.Top / 16, 0
        .Show
        DialogExp ChessMaster(TopSide), "Hello!", "Greet"
    End With
    With ChessMaster(BottomSide)
        .Balloon.FontName = "Comic Sans MS"
        .MoveTo (frmMain.Left + frmMain.Width) / 16, (frmMain.Top + frmMain.Height) / 16, 0
        .Show
    End With
End Sub

Function CBoard(Index As Integer) As Board
    CBoard.X = Int(Index / 10) + 1
    CBoard.Y = Right(Index, 1) + 1
End Function

Function CIndex(XY As Board) As Integer
    CIndex = (XY.X - 1) * 10 + XY.Y - 1
End Function

Function IsCheckmate(Final As Board, side As Integer) As Boolean
    Dim i As Integer, p As Integer
    Dim loopFrom As Integer, loopTo As Integer
    Dim tmpP As ChessPiece, tmpDestroyP As ChessPiece
    loopFrom = IIf(side = TopSide, 16, 0)
    loopTo = IIf(side = TopSide, 31, 15)
    '--Check OUR pieces if any can clear the check
    For p = loopFrom To loopTo
        '--If the piece has not been killed
        If Pieces(p).side = side Then
            '--Save OUR piece for later temporary change
            tmpP = Pieces(p)
Debug.Print "This piece is: "
Debug.Print Space(5) & "Index: " & tmpP.Index
Debug.Print Space(5) & "Name: " & tmpP.Name
Debug.Print Space(5) & "Side: " & tmpP.side
Debug.Print Space(5) & "XY: " & tmpP.XY.X & ", " & tmpP.XY.Y
            '--Check all the positions          <- an example of favoring size over speed, or lazy over speed
            For i = 0 To 89
                '--If not the same piece as OUR piece
                If CBoard(i).X <> tmpP.XY.X Or CBoard(i).Y <> tmpP.XY.Y Then
Debug.Print CBoard(i).X & ", " & CBoard(i).Y
                    '--If any, save other piece on the way
                    tmpDestroyP = GetFilled(CBoard(i).X, CBoard(i).Y)
                    '--If any, move away that piece first
                    If NoPiece <> tmpDestroyP.Name Then frmMain.Pieces_Destroy tmpDestroyP.Index
                    '--If OUR piece can move to that position
                    If IsLegalMove(Pieces(p), Pieces(p).XY, CBoard(i), tmpDestroyP) Then
                        '--Move to there
                        Pieces(p).XY = CBoard(i)
Debug.Print "Moveto: " & Pieces(p).XY.X & ", " & Pieces(p).XY.Y
                        '--If this move can clear the check
                        If Not IsCheck(Final, side) Then
                            IsCheckmate = False
Debug.Print "Not checkmate"
                            '--RESTORE and EXIT
                            '--If saved something, restore it.
                            If tmpDestroyP.Name <> NoPiece Then Pieces(tmpDestroyP.Index) = tmpDestroyP
                            '--Restore OUR piece
                            Pieces(p).XY = tmpP.XY
                            Exit Function
                        End If      '/backto/ if IsCheckmate...
                    End If      '/backto/ if IsLgalMove...
                    '--RESTORE and CONTINUE
                    '--If saved something, restore it.
                    If NoPiece <> tmpDestroyP.Name Then Pieces(tmpDestroyP.Index) = tmpDestroyP
                    '--Restore OUR piece
                    Pieces(p).XY = tmpP.XY
                End If      '/backto/ if Cboard(i)...
            Next
        End If      '/backto/ if Pieces(p).Side <> 0 ...
    Next
    IsCheckmate = True
Debug.Print "Checkmate"
End Function

Function IsCheck(Final As Board, side As Integer) As Boolean
    '--Check right
    IsCheck = IsCheck2_Straight(Final, checkY, checkAcsend, side)
    If IsCheck Then Exit Function
    '--Check left
    IsCheck = IsCheck2_Straight(Final, checkY, checkDesend, side)
    If IsCheck Then Exit Function
    '--Check up
    IsCheck = IsCheck2_Straight(Final, checkX, checkAcsend, side)
    If IsCheck Then Exit Function
    '--Check down
    IsCheck = IsCheck2_Straight(Final, checkX, checkDesend, side)
    If IsCheck Then Exit Function
    IsCheck = IsCheck2_Knight(Final, side)
End Function

Private Function IsCheck2_Straight(Final As Board, Way As Integer, Dir As Integer, side As Integer) As Boolean
Dim pos As Integer, tmp As Integer, tmp2 As Integer, tmpP As ChessPiece
Dim limit As Integer, limit2 As Integer
    pos = IIf(Way = checkX, Final.X, Final.Y)
    tmp = pos
    limit2 = IIf(Way = checkX, 9, 10)
    limit = IIf(Dir = checkAcsend, limit2, 1)
    
    Do While tmp <> limit
        tmp = tmp + Dir
        '--Get following piece
        If Way = checkX Then
            tmpP = GetFilled(tmp, Final.Y)
        Else
            tmpP = GetFilled(Final.X, tmp)
        End If
        
        If NoPiece = tmpP.Name Then GoTo toLoop
        
        '--if the piece is enemy King or Castle
        If (King = tmpP.Name Or Castle = tmpP.Name) And tmpP.side = -side Then
            IsCheck2_Straight = True
            Exit Function
        End If
        
        '--If the piece is enemy Pawn
        If Pawn = tmpP.Name And tmpP.side = -side Then
            '--Check if it is deactivated Pawn
            If Way = checkX Or (Way = checkY And side = -Dir) Then
                '--Atack range of pawn: 1
                If Abs(tmp - pos) = 1 Then
                    IsCheck2_Straight = True
                    Exit Function
                End If
            End If
        End If
        
        '--check if there is Rocket behind previous piece
        tmp2 = tmp
        Do While tmp2 <> limit
            tmp2 = tmp2 + Dir
            '--Get following piece according to check way
            If Way = checkX Then
                tmpP = GetFilled(tmp2, Final.Y)
            Else
                tmpP = GetFilled(Final.X, tmp2)
            End If
            
            If NoPiece <> tmpP.Name Then
                If Rocket = tmpP.Name And tmpP.side = -side Then
                    IsCheck2_Straight = True
                    Exit Function
                Else
                    Exit Do
                End If
            End If
        Loop
        Exit Do
toLoop:
    Loop
    IsCheck2_Straight = False
End Function

Private Function IsCheck2_Knight(Final As Board, side As Integer) As Boolean
Dim tmpP As ChessPiece
    tmpP = GetFilled(Final.X - 1, Final.Y - 2)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    tmpP = GetFilled(Final.X - 2, Final.Y - 1)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    tmpP = GetFilled(Final.X - 2, Final.Y + 1)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    tmpP = GetFilled(Final.X - 1, Final.Y + 2)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    tmpP = GetFilled(Final.X + 1, Final.Y + 2)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    tmpP = GetFilled(Final.X + 2, Final.Y + 1)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    tmpP = GetFilled(Final.X + 2, Final.Y - 1)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    tmpP = GetFilled(Final.X + 1, Final.Y - 2)
    If tmpP.side = -side And Knight = tmpP.Name Then
        IsCheck2_Knight = True
        Exit Function
    End If
    IsCheck2_Knight = False
End Function

Private Function GetFilled(X As Integer, Y As Integer) As ChessPiece
    Dim i As Integer
    For i = 0 To TotalPieces - 1
        If Pieces(i).XY.X = X And Pieces(i).XY.Y = Y Then
            GetFilled = Pieces(i)
            Exit Function
        End If
    Next
End Function

Private Function IsOutOfSq(Final As Board) As Boolean
    IsOutOfSq = IIf((Final.X < 4 Or Final.X > 6) Or (Final.Y > 3 And Final.Y < 8), True, False)
End Function

Private Function IsBlocked(Piece As ChessPiece, Start As Board, Final As Board, _
Optional xDir As Integer) As Integer

Dim diffX As Integer, diffY As Integer
    diffX = Final.X - Start.X
    diffY = Final.Y - Start.Y
    Select Case Piece.Name
        Case Knight, Bishop
            If GetFilled(Final.X - Sgn(diffX), Final.Y - Sgn(diffY)).Name Then
                IsBlocked = True
                Exit Function
            End If
        Case Castle, Rocket
            Dim nBlock As Integer
            '--xDir = 0 when it is going vertically, else going horizontally
            If xDir Then
                Dim i As Integer
                For i = Start.X + Sgn(diffX) To Final.X - Sgn(diffX) Step Sgn(diffX)
                    If GetFilled(i, Start.Y).Name Then
                        nBlock = nBlock + 1
                    End If
                Next
            Else
                Dim j As Integer
                For j = Start.Y + Sgn(diffY) To Final.Y - Sgn(diffY) Step Sgn(diffY)
                    If GetFilled(Start.X, j).Name Then
                        nBlock = nBlock + 1
                    End If
                Next
            End If
    End Select
    IsBlocked = nBlock
End Function

Function IsLegalMove(Piece As ChessPiece, Start As Board, Final As Board, xParam As ChessPiece) As Boolean

    If (Start.X = Final.X And Start.Y = Final.Y) Or Piece.side = xParam.side Then
        IsLegalMove = False
        Exit Function
    End If

Dim diffX As Integer, diffY As Integer
    diffX = Abs(Final.X - Start.X)
    diffY = Abs(Final.Y - Start.Y)
    Select Case Piece.Name
        Case King
            If IsOutOfSq(Final) Then
                IsLegalMove = False
                Exit Function
            End If
            If diffX + diffY <> 1 Then
                IsLegalMove = False
                Exit Function
            End If
        Case Castle
            If diffX * diffY <> 0 Or diffX = diffY Then     'Not (diffX Xor diffY)
                IsLegalMove = False
                Exit Function
            End If
            If IsBlocked(Piece, Start, Final, diffX) Then
                IsLegalMove = False
                Exit Function
            End If
        Case Knight
            If diffX + diffY <> 3 Or diffX = 0 Or diffY = 0 Then
                IsLegalMove = False
                Exit Function
            End If
            If IsBlocked(Piece, Start, Final) Then
                IsLegalMove = False
                Exit Function
            End If
        Case Rocket
            If diffX * diffY <> 0 Or diffX = diffY Then
                IsLegalMove = False
                Exit Function
            End If
            '--If is Normal
            If Normal = xParam.Name Then
                If IsBlocked(Piece, Start, Final, diffX) Then
                    IsLegalMove = False
                    Exit Function
                End If
            Else
                If IsBlocked(Piece, Start, Final, diffX) <> 1 Then
                    IsLegalMove = False
                    Exit Function
                End If
            End If
        Case Scholar
            If IsOutOfSq(Final) Then
                IsLegalMove = False
                Exit Function
            End If
            If diffX + diffY <> 2 Or diffX = 0 Or diffY = 0 Then
                IsLegalMove = False
                Exit Function
            End If
        Case Bishop
            If Piece.side = -FinalSide(Final.Y) Then
                IsLegalMove = False
                Exit Function
            End If
            If diffX <> 2 Or diffY <> 2 Then
                IsLegalMove = False
                Exit Function
            End If
            If IsBlocked(Piece, Start, Final) Then
                IsLegalMove = False
                Exit Function
            End If
        Case Pawn
            If Piece.side = FinalSide(Final.Y) Then
                If diffY <> 1 Then
                    IsLegalMove = False
                    Exit Function
                End If
            End If
            If diffX + diffY <> 1 Then
                IsLegalMove = False
                Exit Function
            End If
            If Final.Y - Start.Y = Piece.side Then
                IsLegalMove = False
                Exit Function
            End If
    End Select
    IsLegalMove = True
End Function

Private Function FinalSide(Y As Integer) As Integer
    If Right(Y - 1, 1) <= 4 Then
        FinalSide = BottomSide
    Else
        FinalSide = TopSide
    End If
End Function

Public Sub DialogExp(Optional Agent As IAgentCtlCharacter, Optional Text As String, Optional Animation As String)
    On Error Resume Next
    If AgentAvail(Agent.Description) Then
        Agent.Play Animation
        Agent.Speak Text
    Else
        If Text <> vbNullString Then MsgBox Text
    End If
End Sub
