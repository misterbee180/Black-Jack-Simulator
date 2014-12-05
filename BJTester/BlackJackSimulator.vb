Public Class BlackJackSimulator

    'Table Rules
    Private mlngDecks As Int32 = 0
    Private mlngUnitBet As Int32 = 1
    Private mlngMaxHands As Int32 = 1
    Private mblnHitSoft17 As Boolean = True

    'Count Variables
    Private mblnAdjustBet = False
    Private mlngMaxBet As Int32 = 200
    Private mIndexLevel As Index = 0
    Private mPosition As Position = 0
    Private mlngPlayers As Int32 = 1
    Private mlngDeckPenatration As Int32 = 0
    Private mlngCount As Int32 = 0

    'Debug
    Private mblnComments As Boolean = False
    Private mblnStatistics As Boolean = True
    Private mCardCountList As New List(Of Int32)
    Private mlngPlayedHands As Int32 = 0
    Private mShuffleCount As Int32 = 0

    'Strategy Tables
    Private mHardStrategy(16, 12) As String
    Private mSoftStrategy(11, 12) As String
    Private mPairStrategy(12, 12) As String

    'Necessary
    Private mlstPlayers As New List(Of Player)
    Private mDealer As New Player()
    Private mRandom As New Random()

#Region "Enumerations"
    Enum Index
        None = 0
        Illustrious18 = 1
    End Enum

    Enum Position
        FirstBase = 0
        SecondBase = 1
        ThirdBase = 2
    End Enum

    Enum CardSuit
        Hearts = 0
        Diamonds = 1
        Clubs = 2
        Spades = 3
    End Enum

    Enum CardType
        Ace = 1
        Two = 2
        Three = 3
        Four = 4
        Five = 5
        Six = 6
        Seven = 7
        Eight = 8
        Nine = 9
        Ten = 10
        Jack = 11
        Queen = 12
        King = 13
        Unknown = 15
    End Enum

    Enum Moves
        M_Hit = 1
        M_Stay = 2
        M_Double = 3
        M_Split = 4
        M_Surrender = 5
    End Enum

    Enum State
        BlackJack = 1
        Bust = 2
        Neither = 3
    End Enum
#End Region

    Public Sub New(ByVal plngDecks As Int32, _
                   ByVal plngUnit As Int32, _
                   ByVal plngHands As Int32, _
                   ByVal plngPenetration As Int32, _
                   ByVal plngPlayers As Int32)

        mlngDecks = plngDecks
        mlngUnitBet = plngUnit
        mlngMaxHands = plngHands
        mlngDeckPenatration = plngPenetration
        mlngPlayers = plngPlayers

        SetUpStrategy()
    End Sub

    Public Function Run(Optional ByRef pTestCase As TestCase = Nothing, _
                        Optional ByVal pblnComments As Boolean = False, _
                        Optional ByVal pblnStatistics As Boolean = False)

        mlngPlayedHands = 0
        mDealer = New Player()
        mblnComments = pblnComments
        mblnStatistics = pblnStatistics
        GeneratePlayers() 'Creates Players
        If mblnStatistics Then
            SetUpStatistics()
        End If

        While mlngPlayedHands < mlngMaxHands
            mlngCount = 0
            If mblnComments Then
                Console.WriteLine("Suffle Deck")
            End If
            If mblnStatistics Then
                mShuffleCount += 1
            End If

            'Get new Shoe
            Dim Shoe As Shoe = pTestCase.mShoe
            If pTestCase.mShoe Is Nothing Then Shoe = New Shoe(mlngDecks, mRandom)

            While mlngPlayedHands < mlngMaxHands And Shoe.Cards.Count() > mlngDeckPenatration
                mlngPlayedHands += 1
                If mblnComments Then
                    Shoe.DisplayDeck()
                End If
                InitialDeal(mlstPlayers, mDealer, Shoe)

                If mblnComments Then
                    Console.WriteLine("Dealers Cards")
                    mDealer.Hands(0).DisplayCards()
                End If

                'Handle dealer 21
                If (mDealer.Hands(0).Cards(1).Type >= CardType.Ten AndAlso _
                    mDealer.Hands(0).Cards(1).Type <= CardType.King) OrElse _
                    mDealer.Hands(0).Cards(1).Type = CardType.Ace Then
                    Dim blnInsurance As Boolean = False
                    If mDealer.Hands(0).Cards(1).Type = CardType.Ace Then
                        'offer chance for insurance
                    End If
                    mDealer.Hands(0).EvaluateState()
                    If mDealer.Hands(0).State = State.BlackJack Then
                        If mblnComments Then
                            Console.WriteLine("Dealer BlackJack")
                        End If

                        For Each Player As Player In mlstPlayers
                            For Each Hand As Hand In Player.Hands
                                Hand.EvaluateState()
                                If Not Hand.insurance AndAlso Hand.State <> State.BlackJack Then
                                    Player.Lose(Hand.GetBet)

                                    If mblnComments Then
                                        Console.WriteLine("Players Hand")
                                        Player.Hands(0).DisplayCards()
                                        Console.WriteLine("Players Loses")
                                        Console.WriteLine("Winnings: " & Player.Winnings)
                                        Console.WriteLine("Losings: " & Player.Losings)
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If

                'Play balckjack if dealer didn't have 21
                If mDealer.Hands(0).State <> State.BlackJack Then
                    For i = 0 To mlngPlayers - 1
                        Dim j = 0
                        Do While j <= mlstPlayers(i).Hands.Count - 1
                            While Not mlstPlayers(i).Hands(j).Done
                                If mlstPlayers(i).Hands(j).Cards.Count() = 1 Then 'handle the case where a split occured and need to add a card to that hand.
                                    mlstPlayers(i).Hands(j).AddCard(Shoe.GetCard(), True)
                                    If mblnStatistics Then
                                        mCardCountList(mlstPlayers(i).Hands(j).Cards(1).Type) += 1
                                    End If
                                End If

                                If mblnComments Then
                                    Console.WriteLine("Player " & i.ToString & " Hand")
                                    mlstPlayers(i).Hands(j).DisplayCards()
                                End If

                                Select Case DetermineMove(mlstPlayers(i).Hands(j), mDealer.Hands(0).Cards(1), mlstPlayers(i).SplitCount)
                                    Case Moves.M_Double
                                        mlstPlayers(i).Hands(j).First = False
                                        mlstPlayers(i).Hands(j).dDouble = True
                                        mlstPlayers(i).Hands(j).Done = True
                                        mlstPlayers(i).Hands(j).AddCard(Shoe.GetCard(), True)

                                        If mblnComments Then
                                            Console.WriteLine("Player Double")
                                            Console.WriteLine("Player " & i.ToString & " Hand")
                                            mlstPlayers(i).Hands(j).DisplayCards()
                                        End If
                                    Case Moves.M_Hit
                                        mlstPlayers(i).Hands(j).First = False
                                        mlstPlayers(i).Hands(j).AddCard(Shoe.GetCard(), True)

                                        If mblnComments Then
                                            Console.WriteLine("Player Hit")
                                        End If
                                    Case Moves.M_Split
                                        If mblnComments Then
                                            Console.WriteLine("Player Split")
                                        End If
                                        mlstPlayers(i).SplitCount += 1
                                        If mlstPlayers(i).Hands(j).Cards(0).Type = CardType.Ace Then
                                            mlstPlayers(i).AddHand(mlstPlayers(i).Hands(j).Cards.Dequeue(), True, True, True, mlstPlayers(i).Hands(j).GetBet(), False)
                                            mlstPlayers(i).Hands(j).SplitAce = True
                                        Else
                                            mlstPlayers(i).AddHand(mlstPlayers(i).Hands(j).Cards.Dequeue(), True, True, False, mlstPlayers(i).Hands(j).GetBet(), False)
                                        End If
                                        mlstPlayers(i).Hands(j).First = True
                                        mlstPlayers(i).Hands(j).Split = True
                                    Case Else
                                        mlstPlayers(i).Hands(j).Done = True
                                End Select

                                If mlstPlayers(i).Hands(j).BestScore = 0 Then 'Busted
                                    mlstPlayers(i).Hands(j).Done = True

                                    If mblnComments Then
                                        Console.WriteLine("Player " & i.ToString & " Hand")
                                        mlstPlayers(i).Hands(j).DisplayCards()
                                    End If
                                End If
                            End While
                            j += 1
                        Loop
                    Next

                    mDealer.Hands(0).Cards(0).RevealCard() 'reveal hidden card

                    'See if there are any hands that didn't bust or blackjack
                    Dim deal As Boolean = False
                    For Each Player As Player In mlstPlayers
                        For Each Hand As Hand In Player.Hands
                            Hand.EvaluateState()
                            If Hand.State = State.Neither Then
                                deal = True
                                Exit For
                            End If
                        Next
                        If deal = True Then
                            Exit For
                        End If
                    Next

                    If deal Then 'get dealer his cards if necessary
                        mDealer.Hands(0).EvaluateState()
                        While mDealer.Hands(0).State = State.Neither
                            If mblnHitSoft17 AndAlso mDealer.Hands(0).isSoft AndAlso mDealer.Hands(0).BestScore = 17 Then
                                mDealer.Hands(0).AddCard(Shoe.GetCard(), True)

                                If mblnComments Then
                                    Console.WriteLine("Dealer Hit")
                                    mDealer.Hands(0).DisplayCards()
                                End If
                            ElseIf mDealer.Hands(0).BestScore < 17 And mDealer.Hands(0).BestScore <> 0 Then
                                mDealer.Hands(0).AddCard(Shoe.GetCard(), True)

                                If mblnComments Then
                                    Console.WriteLine("Dealer Hit")
                                    mDealer.Hands(0).DisplayCards()
                                End If
                            Else
                                Exit While
                            End If
                            mDealer.Hands(0).EvaluateState()
                        End While
                    End If

                    'Compare and assign winnings or losings.
                    For Each Player As Player In mlstPlayers
                        For Each Hand As Hand In Player.Hands
                            If Hand.BestScore() > mDealer.Hands(0).BestScore() Then
                                Player.Win(Hand.GetBet())
                                If mblnComments Then
                                    Console.WriteLine("Player Wins")
                                    Console.WriteLine("Winnings: " & Player.Winnings)
                                    Console.WriteLine("Losings: " & Player.Losings)
                                    Console.WriteLine("Wins: " & Player.Wins)
                                    Console.WriteLine("Losses: " & Player.Losses)
                                    Console.WriteLine("Pushes: " & Player.Pushes)
                                End If
                            ElseIf Hand.BestScore() < mDealer.Hands(0).BestScore() Then
                                Player.Lose(Hand.GetBet())
                                If mblnComments Then
                                    Console.WriteLine("Dealer Wins")
                                    Console.WriteLine("Winnings: " & Player.Winnings)
                                    Console.WriteLine("Losings: " & Player.Losings)
                                    Console.WriteLine("Wins: " & Player.Wins)
                                    Console.WriteLine("Losses: " & Player.Losses)
                                    Console.WriteLine("Pushes: " & Player.Pushes)
                                End If
                            Else
                                'No one wins or loses anytihng
                                Player.Push()
                                If mblnComments Then
                                    Console.WriteLine("Push")
                                    Console.WriteLine("Winnings: " & Player.Winnings)
                                    Console.WriteLine("Losings: " & Player.Losings)
                                    Console.WriteLine("Wins: " & Player.Wins)
                                    Console.WriteLine("Losses: " & Player.Losses)
                                    Console.WriteLine("Pushes: " & Player.Pushes)
                                End If
                            End If
                        Next
                    Next
                End If

                'Remove hands
                For Each Player As Player In mlstPlayers
                    Player.Hands.Clear()
                    Player.SplitCount = 0
                Next
                mDealer.Hands.Clear()

                If mblnComments Then
                    Console.WriteLine()
                End If
            End While
        End While

        Dim z As Int32 = 0
        Dim blnSuccess As Boolean = True
        For Each Player As Player In mlstPlayers
            z += 1
            Console.WriteLine()
            Console.WriteLine("Player " & z & " winnings: " & Player.Winnings)
            Console.WriteLine("Player " & z & " losings: " & Player.Losings)
            Console.WriteLine("Player " & z & " percent: " & (Player.Winnings / Player.Losings) - 1)
            Console.WriteLine()

            If Player.Winnings <> pTestCase.mlngWinnings Or Player.Losings <> pTestCase.mlngLosings Then
                pTestCase.mblnSuccess = False
            End If
        Next

        If mblnStatistics Then
            Console.WriteLine("Shoes Played: " & mShuffleCount)
            For i = 0 To mCardCountList.Count - 1
                If i <> 0 Then
                    Dim card As New Card(CardSuit.Clubs, i)
                    Console.Write(card.Name & " = " & mCardCountList(i).ToString & ", ")
                End If
            Next
            Console.WriteLine()
        End If

        Return blnSuccess
    End Function

    Public Class Card
        Public Suit As New CardSuit
        Public Type As New CardType
        Public mblnIsRevealed As Boolean = False

        Public Sub New(ByVal pSuit As CardSuit, _
                       ByVal pType As CardType)
            Suit = pSuit
            Type = pType
        End Sub

        Public Sub New(ByVal plngValue As Int32)
            Suit = CardSuit.Clubs 'Just a random suit
            Type = ValueToType(plngValue)
        End Sub

        Public Sub RevealCard()
            mblnIsRevealed = True
            'UpdateCount(Type)
        End Sub

        Public Function Name() As String
            Select Case Type
                Case CardType.Ace
                    Return "A"
                Case CardType.Two
                    Return "2"
                Case CardType.Three
                    Return "3"
                Case CardType.Four
                    Return "4"
                Case CardType.Five
                    Return "5"
                Case CardType.Six
                    Return "6"
                Case CardType.Seven
                    Return "7"
                Case CardType.Eight
                    Return "8"
                Case CardType.Nine
                    Return "9"
                Case CardType.Ten
                    Return "10"
                Case CardType.Jack
                    Return "J"
                Case CardType.Queen
                    Return "Q"
                Case Else
                    Return "K"
            End Select
        End Function

        Public Function ValueToType(ByVal plngValue As Int32) As CardType
            Select Case plngValue
                Case 1
                    Return CardType.Ace
                Case 2
                    Return CardType.Two
                Case 3
                    Return CardType.Three
                Case 4
                    Return CardType.Four
                Case 5
                    Return CardType.Five
                Case 6
                    Return CardType.Six
                Case 7
                    Return CardType.Seven
                Case 8
                    Return CardType.Eight
                Case 9
                    Return CardType.Nine
                Case 10
                    Return CardType.Ten
                Case 11
                    Return CardType.Jack
                Case 12
                    Return CardType.Queen
                Case Else
                    Return CardType.King
            End Select
        End Function
    End Class

    Public Class Shoe
        Public Cards As New List(Of Card)
        Public mlngShuffleCount As int32 = 0

        Public Sub New(ByVal plngDecks As Int32, _
                       ByRef pRandom As Random)
            For i = 1 To 13
                For j = 0 To 3
                    For k = 1 To plngDecks
                        Cards.Add(New Card(j, i))
                    Next
                Next
            Next
            Shuffle(pRandom)
            'If mComments Then
            '    DisplayDeck()
            'End If
        End Sub

        Public Sub New(ByVal lngCardArray As Int32())
            For Each cardValue In lngCardArray
                Cards.Add(New Card(cardValue))
            Next
        End Sub

        Public Sub New()
        End Sub

        Public Sub DisplayDeck()
            Console.WriteLine("Deck after Shuffle")
            For i = 0 To Cards.Count - 1
                If i Mod 13 = 0 Then
                    Console.WriteLine()
                End If
                Console.Write(Cards(i).Name & ", ")
            Next
            Console.WriteLine()
            Console.WriteLine()
        End Sub

        Private Sub Shuffle(ByRef pRandom As Random)
            Dim ShuffledCards As New List(Of Card)
            mlngShuffleCount += 1

            For j = 1 To Cards.Count
                Dim k As Int32 = pRandom.Next(Cards.Count)
                ShuffledCards.Add(Cards(k))
                Cards.Remove(Cards(k))
            Next
            Cards = ShuffledCards
        End Sub

        Public Function GetCard() As Card
            Dim z As Card = Cards.First()
            Cards.Remove(Cards.First())
            Return z
        End Function

        Function GetSetCard(ByVal Card As CardType) As Card
            Return New Card(CardSuit.Hearts, Card)
        End Function

    End Class

    Public Class Hand
        Public Cards As New Queue(Of Card)
        Public State As State
        Public Done As Boolean = False
        Public Split As Boolean = False
        Public First As Boolean = False
        Public SplitAce As Boolean = False
        Private Bet As Double
        Public dDouble As Boolean = False
        Public insurance As Boolean = False

        Sub New(ByVal pCard As Card, _
                ByVal pFirst As Boolean, _
                ByVal pSplit As Boolean, _
                ByVal pSplitAce As Boolean, _
                ByVal pBet As Int32)
            Cards.Enqueue(pCard)
            First = pFirst
            Split = pSplit
            Bet = pBet
        End Sub

        Public Function isSoft()
            If HasAce() AndAlso PossibleSum("LT", 11) Then
                Return True
            End If
            Return False
        End Function

        Public Sub DisplayCards()
            For Each Card As Card In Cards
                Console.Write(Card.Name & " ")
            Next
            Console.WriteLine()
        End Sub

        Public Sub AddCard(ByVal pCard As Card, _
                           ByVal pReveal As Boolean)
            Cards.Enqueue(pCard)
            If pReveal Then
                pCard.RevealCard()
            End If
        End Sub

        Public Sub EvaluateState()
            If PossibleSum("EQ", 21) AndAlso First AndAlso Not Split Then
                State = State.BlackJack
            ElseIf BestScore() = 0 Then
                State = State.Bust
            Else
                State = State.Neither
            End If
        End Sub

        Public Function GetBet() As Double
            If dDouble = True Then
                Return Bet * 2
            ElseIf State = State.BlackJack AndAlso First AndAlso Not Split Then
                Return Bet * 1.5
            Else : Return Bet
            End If
        End Function

        Public Function Order() As IEnumerable(Of Card)
            Return Cards.OrderBy(Function(x) x.Type)
        End Function

        Public Function HasAce() As Boolean
            If Order(0).Type = CardType.Ace Then Return True
            Return False
        End Function

        Public Function PossibleSum(ByVal pstrOperator As String, _
                                             ByVal plngValue As Int32) As Boolean
            Dim Sums As New List(Of Int32)
            Sums.Add(0)
            For Each crd As Card In Cards
                Dim NewSums As New List(Of Int32)
                For Each sum As Int32 In Sums
                    If crd.Type = CardType.Jack OrElse crd.Type = CardType.Queen OrElse crd.Type = CardType.King Then
                        NewSums.Add(sum + 10)
                    Else
                        NewSums.Add(sum + crd.Type)
                    End If
                Next
                If crd.Type = CardType.Ace Then
                    Dim totalSums As Int32 = Sums.Count()
                    For i = 0 To totalSums - 1
                        NewSums.Add(NewSums(i) + 10)
                    Next
                End If
                Sums = NewSums
            Next

            If pstrOperator = "LT" Then
                For i = 1 To plngValue - 1
                    If Sums.Contains(i) Then Return True
                Next
            ElseIf pstrOperator = "LTE" Then
                For i = 1 To plngValue
                    If Sums.Contains(i) Then Return True
                Next
            ElseIf pstrOperator = "EQ" Then
                If Sums.Contains(plngValue) Then Return True
            ElseIf pstrOperator = "GT" Then
                For i = plngValue + 1 To 21
                    If Sums.Contains(i) Then Return True
                Next
            Else : Throw New Exception("Not valid operand")
            End If

            Return False
        End Function

        Public Function BestScore() As Int32
            Dim Sums As New List(Of Int32)
            Sums.Add(0)
            For Each crd As Card In Cards
                Dim NewSums As New List(Of Int32)
                For Each sum As Int32 In Sums
                    If crd.Type = CardType.Jack OrElse crd.Type = CardType.Queen OrElse crd.Type = CardType.King Then
                        NewSums.Add(sum + 10)
                    Else
                        NewSums.Add(sum + crd.Type)
                    End If
                Next
                If crd.Type = CardType.Ace Then
                    Dim totalSums As Int32 = Sums.Count()
                    For i = 0 To totalSums - 1
                        NewSums.Add(NewSums(i) + 10)
                    Next
                End If
                Sums = NewSums
            Next
            Dim bestSum As Int32 = 0
            For Each sum As Int32 In Sums
                If sum > bestSum AndAlso sum <= 21 Then
                    bestSum = sum
                End If
            Next
            Return bestSum
        End Function

        Public Function IsPair() As Boolean
            If Cards(0).Type = Cards(1).Type AndAlso Cards.Count = 2 Then
                Return True
            End If
            Return False
        End Function

    End Class

    Public Class Player
        Public Hands As List(Of Hand)
        Public Winnings As Double
        Public Losings As Double
        Public Wins As Int32 = 0
        Public Losses As Int32 = 0
        Public Pushes As Int32 = 0
        Public SplitCount As Int32 = 0

        Public Sub New()
            Winnings = 0
            Losings = 0
            Hands = New List(Of Hand)
        End Sub

        Public Sub Win(ByVal pAmount As Double)
            Winnings += pAmount
            Wins += 1
        End Sub

        Public Sub Lose(ByVal pAmount As Double)
            Losings += pAmount
            Losses += 1
        End Sub

        Public Sub Push()
            Pushes += 1
        End Sub

        Public Sub AddHand(ByVal pStartCard As Card, _
                           ByVal pFirst As Boolean, _
                           ByVal pSplit As Boolean, _
                           ByVal pSplitAce As Boolean, _
                           ByVal pBet As Double, _
                           ByVal pReveal As Boolean)
            Hands.Add(New Hand(pStartCard, pFirst, pSplit, pSplitAce, pBet))
            If pReveal Then
                pStartCard.RevealCard()
            End If
        End Sub
    End Class

#Region "Functions"
    Private Sub GeneratePlayers()
        mlstPlayers.Clear()
        For i = 1 To mlngPlayers
            mlstPlayers.Add(New Player())
        Next
    End Sub

    Private Function getCountMultiplier(ByVal pDeck As Shoe) As Integer
        If mblnAdjustBet = True Then
            Select Case mlngCount / (mlngDecks * pDeck.Cards.Count() / mlngDecks * 52)
                Case Is <= 1
                    Return 1
                Case 2
                    Return 4
                Case 3
                    Return 6
                Case 4
                    Return 16
                Case Else
                    Return 32
            End Select
        Else
            Return 1
        End If

    End Function

    Private Sub UpdateCount(ByVal pType As CardType)
        If pType >= CardType.Two And pType <= CardType.Six Then
            mlngCount += 1
        ElseIf pType = CardType.Ace OrElse pType >= CardType.Ten Then
            mlngCount -= 1
        End If
    End Sub

    Public Sub RestartCount()
        mlngCount = 0
    End Sub

    Private Sub InitialDeal(ByRef pPlayers As List(Of Player), _
                            ByRef pDealer As Player, _
                            ByRef pDeck As Shoe)
        Dim ReceivedCard As Card = pDeck.GetCard()
        If mblnStatistics Then
            mCardCountList(ReceivedCard.Type) += 1
        End If
        pDealer.AddHand(ReceivedCard, _
                        True, _
                        False, _
                        False, _
                        0, _
                        False) 'Add card but do not influence count
        ReceivedCard = pDeck.GetCard()
        If mblnStatistics Then
            mCardCountList(ReceivedCard.Type) += 1
        End If
        pDealer.Hands(0).AddCard(ReceivedCard, _
                                 True)

        For Each Player As Player In pPlayers
            ReceivedCard = pDeck.GetCard()
            If mblnStatistics Then
                mCardCountList(ReceivedCard.Type) += 1
            End If
            Player.AddHand(ReceivedCard, _
                           True, _
                           False, _
                           False, _
                           mlngUnitBet * getCountMultiplier(pDeck), _
                           True)
            ReceivedCard = pDeck.GetCard()
            If mblnStatistics Then
                mCardCountList(ReceivedCard.Type) += 1
            End If
            Player.Hands(0).AddCard(ReceivedCard, _
                                    True)
        Next
    End Sub

    Public Function DetermineMove(ByVal pPlayerHand As Hand, _
                      ByVal pDealerCard As Card, _
                      ByVal pSplitCount As Int32) As Moves

        Dim strMove As String = ""

        If pPlayerHand.SplitAce Then
            Return Moves.M_Stay
        End If

        If pPlayerHand.IsPair() AndAlso pSplitCount < 4 Then
            strMove = BestPlayPair(pPlayerHand.Cards(0).Type, pDealerCard.Type)
        ElseIf pPlayerHand.isSoft() Then
            Dim SoftSum As Int32 = 0
            For i = 1 To pPlayerHand.Cards.Count - 1
                SoftSum += pPlayerHand.Order(i).Type
            Next
            strMove = BestPlaySoft(SoftSum, pDealerCard.Type)
        Else
            strMove = BestPlayHard(pPlayerHand.BestScore, pDealerCard.Type)
        End If

        Select Case strMove
            Case "H"
                Return Moves.M_Hit
            Case "Dh"
                If pPlayerHand.First Then Return Moves.M_Double
                Return Moves.M_Hit
            Case "P"
                Return Moves.M_Split
            Case "S"
                Return Moves.M_Stay
            Case Else
                Throw New Exception("Not valid Move " & strMove)
        End Select
    End Function

    Private Function BestPlayPair(ByVal pPlayerCard As CardType, ByVal pDealersCard As CardType) As String
        If pPlayerCard >= CardType.Two And pPlayerCard <= CardType.King Then
            If pDealersCard >= CardType.Two And pDealersCard <= CardType.King Then
                Return mPairStrategy(pPlayerCard - 2, pDealersCard - 2)
            Else
                Return mPairStrategy(pPlayerCard - 2, 12)
            End If
        Else
            If pDealersCard >= CardType.Two And pDealersCard <= CardType.King Then
                Return mPairStrategy(12, pDealersCard - 2)
            Else
                Return mPairStrategy(12, 12)
            End If
        End If
    End Function

    Private Function BestPlaySoft(ByVal pPlayerCard As CardType, ByVal pDealersCard As CardType) As String
        If pDealersCard >= CardType.Two And pDealersCard <= CardType.King Then
            Return mSoftStrategy(pPlayerCard - 2, pDealersCard - 2)
        Else
            Return mSoftStrategy(pPlayerCard - 2, 12)
        End If
    End Function

    Private Function BestPlayHard(ByVal pHardTotal As Int32, ByVal pDealersCard As CardType) As String
        If pHardTotal = 4 Then Return "H"

        If pDealersCard >= CardType.Two And pDealersCard <= CardType.King Then
            Return mHardStrategy(pHardTotal - 5, pDealersCard - 2)
        Else
            Return mHardStrategy(pHardTotal - 5, 12)
        End If
    End Function

    Private Sub SetUpStrategy()
        GenerateHardStrategy()
        GenerateSoftStrategy()
        GeneratePairStrategy()
    End Sub

    Private Sub SetUpStatistics()
        mCardCountList.Clear()
        mShuffleCount = 0
        For i = 0 To 13
            mCardCountList.Add(0)
        Next
    End Sub

    Public Sub GenerateHardStrategy()
        '5
        mHardStrategy(0, 0) = "H" '2
        mHardStrategy(0, 1) = "H" '3
        mHardStrategy(0, 2) = "H" '4
        mHardStrategy(0, 3) = "H" '5
        mHardStrategy(0, 4) = "H" '6
        mHardStrategy(0, 5) = "H" '7
        mHardStrategy(0, 6) = "H" '8
        mHardStrategy(0, 7) = "H" '9
        mHardStrategy(0, 8) = "H" '10
        mHardStrategy(0, 9) = "H" 'J
        mHardStrategy(0, 10) = "H" 'Q
        mHardStrategy(0, 11) = "H" 'K
        mHardStrategy(0, 12) = "H" 'A

        '6
        mHardStrategy(1, 0) = "H" '2
        mHardStrategy(1, 1) = "H" '3
        mHardStrategy(1, 2) = "H" '4
        mHardStrategy(1, 3) = "H" '5
        mHardStrategy(1, 4) = "H" '6
        mHardStrategy(1, 5) = "H" '7
        mHardStrategy(1, 6) = "H" '8
        mHardStrategy(1, 7) = "H" '9
        mHardStrategy(1, 8) = "H" '10
        mHardStrategy(1, 9) = "H" 'J
        mHardStrategy(1, 10) = "H" 'Q
        mHardStrategy(1, 11) = "H" 'K
        mHardStrategy(1, 12) = "H" 'A

        '7
        mHardStrategy(2, 0) = "H" '2
        mHardStrategy(2, 1) = "H" '3
        mHardStrategy(2, 2) = "H" '4
        mHardStrategy(2, 3) = "H" '5
        mHardStrategy(2, 4) = "H" '6
        mHardStrategy(2, 5) = "H" '7
        mHardStrategy(2, 6) = "H" '8
        mHardStrategy(2, 7) = "H" '9
        mHardStrategy(2, 8) = "H" '10
        mHardStrategy(2, 9) = "H" 'J
        mHardStrategy(2, 10) = "H" 'Q
        mHardStrategy(2, 11) = "H" 'K
        mHardStrategy(2, 12) = "H" 'A

        '8
        mHardStrategy(3, 0) = "H" '2
        mHardStrategy(3, 1) = "H" '3
        mHardStrategy(3, 2) = "H" '4
        mHardStrategy(3, 3) = "H" '5
        mHardStrategy(3, 4) = "H" '6
        mHardStrategy(3, 5) = "H" '7
        mHardStrategy(3, 6) = "H" '8
        mHardStrategy(3, 7) = "H" '9
        mHardStrategy(3, 8) = "H" '10
        mHardStrategy(3, 9) = "H" 'J
        mHardStrategy(3, 10) = "H" 'Q
        mHardStrategy(3, 11) = "H" 'K
        mHardStrategy(3, 12) = "H" 'A

        '9
        mHardStrategy(4, 0) = "H" '2
        mHardStrategy(4, 1) = "Dh" '3
        mHardStrategy(4, 2) = "Dh" '4
        mHardStrategy(4, 3) = "Dh" '5
        mHardStrategy(4, 4) = "Dh" '6
        mHardStrategy(4, 5) = "H" '7
        mHardStrategy(4, 6) = "H" '8
        mHardStrategy(4, 7) = "H" '9
        mHardStrategy(4, 8) = "H" '10
        mHardStrategy(4, 9) = "H" 'J
        mHardStrategy(4, 10) = "H" 'Q
        mHardStrategy(4, 11) = "H" 'K
        mHardStrategy(4, 12) = "H" 'A

        '10
        mHardStrategy(5, 0) = "Dh" '2
        mHardStrategy(5, 1) = "Dh" '3
        mHardStrategy(5, 2) = "Dh" '4
        mHardStrategy(5, 3) = "Dh" '5
        mHardStrategy(5, 4) = "Dh" '6
        mHardStrategy(5, 5) = "Dh" '7
        mHardStrategy(5, 6) = "Dh" '8
        mHardStrategy(5, 7) = "Dh" '9
        mHardStrategy(5, 8) = "H" '10
        mHardStrategy(5, 9) = "H" 'J
        mHardStrategy(5, 10) = "H" 'Q
        mHardStrategy(5, 11) = "H" 'K
        mHardStrategy(5, 12) = "H" 'A

        '11
        mHardStrategy(6, 0) = "Dh" '2
        mHardStrategy(6, 1) = "Dh" '3
        mHardStrategy(6, 2) = "Dh" '4
        mHardStrategy(6, 3) = "Dh" '5
        mHardStrategy(6, 4) = "Dh" '6
        mHardStrategy(6, 5) = "Dh" '7
        mHardStrategy(6, 6) = "Dh" '8
        mHardStrategy(6, 7) = "Dh" '9
        mHardStrategy(6, 8) = "Dh" '10
        mHardStrategy(6, 9) = "Dh" 'J
        mHardStrategy(6, 10) = "Dh" 'Q
        mHardStrategy(6, 11) = "Dh" 'K
        mHardStrategy(6, 12) = "H" 'A

        '12
        mHardStrategy(7, 0) = "H" '2
        mHardStrategy(7, 1) = "H" '3
        mHardStrategy(7, 2) = "S" '4
        mHardStrategy(7, 3) = "S" '5
        mHardStrategy(7, 4) = "S" '6
        mHardStrategy(7, 5) = "H" '7
        mHardStrategy(7, 6) = "H" '8
        mHardStrategy(7, 7) = "H" '9
        mHardStrategy(7, 8) = "H" '10
        mHardStrategy(7, 9) = "H" 'J
        mHardStrategy(7, 10) = "H" 'Q
        mHardStrategy(7, 11) = "H" 'K
        mHardStrategy(7, 12) = "H" 'A

        '13
        mHardStrategy(8, 0) = "S" '2
        mHardStrategy(8, 1) = "S" '3
        mHardStrategy(8, 2) = "S" '4
        mHardStrategy(8, 3) = "S" '5
        mHardStrategy(8, 4) = "S" '6
        mHardStrategy(8, 5) = "H" '7
        mHardStrategy(8, 6) = "H" '8
        mHardStrategy(8, 7) = "H" '9
        mHardStrategy(8, 8) = "H" '10
        mHardStrategy(8, 9) = "H" 'J
        mHardStrategy(8, 10) = "H" 'Q
        mHardStrategy(8, 11) = "H" 'K
        mHardStrategy(8, 12) = "H" 'A

        '14
        mHardStrategy(9, 0) = "S" '2
        mHardStrategy(9, 1) = "S" '3
        mHardStrategy(9, 2) = "S" '4
        mHardStrategy(9, 3) = "S" '5
        mHardStrategy(9, 4) = "S" '6
        mHardStrategy(9, 5) = "H" '7
        mHardStrategy(9, 6) = "H" '8
        mHardStrategy(9, 7) = "H" '9
        mHardStrategy(9, 8) = "H" '10
        mHardStrategy(9, 9) = "H" 'J
        mHardStrategy(9, 10) = "H" 'Q
        mHardStrategy(9, 11) = "H" 'K
        mHardStrategy(9, 12) = "H" 'A

        '15
        mHardStrategy(10, 0) = "S" '2
        mHardStrategy(10, 1) = "S" '3
        mHardStrategy(10, 2) = "S" '4
        mHardStrategy(10, 3) = "S" '5
        mHardStrategy(10, 4) = "S" '6
        mHardStrategy(10, 5) = "H" '7
        mHardStrategy(10, 6) = "H" '8
        mHardStrategy(10, 7) = "H" '9
        mHardStrategy(10, 8) = "H" '10
        mHardStrategy(10, 9) = "H" 'J
        mHardStrategy(10, 10) = "H" 'Q
        mHardStrategy(10, 11) = "H" 'K
        mHardStrategy(10, 12) = "H" 'A

        '16
        mHardStrategy(11, 0) = "S" '2
        mHardStrategy(11, 1) = "S" '3
        mHardStrategy(11, 2) = "S" '4
        mHardStrategy(11, 3) = "S" '5
        mHardStrategy(11, 4) = "S" '6
        mHardStrategy(11, 5) = "H" '7
        mHardStrategy(11, 6) = "H" '8
        mHardStrategy(11, 7) = "H" '9
        mHardStrategy(11, 8) = "H" '10
        mHardStrategy(11, 9) = "H" 'J
        mHardStrategy(11, 10) = "H" 'Q
        mHardStrategy(11, 11) = "H" 'K
        mHardStrategy(11, 12) = "H" 'A

        '17
        mHardStrategy(12, 0) = "S" '2
        mHardStrategy(12, 1) = "S" '3
        mHardStrategy(12, 2) = "S" '4
        mHardStrategy(12, 3) = "S" '5
        mHardStrategy(12, 4) = "S" '6
        mHardStrategy(12, 5) = "S" '7
        mHardStrategy(12, 6) = "S" '8
        mHardStrategy(12, 7) = "S" '9
        mHardStrategy(12, 8) = "S" '10
        mHardStrategy(12, 9) = "S" 'J
        mHardStrategy(12, 10) = "S" 'Q
        mHardStrategy(12, 11) = "S" 'K
        mHardStrategy(12, 12) = "S" 'A

        '18
        mHardStrategy(13, 0) = "S" '2
        mHardStrategy(13, 1) = "S" '3
        mHardStrategy(13, 2) = "S" '4
        mHardStrategy(13, 3) = "S" '5
        mHardStrategy(13, 4) = "S" '6
        mHardStrategy(13, 5) = "S" '7
        mHardStrategy(13, 6) = "S" '8
        mHardStrategy(13, 7) = "S" '9
        mHardStrategy(13, 8) = "S" '10
        mHardStrategy(13, 9) = "S" 'J
        mHardStrategy(13, 10) = "S" 'Q
        mHardStrategy(13, 11) = "S" 'K
        mHardStrategy(13, 12) = "S" 'A

        '19
        mHardStrategy(14, 0) = "S" '2
        mHardStrategy(14, 1) = "S" '3
        mHardStrategy(14, 2) = "S" '4
        mHardStrategy(14, 3) = "S" '5
        mHardStrategy(14, 4) = "S" '6
        mHardStrategy(14, 5) = "S" '7
        mHardStrategy(14, 6) = "S" '8
        mHardStrategy(14, 7) = "S" '9
        mHardStrategy(14, 8) = "S" '10
        mHardStrategy(14, 9) = "S" 'J
        mHardStrategy(14, 10) = "S" 'Q
        mHardStrategy(14, 11) = "S" 'K
        mHardStrategy(14, 12) = "S" 'A

        '20
        mHardStrategy(15, 0) = "S" '2
        mHardStrategy(15, 1) = "S" '3
        mHardStrategy(15, 2) = "S" '4
        mHardStrategy(15, 3) = "S" '5
        mHardStrategy(15, 4) = "S" '6
        mHardStrategy(15, 5) = "S" '7
        mHardStrategy(15, 6) = "S" '8
        mHardStrategy(15, 7) = "S" '9
        mHardStrategy(15, 8) = "S" '10
        mHardStrategy(15, 9) = "S" 'J
        mHardStrategy(15, 10) = "S" 'Q
        mHardStrategy(15, 11) = "S" 'K
        mHardStrategy(15, 12) = "S" 'A

        '21
        mHardStrategy(16, 0) = "S" '2
        mHardStrategy(16, 1) = "S" '3
        mHardStrategy(16, 2) = "S" '4
        mHardStrategy(16, 3) = "S" '5
        mHardStrategy(16, 4) = "S" '6
        mHardStrategy(16, 5) = "S" '7
        mHardStrategy(16, 6) = "S" '8
        mHardStrategy(16, 7) = "S" '9
        mHardStrategy(16, 8) = "S" '10
        mHardStrategy(16, 9) = "S" 'J
        mHardStrategy(16, 10) = "S" 'Q
        mHardStrategy(16, 11) = "S" 'K
        mHardStrategy(16, 12) = "S" 'A
    End Sub

    Public Sub GenerateSoftStrategy()
        'A,2
        mSoftStrategy(0, 0) = "H" '2
        mSoftStrategy(0, 1) = "H" '3
        mSoftStrategy(0, 2) = "H" '4
        mSoftStrategy(0, 3) = "Dh" '5
        mSoftStrategy(0, 4) = "Dh" '6
        mSoftStrategy(0, 5) = "H" '7
        mSoftStrategy(0, 6) = "H" '8
        mSoftStrategy(0, 7) = "H" '9
        mSoftStrategy(0, 8) = "H" '10
        mSoftStrategy(0, 9) = "H" 'J
        mSoftStrategy(0, 10) = "H" 'Q
        mSoftStrategy(0, 11) = "H" 'K
        mSoftStrategy(0, 12) = "H" 'A

        'A,3
        mSoftStrategy(1, 0) = "H" '2
        mSoftStrategy(1, 1) = "H" '3
        mSoftStrategy(1, 2) = "H" '4
        mSoftStrategy(1, 3) = "Dh" '5
        mSoftStrategy(1, 4) = "Dh" '6
        mSoftStrategy(1, 5) = "H" '7
        mSoftStrategy(1, 6) = "H" '8
        mSoftStrategy(1, 7) = "H" '9
        mSoftStrategy(1, 8) = "H" '10
        mSoftStrategy(1, 9) = "H" 'J
        mSoftStrategy(1, 10) = "H" 'Q
        mSoftStrategy(1, 11) = "H" 'K
        mSoftStrategy(1, 12) = "H" 'A

        'A,4
        mSoftStrategy(2, 0) = "H" '2
        mSoftStrategy(2, 1) = "H" '3
        mSoftStrategy(2, 2) = "Dh" '4
        mSoftStrategy(2, 3) = "Dh" '5
        mSoftStrategy(2, 4) = "Dh" '6
        mSoftStrategy(2, 5) = "H" '7
        mSoftStrategy(2, 6) = "H" '8
        mSoftStrategy(2, 7) = "H" '9
        mSoftStrategy(2, 8) = "H" '10
        mSoftStrategy(2, 9) = "H" 'J
        mSoftStrategy(2, 10) = "H" 'Q
        mSoftStrategy(2, 11) = "H" 'K
        mSoftStrategy(2, 12) = "H" 'A

        'A,5
        mSoftStrategy(3, 0) = "H" '2
        mSoftStrategy(3, 1) = "H" '3
        mSoftStrategy(3, 2) = "Dh" '4
        mSoftStrategy(3, 3) = "Dh" '5
        mSoftStrategy(3, 4) = "Dh" '6
        mSoftStrategy(3, 5) = "H" '7
        mSoftStrategy(3, 6) = "H" '8
        mSoftStrategy(3, 7) = "H" '9
        mSoftStrategy(3, 8) = "H" '10
        mSoftStrategy(3, 9) = "H" 'J
        mSoftStrategy(3, 10) = "H" 'Q
        mSoftStrategy(3, 11) = "H" 'K
        mSoftStrategy(3, 12) = "H" 'A

        'A,6
        mSoftStrategy(4, 0) = "H" '2
        mSoftStrategy(4, 1) = "Dh" '3
        mSoftStrategy(4, 2) = "Dh" '4
        mSoftStrategy(4, 3) = "Dh" '5
        mSoftStrategy(4, 4) = "Dh" '6
        mSoftStrategy(4, 5) = "H" '7
        mSoftStrategy(4, 6) = "H" '8
        mSoftStrategy(4, 7) = "H" '9
        mSoftStrategy(4, 8) = "H" '10
        mSoftStrategy(4, 9) = "H" 'J
        mSoftStrategy(4, 10) = "H" 'Q
        mSoftStrategy(4, 11) = "H" 'K
        mSoftStrategy(4, 12) = "H" 'A

        'A,7
        mSoftStrategy(5, 0) = "S" '2
        mSoftStrategy(5, 1) = "Dh" '3
        mSoftStrategy(5, 2) = "Dh" '4
        mSoftStrategy(5, 3) = "Dh" '5
        mSoftStrategy(5, 4) = "Dh" '6
        mSoftStrategy(5, 5) = "S" '7
        mSoftStrategy(5, 6) = "S" '8
        mSoftStrategy(5, 7) = "H" '9
        mSoftStrategy(5, 8) = "H" '10
        mSoftStrategy(5, 9) = "H" 'J
        mSoftStrategy(5, 10) = "H" 'Q
        mSoftStrategy(5, 11) = "H" 'K
        mSoftStrategy(5, 12) = "H" 'A

        'A,8
        mSoftStrategy(6, 0) = "S" '2
        mSoftStrategy(6, 1) = "S" '3
        mSoftStrategy(6, 2) = "S" '4
        mSoftStrategy(6, 3) = "S" '5
        mSoftStrategy(6, 4) = "S" '6
        mSoftStrategy(6, 5) = "S" '7
        mSoftStrategy(6, 6) = "S" '8
        mSoftStrategy(6, 7) = "S" '9
        mSoftStrategy(6, 8) = "S" '10
        mSoftStrategy(6, 9) = "S" 'J
        mSoftStrategy(6, 10) = "S" 'Q
        mSoftStrategy(6, 11) = "S" 'K
        mSoftStrategy(6, 12) = "S" 'A

        'A,9
        mSoftStrategy(7, 0) = "S" '2
        mSoftStrategy(7, 1) = "S" '3
        mSoftStrategy(7, 2) = "S" '4
        mSoftStrategy(7, 3) = "S" '5
        mSoftStrategy(7, 4) = "S" '6
        mSoftStrategy(7, 5) = "S" '7
        mSoftStrategy(7, 6) = "S" '8
        mSoftStrategy(7, 7) = "S" '9
        mSoftStrategy(7, 8) = "S" '10
        mSoftStrategy(7, 9) = "S" 'J
        mSoftStrategy(7, 10) = "S" 'Q
        mSoftStrategy(7, 11) = "S" 'K
        mSoftStrategy(7, 12) = "S" 'A

        'A,10
        mSoftStrategy(8, 0) = "S" '2
        mSoftStrategy(8, 1) = "S" '3
        mSoftStrategy(8, 2) = "S" '4
        mSoftStrategy(8, 3) = "S" '5
        mSoftStrategy(8, 4) = "S" '6
        mSoftStrategy(8, 5) = "S" '7
        mSoftStrategy(8, 6) = "S" '8
        mSoftStrategy(8, 7) = "S" '9
        mSoftStrategy(8, 8) = "S" '10
        mSoftStrategy(8, 9) = "S" 'J
        mSoftStrategy(8, 10) = "S" 'Q
        mSoftStrategy(8, 11) = "S" 'K
        mSoftStrategy(8, 12) = "S" 'A

        'A,J
        mSoftStrategy(9, 0) = "S" '2
        mSoftStrategy(9, 1) = "S" '3
        mSoftStrategy(9, 2) = "S" '4
        mSoftStrategy(9, 3) = "S" '5
        mSoftStrategy(9, 4) = "S" '6
        mSoftStrategy(9, 5) = "S" '7
        mSoftStrategy(9, 6) = "S" '8
        mSoftStrategy(9, 7) = "S" '9
        mSoftStrategy(9, 8) = "S" '10
        mSoftStrategy(9, 9) = "S" 'J
        mSoftStrategy(9, 10) = "S" 'Q
        mSoftStrategy(9, 11) = "S" 'K
        mSoftStrategy(9, 12) = "S" 'A

        'A,Q
        mSoftStrategy(10, 0) = "S" '2
        mSoftStrategy(10, 1) = "S" '3
        mSoftStrategy(10, 2) = "S" '4
        mSoftStrategy(10, 3) = "S" '5
        mSoftStrategy(10, 4) = "S" '6
        mSoftStrategy(10, 5) = "S" '7
        mSoftStrategy(10, 6) = "S" '8
        mSoftStrategy(10, 7) = "S" '9
        mSoftStrategy(10, 8) = "S" '10
        mSoftStrategy(10, 9) = "S" 'J
        mSoftStrategy(10, 10) = "S" 'Q
        mSoftStrategy(10, 11) = "S" 'K
        mSoftStrategy(10, 12) = "S" 'A

        'A,K
        mSoftStrategy(11, 0) = "S" '2
        mSoftStrategy(11, 1) = "S" '3
        mSoftStrategy(11, 2) = "S" '4
        mSoftStrategy(11, 3) = "S" '5
        mSoftStrategy(11, 4) = "S" '6
        mSoftStrategy(11, 5) = "S" '7
        mSoftStrategy(11, 6) = "S" '8
        mSoftStrategy(11, 7) = "S" '9
        mSoftStrategy(11, 8) = "S" '10
        mSoftStrategy(11, 9) = "S" 'J
        mSoftStrategy(11, 10) = "S" 'Q
        mSoftStrategy(11, 11) = "S" 'K
        mSoftStrategy(11, 12) = "S" 'A

    End Sub

    Public Sub GeneratePairStrategy()
        '2
        mPairStrategy(0, 0) = "P" '2
        mPairStrategy(0, 1) = "P" '3
        mPairStrategy(0, 2) = "P" '4
        mPairStrategy(0, 3) = "P" '5
        mPairStrategy(0, 4) = "P" '6
        mPairStrategy(0, 5) = "P" '7
        mPairStrategy(0, 6) = "H" '8
        mPairStrategy(0, 7) = "H" '9
        mPairStrategy(0, 8) = "H" '10
        mPairStrategy(0, 9) = "H" 'J
        mPairStrategy(0, 10) = "H" 'Q
        mPairStrategy(0, 11) = "H" 'K
        mPairStrategy(0, 12) = "H" 'A

        '3
        mPairStrategy(1, 0) = "P" '2
        mPairStrategy(1, 1) = "P" '3
        mPairStrategy(1, 2) = "P" '4
        mPairStrategy(1, 3) = "P" '5
        mPairStrategy(1, 4) = "P" '6
        mPairStrategy(1, 5) = "P" '7
        mPairStrategy(1, 6) = "H" '8
        mPairStrategy(1, 7) = "H" '9
        mPairStrategy(1, 8) = "H" '10
        mPairStrategy(1, 9) = "H" 'J
        mPairStrategy(1, 10) = "H" 'Q
        mPairStrategy(1, 11) = "H" 'K
        mPairStrategy(1, 12) = "H" 'A

        '4
        mPairStrategy(2, 0) = "H" '2
        mPairStrategy(2, 1) = "H" '3
        mPairStrategy(2, 2) = "H" '4
        mPairStrategy(2, 3) = "P" '5
        mPairStrategy(2, 4) = "P" '6
        mPairStrategy(2, 5) = "H" '7
        mPairStrategy(2, 6) = "H" '8
        mPairStrategy(2, 7) = "H" '9
        mPairStrategy(2, 8) = "H" '10
        mPairStrategy(2, 9) = "H" 'J
        mPairStrategy(2, 10) = "H" 'Q
        mPairStrategy(2, 11) = "H" 'K
        mPairStrategy(2, 12) = "H" 'A

        '5
        mPairStrategy(3, 0) = "Dh" '2
        mPairStrategy(3, 1) = "Dh" '3
        mPairStrategy(3, 2) = "Dh" '4
        mPairStrategy(3, 3) = "Dh" '5
        mPairStrategy(3, 4) = "Dh" '6
        mPairStrategy(3, 5) = "Dh" '7
        mPairStrategy(3, 6) = "Dh" '8
        mPairStrategy(3, 7) = "Dh" '9
        mPairStrategy(3, 8) = "H" '10
        mPairStrategy(3, 9) = "H" 'J
        mPairStrategy(3, 10) = "H" 'Q
        mPairStrategy(3, 11) = "H" 'K
        mPairStrategy(3, 12) = "H" 'A

        '6
        mPairStrategy(4, 0) = "P" '2
        mPairStrategy(4, 1) = "P" '3
        mPairStrategy(4, 2) = "P" '4
        mPairStrategy(4, 3) = "P" '5
        mPairStrategy(4, 4) = "P" '6
        mPairStrategy(4, 5) = "H" '7
        mPairStrategy(4, 6) = "H" '8
        mPairStrategy(4, 7) = "H" '9
        mPairStrategy(4, 8) = "H" '10
        mPairStrategy(4, 9) = "H" 'J
        mPairStrategy(4, 10) = "H" 'Q
        mPairStrategy(4, 11) = "H" 'K
        mPairStrategy(4, 12) = "H" 'A

        '2
        mPairStrategy(5, 0) = "P" '2
        mPairStrategy(5, 1) = "P" '3
        mPairStrategy(5, 2) = "P" '4
        mPairStrategy(5, 3) = "P" '5
        mPairStrategy(5, 4) = "P" '6
        mPairStrategy(5, 5) = "P" '7
        mPairStrategy(5, 6) = "H" '8
        mPairStrategy(5, 7) = "H" '9
        mPairStrategy(5, 8) = "H" '10
        mPairStrategy(5, 9) = "H" 'J
        mPairStrategy(5, 10) = "H" 'Q
        mPairStrategy(5, 11) = "H" 'K
        mPairStrategy(5, 12) = "H" 'A

        '8
        mPairStrategy(6, 0) = "P" '2
        mPairStrategy(6, 1) = "P" '3
        mPairStrategy(6, 2) = "P" '4
        mPairStrategy(6, 3) = "P" '5
        mPairStrategy(6, 4) = "P" '6
        mPairStrategy(6, 5) = "P" '7
        mPairStrategy(6, 6) = "P" '8
        mPairStrategy(6, 7) = "P" '9
        mPairStrategy(6, 8) = "P" '10
        mPairStrategy(6, 9) = "P" 'J
        mPairStrategy(6, 10) = "P" 'Q
        mPairStrategy(6, 11) = "P" 'K
        mPairStrategy(6, 12) = "P" 'A

        '9
        mPairStrategy(7, 0) = "P" '2
        mPairStrategy(7, 1) = "P" '3
        mPairStrategy(7, 2) = "P" '4
        mPairStrategy(7, 3) = "P" '5
        mPairStrategy(7, 4) = "P" '6
        mPairStrategy(7, 5) = "S" '7
        mPairStrategy(7, 6) = "H" '8
        mPairStrategy(7, 7) = "H" '9
        mPairStrategy(7, 8) = "S" '10
        mPairStrategy(7, 9) = "S" 'J
        mPairStrategy(7, 10) = "S" 'Q
        mPairStrategy(7, 11) = "S" 'K
        mPairStrategy(7, 12) = "S" 'A

        '10
        mPairStrategy(8, 0) = "S" '2
        mPairStrategy(8, 1) = "S" '3
        mPairStrategy(8, 2) = "S" '4
        mPairStrategy(8, 3) = "S" '5
        mPairStrategy(8, 4) = "S" '6
        mPairStrategy(8, 5) = "S" '7
        mPairStrategy(8, 6) = "S" '8
        mPairStrategy(8, 7) = "S" '9
        mPairStrategy(8, 8) = "S" '10
        mPairStrategy(8, 9) = "S" 'J
        mPairStrategy(8, 10) = "S" 'Q
        mPairStrategy(8, 11) = "S" 'K
        mPairStrategy(8, 12) = "S" 'A

        'J
        mPairStrategy(9, 0) = "S" '2
        mPairStrategy(9, 1) = "S" '3
        mPairStrategy(9, 2) = "S" '4
        mPairStrategy(9, 3) = "S" '5
        mPairStrategy(9, 4) = "S" '6
        mPairStrategy(9, 5) = "S" '7
        mPairStrategy(9, 6) = "S" '8
        mPairStrategy(9, 7) = "S" '9
        mPairStrategy(9, 8) = "S" '10
        mPairStrategy(9, 9) = "S" 'J
        mPairStrategy(9, 10) = "S" 'Q
        mPairStrategy(9, 11) = "S" 'K
        mPairStrategy(9, 12) = "S" 'A

        'Q
        mPairStrategy(10, 0) = "S" '2
        mPairStrategy(10, 1) = "S" '3
        mPairStrategy(10, 2) = "S" '4
        mPairStrategy(10, 3) = "S" '5
        mPairStrategy(10, 4) = "S" '6
        mPairStrategy(10, 5) = "S" '7
        mPairStrategy(10, 6) = "S" '8
        mPairStrategy(10, 7) = "S" '9
        mPairStrategy(10, 8) = "S" '10
        mPairStrategy(10, 9) = "S" 'J
        mPairStrategy(10, 10) = "S" 'Q
        mPairStrategy(10, 11) = "S" 'K
        mPairStrategy(10, 12) = "S" 'A

        'K
        mPairStrategy(11, 0) = "S" '2
        mPairStrategy(11, 1) = "S" '3
        mPairStrategy(11, 2) = "S" '4
        mPairStrategy(11, 3) = "S" '5
        mPairStrategy(11, 4) = "S" '6
        mPairStrategy(11, 5) = "S" '7
        mPairStrategy(11, 6) = "S" '8
        mPairStrategy(11, 7) = "S" '9
        mPairStrategy(11, 8) = "S" '10
        mPairStrategy(11, 9) = "S" 'J
        mPairStrategy(11, 10) = "S" 'Q
        mPairStrategy(11, 11) = "S" 'K
        mPairStrategy(11, 12) = "S" 'A

        '11
        mPairStrategy(12, 0) = "P" '2
        mPairStrategy(12, 1) = "P" '3
        mPairStrategy(12, 2) = "P" '4
        mPairStrategy(12, 3) = "P" '5
        mPairStrategy(12, 4) = "P" '6
        mPairStrategy(12, 5) = "P" '7
        mPairStrategy(12, 6) = "P" '8
        mPairStrategy(12, 7) = "P" '9
        mPairStrategy(12, 8) = "P" '10
        mPairStrategy(12, 9) = "P" 'J
        mPairStrategy(12, 10) = "P" 'Q
        mPairStrategy(12, 11) = "P" 'K
        mPairStrategy(12, 12) = "P" 'A

    End Sub
#End Region
End Class
