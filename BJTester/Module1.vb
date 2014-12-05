Module Module1

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

    Sub Main()
        'Create List of Tests
        Dim CaseList As New List(Of TestCase)
        GenerateTests(CaseList)

        'Create your blackjack simulator
        Dim BlackJackSim As New BlackJackSimulator(plngDecks:=6, _
                                                   plngUnit:=1, _
                                                   plngHands:=1, _
                                                   plngPenetration:=0, _
                                                   plngPlayers:=1)

        'Run Simulation on all test cases
        For Each TCase As TestCase In CaseList
            BlackJackSim.Run(TCase, _
                             TCase.mblnComments, _
                             TCase.mblnStatistics)
            If Not TCase.mblnSuccess Then
                Console.WriteLine("Test Failed: " & TCase.mstrDescription)
            Else
                Console.WriteLine("Test Success: " & TCase.mstrDescription)
            End If
        Next

        Console.ReadLine()
    End Sub

    Public Sub GenerateTests(ByRef pCaseList As List(Of TestCase))
        pCaseList.Add(New TestCase("No one busts. Dealer 7,8,3 Player 7,9,3. Player Wins", _
                                   New BlackJackSimulator.Shoe({7, 8, 7, 9, 3, 3}), _
                                   1, _
                                   0, _
                                   False, _
                                   False))

        pCaseList.Add(New TestCase("No one busts. Dealer 7,8,3 Player 7,9,3. No one wins", _
                                   New BlackJackSimulator.Shoe({7, 8, 7, 8, 3, 3}), _
                                   0, _
                                   0, _
                                   False, _
                                   False))

        pCaseList.Add(New TestCase("No one busts. Dealer 7,8,3 Player 7,9,3. Dealer Wins", _
                                   New BlackJackSimulator.Shoe({7, 9, 7, 8, 3, 3}), _
                                   0, _
                                   1, _
                                   False, _
                                   False))

        'Test Case 1
        pCaseList.Add(New TestCase("Player Splits Ace's receives 10's and dealer draws to 21", _
                                  New BlackJackSimulator.Shoe({1, 6, 1, 1, 10, 10, 4}), _
                                  0, _
                                  0, _
                                  False, _
                                  False))

        'Test Case 2
        pCaseList.Add(New TestCase("Dealer gets 21, everyone else gets 20", _
                                  New BlackJackSimulator.Shoe({1, 10, 10, 10}), _
                                  0, _
                                  1, _
                                  False, _
                                  False))

        pCaseList.Add(New TestCase("Play splits 4 times with 8 and gets 21 each time. Dealer loses", _
                                  New BlackJackSimulator.Shoe({8, 8, 8, 8, 8, 8, 8, 3, 10, 3, 10, 3, 10, 3, 10, 3, 10, 8}), _
                                  10, _
                                  0, _
                                  False, _
                                  False))

        pCaseList.Add(New TestCase("Player gets BJ, dealer gets BJ", _
                                  New BlackJackSimulator.Shoe({10, 1, 10, 1}), _
                                  0, _
                                  0, _
                                  False, _
                                  False))
    End Sub

End Module
