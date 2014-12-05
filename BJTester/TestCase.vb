Public Class TestCase

    Public mstrDescription As String = ""
    Public mShoe As New BlackJackSimulator.Shoe()
    Public mlngWinnings As Int32 = 0
    Public mlngLosings As Int32 = 0
    Public mblnSuccess As Boolean = True
    Public mblnComments As Boolean = False
    Public mblnStatistics As Boolean = False

    Public Sub New(ByVal pstrDescription As String, _
                   ByVal pShoe As BlackJackSimulator.Shoe, _
                   ByVal plngExpectedWinnings As Int32, _
                   ByVal plngExpectedLosings As Int32, _
                   ByVal pblnComments As Boolean, _
                   ByVal pblnStatistics As Boolean)

        mstrDescription = pstrDescription
        mShoe = pShoe
        mlngWinnings = plngExpectedWinnings
        mlngLosings = plngExpectedLosings
        mblnComments = pblnComments
        mblnStatistics = pblnStatistics
    End Sub

End Class
