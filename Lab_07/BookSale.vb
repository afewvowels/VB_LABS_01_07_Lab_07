'Program:		Lab #7
'Programmer:	Keith Smith
'Date:			31 October 2018
'Description:	Describes the BookSale class to be used in the main SalesForm.vb

Option Explicit On
Option Strict On

Public Class BookSale
    ' Define variables for properties
    Private QuantityInteger As Integer
    Private PriceDecimal, ExtendedPriceDecimal As Decimal

    ' Define Properties
    Private Shared AccumulatorSalesTotalDecimal As Decimal

    ' Title is an auto-implemented property
    Public Property Title As String

    ' Quantity and Price use logic to calculate values
    Public Property Quantity As Integer
        Get
            Return QuantityInteger
        End Get
        Set(value As Integer)
            ' Validate that at least 1 book is indicated
            If (value < 1) Then
                Throw New ArgumentOutOfRangeException("Quantity must be greater than 0")
            End If

            QuantityInteger = value
        End Set
    End Property
    Public Property Price As Decimal
        Get
            Return PriceDecimal
        End Get
        Set(value As Decimal)
            ' Validate that positive price is entered
            If (value <= 0) Then
                Throw New ArgumentOutOfRangeException("Price must be greater than $0.00")
            End If

            PriceDecimal = value
        End Set
    End Property
    ' Use read-only since ExtendedPrice is a calculated value
    Public ReadOnly Property ExtendedPrice As Decimal
        Get
            Return ExtendedPriceDecimal
        End Get
    End Property

    ' Use shared to share same variable across all created BookSale objects
    Public Shared ReadOnly Property AccumulatorSalesTotal As Decimal
        Get
            Return AccumulatorSalesTotalDecimal
        End Get
    End Property

    ' Define constructors
    Sub New(ByVal _TitleString As String, ByVal _QuantityInteger As Integer, ByVal _PriceDecimal As Decimal)
        ' Use property get/set methods that contain validation logic
        ' Prevents need to validate before passing values to object
        ' Use property methods, not assigning directly to variables
        Title = _TitleString
        Quantity = _QuantityInteger
        Price = _PriceDecimal

        ' Update calculated extended price value
        CalculateExtendedPrice()

        ' Update total sales accumulator
        AddToAccumulator()
    End Sub


    ' Define methods

    ' Calculate extended price for display and use
    Private Sub CalculateExtendedPrice()
        ExtendedPriceDecimal = Price * Quantity
    End Sub

    Private Sub AddToAccumulator()
        AccumulatorSalesTotalDecimal += ExtendedPrice
    End Sub
End Class
