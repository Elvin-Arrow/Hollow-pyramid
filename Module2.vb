Module Module2
    Sub main()
        Dim Symbol As String
        Dim MaxNumberOfSymbols, NumberOfSpaces, NumberOfSymbols, HollowSpace As Integer
        HollowSpace = 1
        Symbol = ""
        Call SetValues(Symbol, MaxNumberOfSymbols, NumberOfSpaces, NumberOfSymbols)
        Do
            Call OutputSpaces(NumberOfSpaces)
            Call OutPutSymbols(NumberOfSymbols, Symbol, MaxNumberOfSymbols, HollowSpace)
            Call AdjustValuesForNextRow(NumberOfSpaces, NumberOfSymbols)
        Loop Until NumberOfSymbols > MaxNumberOfSymbols
        Console.ReadKey()
    End Sub
    Sub SetValues(ByRef Symbol As String, ByRef MaxNumberOfSymbols As Integer, ByRef NumberOfSpaces As Integer, ByRef NumberOfSymbols As Integer)
        Console.Write("Enter the symbol you want to make a pyramid of: ")
        Symbol = Console.ReadLine
        Symbol = UCase(Symbol)
        Call InputMaxNumberOfSymbols(MaxNumberOfSymbols)
        NumberOfSpaces = (MaxNumberOfSymbols - 1) / 2
        NumberOfSymbols = 1
    End Sub
    Sub InputMaxNumberOfSymbols(ByRef MaxNumberOfSymbols As Integer)
        Console.WriteLine("Please enter an odd number")
        Console.Write("How many of these symbols would you like to make a pyramid of: ")
        MaxNumberOfSymbols = Console.ReadLine
        Do
            If MaxNumberOfSymbols Mod 2 = 0 Then
                Console.WriteLine("Please Enter an Odd number")
                Console.Write("Your new number is: ")
                MaxNumberOfSymbols = Console.ReadLine
            End If
        Loop Until MaxNumberOfSymbols Mod 2 <> 0
    End Sub
    Sub OutputSpaces(ByVal NumberOfSpaces As Integer)
        For x = 1 To NumberOfSpaces
            Console.Write(" ")
        Next
    End Sub
    Sub OutPutSymbols(ByVal NumberOfSymbols As Integer, ByVal Symbol As String, ByVal MaxNumberOfSymbols As Integer, ByRef HollowSpace As Integer)
        If NumberOfSymbols = 1 Or NumberOfSymbols = MaxNumberOfSymbols Then
            For y = 1 To NumberOfSymbols
                Console.Write(Symbol)
            Next
        Else
            Console.Write(Symbol)
            For y = 1 To HollowSpace
                Console.Write(" ")
            Next
            Console.Write(Symbol)
            HollowSpace = HollowSpace + 2
        End If
    End Sub
    Sub AdjustValuesForNextRow(ByRef NumberOfSpaces As Integer, ByRef NumberOfSymbols As Integer)
        Console.WriteLine()
        NumberOfSpaces = NumberOfSpaces - 1
        NumberOfSymbols = NumberOfSymbols + 2
    End Sub
End Module
