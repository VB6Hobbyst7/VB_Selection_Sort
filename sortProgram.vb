Module sortProgram

    'Main subroutine for program
    Sub Main()

        Dim numberArray(100) 'Stores values to be sorted
        Dim userInput 'Stores user action
        Call generateValues(numberArray)

        Console.WriteLine("Unsorted Values:")
        Call printValues(numberArray)

        userInput = "a"
        While (userInput <> "q" Or userInput <> "s") 'Loop used to continually prompt menu if invalid command is given
            Console.WriteLine("Press 's' to sort or 'q' to exit program")
            userInput = Console.ReadLine()

            If userInput = "q" Then
                End
            ElseIf userInput = "s" Then
                Call sortValues(numberArray)
                Call printValues(numberArray)
                Console.ReadLine()
                End
            Else
                Console.WriteLine("Invalid command")
            End If
        End While

    End Sub

    'Subroutine used to fill array with values
    Sub generateValues(ByRef emptyArray)

        For i = 0 To UBound(emptyArray)
            emptyArray(i) = Int(Rnd() * 100) + 0 'Generates number between 0 and 100 and stores in array index
        Next

    End Sub

    'Subroutine used to print array
    Sub printValues(ByRef printArray)

        For i = 0 To UBound(printArray)
            Console.WriteLine(printArray(i))
        Next

    End Sub

    'Subroutine used to sort array, uses selection sort method
    Sub sortValues(ByRef unsortedArray)
        Dim smallestValue 'Stores smallest value in the array
        Dim savedIndex 'Used for the swap portion of the sort

        For i = 0 To UBound(unsortedArray)
            smallestValue = unsortedArray(i) 'Start with assumption that first index checked is smallest value
            For j = i + 1 To UBound(unsortedArray)
                If unsortedArray(j) < smallestValue Then 'Checks to see if next index's value is smaller
                    smallestValue = unsortedArray(j)
                    savedIndex = j
                End If
            Next
            'Before moving onto the next array index swap the values
            unsortedArray(savedIndex) = unsortedArray(i)
            unsortedArray(i) = smallestValue
        Next
    End Sub

End Module
